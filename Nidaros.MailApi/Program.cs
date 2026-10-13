using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

var builder = WebApplication.CreateBuilder(args);

var app = builder.Build();

// HTML не кэшируем, чтобы Outlook всегда брал свежую версию панели
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        if (ctx.File.Name.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
        {
            ctx.Context.Response.Headers.CacheControl = "no-store, no-cache, must-revalidate";
        }
    }
});

app.MapGet("/health", () => Results.Ok(new
{
    status = "OK",
    service = "Nidaros Mail API"
}));


app.MapPost("/api/mail/analyze", (MailRequest mail) =>
{
    var (response, error) = MailAnalyzer.Analyze(mail);

    return error is null
        ? Results.Ok(response)
        : Results.BadRequest(new { error });
});

app.Run();


/// <summary>
/// Вся логика анализа письма. Вынесена из эндпоинта, чтобы её можно было вызывать из тестов.
/// </summary>
public static class MailAnalyzer
{
    public const long MaxFileBytes = 10 * 1024 * 1024; // 10 MB на файл

    /// Ошибка запроса → (null, текст ошибки); иначе (ответ, null)
    public static (AnalyzeResponse? Response, string? Error) Analyze(MailRequest mail)
    {
        var uploadedFiles = mail.Files ?? [];

        if (string.IsNullOrWhiteSpace(mail.Subject) &&
            string.IsNullOrWhiteSpace(mail.Body) &&
            uploadedFiles.Length == 0)
        {
            return (null, "Email subject, body and files are empty.");
        }

        // Какие поля искать. Пусто/null — все поля.
        // Ключи приводим к каноническому написанию ("IBAN" → "iban")
        var knownKeys = FieldSpecs().Select(f => f.Key).ToArray();
        var requested = (mail.RequestedFields ?? [])
            .Where(k => !string.IsNullOrWhiteSpace(k))
            .Select(k => knownKeys.FirstOrDefault(known => known.Equals(k.Trim(), StringComparison.OrdinalIgnoreCase)) ?? k.Trim())
            .Distinct()
            .ToArray();

        var unknown = requested.Where(k => !knownKeys.Contains(k)).ToArray();

        if (unknown.Length > 0)
        {
            return (null, $"Unknown fields: {string.Join(", ", unknown)}.");
        }

        var keys = requested.Length > 0 ? requested : knownKeys;

        // Каждый источник (тело письма, каждый файл) анализируется ОТДЕЛЬНО:
        // реквизиты двух разных фактур не должны смешиваться в одном тексте
        var sources = new List<SourceText>();
        var fileResults = new List<FileProcessResult>();

        if (!string.IsNullOrWhiteSpace(mail.Body))
        {
            sources.Add(new SourceText("email", "E-mail", NormalizeText(mail.Body), PairMode.Text));
        }

        for (var i = 0; i < uploadedFiles.Length; i++)
        {
            var file = uploadedFiles[i];
            var name = string.IsNullOrWhiteSpace(file.Name) ? $"bestand {i + 1}" : file.Name;
            var extracted = ExtractFileText(file, MaxFileBytes);

            // ok — текст прочитан; skipped — тип файла не поддерживается (логотип в подписи и т.п.);
            // error — файл нужного типа, но прочитать не удалось (битый, пустой, скан без текста, слишком большой)
            var status = extracted.Error is null ? "ok" : extracted.Unsupported ? "skipped" : "error";

            fileResults.Add(new FileProcessResult(name, file.Source ?? "", extracted.Text.Length, status, extracted.Error));

            if (extracted.Text.Length > 0)
            {
                sources.Add(new SourceText("file", name, NormalizeText(extracted.Text), extracted.Pairs));
            }
        }

        // Номер фактуры, итог и реквизиты поставщика нужны для определения "какая это фактура",
        // даже если их не запросили
        var identityKeys = new[] { "invoiceNumber", "totalInclVat", "kvk", "vat", "name", "customerName" };
        var extractKeys = keys.Union(identityKeys).ToArray();
        var invoiceKeys = new List<InvoiceKey?>();
        var allFieldsPerDocument = new List<Dictionary<string, FieldResult>>();

        var documents = sources.Select(source =>
        {
            // Структура документа: таблица позиций и адресные блоки без меток
            var model = InvoiceStructure.Analyze(source.Text);
            var allFields = ExtractFields(model, extractKeys, source.Pairs)
                .ToDictionary(kv => kv.Key, kv => kv.Value with { Source = source.Name });

            invoiceKeys.Add(InvoiceKeyOf(source, allFields, model));
            allFieldsPerDocument.Add(allFields);

            var docFields = keys.ToDictionary(k => k, k => allFields[k]);

            return new DocumentResult(
                source.Kind,
                source.Name,
                null,
                docFields,
                model.Lines,
                docFields.Values.All(f => f.Status == "found"),
                source.Text.Length > 5000 ? source.Text[..5000] + "…" : source.Text);
        }).ToList();

        var invoiceIds = AssignInvoiceIds(invoiceKeys);
        documents = documents.Select((d, i) => d with { InvoiceId = invoiceIds[i] }).ToList();

        // Причины, по которым отправлять нельзя (issues) и на что стоит посмотреть (warnings)
        var issues = new List<string>();
        var warnings = new List<string>();

        // Разные фактуры (разные номера) НЕ сливаются, даже если IBAN/KVK/BTW у них совпадают:
        // совпадение реквизитов одного поставщика не делает две фактуры одной
        var invoices = documents
            .Where(d => d.InvoiceId is not null)
            .GroupBy(d => d.InvoiceId)
            .ToList();

        Dictionary<string, FieldResult> fields;
        List<InvoiceLine> lines;

        if (invoices.Count > 1)
        {
            fields = SeparateInvoices(documents, keys, invoices.Count);
            lines = [];
            issues.Add($"Het bericht bevat {invoices.Count} verschillende facturen " +
                       $"({string.Join(", ", invoices.Select(g => $"'{string.Join(" + ", g.Select(d => d.Name))}'"))}); " +
                       "gegevens worden niet samengevoegd, verwerk per factuur (documents[]).");
        }
        else
        {
            // Одна фактура (возможно, в нескольких копиях: PDF + DOCX).
            // Реквизиты берутся ТОЛЬКО из неё: в тексте письма почти всегда подпись — отправителя
            // или того, кто переслал письмо (Nidaros), и её KVK/адрес не должны ни подменять,
            // ни блокировать данные фактуры. Без фактуры письмо — обычный источник.
            var invoiceDocs = documents.Where(d => d.InvoiceId is not null).ToList();
            var contextDocs = documents.Where(d => d.InvoiceId is null).ToList();

            if (invoiceDocs.Count > 0)
            {
                fields = MergeFields(invoiceDocs, keys);
                AddContext(fields, contextDocs, keys, warnings);
            }
            else
            {
                fields = MergeFields(documents, keys);
                NameFromKlantnaam(fields, keys, allFieldsPerDocument);
            }

            // Позиции — только если они есть в одном документе или одинаковые во всех
            var lineSets = (invoiceDocs.Count > 0 ? invoiceDocs : documents)
                .Where(d => d.Lines.Count > 0).Select(d => d.Lines).ToList();
            lines = lineSets.Count > 0 && lineSets.All(l => l.SequenceEqual(lineSets[0])) ? lineSets[0] : [];

            if (lineSets.Count > 1 && lines.Count == 0)
            {
                issues.Add("Factuurregels verschillen tussen documenten; zie documents[].lines.");
            }
        }

        foreach (var file in fileResults)
        {
            if (file.Status == "error") issues.Add($"Bestand '{file.Name}' kon niet worden gelezen: {file.Error}");
            if (file.Status == "skipped") warnings.Add($"Bestand '{file.Name}' overgeslagen: {file.Error}");
        }

        if (sources.Count == 0)
        {
            issues.Add("Geen leesbare tekst in e-mail of bestanden.");
        }

        var conflicts = fields.Where(f => f.Value.Status == "conflict").Select(f => f.Key).ToList();

        if (conflicts.Count > 0)
        {
            issues.Add($"Documenten bevatten verschillende waarden voor: {string.Join(", ", conflicts)}.");
        }

        // Всё, что не "found": missing, invalid, guessed (без метки), review (сомнительное значение)
        var incomplete = fields
            .Where(f => f.Value.Status is not ("found" or "conflict"))
            .Select(f => $"{f.Key} ({f.Value.Status})")
            .ToList();

        if (incomplete.Count > 0)
        {
            issues.Add($"Niet gevonden, ongeldig of te controleren: {string.Join(", ", incomplete)}.");
        }

        var response = new AnalyzeResponse(
            Success: true,
            mail.ItemId,
            mail.Subject,
            mail.Sender,
            fields,
            // Позиции счёта по строкам (article/quantity в fields — сводка через "; ")
            lines,
            // Результат по каждому источнику отдельно
            documents,
            mail.Attachments,
            AttachmentCount: mail.Attachments?.Length ?? 0,
            Files: fileResults,
            RequestedFields: fields.Keys.ToList(),
            // Готово: все поля найдены и проверены, документы не противоречат друг другу, все файлы прочитаны
            ReadyToSend: issues.Count == 0,
            issues,
            warnings,
            // Для отладки: какой текст сервер реально увидел (по источникам)
            ExtractedText: string.Join("\n\n", documents.Select(d => $"===== {d.Name} =====\n{d.ExtractedText}")),
            ProcessedAt: DateTimeOffset.UtcNow);

        return (response, null);
    }


    // ─────────────────────────────────────────────────────────────
    // Извлечение полей
    // ─────────────────────────────────────────────────────────────

    static FieldSpec[] FieldSpecs() =>
    [
        // Метки: длинные варианты ПЕРВЫМИ ("btw nummer" раньше "btw").
        // name/address/postalCode/city — сторона, чьи KVK/BTW/IBAN (на фактуре — поставщик);
        // без метки берутся из адресного блока отправителя. customer* — клиент / получатель.
        // Метки, которые явно называют клиента ("Klantnaam", "Factuuradres"), относятся
        // ТОЛЬКО к customer*: иначе имя клиента попадёт в name рядом с KVK/IBAN поставщика.
        new("name",
            Labels: "naam leverancier|leverancier|bedrijfsnaam|volledige naam|naam|company name|supplier|full name|name"),

        new("address",
            Labels: "straat en huisnummer|adres|address|street"),

        new("postalCode",
            Labels: "postcode|postal code|zip code|zip",
            ValuePattern: @"\d{4}[ \t]?[A-Za-z]{2}",
            Validate: ValidatePostalCode,
            FallbackPattern: @"\b(?<value>[1-9]\d{3} ?[A-Z]{2})\b"),

        new("city",
            Labels: "woonplaats|plaats|stad|city|town"),

        new("iban",
            Labels: "iban[- ]?nummer|bankrekeningnummer|rekeningnummer|iban",
            ValuePattern: @"[A-Za-z]{2}\d{2}[A-Za-z0-9 ]{10,40}?",
            Validate: ValidateIban,
            FallbackPattern: @"\b(?<value>[A-Z]{2}\d{2}[A-Z0-9]{11,30})\b"),

        new("customerType",
            Labels: "nieuwe klant of bestaande klant|klanttype|soort klant|klant|customer type|customer",
            ValuePattern: @"nieuwe klant|bestaande klant|nieuw|bestaand|new|existing",
            Validate: ValidateCustomerType,
            Strict: true),

        new("kvk",
            Labels: @"kvk[- ]?nummer|kvk[- ]?nr\.?|kamer van koophandel|kvk|coc number|coc",
            ValuePattern: @"[\d .]{8,12}",
            Validate: ValidateKvk),

        new("vat",
            Labels: @"btw[- ]?nummer|btw[- ]?id|btw[- ]?nr\.?|btw|vat number|vat id|vat",
            Validate: ValidateVat,
            FallbackPattern: @"\b(?<value>NL\d{9}B\d{2})\b"),

        new("invoiceNumber",
            Labels: @"factuurnummer|factuur[- ]?nr\.?|factuur|invoice number|invoice no\.?|invoice",
            // Номер счёта — одно "слово" с хотя бы одной цифрой (иначе "Factuur" + следующая строка)
            ValuePattern: @"(?=[\w\-/.]*\d)[A-Za-z0-9][\w\-/.]*",
            Strict: true),

        new("invoiceDate",
            Labels: "factuurdatum|datum factuur|invoice date",
            Validate: ValidateDate),

        new("customerNumber",
            Labels: @"klantnummer|klantnr\.?|debiteurnummer|customer number|customer no\.?",
            ValuePattern: @"[A-Za-z0-9][\w\-/.]*",
            Strict: true),

        new("customerName",
            Labels: "klantnaam|naam klant|naam debiteur|debiteur|customer name"),
        new("customerAddress",
            Labels: "factuuradres|adres klant|customer address"),
        new("customerPostalCode", Labels: null, Validate: ValidatePostalCode),
        new("customerCity", Labels: null),

        new("article",
            Labels: "artikelnaam|artikel|product|omschrijving|article|item"),

        new("quantity",
            Labels: "hoeveelheid|aantal|quantity|qty",
            Validate: ValidateQuantity),

        new("deliveryDate",
            Labels: "datum afname|afnamedatum|leverdatum|delivery date",
            Validate: ValidateDate),

        new("paymentDueDate",
            Labels: "uiterste betalingsdatum|vervaldatum|betaaldatum|payment due date|due date",
            Validate: ValidateDate,
            // "Te betalen binnen 31 dagen, voor 31-10-2026 op rekeningnummer ..."
            Context: [@"betal\p{L}*\b[^\n]{0,80}?\b(?:vóór|voor|uiterlijk|before)\s+(?<value>\d{1,2}[-/.]\d{1,2}[-/.]\d{4})"]),

        new("paymentTermDays",
            Labels: "betalingstermijn|betaaltermijn|payment term",
            Validate: ValidateDays,
            Context: [@"binnen\s+(?<value>\d{1,3})\s+dagen", @"within\s+(?<value>\d{1,3})\s+days"]),

        new("totalExclVat",
            Labels: @"totaal excl\.? btw|subtotaal|total excl\.? vat|subtotal",
            ValuePattern: FieldSpec.AmountPattern,
            Validate: ValidateAmount,
            Strict: true),

        new("vatAmount",
            Labels: @"\d{1,2}(?:[.,]\d+)? ?% btw|btw[- ]?bedrag|vat amount|\d{1,2}(?:[.,]\d+)? ?% vat",
            ValuePattern: FieldSpec.AmountPattern,
            Validate: ValidateAmount,
            Strict: true),

        new("totalInclVat",
            Labels: @"totaal incl\.? btw|totaal te betalen|totaalbedrag|total incl\.? vat|total due|amount due",
            ValuePattern: FieldSpec.AmountPattern,
            Validate: ValidateAmount,
            Strict: true),

        new("paymentArrears",
            Labels: "betalingsachterstanden|betalingsachterstand|achterstand|payment arrears|arrears")
    ];


    /// Значение на строке ПОД меткой ("Naam:" / "Jan Jansen") берётся в зависимости от источника:
    ///   • PDF (Layout) — никогда: пары уже собраны по координатам, метка без значения — пустое поле;
    ///   • DOCX (Paragraphs) — только если документ оформлен как вертикальная форма, тогда found;
    ///     одиночная пустая "Naam:" перед обычным абзацем остаётся пустой;
    ///   • письмо, XML (Text) — в вертикальной форме found, иначе guessed (надо проверить).
    static Dictionary<string, FieldResult> ExtractFields(
        DocumentModel document, IReadOnlyCollection<string> keys, PairMode pairs)
    {
        // Метки, которые не являются нашими полями, но тоже не могут быть значением
        // (иначе при пустом "Naam" значением станет следующая строка "BSN")
        const string otherLabels =
            "bsn|type dossier|e-?mail(?:adres)?|telefoon(?:nummer)?|tel|geboortedatum";

        var specs = FieldSpecs();
        var allLabels = string.Join("|", specs
            .Where(s => s.Labels is not null)
            .Select(s => s.Labels)
            .Append(otherLabels));

        var derived = InvoiceStructure.DerivedValues(document);

        var vertical = pairs != PairMode.Layout && IsVerticalLayout(document.FieldText, allLabels);

        string? nextLineStatus = pairs switch
        {
            PairMode.Layout => null,
            PairMode.Paragraphs => vertical ? "found" : null,
            _ => vertical ? "found" : "guessed"
        };

        // Метки всех полей участвуют в allLabels, даже если поле не запрошено
        var fields = specs
            .Where(s => keys.Contains(s.Key))
            .ToDictionary(s => s.Key, s => Field(document.FieldText, s, allLabels, derived, nextLineStatus));

        InvoiceStructure.CheckTotals(fields, document);

        return fields;
    }


    /// <summary>
    /// Текст оформлен как вертикальная форма "Метка: / значение / Метка: / значение"
    /// (так Outlook часто отдаёт HTML-таблицы письма текстом, так же выглядят формы в Word).
    ///
    /// Считаем:
    ///   vertical — строк "только метка", за которыми идёт строка-значение (не метка и не "x: y");
    ///   inline   — строк "Метка: значение" в одну строку.
    /// Пустые поля (метка, за которой сразу другая метка) не считаются ни за, ни против:
    /// одно-два незаполненных поля не должны превращать остальные поля формы в guessed.
    /// Вертикальная форма: vertical ≥ 3 и таких пар больше, чем однострочных.
    /// Документ "Метка: значение" с одной меткой, под которой обычный текст, — не вертикальный.
    /// </summary>
    static bool IsVerticalLayout(string text, string allLabels)
    {
        const string endOfWord = @"(?![\p{L}\p{N}])";
        const string labelTail = @"(?:[ \t]+[\p{L} ]{1,40}?)?";

        var labelOnly = new Regex($@"^[ \t]*(?>{allLabels}){endOfWord}{labelTail}[ \t]*:?[ \t]*$",
            RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(500));
        var inlinePair = new Regex($@"^[ \t]*(?>{allLabels}){endOfWord}{labelTail}[ \t]*:[ \t]*\S",
            RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(500));

        var lines = text.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0).ToList();
        int vertical = 0, inline = 0;

        for (var i = 0; i < lines.Count; i++)
        {
            if (inlinePair.IsMatch(lines[i]))
            {
                inline++;
            }
            else if (labelOnly.IsMatch(lines[i]) && i + 1 < lines.Count &&
                     !labelOnly.IsMatch(lines[i + 1]) && !lines[i + 1].Contains(':'))
            {
                vertical++;
            }
        }

        return vertical >= 3 && vertical > inline;
    }


    /// <summary>
    /// Признаки фактуры в документе.
    ///   • Текст письма фактурой НЕ считается никогда: "Factuurnummer: …" или "Totaal: …" в письме —
    ///     это ссылка на вложение, а не вторая фактура. Письмо — контекст: его поля сверяются
    ///     с фактурой и дают conflict, если расходятся.
    ///   • Файл с номером фактуры, позициями или итоговой суммой — фактура.
    ///   • Файл без этих признаков (условия поставки и т.п.) — тоже контекст: null.
    /// Supplier — реквизиты, которые отличают одну компанию от другой: KVK, BTW, название.
    /// IBAN сюда не входит: другой IBAN у той же компании и того же номера — это конфликт
    /// (возможная подмена реквизитов), а не другая фактура.
    /// </summary>
    static InvoiceKey? InvoiceKeyOf(SourceText source, Dictionary<string, FieldResult> fields, DocumentModel model)
    {
        if (source.Kind == "email")
            return null;

        static string? Known(FieldResult f) =>
            f.Status is "found" or "review" && f.Value.Length > 0
                ? Regex.Replace(f.Value, @"[^\p{L}\p{N}]", "").ToUpperInvariant()
                : null;

        var number = Known(fields["invoiceNumber"]);
        var looksLikeInvoice = number is not null || model.Lines.Count > 0 || fields["totalInclVat"].Status != "missing";

        if (!looksLikeInvoice)
            return null;

        var supplier = new Dictionary<string, string>();
        foreach (var key in new[] { "kvk", "vat", "name" })
        {
            if (Known(fields[key]) is { } value) supplier[key] = value;
        }

        return new InvoiceKey(number, supplier);
    }


    /// <summary>
    /// Раскладывает документы по фактурам. Два документа — копии одной фактуры, только если
    ///   • у обоих есть номер фактуры и он совпадает,
    ///   • и реквизиты поставщика не противоречат: KVK / BTW / название, известные в обоих, равны.
    /// Одинаковый номер у двух разных компаний (2025-001 бывает у всех) — две фактуры.
    /// Фактура без номера — всегда отдельная: доказать, что это копия, нечем.
    /// Результат: id фактуры для каждого документа, null — контекст (письмо, приложение).
    /// </summary>
    static List<string?> AssignInvoiceIds(List<InvoiceKey?> keys)
    {
        var clusters = new List<(string Number, Dictionary<string, string> Supplier, string Id)>();
        var ids = new List<string?>();

        for (var i = 0; i < keys.Count; i++)
        {
            var key = keys[i];

            if (key is null)
            {
                ids.Add(null);
                continue;
            }

            if (key.Number is null)
            {
                ids.Add($"doc:{i}");
                continue;
            }

            var match = clusters.FindIndex(c =>
                c.Number == key.Number &&
                key.Supplier.All(kv => !c.Supplier.TryGetValue(kv.Key, out var known) || known == kv.Value));

            if (match < 0)
            {
                var sameNumber = clusters.Count(c => c.Number == key.Number);
                var id = sameNumber == 0 ? $"nr:{key.Number}" : $"nr:{key.Number}#{sameNumber + 1}";
                clusters.Add((key.Number, new Dictionary<string, string>(key.Supplier), id));
                ids.Add(id);
            }
            else
            {
                // Копия: дополняем реквизиты кластера тем, что известно из этой копии
                foreach (var kv in key.Supplier) clusters[match].Supplier.TryAdd(kv.Key, kv.Value);
                ids.Add(clusters[match].Id);
            }
        }

        return ids;
    }


    /// Поля, которых на фактуре не бывает: их можно взять из текста письма, даже если есть фактура
    static readonly HashSet<string> ContextOnlyFields = ["customerType", "paymentArrears"];

    /// <summary>
    /// Письмо рядом с фактурой.
    ///   • Поля фактуры (реквизиты, суммы, позиции) из письма НЕ берутся: ни для заполнения пустых
    ///     (KVK из подписи Nidaros не должен появиться у фактуры без KVK), ни для конфликта.
    ///   • Поля, которых на фактуре не бывает (customerType, paymentArrears), берутся из письма.
    ///   • Если в письме по метке указано другое значение ("IBAN: ..."), это warning:
    ///     не блокирует, но видно (например, письмо сообщает новый IBAN).
    /// </summary>
    static void AddContext(
        Dictionary<string, FieldResult> fields, List<DocumentResult> contextDocs, IEnumerable<string> keys, List<string> warnings)
    {
        if (contextDocs.Count == 0)
            return;

        var context = MergeFields(contextDocs, keys);

        foreach (var key in keys)
        {
            var fromContext = context[key];
            var fromInvoice = fields[key];

            if (fromContext.Status is "missing" or "conflict")
                continue;

            if (ContextOnlyFields.Contains(key) && fromInvoice.Status == "missing")
            {
                fields[key] = fromContext;
            }
            else if (fromContext.Method is "label" or "context" && fromInvoice.Status != "missing" &&
                     Regex.Replace(fromContext.Value, @"\s+", "").ToUpperInvariant() !=
                     Regex.Replace(fromInvoice.Value, @"\s+", "").ToUpperInvariant())
            {
                warnings.Add($"{fromContext.Source} noemt bij {key} '{fromContext.Value}', de factuur '{fromInvoice.Value}'; " +
                             "de waarde van de factuur is gebruikt.");
            }
        }
    }


    /// <summary>
    /// Без фактуры (старая форма в письме) "Klantnaam: Jan Jansen" — это и есть name:
    /// запись в форме описывает клиента. На фактуре "Klantnaam" — только customerName,
    /// чтобы имя клиента не встало рядом с KVK/IBAN поставщика.
    /// Берётся только значение по метке и только если все документы называют одно и то же имя.
    /// </summary>
    static void NameFromKlantnaam(
        Dictionary<string, FieldResult> fields, IEnumerable<string> keys, List<Dictionary<string, FieldResult>> allFieldsPerDocument)
    {
        if (!keys.Contains("name") || fields["name"].Status != "missing")
            return;

        var names = allFieldsPerDocument
            .Select(f => f["customerName"])
            .Where(f => f.Method == "label" && f.Status is "found" or "review" && f.Value.Length > 0)
            .ToList();

        if (names.Count > 0 && names.Select(f => f.Value.ToUpperInvariant()).Distinct().Count() == 1)
        {
            fields["name"] = names[0] with { Note = "Uit 'Klantnaam' (geen factuur in dit bericht)." };
        }
    }


    /// Несколько фактур в одном письме: ни одно поле не сливается, даже совпадающее.
    /// Все варианты — в Alternatives, обработка — по documents[].
    static Dictionary<string, FieldResult> SeparateInvoices(
        List<DocumentResult> documents, IEnumerable<string> keys, int invoiceCount)
    {
        return keys.ToDictionary(key => key, key => new FieldResult(
            "",
            1,
            "conflict",
            $"Bericht bevat {invoiceCount} facturen; kies per factuur in documents[].",
            Source: null,
            Alternatives: documents
                .Select(d => d.Fields[key])
                .Where(f => f.Status != "missing")
                .Select(f => new FieldAlternative(f.Value, f.Source ?? "", f.Status))
                .ToList(),
            Method: "none"));
    }


    /// <summary>
    /// Сводит результаты документов ОДНОЙ фактуры (копии + письмо как контекст) в один набор полей.
    ///   • found, review и invalid — значения по метке/структуре, они сравниваются между собой;
    ///     guessed (без метки / со следующей строки) используется, только если других нет.
    ///   • одинаковые значения → поле с перечнем источников; худший статус побеждает
    ///     (invalid > review > found): если копия говорит "плохо", поле плохое;
    ///   • разные → status "conflict", пустое значение и варианты в Alternatives:
    ///     вставлять в форму наугад одно из двух нельзя.
    /// </summary>
    static Dictionary<string, FieldResult> MergeFields(List<DocumentResult> documents, IEnumerable<string> keys)
    {
        static string Canonical(string value) =>
            Regex.Replace(value, @"\s+", "").ToUpperInvariant();

        static int Severity(string status) => status switch
        {
            "invalid" => 0,
            "review" => 1,
            _ => 2
        };

        var result = new Dictionary<string, FieldResult>();

        foreach (var key in keys)
        {
            var candidates = documents
                .Select(d => d.Fields[key])
                .Where(f => f.Status != "missing")
                .ToList();

            if (candidates.Count == 0)
            {
                result[key] = new FieldResult("", 1, "missing", "Niet gevonden in e-mail of bestanden.", Method: "none");
                continue;
            }

            var strong = candidates.Where(f => f.Status is "found" or "review" or "invalid").ToList();
            var pool = strong.Count > 0 ? strong : candidates;

            var groups = pool.GroupBy(f => Canonical(f.Value)).ToList();

            if (groups.Count == 1)
            {
                var worst = pool.OrderBy(f => Severity(f.Status)).ThenByDescending(f => f.Confidence).First();
                result[key] = worst with { Source = string.Join(", ", pool.Select(f => f.Source).Distinct()) };
                continue;
            }

            result[key] = new FieldResult(
                "",
                1,
                "conflict",
                $"Verschillende waarden in {pool.Select(f => f.Source).Distinct().Count()} documenten, kies de juiste.",
                Source: null,
                Alternatives: pool
                    .Select(f => new FieldAlternative(f.Value, f.Source ?? "", f.Status))
                    .ToList(),
                Method: "none");
        }

        return result;
    }


    /// Порядок поиска:
    ///   1. по метке: "Naam: Test Persoon" / "Naam - Test Persoon" / "Naam Test Persoon" (одна строка);
    ///      "KVK: <что угодно>" — берём что угодно и валидируем (так "IBAN: xxxx" станет invalid, а не missing);
    ///      "Naam" и "Test Persoon" на соседних строках (PDF-таблица, прочитанная по колонкам)
    ///   2. по контексту фразы ("Te betalen ... voor 31-10-2026")
    ///   3. из структуры документа (таблица позиций, адресный блок)
    ///   4. fallback без метки (только для форматов, которые узнаются сами: IBAN, postcode, BTW)
    static FieldResult Field(
        string text, FieldSpec spec, string allLabels, IReadOnlyDictionary<string, Derived> derived, string? nextLineStatus)
    {
        var timeout = TimeSpan.FromMilliseconds(500);
        var options = RegexOptions.IgnoreCase | RegexOptions.Multiline;

        if (spec.Labels is not null)
        {
            // (?>...) — атомарная группа: если совпала длинная метка "BTW nummer",
            // движок не откатится к короткой "BTW" и не возьмёт "nummer" как значение.
            // (?![\p{L}\p{N}]) — конец слова; работает и после точки ("KVK nr.").
            const string endOfWord = @"(?![\p{L}\p{N}])";

            // Метка в начале строки или после таба (колонки в одной строке разделены табом)
            var label = $@"(?:^|\t)[ \t]*(?>{spec.Labels}){endOfWord}";

            // Хвост метки до двоеточия: "Betalingsachterstanden binnen de eigen company:"
            const string labelTail = @"[ \t]+[\p{L} ]{1,40}?:";

            // Значение не может быть чужой меткой ("BSN") или готовой парой "Klantnummer: 123"
            var notALabel = $@"(?!(?:{allLabels}){endOfWord})(?![\p{{L}}][\p{{L}} .\-]{{0,40}}:)";

            // Значение заканчивается на конце строки или на табе (следующая колонка)
            var end = @"(?=[ \t]*(?:\t|$))";
            var value = $@"{notALabel}(?<value>{spec.ValuePattern}){end}";

            var sameLine = $@"{label}(?:[ \t]*[:\-–][ \t]*|{labelTail}[ \t]*|[ \t]+){value}";
            var afterColon = $@"{label}[ \t]*:[ \t]*{notALabel}(?<value>[^\n\t]*?\S){end}";
            var nextLine = $@"{label}(?:{labelTail}|[ \t]*[:\-–]?)[ \t]*\n(?:[ \t]*\n)*[ \t]*{value}";

            // (шаблон, метод, статус). Следующая строка: null — не берём (PDF/DOCX),
            // "found" — вертикальная форма в письме, "guessed" — догадка, её надо проверить
            var patterns = new List<(string Pattern, string Method, string Status)> { (sameLine, "label", "found") };
            if (!spec.Strict) patterns.Add((afterColon, "label", "found"));
            if (nextLineStatus is not null) patterns.Add((nextLine, "nextLine", nextLineStatus));

            foreach (var (pattern, method, status) in patterns)
            {
                var match = Regex.Match(text, pattern, options, timeout);

                if (match.Success)
                {
                    return Score(match.Groups["value"].Value, spec.Validate, status == "found" ? 100 : 70, status, method,
                        status == "guessed" ? "Waarde van de volgende regel, controleer." : null);
                }
            }
        }

        foreach (var pattern in spec.Context ?? [])
        {
            var match = Regex.Match(text, pattern, options, timeout);

            if (match.Success)
            {
                return Score(match.Groups["value"].Value, spec.Validate, 100, "found", "context");
            }
        }

        if (derived.TryGetValue(spec.Key, out var fromStructure))
        {
            return Score(fromStructure.Value, spec.Validate, fromStructure.Confidence,
                fromStructure.NeedsReview ? "review" : "found", "structure", fromStructure.Note);
        }

        if (spec.FallbackPattern is not null)
        {
            // Без IgnoreCase: иначе "2026 Al..." похоже на почтовый индекс
            var match = Regex.Match(text, spec.FallbackPattern, RegexOptions.Multiline, timeout);

            if (match.Success)
            {
                return Score(match.Groups["value"].Value, spec.Validate, 85, "guessed", "pattern",
                    "Gevonden zonder label, controleer.");
            }
        }

        return new FieldResult("", 1, "missing", "Niet gevonden in e-mail of bestanden.", Method: "none");
    }


    /// <summary>
    /// Валидация найденного значения.
    ///   Problem → "invalid" (значение неверное);
    ///   Warning → "review" (значение похоже на правду, но сомнительное: человек должен проверить).
    ///   review, как и invalid, блокирует readyToSend — предупреждение не прячется внутри "found".
    ///
    /// Confidence — порядковая оценка СПОСОБА извлечения (метка 100, структура 95, шаблон 85,
    /// следующая строка 70, ошибка ≤ 75), а не вероятность правильности. Для решений — Status и Method.
    /// </summary>
    static FieldResult Score(
        string raw,
        Func<string, Check>? validate,
        int baseConfidence,
        string status,
        string method,
        string? note = null)
    {
        var value = raw.Trim();
        var check = validate?.Invoke(value) ?? new Check(value);

        if (check.Problem is not null)
        {
            return new FieldResult(check.Value, Math.Min(baseConfidence, 75), "invalid", check.Problem, Method: method);
        }

        if (check.Warning is not null)
        {
            // guessed остаётся guessed (догадка хуже сомнения), found → review
            return new FieldResult(check.Value, Math.Min(baseConfidence, 90), status == "found" ? "review" : status,
                note is null ? check.Warning : $"{note} {check.Warning}", Method: method);
        }

        return new FieldResult(check.Value, baseConfidence, status, note, Method: method);
    }


    // ─────────────────────────────────────────────────────────────
    // Валидация и нормализация
    // ─────────────────────────────────────────────────────────────

    static Check ValidatePostalCode(string value)
    {
        // "7901 EB" или "7901 EB Hoogeveen" (город после индекса отбрасываем)
        var m = Regex.Match(value, @"^(\d{4})\s?([A-Za-z]{2})(?![A-Za-z])");

        if (!m.Success)
            return new(value, "Ongeldige postcode.");

        var normalized = $"{m.Groups[1].Value} {m.Groups[2].Value.ToUpperInvariant()}";

        return m.Groups[1].Value[0] == '0'
            ? new(normalized, "Nederlandse postcode begint niet met 0.")
            : new(normalized);
    }


    static Check ValidateIban(string value)
    {
        var iban = Regex.Replace(value, @"\s", "").ToUpperInvariant();

        if (!Regex.IsMatch(iban, @"^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$"))
            return new(iban, "Ongeldig IBAN-formaat.");

        if (iban.StartsWith("NL") && iban.Length != 18)
            return new(FormatIban(iban), "Nederlandse IBAN moet 18 tekens hebben.");

        // ISO 13616 mod-97: eerste 4 tekens naar achteren, letters → getallen (A=10)
        var rearranged = iban[4..] + iban[..4];
        var remainder = 0;

        foreach (var ch in rearranged)
        {
            var digits = char.IsLetter(ch) ? (ch - 'A' + 10).ToString() : ch.ToString();

            foreach (var d in digits)
            {
                remainder = (remainder * 10 + (d - '0')) % 97;
            }
        }

        return remainder == 1
            ? new(FormatIban(iban))
            : new(FormatIban(iban), "IBAN-controlegetal klopt niet.");
    }


    static string FormatIban(string iban) =>
        string.Join(" ", Enumerable.Range(0, (iban.Length + 3) / 4)
            .Select(i => iban.Substring(i * 4, Math.Min(4, iban.Length - i * 4))));


    static Check ValidateCustomerType(string value) =>
        value.Trim().ToLowerInvariant() switch
        {
            "nieuwe klant" or "nieuw" or "new" => new("new"),
            "bestaande klant" or "bestaand" or "existing" => new("existing"),
            _ => new(value, "Onbekend klanttype.")
        };


    static Check ValidateKvk(string value)
    {
        var digits = Regex.Replace(value, @"[\s.]", "");

        return Regex.IsMatch(digits, @"^\d{8}$")
            ? new(digits)
            : new(digits, "KVK-nummer moet 8 cijfers hebben.");
    }


    static Check ValidateVat(string value)
    {
        var vat = Regex.Replace(value, @"[\s.]", "").ToUpperInvariant();

        if (Regex.IsMatch(vat, @"^NL\d{9}B\d{2}$"))
            return new(vat);

        // "59902458B01" — похоже на NL BTW-id, но без "NL" и/или не 9 цифр: значение оставляем, предупреждаем
        if (Regex.IsMatch(vat, @"^(?:NL)?\d{8,9}B\d{2}$"))
            return new(vat, Warning: "Afwijkend BTW-formaat (verwacht NL123456789B01), controleer.");

        return Regex.IsMatch(vat, @"^[A-Z]{2}[A-Z0-9]{2,13}$")
            ? new(vat, Warning: "Geen Nederlands BTW-nummer, controleer.")
            : new(vat, "Ongeldig BTW-nummer.");
    }


    static Check ValidateQuantity(string value)
    {
        return Regex.IsMatch(value, @"^\d+(?:[.,]\d+)?\b")
            ? new(value)
            : new(value, "Hoeveelheid begint niet met een getal.");
    }


    static Check ValidateDays(string value)
    {
        var m = Regex.Match(value, @"^\d{1,3}");

        return m.Success
            ? new(m.Value)
            : new(value, "Betalingstermijn niet herkend.");
    }


    static Check ValidateAmount(string value)
    {
        var amount = InvoiceStructure.ParseAmount(value);

        return amount is null
            ? new(value, "Bedrag niet herkend.")
            : new(amount.Value.ToString("0.00", CultureInfo.InvariantCulture));
    }


    static Check ValidateDate(string value)
    {
        string[] formats =
        [
            "dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy", "d/M/yyyy",
            "dd.MM.yyyy", "d.M.yyyy", "yyyy-MM-dd",
            "d MMMM yyyy", "dd MMMM yyyy", "d MMM yyyy"
        ];

        foreach (var culture in new[] { GetCulture("nl-NL"), CultureInfo.InvariantCulture })
        {
            if (DateTime.TryParseExact(value.Trim(), formats, culture,
                    DateTimeStyles.AllowWhiteSpaces, out var date))
            {
                return new(date.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture));
            }
        }

        return new(value, "Datum niet herkend.");
    }


    static CultureInfo GetCulture(string name)
    {
        // В Docker с InvariantGlobalization культуры может не быть
        try { return CultureInfo.GetCultureInfo(name); }
        catch (CultureNotFoundException) { return CultureInfo.InvariantCulture; }
    }


    static string NormalizeText(string text) =>
        text.Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace(' ', ' ')   // non-breaking space из PDF/Word
            .Replace(' ', ' ')
            .Replace(' ', ' ')
            .Trim();


    // ─────────────────────────────────────────────────────────────
    // Чтение файлов
    // ─────────────────────────────────────────────────────────────

    static ExtractedFile ExtractFileText(UploadedFile file, long maxBytes)
    {
        static ExtractedFile Failed(string error) => new("", error, Unsupported: false, PairMode.Text);

        byte[] bytes;

        try
        {
            bytes = Convert.FromBase64String(file.Base64 ?? "");
        }
        catch (FormatException)
        {
            return Failed("Invalid base64 content.");
        }

        if (bytes.Length == 0)
        {
            return Failed("File is empty.");
        }

        if (bytes.Length > maxBytes)
        {
            return Failed($"File too large (max {maxBytes / 1024 / 1024} MB).");
        }

        var name = file.Name?.ToLowerInvariant() ?? "";
        var type = file.ContentType?.ToLowerInvariant() ?? "";

        try
        {
            if (name.EndsWith(".pdf") || type.Contains("pdf"))
            {
                var (text, error) = ExtractPdfText(bytes);
                return new ExtractedFile(text, error, Unsupported: false, PairMode.Layout);
            }

            // DOCX проверяем ДО XML: его MIME-тип содержит "openxmlformats" → "xml"
            if (name.EndsWith(".docx") || type.Contains("wordprocessingml"))
            {
                var (text, error) = DocxText.Extract(bytes);
                // Абзацы и ячейки Word задают пары; значение из соседнего абзаца — только в вертикальной форме
                return new ExtractedFile(text, error, Unsupported: false, PairMode.Paragraphs);
            }

            if (name.EndsWith(".doc") || type == "application/msword")
            {
                // Документ, а не картинка из подписи — прочитать нужно, поэтому это ошибка, а не пропуск
                return Failed("Old .doc format is not supported. Save the file as .docx.");
            }

            if (name.EndsWith(".xml") || type.Contains("xml"))
            {
                var (text, error) = ExtractXmlText(bytes);
                return new ExtractedFile(text, error, Unsupported: false, PairMode.Text);
            }
        }
        catch (Exception ex)
        {
            return Failed($"Could not read file: {ex.Message}");
        }

        return new ExtractedFile("", "Unsupported file type (only PDF, DOCX and XML).",
            Unsupported: true, PairMode.Text);
    }


    static (string Text, string? Error) ExtractPdfText(byte[] bytes)
    {
        using var pdf = PdfDocument.Open(bytes);

        var pages = new List<string>();

        foreach (var page in pdf.GetPages())
        {
            // PdfPig: ось Y вверх (Top > Bottom). PdfLayout работает "сверху вниз".
            var height = page.Height;
            var words = page.GetWords().Select(w => new WordBox(
                w.Text,
                w.BoundingBox.Left,
                w.BoundingBox.Right,
                height - w.BoundingBox.Top,
                height - w.BoundingBox.Bottom));

            // Склеиваем "метка:" со значением под ней или справа — по координатам,
            // а не по порядку текста (поля могут быть разбросаны по колонкам)
            var text = PdfLayout.BuildText(words);

            pages.Add(text.Length > 0 ? text : ContentOrderTextExtractor.GetText(page));
        }

        var result = string.Join("\n\n", pages).Trim();

        return result.Length > 0
            ? (result, null)
            : ("", "PDF has no text layer (scanned document?).");
    }


    static (string Text, string? Error) ExtractXmlText(byte[] bytes)
    {
        // XDocument.Load(Stream) сам определяет кодировку; DTD по умолчанию запрещены (защита от XXE)
        using var stream = new MemoryStream(bytes);
        var doc = XDocument.Load(stream);

        // <PostalCode>1234AB</PostalCode> → "Postal Code: 1234AB"
        var lines = doc.Descendants()
            .Where(e => !e.HasElements && !string.IsNullOrWhiteSpace(e.Value))
            .Select(e => $"{HumanizeName(e.Name.LocalName)}: {e.Value.Trim()}");

        var text = string.Join("\n", lines);

        return text.Length > 0
            ? (text, null)
            : ("", "XML contains no values.");
    }


    static string HumanizeName(string name)
    {
        var spaced = Regex.Replace(name, "([a-z0-9])([A-Z])", "$1 $2");
        return spaced.Replace('_', ' ').Replace('-', ' ');
    }
}


// ─────────────────────────────────────────────────────────────
// DOCX: абзацы, таблицы, текстовые блоки (text boxes), колонтитулы
// ─────────────────────────────────────────────────────────────

/// <summary>
/// В счетах из Word реквизиты часто лежат в текстовых блоках (надписях).
/// Их текст вложен в абзац-якорь, и наивный обход Descendants() склеивает всё в одну строку,
/// причём дважды: Word хранит копию блока для старых версий (mc:Fallback, VML).
///
/// Здесь:
///   • текст абзаца собирается без вложенных блоков;
///   • блоки выводятся отдельно, каждый абзац — своей строкой, в порядке на странице (сверху вниз, слева направо);
///   • из mc:AlternateContent берётся одна ветка (Choice, иначе Fallback);
///   • таблицы с заголовком (3+ колонки) — ячейки через таб, таблицы "метка | значение" — "метка: значение".
/// </summary>
public static class DocxText
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

    public static (string Text, string? Error) Extract(byte[] bytes)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);

        if (zip.GetEntry("word/document.xml") is null)
        {
            return ("", "Not a valid DOCX file.");
        }

        // Колонтитулы тоже: там часто стоят KVK / BTW / IBAN
        var headers = zip.Entries.Select(e => e.FullName)
            .Where(n => Regex.IsMatch(n, @"^word/header\d*\.xml$")).Order();
        var footers = zip.Entries.Select(e => e.FullName)
            .Where(n => Regex.IsMatch(n, @"^word/footer\d*\.xml$")).Order();

        var parts = headers.Append("word/document.xml").Concat(footers)
            .Select(path => ReadPart(zip, path))
            .Where(t => t.Length > 0)
            .ToList();

        // Повторяющиеся колонтитулы (первая страница / остальные) — один раз
        var text = string.Join("\n\n", parts.Distinct());

        return text.Length > 0
            ? (text, null)
            : ("", "DOCX contains no text.");
    }

    private static string ReadPart(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path);

        // Защита от zip-бомб
        if (entry is null || entry.Length > 50 * 1024 * 1024)
        {
            return "";
        }

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);

        // document.xml: w:document/w:body; header/footer: w:hdr / w:ftr
        var root = doc.Root?.Element(W + "body") ?? doc.Root;

        if (root is null)
        {
            return "";
        }

        var lines = new List<string>();
        Walk(root, lines);

        // Пустые строки оставляем как разделители блоков, но не больше одной подряд
        var result = new StringBuilder();
        var blank = true;

        foreach (var line in lines.SelectMany(l => l.Split('\n')))
        {
            if (line.Trim().Length == 0)
            {
                if (!blank) result.Append('\n');
                blank = true;
                continue;
            }

            result.Append(line.TrimEnd()).Append('\n');
            blank = false;
        }

        return result.ToString().Trim();
    }

    private static void Walk(XElement container, List<string> lines)
    {
        foreach (var child in container.Elements())
        {
            if (child.Name == W + "p")
            {
                lines.Add(ParagraphText(child));

                // Текстовые блоки, привязанные к абзацу — после него, в порядке расположения
                foreach (var box in TextBoxes(child).OrderBy(b => b.Top).ThenBy(b => b.Left))
                {
                    lines.Add("");
                    Walk(box.Content, lines);
                    lines.Add("");
                }
            }
            else if (child.Name == W + "tbl")
            {
                Table(child, lines);
            }
            else if (child.Name == W + "sectPr" || child.Name == W + "tblPr" || child.Name == W + "pPr")
            {
                // настройки, текста нет
            }
            else
            {
                // w:sdt/w:sdtContent (поля форм), w:customXml, w:ins и т.п.
                Walk(child, lines);
            }
        }
    }

    private static void Table(XElement table, List<string> lines)
    {
        var rows = new List<List<string>>();

        foreach (var row in table.Elements(W + "tr"))
        {
            // Пустые ячейки НЕ выбрасываем: пустая ячейка — это колонка без значения.
            // Если её убрать, значения справа съезжают под чужие заголовки.
            // Объединённая ячейка (gridSpan) и пропуск в начале строки (gridBefore) занимают колонки сетки.
            var cells = new List<string>();
            cells.AddRange(Enumerable.Repeat("", GridValue(row.Element(W + "trPr"), "gridBefore")));

            foreach (var tc in row.Elements(W + "tc"))
            {
                var cellLines = new List<string>();
                Walk(tc, cellLines);
                cells.Add(string.Join(" ", cellLines.Select(l => l.Trim()).Where(l => l.Length > 0)));
                cells.AddRange(Enumerable.Repeat("", Math.Max(GridValue(tc.Element(W + "tcPr"), "gridSpan"), 1) - 1));
            }

            if (cells.Any(c => c.Length > 0))
            {
                rows.Add(cells);
            }
        }

        if (rows.Count == 0)
        {
            return;
        }

        // Все строки — одной ширины (сетка таблицы)
        var width = rows.Max(r => r.Count);
        rows.ForEach(r => r.AddRange(Enumerable.Repeat("", width - r.Count)));

        var header = rows[0];

        lines.Add("");

        if (width == 2)
        {
            // "Naam | Jan Jansen" → "Naam: Jan Jansen"; "Naam | (пусто)" → "Naam:" — пустое поле
            rows.ForEach(r => lines.Add(Pair(r[0], r[1])));
        }
        else if (rows.Count == 2 &&
                 header.All(c => c.Length > 0 && !c.Any(char.IsDigit)) &&
                 !InvoiceStructure.IsItemTableHeader(header))
        {
            // Форма "заголовки сверху, значения под ними" (Naam | KVK | IBAN / Jan | 123 | NL..):
            // пары по колонкам, иначе значения достаются не тем меткам
            for (var c = 0; c < width; c++)
            {
                lines.Add(Pair(header[c], rows[1][c]));
            }
        }
        else
        {
            // Таблица позиций и прочие сетки: колонки через таб, пустые ячейки сохраняются ("a\t\tc")
            rows.ForEach(r => lines.Add(string.Join("\t", r)));
        }

        lines.Add("");
    }

    private static string Pair(string label, string value)
    {
        if (label.Length == 0)
        {
            return value;
        }

        var name = label.TrimEnd(':', ' ');
        return value.Length == 0 ? $"{name}:" : $"{name}: {value}";
    }

    private static int GridValue(XElement? properties, string name) =>
        int.TryParse(properties?.Element(W + name)?.Attribute(W + "val")?.Value, out var n) ? n : 0;

    private static string ParagraphText(XElement paragraph)
    {
        var sb = new StringBuilder();

        void Collect(XElement element)
        {
            foreach (var node in element.Elements())
            {
                var name = node.Name;

                if (name == W + "t") sb.Append(node.Value);
                else if (name == W + "tab") sb.Append('\t');
                else if (name == W + "br" || name == W + "cr") sb.Append('\n');
                else if (name == W + "noBreakHyphen") sb.Append('-');
                // Текстовые блоки и рисунки обрабатываются отдельно; удалённый текст и коды полей не нужны
                else if (name == W + "drawing" || name == W + "pict" || name == Mc + "AlternateContent" ||
                         name == W + "txbxContent" || name == W + "del" || name == W + "instrText" ||
                         name == W + "rPr" || name == W + "pPr") { }
                else Collect(node);
            }
        }

        Collect(paragraph);
        return sb.ToString();
    }

    private sealed record TextBox(XElement Content, long Top, long Left);

    private static List<TextBox> TextBoxes(XElement paragraph)
    {
        var boxes = new List<TextBox>();

        void Find(XElement element)
        {
            foreach (var node in element.Elements())
            {
                if (node.Name == Mc + "AlternateContent")
                {
                    // Одна ветка: Choice (DrawingML), если в ней есть текст, иначе Fallback (VML)
                    var choice = node.Elements(Mc + "Choice")
                        .FirstOrDefault(c => c.Descendants(W + "txbxContent").Any());
                    var branch = choice ?? node.Element(Mc + "Fallback");

                    if (branch is not null) Find(branch);
                }
                else if (node.Name == W + "txbxContent")
                {
                    var anchor = node.Ancestors(Wp + "anchor").FirstOrDefault();
                    boxes.Add(new TextBox(node, Offset(anchor, "positionV"), Offset(anchor, "positionH")));
                }
                else if (node.Name != W + "t" && node.Name != W + "rPr" && node.Name != W + "pPr")
                {
                    Find(node);
                }
            }
        }

        Find(paragraph);
        return boxes;
    }

    private static long Offset(XElement? anchor, string axis)
    {
        var value = anchor?.Element(Wp + axis)?.Element(Wp + "posOffset")?.Value;
        return long.TryParse(value, out var emu) ? emu : 0;
    }
}


// ─────────────────────────────────────────────────────────────
// Структура счёта: таблица позиций и адресные блоки без меток
// ─────────────────────────────────────────────────────────────

/// <summary>
/// Работает по уже извлечённому тексту (письмо + PDF + DOCX), поэтому одинаково для всех форматов.
///
/// Таблица позиций: строка-заголовок, где каждая ячейка — известное название колонки
/// ("Datum", "Aantal", "Omschrijving", "Prijs per stuk", "Totaal"). В PDF соседние заголовки могут
/// слипнуться в одну ячейку ("Aantal Omschrijving") — такие ячейки делятся по словарю.
/// Дальше берутся строки с тем же числом ячеек. Строки таблицы убираются из текста для поиска
/// по меткам, иначе "Product 1" превращается в article = "1".
///
/// Адресный блок: "Название / Улица 12 / 1234 AB Город", строки "Метка: значение" между ними
/// пропускаются (в PDF соседняя колонка "Klantnummer: ..." вклинивается между строками адреса).
/// Отправитель — блок с именем из "t.n.v. ..." или блок прямо над KVK/BTW/IBAN, иначе первый.
/// </summary>
public static class InvoiceStructure
{
    private static readonly (string Term, string Key)[] ColumnTerms = new (string Term, string Key)[]
    {
        ("prijs per stuk", "unitPrice"), ("prijs per eenheid", "unitPrice"), ("stukprijs", "unitPrice"),
        ("eenheidsprijs", "unitPrice"), ("unit price", "unitPrice"), ("price per unit", "unitPrice"),
        ("prijs", "unitPrice"), ("price", "unitPrice"), ("tarief", "unitPrice"),
        ("omschrijving", "article"), ("artikelnaam", "article"), ("artikel", "article"), ("product", "article"),
        ("description", "article"), ("article", "article"), ("item", "article"),
        ("hoeveelheid", "quantity"), ("aantal", "quantity"), ("stuks", "quantity"), ("quantity", "quantity"),
        ("qty", "quantity"),
        ("leverdatum", "date"), ("datum", "date"), ("date", "date"),
        ("totaalbedrag", "total"), ("subtotaal", "total"), ("totaal", "total"), ("bedrag", "total"),
        ("subtotal", "total"), ("total", "total"), ("amount", "total"),
        ("btw %", "vatRate"), ("btw%", "vatRate"), ("btw", "vatRate"), ("vat", "vatRate"),
        ("artikelnummer", "code"), ("artikelnr.", "code"), ("art.nr.", "code"), ("code", "code"), ("sku", "code"),
        ("eenheid", "unit"), ("unit", "unit"),
    }.OrderByDescending(t => t.Term.Length).ToArray();

    // Индекс как отдельная ячейка: "7901 EB Hoogeveen" (без IgnoreCase — буквы индекса заглавные)
    private static readonly Regex PostcodeCity =
        new(@"^(?<pc>[1-9]\d{3} ?[A-Z]{2})[ ]+(?<city>\p{Lu}[\p{L}'’\-. ]{1,40})$");

    // "Molenstraat 129", "Van Goghlaan 12-A", "Postbus 123"
    private static readonly Regex Street =
        new(@"^\p{L}[\p{L}\p{M}'’. \-]{1,60} \d{1,5}(?: ?[-/]? ?[A-Za-z0-9]{1,4})?$");

    private static readonly string[] NotAName =
        ["factuur", "invoice", "creditnota", "nederland", "netherlands", "the netherlands", "belgië", "belgie"];

    public static DocumentModel Analyze(string text)
    {
        var lines = text.Split('\n');
        var (items, tableLines) = FindTable(lines);
        var blocks = FindAddressBlocks(lines);
        var (supplier, customer) = AssignParties(text, lines, blocks);

        // Для поиска по меткам — текст без строк таблицы (номера строк сохраняем)
        var fieldText = string.Join("\n", lines.Select((l, i) => tableLines.Contains(i) ? "" : l));

        return new DocumentModel(text, fieldText, items, supplier, customer);
    }

    public static Dictionary<string, Derived> DerivedValues(DocumentModel doc)
    {
        var result = new Dictionary<string, Derived>();

        void Party(AddressBlock? block, string prefix, string who)
        {
            if (block is null) return;

            var note = $"Uit adresblok ({who}).";
            result[prefix is "" ? "name" : prefix + "Name"] = new(block.Name, 95, note);
            result[prefix is "" ? "address" : prefix + "Address"] = new(block.Street, 95, note);
            result[prefix is "" ? "postalCode" : prefix + "PostalCode"] = new(block.PostalCode, 95, note);
            result[prefix is "" ? "city" : prefix + "City"] = new(block.City, 95, note);
        }

        Party(doc.Supplier, "", "afzender factuur");
        Party(doc.Customer, "customer", "klant");

        if (doc.Lines.Count > 0)
        {
            var note = doc.Lines.Count == 1 ? "Uit factuurregel." : $"Uit {doc.Lines.Count} factuurregels.";

            var articles = doc.Lines.Select(l => l.Article).OfType<string>().ToList();
            if (articles.Count > 0)
                result["article"] = new(string.Join("; ", articles), 100, note);

            var quantities = doc.Lines.Select(l => l.Quantity).OfType<string>().ToList();
            if (quantities.Count > 0)
                result["quantity"] = new(string.Join("; ", quantities), 100, note);

            var dates = doc.Lines.Select(l => l.Date).OfType<string>().Distinct().ToList();
            if (dates.Count > 0)
                // Несколько разных дат — какую ставить, решает человек
                result["deliveryDate"] = dates.Count == 1
                    ? new(dates[0], 100, note)
                    : new(dates[0], 90, $"Meerdere data: {string.Join(", ", dates)}, controleer.", NeedsReview: true);
        }

        return result;
    }

    /// Суммы должны сходиться: сумма строк = excl. btw, excl. + btw = incl.
    public static void CheckTotals(Dictionary<string, FieldResult> fields, DocumentModel doc)
    {
        decimal? Get(string key) =>
            fields.TryGetValue(key, out var f) && f.Status == "found"
                ? decimal.Parse(f.Value, CultureInfo.InvariantCulture)
                : null;

        void Warn(string key, string message)
        {
            var f = fields[key];
            fields[key] = f with
            {
                Confidence = Math.Min(f.Confidence, 75),
                Status = "invalid",
                Note = message
            };
        }

        var excl = Get("totalExclVat");
        var vat = Get("vatAmount");
        var incl = Get("totalInclVat");

        var lineTotals = doc.Lines.Select(l => l.Total).ToList();

        if (excl is not null && lineTotals.Count > 0 && lineTotals.All(t => t is not null))
        {
            var sum = lineTotals.Sum(t => t!.Value);

            if (Math.Abs(sum - excl.Value) > 0.01m)
                Warn("totalExclVat", $"Som van de regels ({sum:0.00}) ≠ totaal excl. btw.");
        }

        if (excl is not null && vat is not null && incl is not null &&
            Math.Abs(excl.Value + vat.Value - incl.Value) > 0.01m)
        {
            Warn("totalInclVat", $"Excl. btw + btw ({excl + vat:0.00}) ≠ totaal incl. btw.");
        }
    }

    /// "€ 1.234,56" / "1234.56" / "EUR 30,00" / "-5,00" → decimal
    public static decimal? ParseAmount(string value)
    {
        var s = Regex.Replace(value, @"€|eur|\s", "", RegexOptions.IgnoreCase);

        if (!Regex.IsMatch(s, @"^-?\d[\d.,]*$"))
            return null;

        var lastComma = s.LastIndexOf(',');
        var lastDot = s.LastIndexOf('.');

        if (lastComma >= 0 && lastDot >= 0)
        {
            // Последний разделитель — десятичный
            s = lastComma > lastDot
                ? s.Replace(".", "").Replace(',', '.')
                : s.Replace(",", "");
        }
        else if (lastComma >= 0)
        {
            s = s.Replace(',', '.');
        }
        else if (Regex.IsMatch(s, @"^-?\d{1,3}(\.\d{3})+$"))
        {
            // "1.234" — тысячи (NL)
            s = s.Replace(".", "");
        }

        return decimal.TryParse(s, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture, out var amount)
            ? amount
            : null;
    }

    /// Непустые ячейки строки — для поиска адресных блоков (позиция колонки не важна)
    private static List<string> NonEmptyCells(string line) =>
        Regex.Split(line, @"\t| {2,}| \| ")
            .Select(c => c.Trim())
            .Where(c => c.Length > 0)
            .ToList();

    /// Ячейки с сохранением позиций — для таблиц. Строка с табами (DOCX-сетка, PDF-колонки):
    /// пустые ячейки остаются на своих местах. Без табов (текст письма): разрыв из 2+ пробелов или " | ".
    private static List<string> GridCells(string line) =>
        line.Contains('\t')
            ? line.Split('\t').Select(c => c.Trim()).ToList()
            : NonEmptyCells(line);

    /// Строка — заголовок таблицы позиций (Datum | Aantal | Omschrijving | ...)
    public static bool IsItemTableHeader(IEnumerable<string> cells) => ItemColumns(cells.ToList()) is not null;

    /// Ключи колонок по заголовку; пустая ячейка заголовка — колонка "_" (без смысла, но на своём месте)
    private static List<string>? ItemColumns(List<string> headerCells)
    {
        if (headerCells.All(c => c.Length == 0))
            return null;

        var columns = new List<string>();

        foreach (var cell in headerCells)
        {
            if (cell.Length == 0)
            {
                columns.Add("_");
                continue;
            }

            var keys = HeaderKeys(cell);

            if (keys is null)
                return null;

            columns.AddRange(keys);
        }

        return columns.Count(c => c != "_") >= 3 && (columns.Contains("article") || columns.Contains("quantity"))
            ? columns
            : null;
    }

    /// "Aantal Omschrijving" → [quantity, article]; не заголовок → null
    private static List<string>? HeaderKeys(string cell)
    {
        var rest = cell.Trim().ToLowerInvariant();
        var keys = new List<string>();

        while (rest.Length > 0)
        {
            var hit = ColumnTerms.FirstOrDefault(t =>
                rest.StartsWith(t.Term, StringComparison.Ordinal) &&
                (rest.Length == t.Term.Length || !char.IsLetterOrDigit(rest[t.Term.Length])));

            if (hit.Term is null)
                return null;

            keys.Add(hit.Key);
            rest = rest[hit.Term.Length..].TrimStart(' ', '.', ':', '(', ')', '€', '%');
        }

        return keys;
    }

    private static (List<InvoiceLine> Items, HashSet<int> Used) FindTable(string[] lines)
    {
        var items = new List<InvoiceLine>();
        var used = new HashSet<int>();

        for (var i = 0; i < lines.Length; i++)
        {
            var columns = ItemColumns(GridCells(lines[i]));

            if (columns is null)
                continue;

            var rows = new List<int>();

            for (var j = i + 1; j < lines.Length; j++)
            {
                var cells = GridCells(lines[j]);

                if (cells.All(c => c.Length == 0))
                {
                    // одна пустая строка внутри таблицы допустима
                    if (j + 1 < lines.Length && GridCells(lines[j + 1]).Count == columns.Count) continue;
                    break;
                }

                // Пустая ячейка = пустое значение (null), а не сдвиг соседних колонок
                string? Cell(string key)
                {
                    var index = columns.IndexOf(key);
                    return index >= 0 && cells[index].Length > 0 ? cells[index] : null;
                }

                // Строка позиции: та же ширина, есть товар или количество, есть цифры.
                // "Totaal excl. btw | | | | € 235,00" — не позиция: ни товара, ни количества
                if (cells.Count != columns.Count ||
                    (Cell("article") is null && Cell("quantity") is null) ||
                    !cells.Any(c => c.Any(char.IsDigit)))
                    break;

                rows.Add(j);

                var date = Cell("date");

                items.Add(new InvoiceLine(
                    date is null ? null : NormalizeDate(date),
                    Cell("quantity"),
                    Cell("article"),
                    Cell("unitPrice") is { } price ? ParseAmount(price) : null,
                    Cell("total") is { } total ? ParseAmount(total) : null));
            }

            if (rows.Count == 0)
                continue;

            used.Add(i);
            used.UnionWith(rows);
            i = rows[^1];
        }

        return (items, used);
    }

    private static string NormalizeDate(string value)
    {
        string[] formats = ["dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy", "d/M/yyyy", "dd.MM.yyyy", "d.M.yyyy", "yyyy-MM-dd"];

        return DateTime.TryParseExact(value.Trim(), formats, CultureInfo.InvariantCulture,
            DateTimeStyles.None, out var date)
            ? date.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture)
            : value;
    }

    private static List<AddressBlock> FindAddressBlocks(string[] lines)
    {
        var blocks = new List<AddressBlock>();

        for (var i = 0; i < lines.Length; i++)
        {
            foreach (var cell in NonEmptyCells(lines[i]))
            {
                var m = PostcodeCity.Match(cell);

                if (!m.Success)
                    continue;

                // Вверх от индекса: сначала улица, потом название. "Метка: значение" — пропускаем.
                string? street = null, name = null;
                var steps = 0;

                for (var j = i - 1; j >= 0 && steps < 6 && name is null; j--)
                {
                    var cells = NonEmptyCells(lines[j]);

                    if (cells.Count == 0)
                        continue;

                    steps++;

                    if (cells.All(c => c.Contains(':')))
                        continue;

                    if (street is null)
                    {
                        street = cells.FirstOrDefault(c => Street.IsMatch(c));
                        if (street is null) break;
                    }
                    else
                    {
                        name = cells.FirstOrDefault(IsName);
                        if (name is null) break;
                    }
                }

                if (street is not null && name is not null)
                {
                    var postalCode = m.Groups["pc"].Value;
                    postalCode = postalCode.Length == 6 ? $"{postalCode[..4]} {postalCode[4..]}" : postalCode;

                    blocks.Add(new AddressBlock(name, street, postalCode, m.Groups["city"].Value.Trim(), i));
                }
            }
        }

        // Один и тот же блок может встретиться дважды (письмо + вложение, PDF + DOCX)
        return blocks
            .GroupBy(b => (Key(b.Name), b.PostalCode))
            .Select(g => g.First())
            .ToList();
    }

    private static bool IsName(string cell) =>
        cell.Length is >= 2 and <= 60 &&
        cell.Any(char.IsLetter) &&
        !cell.Contains(':') && !cell.Contains('@') &&
        !cell.Contains("www.", StringComparison.OrdinalIgnoreCase) &&
        !Street.IsMatch(cell) && !PostcodeCity.IsMatch(cell) &&
        cell.Count(char.IsDigit) * 2 < cell.Length &&
        !NotAName.Contains(cell.ToLowerInvariant());

    private static string Key(string name) =>
        Regex.Replace(name.ToLowerInvariant(), @"[^\p{L}\p{N}]+", " ").Trim();

    private static (AddressBlock? Supplier, AddressBlock? Customer) AssignParties(
        string text, string[] lines, List<AddressBlock> blocks)
    {
        if (blocks.Count == 0)
            return (null, null);

        AddressBlock? supplier = null;

        // 1. "... op rekeningnummer t.n.v. Techflow BV onder vermelding ..." — получатель платежа
        foreach (Match m in Regex.Matches(text,
                     @"t\.\s?n\.\s?v\.\s+(?<name>[^\[\s][^\n]*?)(?=\s+onder\b|\s*[,;]|\s*\.?[ \t]*$)",
                     RegexOptions.IgnoreCase | RegexOptions.Multiline))
        {
            var payee = Key(m.Groups["name"].Value);
            supplier = blocks.FirstOrDefault(b => Key(b.Name) == payee)
                       ?? blocks.FirstOrDefault(b => payee.Length > 2 && (Key(b.Name).Contains(payee) || payee.Contains(Key(b.Name))));

            if (supplier is not null) break;
        }

        // 2. Блок, после которого идут KVK / BTW / IBAN (реквизиты отправителя)
        if (supplier is null)
        {
            var details = Array.FindIndex(lines, l =>
                Regex.IsMatch(l, @"(?:^|\t)\s*(?:kvk|btw|iban|vat|coc)\b", RegexOptions.IgnoreCase));

            if (details >= 0)
                supplier = blocks.LastOrDefault(b => b.Line < details);
        }

        // 3. Шапка письма/счёта — обычно первый блок
        supplier ??= blocks[0];

        var customer = blocks.FirstOrDefault(b => b != supplier);

        return (supplier, customer);
    }
}


// ─────────────────────────────────────────────────────────────
// Раскладка PDF: пары "метка → значение" по координатам
// ─────────────────────────────────────────────────────────────

/// Слово на странице. Координаты "сверху вниз": Top < Bottom.
public sealed record WordBox(string Text, double Left, double Right, double Top, double Bottom);

/// <summary>
/// Превращает слова страницы в текст, удобный для поиска полей.
///
/// Проблема: в PDF поля могут стоять в несколько колонок ("Naam:" слева, "KVK nummer:" справа,
/// значения под ними). Обычное извлечение текста перемешивает колонки:
/// "Naam: / KVK nummer: / Test Persoon / 00000000".
///
/// Решение:
///   1. слова → строки (по вертикали);
///   2. строки → сегменты: большой горизонтальный разрыв = другая колонка; кроме того,
///      сегмент режется перед новой меткой ("KVK: 123 BTW: 456" → "KVK: 123" | "BTW: 456");
///   3. каждой метке "...:" подбираем значение — сегмент ПОД ней (предпочтительно) или СПРАВА:
///      • снизу — только ближайший: между меткой и значением не должно быть других сегментов
///        (пустая "KVK:" не "перепрыгивает" через "BTW: ..." к следующей строке);
///      • справа — только соседний сегмент той же строки, и не из "свободной" колонки текста:
///        если в выровненной колонке есть строка без метки слева (адресный блок, абзац),
///        это отдельный блок, а не значения меток;
///      пары выбираются глобально по наименьшему расстоянию, чтобы соседние колонки
///      не "украли" значения друг у друга;
///   4. на выходе: "Naam: Test Persoon" — по строке на пару; метка без значения — отдельной строкой
///      (раскладка решила, что значения нет, — поиск по тексту не должен подставить чужое).
/// </summary>
public static class PdfLayout
{
    private sealed record Segment(string Text, double Left, double Right, double Top, double Bottom, int Line)
    {
        public double Height => Math.Max(Bottom - Top, 1);
        public bool IsLabel => Text.Length > 1 && Text.EndsWith(':') && Text.Any(char.IsLetter);

        /// Уже готовая пара "Datum afname: 12-12-2026" — не может быть значением чужой метки
        public bool IsInlinePair => Regex.IsMatch(Text, @"^\p{L}[\p{L} .\-]{0,60}:\s+\S");

        public bool OverlapsHorizontally(double left, double right) =>
            Math.Min(Right, right) - Math.Max(Left, left) > 0;
    }

    public static string BuildText(IEnumerable<WordBox> words)
    {
        var segments = ToSegments(words);

        if (segments.Count == 0)
        {
            return "";
        }

        var pairs = PairLabels(segments);
        var usedAsValue = pairs.Values.ToHashSet();

        var lines = new List<string>();

        foreach (var line in segments.GroupBy(s => s.Line).OrderBy(g => g.Key))
        {
            var rest = new List<string>();

            foreach (var segment in line.OrderBy(s => s.Left))
            {
                if (usedAsValue.Contains(segment))
                {
                    continue;
                }

                if (pairs.TryGetValue(segment, out var value))
                {
                    lines.Add($"{segment.Text} {value.Text}");
                }
                else if (segment.IsLabel)
                {
                    // Метка без значения — своей строкой, без соседей по табу
                    lines.Add(segment.Text);
                }
                else
                {
                    rest.Add(segment.Text);
                }
            }

            // Колонки одной строки без меток — через таб, чтобы не склеились в одно значение
            if (rest.Count > 0)
            {
                lines.Add(string.Join("\t", rest));
            }
        }

        return string.Join("\n", lines);
    }

    private static List<Segment> ToSegments(IEnumerable<WordBox> words)
    {
        var sorted = words
            .Where(w => !string.IsNullOrWhiteSpace(w.Text))
            .OrderBy(w => w.Top)
            .ThenBy(w => w.Left)
            .ToList();

        // 1. Строки: слово попадает в строку, если его середина внутри её высоты
        var lines = new List<List<WordBox>>();

        foreach (var word in sorted)
        {
            var middle = (word.Top + word.Bottom) / 2;
            var line = lines.FindLast(l => middle >= l.Min(w => w.Top) && middle <= l.Max(w => w.Bottom));

            if (line is null)
            {
                lines.Add([word]);
            }
            else
            {
                line.Add(word);
            }
        }

        // 2. Сегменты
        var segments = new List<Segment>();

        for (var lineNo = 0; lineNo < lines.Count; lineNo++)
        {
            var lineWords = lines[lineNo].OrderBy(w => w.Left).ToList();
            var current = new List<WordBox>();

            for (var i = 0; i < lineWords.Count; i++)
            {
                var word = lineWords[i];

                if (current.Count > 0)
                {
                    var previous = current[^1];
                    var fontHeight = Math.Max(previous.Bottom - previous.Top, word.Bottom - word.Top);

                    // Разрыв больше 1.5 высоты шрифта = другая колонка;
                    // новая метка после уже начатой метки/пары = другое поле, даже при обычном пробеле
                    // ("KVK: 123 BTW: 456" → "KVK: 123" | "BTW: 456"; "KVK: BTW: 456" → "KVK:" | "BTW: 456")
                    var columnBreak = word.Left - previous.Right > 1.5 * fontHeight;
                    var nextLabel = current.Any(w => w.Text.EndsWith(':')) && StartsLabel(lineWords, i);

                    if (columnBreak || nextLabel)
                    {
                        segments.Add(ToSegment(current, lineNo));
                        current = [];
                    }
                }

                current.Add(word);
            }

            if (current.Count > 0)
            {
                segments.Add(ToSegment(current, lineNo));
            }
        }

        return segments;
    }

    /// С этого слова начинается метка: 1–3 слова из букв, последнее кончается на ':' ("BTW:", "KVK nummer:")
    private static bool StartsLabel(List<WordBox> words, int start)
    {
        for (var i = start; i < Math.Min(start + 3, words.Count); i++)
        {
            var text = words[i].Text;

            if (!Regex.IsMatch(text, @"^\p{L}[\p{L}.\-]*:?$"))
                return false;

            if (text.EndsWith(':'))
                return true;
        }

        return false;
    }

    private static Segment ToSegment(List<WordBox> words, int line) => new(
        string.Join(" ", words.Select(w => w.Text)).Trim(),
        words.Min(w => w.Left),
        words.Max(w => w.Right),
        words.Min(w => w.Top),
        words.Max(w => w.Bottom),
        line);

    private static Dictionary<Segment, Segment> PairLabels(List<Segment> segments)
    {
        var freeColumn = FreeTextColumnSegments(segments);
        var candidates = new List<(Segment Label, Segment Value, double Cost)>();

        foreach (var label in segments.Where(s => s.IsLabel))
        {
            var h = label.Height;

            foreach (var value in segments.Where(s => !s.IsLabel && !s.IsInlinePair && s != label))
            {
                // Значение ПОД меткой: близко по вертикали, в той же колонке и без сегментов между ними
                var gapBelow = value.Top - label.Bottom;
                var leftShift = Math.Abs(value.Left - label.Left);

                if (value.Top > label.Top + h / 2 && gapBelow <= 2.5 * h &&
                    (value.OverlapsHorizontally(label.Left, label.Right) || leftShift <= 2 * h) &&
                    !IsBlockedVertically(segments, label, value))
                {
                    candidates.Add((label, value, Math.Max(gapBelow, 0) + 0.1 * leftShift));
                }

                // Значение СПРАВА на той же строке (дороже, чем снизу): ближайший сосед,
                // не из свободной колонки текста
                var gapRight = value.Left - label.Right;

                if (value.Line == label.Line && gapRight > 0 && gapRight <= 250 &&
                    !freeColumn.Contains(value) &&
                    !segments.Any(s => s.Line == label.Line && s.Left >= label.Right && s.Right <= value.Left && s != value))
                {
                    candidates.Add((label, value, 1000 + gapRight));
                }
            }
        }

        // Жадно по наименьшему расстоянию: каждая метка и каждое значение — максимум в одной паре
        var pairs = new Dictionary<Segment, Segment>();
        var taken = new HashSet<Segment>();

        foreach (var (label, value, _) in candidates.OrderBy(c => c.Cost))
        {
            if (!pairs.ContainsKey(label) && taken.Add(value))
            {
                pairs[label] = value;
            }
        }

        return pairs;
    }

    /// Есть ли другой сегмент между меткой и значением под ней (в той же полосе по горизонтали)
    private static bool IsBlockedVertically(List<Segment> segments, Segment label, Segment value)
    {
        var left = Math.Min(label.Left, value.Left);
        var right = Math.Max(label.Right, value.Right);

        return segments.Any(s =>
            s != label && s != value &&
            s.Top >= label.Bottom - 0.5 && s.Bottom <= value.Top + 0.5 &&
            s.OverlapsHorizontally(left, right));
    }

    /// <summary>
    /// Сегменты "свободных" колонок: выровненные по левому краю строки подряд,
    /// где хотя бы у одной строки слева нет метки. Это адресный блок или абзац —
    /// его строки не могут быть значениями меток слева, даже если стоят на той же высоте.
    /// </summary>
    private static HashSet<Segment> FreeTextColumnSegments(List<Segment> segments)
    {
        bool HasLabelOnLeft(Segment s) =>
            segments
                .Where(o => o.Line == s.Line && o.Right <= s.Left)
                .OrderByDescending(o => o.Right)
                .FirstOrDefault() is { IsLabel: true };

        var candidates = segments.Where(s => !s.IsLabel && !s.IsInlinePair).ToList();
        var result = new HashSet<Segment>();
        var visited = new HashSet<Segment>();

        foreach (var start in candidates)
        {
            if (!visited.Add(start))
                continue;

            // Собираем колонку: соседние строки, левый край совпадает
            var column = new List<Segment> { start };

            for (var i = 0; i < column.Count; i++)
            {
                var s = column[i];

                foreach (var other in candidates)
                {
                    if (visited.Contains(other))
                        continue;

                    var h = Math.Max(s.Height, other.Height);
                    var verticalGap = Math.Max(other.Top - s.Bottom, s.Top - other.Bottom);

                    if (Math.Abs(other.Left - s.Left) <= h && verticalGap <= 1.2 * h && other.Line != s.Line)
                    {
                        visited.Add(other);
                        column.Add(other);
                    }
                }
            }

            if (column.Count > 1 && column.Any(s => !HasLabelOnLeft(s)))
            {
                result.UnionWith(column);
            }
        }

        return result;
    }
}


// ─────────────────────────────────────────────────────────────
// Модели
// ─────────────────────────────────────────────────────────────

public record MailRequest(
    string? ItemId,
    string? Subject,
    string? Sender,
    string? Body,
    MailAttachment[]? Attachments,
    UploadedFile[]? Files,
    string[]? RequestedFields // какие поля обязательно вернуть; пусто = все
);

public record MailAttachment(
    string Id,
    string Name,
    long Size,
    string ContentType
);

public record UploadedFile(
    string? Name,
    string? ContentType,
    string? Base64,
    string? Source // "attachment" или "upload"
);

/// Status: ok | skipped (тип не поддерживается) | error (не удалось прочитать — блокирует readyToSend)
public record FileProcessResult(
    string Name,
    string Source,
    int Characters,
    string Status,
    string? Error
);

/// Как в источнике связаны метка и значение (см. ExtractFields)
public enum PairMode
{
    /// PDF: пары собраны по координатам
    Layout,
    /// DOCX: абзацы и ячейки Word
    Paragraphs,
    /// Текст письма, XML
    Text
}

public record ExtractedFile(string Text, string? Error, bool Unsupported, PairMode Pairs);

/// Kind: email | file
public record SourceText(string Kind, string Name, string Text, PairMode Pairs);

/// InvoiceId: "nr:<номер>" (или "nr:<номер>#2" — тот же номер у другой компании) | "doc:<n>" (фактура без номера)
///            | null (не фактура: письмо, приложение)
public record DocumentResult(
    string Kind,
    string Name,
    string? InvoiceId,
    Dictionary<string, FieldResult> Fields,
    List<InvoiceLine> Lines,
    bool ReadyToSend,
    string ExtractedText
);

/// Labels = null — поле только из структуры документа (адресный блок, таблица)
/// Strict = true — значение должно совпасть с ValuePattern даже после двоеточия
/// Context — шаблоны фраз без метки, группа (?<value>...)
public record FieldSpec(
    string Key,
    string? Labels,
    string ValuePattern = @"[^\n\t]+?",
    Func<string, Check>? Validate = null,
    string? FallbackPattern = null,
    bool Strict = false,
    string[]? Context = null
)
{
    public const string AmountPattern = @"(?:€|EUR)?[ \t]*-?\d[\d.,]*";
}

// Результат проверки: Problem → invalid, Warning → значение принято, но с пометкой
public record Check(string Value, string? Problem = null, string? Warning = null);

// Status: found | review (сомнительно, проверить) | guessed (без метки) | invalid | missing
//         | conflict (документы противоречат друг другу или это разные фактуры)
// Method: label | nextLine | context | structure | pattern | none — КАК найдено значение
// Confidence: порядковая оценка способа извлечения, НЕ вероятность. readyToSend её не использует.
// Source — из какого документа значение ("E-mail", имя файла; несколько — через запятую)
public record FieldResult(
    string Value,
    int Confidence,
    string Status,
    string? Note,
    string? Source = null,
    List<FieldAlternative>? Alternatives = null,
    string Method = "none"
);

public record FieldAlternative(string Value, string Source, string Status);

public record Derived(string Value, int Confidence, string Note, bool NeedsReview = false);

// Признаки фактуры: номер (нормализованный) и реквизиты поставщика (kvk / vat / name)
public record InvoiceKey(string? Number, Dictionary<string, string> Supplier);

public record InvoiceLine(string? Date, string? Quantity, string? Article, decimal? UnitPrice, decimal? Total);

public record AddressBlock(string Name, string Street, string PostalCode, string City, int Line);

// FieldText — текст без строк таблицы позиций (для поиска по меткам)
public record DocumentModel(
    string Text,
    string FieldText,
    List<InvoiceLine> Lines,
    AddressBlock? Supplier,
    AddressBlock? Customer
);

public record AnalyzeResponse(
    bool Success,
    string? ItemId,
    string? Subject,
    string? Sender,
    Dictionary<string, FieldResult> Fields,
    List<InvoiceLine> Lines,
    List<DocumentResult> Documents,
    MailAttachment[]? Attachments,
    int AttachmentCount,
    List<FileProcessResult> Files,
    List<string> RequestedFields,
    bool ReadyToSend,
    List<string> Issues,
    List<string> Warnings,
    string ExtractedText,
    DateTimeOffset ProcessedAt
);
