using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

const long MaxFileBytes = 10 * 1024 * 1024; // 10 MB на файл

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
    var uploadedFiles = mail.Files ?? [];

    if (string.IsNullOrWhiteSpace(mail.Subject) &&
        string.IsNullOrWhiteSpace(mail.Body) &&
        uploadedFiles.Length == 0)
    {
        return Results.BadRequest(new
        {
            error = "Email subject, body and files are empty."
        });
    }

    // Текст письма + текст всех файлов
    var combined = new StringBuilder(mail.Body?.Trim() ?? "");
    var fileResults = new List<FileProcessResult>();

    foreach (var file in uploadedFiles)
    {
        var (fileText, error) = ExtractFileText(file, MaxFileBytes);

        fileResults.Add(new FileProcessResult(
            file.Name ?? "",
            file.Source ?? "",
            fileText.Length,
            error
        ));

        if (fileText.Length > 0)
        {
            combined.Append("\n\n").Append(fileText);
        }
    }

    var text = NormalizeText(combined.ToString());

    var fields = ExtractFields(text);

    return Results.Ok(new
    {
        success = true,
        mail.ItemId,
        mail.Subject,
        mail.Sender,
        fields,
        attachments = mail.Attachments,
        attachmentCount = mail.Attachments?.Length ?? 0,
        files = fileResults,
        readyToSend = fields.Values.All(field => field.Confidence == 100),
        // Для отладки: какой текст сервер реально увидел
        extractedText = text.Length > 5000 ? text[..5000] + "…" : text,
        processedAt = DateTimeOffset.UtcNow
    });
});

app.Run();


// ─────────────────────────────────────────────────────────────
// Извлечение полей
// ─────────────────────────────────────────────────────────────

static FieldSpec[] FieldSpecs() =>
[
    // Метки: длинные варианты ПЕРВЫМИ ("btw nummer" раньше "btw")
    new("name",
        Labels: "klantnaam|naam klant|volledige naam|naam|full name|name"),

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
        Labels: "iban[- ]?nummer|rekeningnummer|iban",
        ValuePattern: @"[A-Za-z]{2}\d{2}[A-Za-z0-9 ]{10,40}?",
        Validate: ValidateIban,
        FallbackPattern: @"\b(?<value>[A-Z]{2}\d{2}[A-Z0-9]{11,30})\b"),

    new("customerType",
        Labels: "nieuwe klant of bestaande klant|klanttype|soort klant|klant|customer type|customer",
        ValuePattern: @"nieuwe klant|bestaande klant|nieuw|bestaand|new|existing",
        Validate: ValidateCustomerType),

    new("kvk",
        Labels: @"kvk[- ]?nummer|kvk[- ]?nr\.?|kamer van koophandel|kvk|coc number|coc",
        ValuePattern: @"[\d .]{8,12}",
        Validate: ValidateKvk),

    new("vat",
        Labels: @"btw[- ]?nummer|btw[- ]?id|btw[- ]?nr\.?|btw|vat number|vat id|vat",
        Validate: ValidateVat,
        FallbackPattern: @"\b(?<value>NL\d{9}B\d{2})\b"),

    new("invoiceNumber",
        Labels: @"factuurnummer|factuur[- ]?nr\.?|factuur|invoice number|invoice no\.?|invoice"),

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
        Validate: ValidateDate),

    new("paymentArrears",
        Labels: "betalingsachterstanden|betalingsachterstand|achterstand|payment arrears|arrears")
];


static Dictionary<string, FieldResult> ExtractFields(string text)
{
    // Метки, которые не являются нашими полями, но тоже не могут быть значением
    // (иначе при пустом "Naam" значением станет следующая строка "BSN")
    const string otherLabels =
        "bsn|type dossier|e-?mail(?:adres)?|telefoon(?:nummer)?|geboortedatum";

    var specs = FieldSpecs();
    var allLabels = string.Join("|", specs.Select(s => s.Labels).Append(otherLabels));

    return specs.ToDictionary(s => s.Key, s => Field(text, s, allLabels));
}


/// Ищет значение по метке в трёх вариантах:
///   1. "Naam: Test Persoon" / "Naam - Test Persoon" / "Naam Test Persoon" (одна строка)
///   2. "Naam" и "Test Persoon" на соседних строках (PDF-таблица, прочитанная по колонкам)
///   3. fallback без метки (только для форматов, которые узнаются сами: IBAN, postcode, BTW)
static FieldResult Field(string text, FieldSpec spec, string allLabels)
{
    var timeout = TimeSpan.FromMilliseconds(500);
    var options = RegexOptions.IgnoreCase | RegexOptions.Multiline;

    // (?>...) — атомарная группа: если совпала длинная метка "BTW nummer",
    // движок не откатится к короткой "BTW" и не возьмёт "nummer" как значение.
    // (?![\p{L}\p{N}]) — конец слова; работает и после точки ("KVK nr.").
    const string endOfWord = @"(?![\p{L}\p{N}])";

    // Метка в начале строки или после таба (колонки в одной строке разделены табом)
    var label = $@"(?:^|\t)[ \t]*(?>{spec.Labels}){endOfWord}";

    // Хвост метки до двоеточия: "Betalingsachterstanden binnen de eigen company:"
    const string labelTail = @"[ \t]+[\p{L} ]{1,40}?:";

    var notALabel = $@"(?!(?:{allLabels}){endOfWord})";

    // Значение заканчивается на конце строки или на табе (следующая колонка)
    var value = $@"{notALabel}(?<value>{spec.ValuePattern})(?=[ \t]*(?:\t|$))";

    var sameLine = $@"{label}(?:[ \t]*[:\-–][ \t]*|{labelTail}[ \t]*|[ \t]+){value}";
    var nextLine = $@"{label}(?:{labelTail}|[ \t]*[:\-–]?)[ \t]*\n(?:[ \t]*\n)*[ \t]*{value}";

    foreach (var pattern in new[] { sameLine, nextLine })
    {
        var match = Regex.Match(text, pattern, options, timeout);

        if (match.Success)
        {
            return Score(match.Groups["value"].Value, spec.Validate, 100, "found");
        }
    }

    if (spec.FallbackPattern is not null)
    {
        // Без IgnoreCase: иначе "2026 Al..." похоже на почтовый индекс
        var match = Regex.Match(text, spec.FallbackPattern, RegexOptions.Multiline, timeout);

        if (match.Success)
        {
            return Score(match.Groups["value"].Value, spec.Validate, 85, "guessed");
        }
    }

    return new FieldResult("", 1, "missing", "Niet gevonden in e-mail of bestanden.");
}


static FieldResult Score(
    string raw,
    Func<string, (string Value, string? Problem)>? validate,
    int baseConfidence,
    string status)
{
    var value = raw.Trim();

    if (validate is null)
    {
        return new FieldResult(value, baseConfidence, status,
            status == "guessed" ? "Gevonden zonder label, controleer." : null);
    }

    var (normalized, problem) = validate(value);

    if (problem is not null)
    {
        return new FieldResult(normalized, Math.Min(baseConfidence, 75), "invalid", problem);
    }

    return new FieldResult(normalized, baseConfidence, status,
        status == "guessed" ? "Gevonden zonder label, controleer." : null);
}


// ─────────────────────────────────────────────────────────────
// Валидация и нормализация
// ─────────────────────────────────────────────────────────────

static (string, string?) ValidatePostalCode(string value)
{
    var m = Regex.Match(value, @"^(\d{4})\s?([A-Za-z]{2})$");

    if (!m.Success)
        return (value, "Ongeldige postcode.");

    var normalized = $"{m.Groups[1].Value} {m.Groups[2].Value.ToUpperInvariant()}";

    return m.Groups[1].Value[0] == '0'
        ? (normalized, "Nederlandse postcode begint niet met 0.")
        : (normalized, null);
}


static (string, string?) ValidateIban(string value)
{
    var iban = Regex.Replace(value, @"\s", "").ToUpperInvariant();

    if (!Regex.IsMatch(iban, @"^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$"))
        return (iban, "Ongeldig IBAN-formaat.");

    if (iban.StartsWith("NL") && iban.Length != 18)
        return (FormatIban(iban), "Nederlandse IBAN moet 18 tekens hebben.");

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
        ? (FormatIban(iban), null)
        : (FormatIban(iban), "IBAN-controlegetal klopt niet.");
}


static string FormatIban(string iban) =>
    string.Join(" ", Enumerable.Range(0, (iban.Length + 3) / 4)
        .Select(i => iban.Substring(i * 4, Math.Min(4, iban.Length - i * 4))));


static (string, string?) ValidateCustomerType(string value) =>
    value.Trim().ToLowerInvariant() switch
    {
        "nieuwe klant" or "nieuw" or "new" => ("new", null),
        "bestaande klant" or "bestaand" or "existing" => ("existing", null),
        _ => (value, "Onbekend klanttype.")
    };


static (string, string?) ValidateKvk(string value)
{
    var digits = Regex.Replace(value, @"[\s.]", "");

    return Regex.IsMatch(digits, @"^\d{8}$")
        ? (digits, null)
        : (digits, "KVK-nummer moet 8 cijfers hebben.");
}


static (string, string?) ValidateVat(string value)
{
    var vat = Regex.Replace(value, @"[\s.]", "").ToUpperInvariant();

    if (Regex.IsMatch(vat, @"^NL\d{9}B\d{2}$"))
        return (vat, null);

    return Regex.IsMatch(vat, @"^[A-Z]{2}[A-Z0-9]{2,13}$")
        ? (vat, "Geen Nederlands BTW-formaat (NL123456789B01).")
        : (vat, "Ongeldig BTW-nummer.");
}


static (string, string?) ValidateQuantity(string value)
{
    return Regex.IsMatch(value, @"^\d+(?:[.,]\d+)?\b")
        ? (value, null)
        : (value, "Hoeveelheid begint niet met een getal.");
}


static (string, string?) ValidateDate(string value)
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
            return (date.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture), null);
        }
    }

    return (value, "Datum niet herkend.");
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
        .Replace('\u00A0', ' ')   // non-breaking space из PDF/Word
        .Trim();


// ─────────────────────────────────────────────────────────────
// Чтение файлов
// ─────────────────────────────────────────────────────────────

static (string Text, string? Error) ExtractFileText(UploadedFile file, long maxBytes)
{
    byte[] bytes;

    try
    {
        bytes = Convert.FromBase64String(file.Base64 ?? "");
    }
    catch (FormatException)
    {
        return ("", "Invalid base64 content.");
    }

    if (bytes.Length == 0)
    {
        return ("", "File is empty.");
    }

    if (bytes.Length > maxBytes)
    {
        return ("", $"File too large (max {maxBytes / 1024 / 1024} MB).");
    }

    var name = file.Name?.ToLowerInvariant() ?? "";
    var type = file.ContentType?.ToLowerInvariant() ?? "";

    try
    {
        if (name.EndsWith(".pdf") || type.Contains("pdf"))
        {
            return ExtractPdfText(bytes);
        }

        // DOCX проверяем ДО XML: его MIME-тип содержит "openxmlformats" → "xml"
        if (name.EndsWith(".docx") || type.Contains("wordprocessingml"))
        {
            return ExtractDocxText(bytes);
        }

        if (name.EndsWith(".doc") || type == "application/msword")
        {
            return ("", "Old .doc format is not supported. Save the file as .docx.");
        }

        if (name.EndsWith(".xml") || type.Contains("xml"))
        {
            return ExtractXmlText(bytes);
        }
    }
    catch (Exception ex)
    {
        return ("", $"Could not read file: {ex.Message}");
    }

    return ("", "Unsupported file type (only PDF, DOCX and XML).");
}


static (string Text, string? Error) ExtractDocxText(byte[] bytes)
{
    // DOCX = ZIP-архив, текст лежит в word/document.xml
    using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);

    var entry = zip.GetEntry("word/document.xml");

    if (entry is null)
    {
        return ("", "Not a valid DOCX file.");
    }

    // Защита от zip-бомб
    if (entry.Length > 50 * 1024 * 1024)
    {
        return ("", "DOCX content too large.");
    }

    using var stream = entry.Open();
    var doc = XDocument.Load(stream);

    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var body = doc.Root?.Element(w + "body");

    if (body is null)
    {
        return ("", "DOCX has no body.");
    }

    var lines = new List<string>();

    void Walk(XElement element)
    {
        foreach (var child in element.Elements())
        {
            if (child.Name == w + "p")
            {
                lines.Add(ParagraphText(child, w));
            }
            else if (child.Name == w + "tbl")
            {
                foreach (var row in child.Elements(w + "tr"))
                {
                    var cells = row.Elements(w + "tc")
                        .Select(tc => string.Join(" ",
                            tc.Elements(w + "p").Select(p => ParagraphText(p, w)))
                            .Trim())
                        .Where(c => c.Length > 0)
                        .ToList();

                    // Таблица "Naam | Jan Jansen" → "Naam: Jan Jansen"
                    lines.Add(cells.Count == 2
                        ? $"{cells[0].TrimEnd(':', ' ')}: {cells[1]}"
                        : string.Join(" | ", cells));
                }
            }
            else if (child.Name == w + "sdt")
            {
                // Content controls (поля форм в Word)
                var content = child.Element(w + "sdtContent");
                if (content is not null) Walk(content);
            }
        }
    }

    Walk(body);

    var text = string.Join("\n", lines.Where(l => l.Trim().Length > 0));

    return text.Length > 0
        ? (text, null)
        : ("", "DOCX contains no text.");
}


static string ParagraphText(XElement paragraph, XNamespace w)
{
    var sb = new StringBuilder();

    foreach (var node in paragraph.Descendants())
    {
        if (node.Name == w + "t") sb.Append(node.Value);
        else if (node.Name == w + "tab") sb.Append('\t');
        else if (node.Name == w + "br") sb.Append('\n');
    }

    return sb.ToString();
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
///   2. строки → сегменты (большой горизонтальный разрыв = другая колонка);
///   3. каждой метке "...:" подбираем значение — сегмент ПОД ней (предпочтительно) или СПРАВА;
///      пары выбираются глобально по наименьшему расстоянию, чтобы соседние колонки
///      не "украли" значения друг у друга;
///   4. на выходе: "Naam: Test Persoon" — по строке на пару, остальное как есть.
/// </summary>
public static class PdfLayout
{
    private sealed record Segment(string Text, double Left, double Right, double Top, double Bottom, int Line)
    {
        public double Height => Math.Max(Bottom - Top, 1);
        public bool IsLabel => Text.Length > 1 && Text.EndsWith(':');

        /// Уже готовая пара "Datum afname: 12-12-2026" — не может быть значением чужой метки
        public bool IsInlinePair => Regex.IsMatch(Text, @"^\p{L}[\p{L} .\-]{0,60}:\s+\S");
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

        // 2. Сегменты: разрыв больше 1.5 высоты шрифта = другая колонка
        var segments = new List<Segment>();

        for (var lineNo = 0; lineNo < lines.Count; lineNo++)
        {
            var current = new List<WordBox>();

            foreach (var word in lines[lineNo].OrderBy(w => w.Left))
            {
                if (current.Count > 0)
                {
                    var previous = current[^1];
                    var fontHeight = Math.Max(previous.Bottom - previous.Top, word.Bottom - word.Top);

                    if (word.Left - previous.Right > 1.5 * fontHeight)
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

    private static Segment ToSegment(List<WordBox> words, int line) => new(
        string.Join(" ", words.Select(w => w.Text)).Trim(),
        words.Min(w => w.Left),
        words.Max(w => w.Right),
        words.Min(w => w.Top),
        words.Max(w => w.Bottom),
        line);

    private static Dictionary<Segment, Segment> PairLabels(List<Segment> segments)
    {
        var candidates = new List<(Segment Label, Segment Value, double Cost)>();

        foreach (var label in segments.Where(s => s.IsLabel))
        {
            var h = label.Height;

            foreach (var value in segments.Where(s => !s.IsLabel && !s.IsInlinePair && s != label))
            {
                // Значение ПОД меткой: близко по вертикали и в той же колонке
                var gapBelow = value.Top - label.Bottom;
                var overlap = Math.Min(label.Right, value.Right) - Math.Max(label.Left, value.Left);
                var leftShift = Math.Abs(value.Left - label.Left);

                if (value.Top > label.Top + h / 2 && gapBelow <= 2.5 * h && (overlap > 0 || leftShift <= 2 * h))
                {
                    candidates.Add((label, value, Math.Max(gapBelow, 0) + 0.1 * leftShift));
                }

                // Значение СПРАВА на той же строке (дороже, чем снизу)
                var gapRight = value.Left - label.Right;

                if (value.Line == label.Line && gapRight > 0 && gapRight <= 250)
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
    UploadedFile[]? Files
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

public record FileProcessResult(
    string Name,
    string Source,
    int Characters,
    string? Error
);

public record FieldSpec(
    string Key,
    string Labels,
    string ValuePattern = @"[^\n\t]+?",
    Func<string, (string Value, string? Problem)>? Validate = null,
    string? FallbackPattern = null
);

/// Status: found | guessed | invalid | missing
public record FieldResult(
    string Value,
    int Confidence,
    string Status,
    string? Note
);