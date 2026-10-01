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

    // Текст письма + текст всех PDF/XML файлов
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
            combined.AppendLine().AppendLine().AppendLine(fileText);
        }
    }

    var text = combined.ToString().Trim();

    var customerType = Extract(
        text,
        @"^\s*(?:nieuwe klant of bestaande klant|klant|customer)\s*[:\-]\s*(?<value>nieuwe klant|bestaande klant|new|existing)\s*$"
    );

    customerType = customerType with
    {
        Value = customerType.Value.ToLowerInvariant() switch
        {
            "nieuwe klant" or "new" => "new",
            "bestaande klant" or "existing" => "existing",
            _ => ""
        }
    };

    var fields = new Dictionary<string, FieldResult>
    {
        ["name"] = Extract(
            text,
            @"^\s*(?:naam|name)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["address"] = Extract(
            text,
            @"^\s*(?:adres|address)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["postalCode"] = Extract(
            text,
            @"^\s*(?:postcode|postal code)\s*[:\-]\s*(?<value>\d{4}\s?[A-Z]{2})",
            @"\b(?<value>\d{4}\s?[A-Z]{2})\b"
        ),

        ["city"] = Extract(
            text,
            @"^\s*(?:woonplaats|city)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["iban"] = Extract(
            text,
            @"^\s*iban\s*[:\-]\s*(?<value>[A-Z]{2}\d{2}[A-Z0-9 ]{10,30})",
            @"\b(?<value>[A-Z]{2}\d{2}[A-Z0-9]{11,30})\b"
        ),

        ["customerType"] = customerType,

        ["kvk"] = Extract(
            text,
            @"^\s*(?:kvk nummer|kvk-nummer|kvk)\s*[:\-]\s*(?<value>\d{8})"
        ),

        ["vat"] = Extract(
            text,
            @"^\s*(?:btw nummer|btw-nummer|vat number|vat)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["invoiceNumber"] = Extract(
            text,
            @"^\s*(?:factuurnummer|invoice number)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["article"] = Extract(
            text,
            @"^\s*(?:artikel|article|item)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["quantity"] = Extract(
            text,
            @"^\s*(?:hoeveelheid|quantity)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["deliveryDate"] = Extract(
            text,
            @"^\s*(?:datum afname|delivery date)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["paymentDueDate"] = Extract(
            text,
            @"^\s*(?:uiterste betalingsdatum|payment due date)\s*[:\-]\s*(?<value>[^\r\n]+)"
        ),

        ["paymentArrears"] = Extract(
            text,
            @"^\s*(?:betalingsachterstanden|payment arrears)\s*[:\-]\s*(?<value>[^\r\n]+)"
        )
    };

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
        processedAt = DateTimeOffset.UtcNow
    });
});

app.Run();


static FieldResult Extract(
    string text,
    string pattern,
    string? fallbackPattern = null)
{
    var match = Regex.Match(
        text,
        pattern,
        RegexOptions.IgnoreCase | RegexOptions.Multiline
    );

    if (match.Success)
    {
        return new FieldResult(
            match.Groups["value"].Value.Trim(),
            100
        );
    }

    if (fallbackPattern is null)
    {
        return new FieldResult("", 1);
    }

    match = Regex.Match(
        text,
        fallbackPattern,
        RegexOptions.IgnoreCase | RegexOptions.Multiline
    );

    return match.Success
        ? new FieldResult(match.Groups["value"].Value.Trim(), 85)
        : new FieldResult("", 1);
}


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

                    // Таблица "Naam | Jan Jansen" → "Naam: Jan Jansen", чтобы сработали регулярки
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

    var text = string.Join(Environment.NewLine, lines.Where(l => l.Trim().Length > 0));

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

    var sb = new StringBuilder();

    foreach (var page in pdf.GetPages())
    {
        // Сохраняет порядок строк, так что "Naam: ..." остаётся на одной строке
        sb.AppendLine(ContentOrderTextExtractor.GetText(page));
    }

    var text = sb.ToString().Trim();

    return text.Length > 0
        ? (text, null)
        : ("", "PDF has no text layer (scanned document?).");
}


static (string Text, string? Error) ExtractXmlText(byte[] bytes)
{
    // XDocument.Load(Stream) сам определяет кодировку; DTD по умолчанию запрещены (защита от XXE)
    using var stream = new MemoryStream(bytes);
    var doc = XDocument.Load(stream);

    // Каждый конечный элемент -> строка "label: value",
    // чтобы сработали те же регулярки, что и для письма.
    // <PostalCode>1234AB</PostalCode> -> "Postal Code: 1234AB"
    var lines = doc.Descendants()
        .Where(e => !e.HasElements && !string.IsNullOrWhiteSpace(e.Value))
        .Select(e => $"{HumanizeName(e.Name.LocalName)}: {e.Value.Trim()}");

    var text = string.Join(Environment.NewLine, lines);

    return text.Length > 0
        ? (text, null)
        : ("", "XML contains no values.");
}


static string HumanizeName(string name)
{
    var spaced = Regex.Replace(name, "([a-z0-9])([A-Z])", "$1 $2");
    return spaced.Replace('_', ' ').Replace('-', ' ');
}


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

public record FieldResult(
    string Value,
    int Confidence
);