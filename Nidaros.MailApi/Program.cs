using System.Text.RegularExpressions;

var builder = WebApplication.CreateBuilder(args);

var app = builder.Build();

app.UseStaticFiles();

app.MapGet("/health", () => Results.Ok(new
{
    status = "OK",
    service = "Nidaros Mail API"
}));

app.MapPost("/api/mail/analyze", (MailRequest mail) =>
{
    var role = mail.Role?.Trim().ToLowerInvariant();

    if (role is not ("student" or "entrepreneur" or "worker"))
    {
        return Results.BadRequest(new
        {
            error = "Unknown role."
        });
    }

    if (string.IsNullOrWhiteSpace(mail.Subject) &&
        string.IsNullOrWhiteSpace(mail.Body))
    {
        return Results.BadRequest(new
        {
            error = "Email subject and body are empty."
        });
    }

    var text = mail.Body?.Trim() ?? "";

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

        ["bsn"] = Extract(
            text,
            @"^\s*bsn\s*[:\-]\s*(?<value>\d{9})",
            @"\b(?<value>\d{9})\b"
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

    if (role == "entrepreneur")
    {
        fields["kvk"] = Extract(
            text,
            @"^\s*(?:kvk nummer|kvk-nummer|kvk)\s*[:\-]\s*(?<value>\d{8})"
        );
    }

    return Results.Ok(new
    {
        success = true,
        role,
        mail.ItemId,
        mail.Subject,
        mail.Sender,
        fields,
        attachments = mail.Attachments,
        attachmentCount = mail.Attachments?.Length ?? 0,
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


public record MailRequest(
    string? Role,
    string? ItemId,
    string? Subject,
    string? Sender,
    string? Body,
    MailAttachment[]? Attachments
);

public record MailAttachment(
    string Id,
    string Name,
    long Size,
    string ContentType
);

public record FieldResult(
    string Value,
    int Confidence
);