
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
    if (string.IsNullOrWhiteSpace(mail.Subject) &&
        string.IsNullOrWhiteSpace(mail.Body))
    {
        return Results.BadRequest(new
        {
            error = "Email subject and body are empty."
        });
    }

    var text = Regex.Replace(
        mail.Body ?? "",
        "<[^>]*>",
        " "
    );

    return Results.Ok(new
    {
        success = true,
        mail.ItemId,
        mail.Subject,
        mail.Sender,
        text = text.Trim(),
        attachments = mail.Attachments,
        attachmentCount = mail.Attachments?.Length ?? 0,
        processedAt = DateTimeOffset.UtcNow
    });
});

app.Run();

public record MailRequest(
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
