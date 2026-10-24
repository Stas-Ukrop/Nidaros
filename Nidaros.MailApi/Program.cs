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
