var builder = WebApplication.CreateBuilder(args);

builder.Services.AddHttpClient<AiHubAuthService>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(20);
});
builder.Services.AddSingleton<AuthSessionStore>();
builder.Services.AddHttpClient<AiHubPaymentService>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(25);
});

var app = builder.Build();

app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        if (ctx.File.Name.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
            ctx.Context.Response.Headers.CacheControl = "no-store, no-cache, must-revalidate";
    }
});

app.MapGet("/health", () => Results.Ok(new { status = "OK", service = "Nidaros Mail API" }));

app.MapPost("/api/auth/login", async (LoginRequest request, AiHubAuthService auth, AuthSessionStore sessions, HttpContext ctx) =>
{
    if (string.IsNullOrWhiteSpace(request.UserName) || string.IsNullOrWhiteSpace(request.Password))
        return Results.BadRequest(new { error = "Gebruikersnaam en wachtwoord zijn verplicht." });
    if (!sessions.AllowAttempt(ctx.Connection.RemoteIpAddress?.ToString() ?? "unknown"))
        return Results.Json(new { error = "Te veel pogingen. Probeer later opnieuw." }, statusCode: 429);
    var result = await auth.LoginAsync(request.UserName, request.Password);
    if (result.Value is null)
        return Results.Json(new { error = result.Error }, statusCode: result.StatusCode);
    var id = sessions.AddChallenge(result.Value);
    return Results.Ok(new { challengeId = id });
});

app.MapPost("/api/auth/verify", async (VerifyRequest request, AiHubAuthService auth, AuthSessionStore sessions) =>
{
    if (string.IsNullOrWhiteSpace(request.ChallengeId) || request.Code is null ||
        !System.Text.RegularExpressions.Regex.IsMatch(request.Code, @"^\d{6}$"))
        return Results.BadRequest(new { error = "Voer een code van 6 cijfers in." });
    if (!sessions.TryUseChallenge(request.ChallengeId, out var remoteId))
        return Results.Json(new { error = "Verificatie verlopen of te veel pogingen." }, statusCode: 401);
    var result = await auth.VerifyAsync(remoteId!, request.Code);
    if (result.Value is null)
        return Results.Json(new { error = result.Error }, statusCode: result.StatusCode);
    var token = sessions.AddSession(result.Value);
    return Results.Ok(new { accessToken = token, expiresIn = 1800 });
});

// Controleren of de lokale server-sessie nog bestaat (geen nieuw token aanmaken).
app.MapGet("/api/auth/session", (HttpContext ctx, AuthSessionStore sessions) =>
{
    if (!sessions.ValidSession(Bearer(ctx))) return Results.Unauthorized();
    return Results.Ok(new { valid = true });
});

app.MapPost("/api/auth/logout", (HttpContext ctx, AuthSessionStore sessions) =>
{
    sessions.RemoveSession(Bearer(ctx));
    return Results.Ok(new { success = true });
});

app.MapPost("/api/mail/analyze", (HttpContext ctx, MailRequest mail, AuthSessionStore sessions) =>
{
    if (!sessions.ValidSession(Bearer(ctx)))
        return Results.Unauthorized();
    var (response, error) = MailAnalyzer.Analyze(mail);
    return error is null ? Results.Ok(response) : Results.BadRequest(new { error });
});

// Payment requests are only created after explicit user confirmation.
app.MapPost("/api/payments/v3/create", async (HttpContext ctx, CreatePaymentRequest request,
    AuthSessionStore sessions, AiHubPaymentService payments) =>
{
    if (!sessions.TryGetAiHubToken(Bearer(ctx), out var aiHubToken)) return Results.Unauthorized();
    var outcome = await payments.CreateAsync(request, aiHubToken!);
    return Results.Json(outcome.Body, statusCode: outcome.StatusCode);
});

app.Run();

static string? Bearer(HttpContext ctx)
{
    var value = ctx.Request.Headers.Authorization.ToString();
    return value.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) ? value[7..].Trim() : null;
}

public record LoginRequest(string? UserName, string? Password);
public record VerifyRequest(string? ChallengeId, string? Code);
