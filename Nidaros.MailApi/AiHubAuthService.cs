using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text.Json;

// URL методов задаются через переменные Render Environment. Пароли не записываются в лог.
public sealed class AiHubAuthService(HttpClient http, IConfiguration config)
{
    public async Task<AuthReply<string>> LoginAsync(string user, string password)
    {
        var url = Url("/v1/Login");
        if (!IsTrustedUrl(url)) return new(null, "AIHUB_BASE_URL niet geconfigureerd.", 503);
        try
        {
            using var body = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["userName"] = user,
                ["password"] = password
            });
            using var response = await http.PostAsync(url, body);
            if (!response.IsSuccessStatusCode) return new(null, "AI-Hub login mislukt.", 401);
            using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
            var id = FindString(doc.RootElement, "TwoFactorAuthenticationId");
            return string.IsNullOrWhiteSpace(id)
                ? new(null, "AI-Hub gaf geen TwoFactorAuthenticationId terug.", 502)
                : new(id, null, 200);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
        { return new(null, "AI-Hub is niet bereikbaar of antwoordt ongeldig.", 502); }
    }

    public async Task<AuthReply<string>> VerifyAsync(string authId, string code)
    {
        var url = Url("/v1/Token");
        if (!IsTrustedUrl(url)) return new(null, "AIHUB_BASE_URL niet geconfigureerd.", 503);
        try
        {
            // De naam AuthenticationId komt exact overeen met de Postman-collectie.
            using var body = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "password",
                ["AuthenticationId"] = authId,
                ["AuthenticationCode"] = code
            });
            using var response = await http.PostAsync(url, body);
            if (!response.IsSuccessStatusCode) return new(null, "Verificatiecode ongeldig of verlopen.", 401);
            using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
            var token = FindString(doc.RootElement, "access_token") ?? FindString(doc.RootElement, "accessToken") ?? FindString(doc.RootElement, "token");
            return string.IsNullOrWhiteSpace(token)
                ? new(null, "AI-Hub gaf geen token terug.", 502)
                : new(token, null, 200);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
        { return new(null, "AI-Hub verificatie niet beschikbaar.", 502); }
    }

    private string? Url(string path)
    {
        var baseUrl = config["AIHUB_BASE_URL"]?.TrimEnd('/');
        return IsTrustedUrl(baseUrl) ? baseUrl + path : null;
    }

    private static bool IsTrustedUrl(string? url) =>
        Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
        uri.Scheme == Uri.UriSchemeHttps &&
        uri.Host.Equals("api.ais-test.nidaros-appserver.com", StringComparison.OrdinalIgnoreCase);

    private static string? FindString(JsonElement obj, string name)
    {
        if (obj.ValueKind != JsonValueKind.Object) return null;
        foreach (var prop in obj.EnumerateObject())
            if (prop.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                return prop.Value.ValueKind == JsonValueKind.String ? prop.Value.GetString() : prop.Value.ToString();
        return null;
    }
}

public record AuthReply<T>(T? Value, string? Error, int StatusCode);

// Локальные краткоживущие сессии: не передаем AI-Hub token в браузер.
// Для нескольких инстансов Render понадобится Redis/общая БД.
public sealed class AuthSessionStore
{
    private readonly ConcurrentDictionary<string, (string RemoteId, DateTimeOffset Until, int Attempts)> challenges = new();
    private readonly ConcurrentDictionary<string, (string AiHubToken, DateTimeOffset Until)> sessions = new();
    private readonly ConcurrentDictionary<string, (int Count, DateTimeOffset Until)> attempts = new();

    public bool AllowAttempt(string ip)
    {
        var now = DateTimeOffset.UtcNow;
        var val = attempts.AddOrUpdate(ip, _ => (1, now.AddMinutes(10)), (_, old) =>
            old.Until <= now ? (1, now.AddMinutes(10)) : (old.Count + 1, old.Until));
        return val.Count <= 10;
    }
    public string AddChallenge(string remoteId)
    {
        var id = NewToken();
        challenges[id] = (remoteId, DateTimeOffset.UtcNow.AddMinutes(5), 0);
        return id;
    }
    public bool TryUseChallenge(string? id, out string? remoteId)
    {
        remoteId = null;
        if (id is null || !challenges.TryGetValue(id, out var c) || c.Until <= DateTimeOffset.UtcNow)
            return false;
        if (c.Attempts >= 5) { challenges.TryRemove(id, out _); return false; }
        // Разовый challenge: повторные запросы требуют нового Login.
        if (!challenges.TryRemove(id, out _)) return false;
        remoteId = c.RemoteId;
        return true;
    }
    public string AddSession(string aiHubToken)
    {
        var token = NewToken();
        sessions[token] = (aiHubToken, DateTimeOffset.UtcNow.AddMinutes(30));
        return token;
    }
    public bool TryGetAiHubToken(string? token, out string? aiHubToken)
    {
        aiHubToken = null;
        if (token is null || !sessions.TryGetValue(token, out var session)) return false;
        if (session.Until <= DateTimeOffset.UtcNow) { sessions.TryRemove(token, out _); return false; }
        aiHubToken = session.AiHubToken;
        return true;
    }
    public bool ValidSession(string? token) => TryGetAiHubToken(token, out _);
    public void RemoveSession(string? token) { if (token is not null) sessions.TryRemove(token, out _); }
    private static string NewToken() => Convert.ToHexString(RandomNumberGenerator.GetBytes(32));
}
