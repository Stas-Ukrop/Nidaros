using System.Globalization;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.RegularExpressions;

public sealed record CreatePaymentRequest(string? Amount, string? Description, string? ReturnUrl, int? ValidityMinutes);
public sealed record PaymentReply(object Body, int StatusCode);

public sealed class AiHubPaymentService(HttpClient http, IConfiguration config)
{
    public async Task<PaymentReply> CreateAsync(CreatePaymentRequest request, string accessToken)
    {
        if (!decimal.TryParse(request.Amount, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var amount)
            || amount <= 0 || amount > 9999999.99m || decimal.Round(amount, 2) != amount)
            return new(new { error = "Gebruik een geldig bedrag met maximaal twee decimalen (punt)." }, 400);
        if (string.IsNullOrWhiteSpace(request.Description) || request.Description.Length > 250)
            return new(new { error = "Omschrijving moet 1 tot 250 tekens bevatten." }, 400);
        if (request.ValidityMinutes is < 1 or > 525600)
            return new(new { error = "Ongeldige geldigheidsduur in minuten." }, 400);
        if (!string.IsNullOrWhiteSpace(request.ReturnUrl) &&
            (!Uri.TryCreate(request.ReturnUrl, UriKind.Absolute, out var returnUri) || returnUri.Scheme != "https"))
            return new(new { error = "ReturnUrl moet een geldige HTTPS-link zijn." }, 400);

        var baseUrl = config["AIHUB_BASE_URL"]?.TrimEnd('/');
        if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri) ||
            uri.Scheme != Uri.UriSchemeHttps ||
            uri.Host != "api.ais-test.nidaros-appserver.com")
            return new(new { error = "AIHUB_BASE_URL niet geconfigureerd." }, 503);

        var values = new Dictionary<string, string>
        {
            ["amount"] = amount.ToString("0.00", CultureInfo.InvariantCulture),
            ["description"] = request.Description.Trim()
        };
        if (!string.IsNullOrWhiteSpace(request.ReturnUrl)) values["returnUrl"] = request.ReturnUrl.Trim();
        if (request.ValidityMinutes.HasValue) values["validityMinutes"] = request.ValidityMinutes.Value.ToString(CultureInfo.InvariantCulture);

        using var message = new HttpRequestMessage(HttpMethod.Post, baseUrl + "/v3/payment-request")
        {
            Content = new FormUrlEncodedContent(values)
        };
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        try
        {
            using var response = await http.SendAsync(message);
            var payload = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                // Avoid exposing upstream response or internal data to clients.
                return new(new { error = response.StatusCode switch
                {
                    System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden => "AI-Hub weigert toegang. Log opnieuw in of controleer rechten.",
                    System.Net.HttpStatusCode.BadRequest => "AI-Hub heeft de betaalgegevens afgekeurd. Controleer bedrag, geldigheid en returnUrl.",
                    _ => "AI-Hub kon het betaalverzoek niet verwerken. Controleer de status in AI-Hub voordat u opnieuw probeert."
                } }, (int)response.StatusCode);
            }
            using var document = JsonDocument.Parse(payload);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return new(new { error = "Onverwacht antwoord van AI-Hub. Controleer de status voordat u opnieuw probeert." }, 502);
            string? Get(string key) => root.EnumerateObject().FirstOrDefault(p => p.Name.Equals(key, StringComparison.OrdinalIgnoreCase)).Value.ToString();
            var id = Get("PaymentRequestId");
            var url = Get("PaymentRequestUrl") ?? Get("AuthorizationUrl");
            var until = Get("ValidUntilUtc");
            if (string.IsNullOrWhiteSpace(id) && string.IsNullOrWhiteSpace(url))
                return new(new { error = "Betaalverzoek mogelijk aangemaakt, maar antwoord bevat geen id of link. Controleer eerst AI-Hub." }, 502);
            return new(new { paymentRequestId = id, paymentRequestUrl = url, validUntilUtc = until }, 200);
        }
        catch (TaskCanceledException)
        {
            return new(new { error = "Time-out. Het verzoek kan zijn aangemaakt. Controleer AI-Hub vóór opnieuw proberen." }, 504);
        }
        catch (HttpRequestException)
        {
            return new(new { error = "Netwerkfout. Controleer in AI-Hub of een verzoek is aangemaakt." }, 502);
        }
        catch (JsonException)
        {
            return new(new { error = "Antwoord van AI-Hub is ongeldig. Controleer eerst of er een verzoek is aangemaakt." }, 502);
        }
    }
}
