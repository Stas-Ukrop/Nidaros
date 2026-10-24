# Nidaros — AI-Hub testomgeving (Postman-compatible)

## Bestanden
- `Program.cs`: POST /api/auth/login, POST /api/auth/verify, POST /api/auth/logout, beveiligd POST /api/mail/analyze.
- `AiHubAuthService.cs`: AI-Hub V1 Login en Token via `application/x-www-form-urlencoded`.
- `wwwroot/taskpane.html`: bestaande Outlook-interface met login- en OTP-panelen.
- Bestaande `MailAnalyzer.cs` ongewijzigd laten.

## Render Environment

`AIHUB_BASE_URL=https://api.ais-test.nidaros-appserver.com`

## Postman contract (uit aangeleverde collectie)

- POST `/v1/Login`: `userName`, `password` (form-url-encoded). Antwoord bevat `TwoFactorAuthenticationId`.
- POST `/v1/Token`: `grant_type=password`, `AuthenticationId`, `AuthenticationCode` (form-url-encoded). Antwoord bevat `access_token`.

TOTP wordt **niet** op de server gegenereerd en seed wordt niet opgeslagen: gebruiker voert 6-cijferige code handmatig in Outlook in.

## Plaatsing

1. Kopieer `Program.cs` en `AiHubAuthService.cs` naar dezelfde C# projectmap als `MailAnalyzer.cs`.
2. Vervang `wwwroot/taskpane.html`.
3. Stel alleen `AIHUB_BASE_URL` in op Render. Verwijder oude `AIHUB_LOGIN_URL` / `AIHUB_TOKEN_URL` indien aanwezig.
4. Deploy en controleer `/health`, inloggen, verificatie, documentanalyse en uitloggen.

Beveiliging: analyse-endpoint valideert een lokale sessie; alleen UI verbergen is onvoldoende. Sessies zitten in het geheugen (verdwijnen bij Render-restart; voor meerdere replicas is gedeelde sessieopslag nodig). Geen wachtwoorden, TOTP-seeds of AI-Hub tokens naar logs of browser storage. Rate limits en HTTPS blijven noodzakelijk in productie.
