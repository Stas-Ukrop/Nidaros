
# Nidaros Mail Assistant

Een Outlook-invoegtoepassing die gegevens van het geopende e-mailbericht naar een lokale ASP.NET Core-server stuurt en het verwerkingsresultaat in het zijpaneel van Outlook (Task Pane) toont.

## Huidige functionaliteit

- Het zijpaneel openen vanuit Outlook.
- Het huidige bericht uitlezen via Office.js: bericht-ID, onderwerp, e-mailadres van de afzender en berichttekst.
- Metagegevens van bijlagen doorgeven (ID, naam, grootte en MIME-type), voor zover beschikbaar via Office.js.
- Berichtgegevens naar de C# API sturen en het JSON-antwoord in het Task Pane tonen.
- De serverstatus controleren via `GET /health`.

**Let op:** deze versie downloadt de inhoud van bijlagen niet, leest de mailbox niet rechtstreeks via Microsoft Graph of Exchange en analyseert berichten niet automatisch zonder actie van de gebruiker.

## Architectuur

```text
Outlook → Office.js / Task Pane → HTTPS → ASP.NET Core API → JSON → Task Pane
```

Tijdens lokale ontwikkeling worden de interface en de API door dezelfde applicatie aangeboden via `https://localhost:3000`. Voor verzoeken vanuit het Task Pane naar deze API is daarom geen aparte CORS-configuratie nodig.

## Technologieën

- C# en ASP.NET Core Minimal API
- HTML, CSS en JavaScript
- Office.js (Outlook-invoegtoepassing)
- HTTPS voor lokale ontwikkeling

## Projectstructuur

```text
Nidaros.MailApi/
├── Program.cs
├── Properties/
│   └── launchSettings.json
└── wwwroot/
    ├── taskpane.html
    └── commands.html
```

Het manifest van de Outlook-invoegtoepassing wordt apart beheerd en verwijst naar `https://localhost:3000/taskpane.html` en `https://localhost:3000/commands.html`.

## Lokaal starten

Benodigd: de .NET SDK, Outlook met ondersteuning voor invoegtoepassingen en een vertrouwd lokaal HTTPS-certificaat.

1. Vertrouw het ontwikkelcertificaat vanuit de projectmap:

   ```powershell
   dotnet dev-certs https --trust
   ```

2. Controleer of het HTTPS-profiel in `Properties/launchSettings.json` de instelling `"applicationUrl": "https://localhost:3000"` bevat.

3. Start de server:

   ```powershell
   dotnet run --launch-profile https
   ```

4. Controleer `https://localhost:3000/health` en `https://localhost:3000/taskpane.html` in de browser.

5. Installeer het manifest in Outlook, open een e-mailbericht, start **Nidaros Mail Assistant** en klik op **Analyze mail**.

In een gewone browser heeft het Task Pane geen toegang tot e-mailberichten: Office.js krijgt de context van het huidige bericht alleen binnen Outlook.

## API

### `GET /health`

Geeft de status van de server terug.

### `POST /api/mail/analyze`

Ontvangt JSON met de velden `itemId`, `subject`, `sender`, `body` en `attachments`. Geeft de verwerkingsstatus, berichtgegevens, tekst, metagegevens en het aantal bijlagen en het verwerkingstijdstip terug.

Voorbeeldverzoek:

```json
{
  "itemId": "test-001",
  "subject": "Test Nidaros",
  "sender": "test@example.com",
  "body": "Hello! This is a test email.",
  "attachments": []
}
```

De huidige implementatie wijst een verzoek af als zowel het onderwerp als de berichttekst leeg zijn. Het Task Pane leest de berichttekst met `Office.CoercionType.Text`.

## Getest resultaat

De volledige keten is lokaal getest: Outlook opende het Task Pane, Office.js las het geselecteerde bericht, de C# API gaf een succesvol JSON-antwoord terug en het paneel toonde het resultaat. Bij een testbericht zonder berichttekst gaf de server een leeg veld `text` terug.

## Volgende stappen

- De weergave verbeteren voor de lichte en donkere Outlook-thema's.
- Ontvangers, verzenddatum en aanvullende berichtvelden toevoegen.
- De inhoud van bijlagen ophalen, rekening houdend met de beperkingen van de Outlook-client en de Mailbox API-versie.
- Authenticatie en API-beveiliging toevoegen voordat de server buiten de lokale computer beschikbaar wordt gemaakt.
- Berichttekst en bestanden inhoudelijk verwerken op de server.

## Beveiliging

Dit project is momenteel bedoeld voor lokale ontwikkeling. Publiceer de API niet op internet zonder authenticatie, autorisatie, limieten voor verzoekgroottes en passende bescherming van e-mailgegevens. Plaats geen echte e-mails, tokens of geheimen in de Git-repository.



# Nidaros Mail Assistant

Надстройка для Outlook, которая передаёт данные открытого электронного письма на локальный сервер ASP.NET Core и отображает результат обработки в боковой панели Outlook (Task Pane).

## Текущие возможности

- Открытие боковой панели непосредственно в Outlook.
- Чтение текущего письма через Office.js: идентификатор письма, тема, электронный адрес отправителя и текст сообщения.
- Передача метаданных вложений (ID, имя, размер и MIME-тип), если они доступны через Office.js.
- Отправка данных письма на C# API и отображение JSON-ответа в Task Pane.
- Проверка состояния сервера через `GET /health`.

**Важно:** текущая версия не загружает содержимое вложений, не читает почтовый ящик напрямую через Microsoft Graph или Exchange и не анализирует письма автоматически без действия пользователя.

## Архитектура

    Outlook → Office.js / Task Pane → HTTPS → ASP.NET Core API → JSON → Task Pane

При локальной разработке интерфейс и API обслуживаются одним приложением по адресу `https://localhost:3000`. Поэтому для запросов из Task Pane к этому API отдельная настройка CORS не требуется.

## Используемые технологии

- C# и ASP.NET Core Minimal API
- HTML, CSS и JavaScript
- Office.js (надстройка Outlook)
- HTTPS для локальной разработки

## Структура проекта

    Nidaros.MailApi/
    ├── Program.cs
    ├── Properties/
    │   └── launchSettings.json
    └── wwwroot/
        ├── taskpane.html
        └── commands.html

Манифест надстройки Outlook хранится отдельно и ссылается на `https://localhost:3000/taskpane.html` и `https://localhost:3000/commands.html`.

## Локальный запуск

Необходимы .NET SDK, Outlook с поддержкой надстроек и доверенный локальный HTTPS-сертификат.

1. Установите доверие к сертификату разработки из папки проекта:

       dotnet dev-certs https --trust

2. Убедитесь, что HTTPS-профиль в `Properties/launchSettings.json` содержит настройку `"applicationUrl": "https://localhost:3000"`.

3. Запустите сервер:

       dotnet run --launch-profile https

4. Проверьте в браузере адреса `https://localhost:3000/health` и `https://localhost:3000/taskpane.html`.

5. Установите манифест в Outlook, откройте электронное письмо, запустите **Nidaros Mail Assistant** и нажмите **Analyze mail**.

В обычном браузере Task Pane не имеет доступа к электронным письмам: Office.js получает контекст текущего письма только внутри Outlook.

## API

### `GET /health`

Возвращает состояние сервера.

### `POST /api/mail/analyze`

Принимает JSON с полями `itemId`, `subject`, `sender`, `body` и `attachments`.

Возвращает статус обработки, данные письма, текст, метаданные и количество вложений, а также время обработки.

Пример запроса:

    {
      "itemId": "test-001",
      "subject": "Test Nidaros",
      "sender": "test@example.com",
      "body": "Hello! This is a test email.",
      "attachments": []
    }

Текущая реализация отклоняет запрос, если одновременно отсутствуют тема и текст письма. Task Pane получает текст сообщения с помощью `Office.CoercionType.Text`.

## Результаты тестирования

Вся цепочка была проверена локально: Outlook открыл Task Pane, Office.js прочитал выбранное письмо, C# API вернул успешный JSON-ответ, а панель отобразила результат.

При тестировании письма без текста сервер вернул пустое поле `text`.

## Следующие этапы разработки

- Улучшить отображение интерфейса для светлой и тёмной тем Outlook.
- Добавить получателей, дату отправки и дополнительные поля письма.
- Реализовать получение содержимого вложений с учётом ограничений клиента Outlook и версии Mailbox API.
- Добавить аутентификацию и защиту API перед публикацией сервера за пределами локального компьютера.
- Реализовать содержательную обработку текста писем и файлов на сервере.

## Безопасность

В настоящее время проект предназначен для локальной разработки.

Не публикуйте API в интернете без аутентификации, авторизации, ограничений размера запросов и надлежащей защиты данных электронной почты.

Не размещайте настоящие письма, токены или секретные ключи в Git-репозитории.
