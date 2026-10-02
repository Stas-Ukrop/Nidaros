# Nidaros Mail Assistant

Outlook-invoegtoepassing die gegevens uit een e-mail en de bijlage haalt via de Nidaros-server.

Na **Analyze mail** opent het paneel. Bij **Gegevens ophalen** stuurt het paneel de e-mail en de gekozen bijlage naar `POST /api/mail/analyze` op dezelfde server en toont het antwoord.

```text
Outlook → manifest.xml → taskpane.html → /api/mail/analyze → antwoord in het paneel
```

## Manifest installeren

1. Open https://aka.ms/olksideload (Outlook op het web of de nieuwe Outlook).
   In de klassieke Outlook: **Start → Invoegtoepassingen ophalen**.
2. Ga naar **Mijn invoegtoepassingen → Aangepaste invoegtoepassingen**.
3. Kies **Een aangepaste invoegtoepassing toevoegen → Toevoegen vanuit bestand** en selecteer `manifest.xml`.
4. Open een e-mail en klik op **Analyze mail** in het lint.

---

# Nidaros Mail Assistant (RU)

Надстройка Outlook, которая извлекает данные из письма и вложения через сервер Nidaros.

По кнопке **Analyze mail** открывается панель. По **Gegevens ophalen** панель отправляет письмо и выбранное вложение на `POST /api/mail/analyze` того же сервера и показывает ответ.

```text
Outlook → manifest.xml → taskpane.html → /api/mail/analyze → ответ в панели
```

## Установка манифеста

1. Откройте https://aka.ms/olksideload (Outlook в браузере или новый Outlook).
   В классическом Outlook: **Главная → Получить надстройки** (Get Add-ins).
2. Перейдите в **Мои надстройки → Пользовательские надстройки** (My add-ins → Custom add-ins).
3. Выберите **Добавить пользовательскую надстройку → Добавить из файла** (Add from file) и укажите `manifest.xml`.
4. Откройте письмо и нажмите **Analyze mail** на ленте.