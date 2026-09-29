
# Сборка ASP.NET Core приложения
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build

WORKDIR /src

COPY . .

RUN dotnet publish "Nidaros.MailApi.csproj" \
    -c Release \
    -o /app/publish \
    /p:UseAppHost=false

# Запуск приложения
FROM mcr.microsoft.com/dotnet/aspnet:10.0

WORKDIR /app

COPY --from=build /app/publish .

ENV ASPNETCORE_URLS=http://0.0.0.0:10000

EXPOSE 10000

ENTRYPOINT ["dotnet", "Nidaros.MailApi.dll"]
