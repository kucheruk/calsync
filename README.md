# CalSync - Утилита синхронизации календарей

Утилита командной строки для синхронизации событий из .ics файлов с Microsoft Exchange Server 2019 (on-premise) через EWS (Exchange Web Services).

## Описание

CalSync автоматически:
- Загружает .ics файл по указанному URL
- Парсит календарные события из файла
- Подключается к Microsoft Exchange Server через EWS
- Синхронизирует события в заданном календаре:
  - Создает новые события
  - Обновляет существующие события
  - Удаляет события, которые были удалены из источника

## Требования

- .NET 9.0 или выше
- Доступ к Microsoft Exchange Server 2019 (on-premise)
- Учетные данные для подключения к Exchange
- Сетевой доступ к источнику .ics файла

## Установка

```bash
git clone <repository-url>
cd calsync
dotnet build
```

## Использование

```bash
# Базовая синхронизация
dotnet run -- --url "https://example.com/calendar.ics" --exchange-url "https://exchange.company.com/EWS/Exchange.asmx" --username "user@company.com" --calendar "Calendar"

# С дополнительными параметрами
dotnet run -- --url "https://example.com/calendar.ics" \
              --exchange-url "https://exchange.company.com/EWS/Exchange.asmx" \
              --username "user@company.com" \
              --password "password" \
              --calendar "My Calendar" \
              --sync-interval 3600 \
              --log-level Debug
```

## Параметры командной строки

| Параметр | Описание | Обязательный |
|----------|----------|--------------|
| `--url` | URL .ics файла для загрузки | Да |
| `--exchange-url` | URL EWS сервера Exchange | Да |
| `--username` | Имя пользователя для Exchange | Да |
| `--password` | Пароль (если не указан, будет запрошен) | Нет |
| `--calendar` | Имя календаря для синхронизации | Да |
| `--sync-interval` | Интервал синхронизации в секундах | Нет |
| `--log-level` | Уровень логирования (Debug, Info, Warning, Error) | Нет |
| `--dry-run` | Режим тестирования без изменений | Нет |

## Конфигурация

Утилита поддерживает конфигурационный файл `appsettings.json`:

```json
{
  "CalSync": {
    "DefaultSyncInterval": 3600,
    "LogLevel": "Info",
    "Exchange": {
      "Domain": "company.com",
      "Version": "Exchange2016_SP1"
    },
    "IcsSettings": {
      "TimeoutSeconds": 30,
      "RetryAttempts": 3
    }
  }
}
```

## Логирование

Утилита ведет подробные логи всех операций:
- Загрузка .ics файлов
- Подключение к Exchange
- Создание, обновление и удаление событий
- Ошибки и предупреждения

## Безопасность

- Пароли не сохраняются в конфигурационных файлах
- Поддержка Windows Authentication
- SSL/TLS соединения с Exchange Server
- Валидация SSL сертификатов

## Разработка

См. [spec.md](spec.md) для технических требований и плана разработки.

## Лицензия

MIT License

## Поддержка

Для вопросов и багрепортов используйте Issues в данном репозитории. 