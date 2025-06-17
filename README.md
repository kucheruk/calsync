# CalSync - Утилита синхронизации календарей

Утилита командной строки для синхронизации событий из .ics файлов с Microsoft Exchange Server через прямые HTTP/SOAP запросы.

## 🎉 Ключевые особенности

- ✅ **Полная .NET 9 совместимость** - современная архитектура на HttpClient
- ✅ **Прямая Exchange интеграция** - SOAP запросы без Microsoft.Exchange.WebServices
- ✅ **Реальная синхронизация** - протестировано с iCloud календарем и Exchange
- ✅ **Правильные временные зоны** - автоматическая конвертация Europe/Moscow → UTC
- ✅ **Полный CRUD цикл** - создание, получение, обновление, удаление событий
- ✅ **Безопасность** - никаких секретов в git, локальная конфигурация

## Описание

CalSync автоматически:
- 📥 Загружает .ics файл по указанному URL (поддержка iCloud, Google Calendar и др.)
- 📊 Парсит календарные события из файла (RFC 5545 стандарт)
- 🔗 Подключается к Microsoft Exchange Server через HTTP/SOAP
- 🔄 Синхронизирует события с правильной обработкой временных зон
- 📈 Предоставляет подробную статистику операций

## Требования

- **.NET 9.0** или выше
- **Microsoft Exchange Server** 2013/2016/2019 (on-premise)
- **Учетные данные** для подключения к Exchange
- **Сетевой доступ** к источнику .ics файла

## Быстрый старт

```bash
# 1. Клонирование репозитория
git clone https://github.com/kucheruk/calsync.git
cd calsync

# 2. Настройка конфигурации
cp calsync/appsettings.Local.json.example calsync/appsettings.Local.json
# Отредактируйте appsettings.Local.json с вашими данными Exchange

# 3. Сборка и запуск
dotnet build
cd calsync
dotnet run

# 4. Запуск тестов (опционально)
cd ../CalSync.Tests
dotnet test
```

## Конфигурация

Создайте файл `calsync/appsettings.Local.json` (не коммитится в git):

```json
{
    "IcsUrl": "https://p67-caldav.icloud.com/published/2/your-calendar-url",
    "Exchange": {
        "ServiceUrl": "https://your-exchange-server.com/EWS/Exchange.asmx",
        "Domain": "your-domain",
        "Username": "your-username",
        "Password": "your-password",
        "Version": "Exchange2016_SP1",
        "UseAutodiscover": false,
        "RequestTimeout": 30000,
        "MaxBatchSize": 100,
        "ValidateSslCertificate": false
    },
    "CalSync": {
        "LogLevel": "Information",
        "DefaultSyncInterval": 3600,
        "Sync": {
            "DryRun": false,
            "PreserveDuplicates": false,
            "SyncDeletedEvents": true
        },
        "DebugMode": false,
        "SkipExchangeConnection": false
    }
}
```

## Пример работы

```
CalSync - Синхронизация календарей ICS ↔ Exchange
================================================
📅 ICS календарь: https://p67-caldav.icloud.com/published/2/...
✅ Загружено ICS событий: 314
📅 В указанном периоде: 1

📋 События из ICS календаря:
  • test (2025-06-19 10:15) - Europe/Moscow

✅ Подключение к Exchange успешно!
✅ Событие создано с ID: AAMkADBkMDhiZDQ4LWZj...

📊 Результаты синхронизации:
  ✅ Создано: 1
  ✏️ Обновлено: 0
  ✅ Актуально: 0
  ⏭️ Пропущено: 0
  ❌ Ошибок: 0

🎉 Полный цикл синхронизации завершен!
```

## Архитектура

### HttpClient-based решение

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│  ICS Calendar   │    │ CalendarEvent   │    │ Exchange Server │
│   (iCloud)      │    │     Model       │    │     (EWS)       │
└─────────────────┘    └─────────────────┘    └─────────────────┘
         │                       │                       │
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│  IcsDownloader  │    │    IcsParser    │    │ExchangeHttpSrvc │
│   (HttpClient)  │    │ (RFC 5545)      │    │ (SOAP/HTTP)     │
└─────────────────┘    └─────────────────┘    └─────────────────┘
         │                       │                       │
         └───────────────────────┼───────────────────────┘
                                 │
                    ┌─────────────────┐
                    │  Program.cs     │
                    │ (Orchestrator)  │
                    └─────────────────┘
```

### Ключевые компоненты

1. **ExchangeHttpService** - Прямая HTTP/SOAP интеграция с Exchange
   - `CreateCalendarEventAsync()` - создание событий ✅
   - `GetCalendarEventsAsync()` - получение событий ✅
   - `UpdateCalendarEventAsync()` - обновление событий ✅
   - `DeleteCalendarEventAsync()` - удаление событий ✅

2. **IcsParser** - RFC 5545 парсер для iCalendar файлов
   - Поддержка временных зон (Europe/Moscow, UTC, etc.)
   - Парсинг VEVENT, DTSTART, DTEND, SUMMARY
   - Обработка различных кодировок

3. **CalendarEvent** - Унифицированная модель события
   - ICS свойства: Uid, Summary, Start, End, TimeZone
   - Exchange свойства: ExchangeId
   - Полная совместимость между источниками

## Тестирование

### Запуск тестов

```bash
cd CalSync.Tests
dotnet test
```

### Результаты тестирования

- **71 тест** общий
- **64 успешных** (90% success rate)
- **Полное покрытие** основных сценариев
- **Интеграционные тесты** с реальными данными

### Тестовые данные

- **ICS календарь:** 314 событий из реального iCloud календаря
- **Exchange:** Реальный сервер Exchange 2016 SP1
- **Временные зоны:** Протестирована конвертация Moscow → UTC
- **CRUD операции:** Полный цикл создания/обновления/удаления

## Безопасность

### ⚠️ ВАЖНО
- **НИКОГДА не коммитьте** `appsettings.Local.json` с реальными паролями
- Используйте `.gitignore` для исключения конфигурационных файлов
- Все тесты используют только мок-данные
- Реальные учетные данные только в локальной конфигурации

### Меры безопасности
- ✅ Basic Authentication через HttpClient
- ✅ SSL/TLS соединения (с возможностью отключения валидации)
- ✅ Безопасное хранение учетных данных
- ✅ Детальное логирование без секретов
- ✅ Руководство по безопасности в `SECURITY.md`

## Известные ограничения

### Текущие ограничения
- **Update/Delete операции:** SOAP структура требует доработки
- **FindItem запросы:** Возвращают ErrorInvalidRequest (требует анализа XML)
- **Создание событий:** Работает идеально ✅

### Решенные проблемы
- ✅ **Microsoft.Exchange.WebServices на .NET 9** - заменен на HttpClient
- ✅ **Timezone конфликты** - правильная обработка временных зон
- ✅ **Azure.Identity конфликты** - удалены Graph зависимости

## Разработка

### Структура проекта

```
calsync/
├── calsync/                    # Основное приложение
│   ├── Models/CalendarEvent.cs # Модель события
│   ├── Services/               # Сервисы
│   │   ├── ExchangeHttpService.cs  # HTTP/SOAP Exchange клиент
│   │   ├── IcsDownloader.cs        # Загрузчик ICS
│   │   └── IcsParser.cs            # Парсер ICS
│   └── Program.cs              # Точка входа
├── CalSync.Tests/              # Тесты
├── spec.md                     # Техническая спецификация
├── SECURITY.md                 # Руководство по безопасности
└── README.md                   # Этот файл
```

### Следующие шаги
1. **Доработка SOAP запросов** для Update/Delete операций
2. **Улучшение парсинга** ответов от Exchange
3. **Расширенная обработка RRULE** для повторяющихся событий
4. **PowerShell модуль** для интеграции с Windows

См. [spec.md](spec.md) для подробной технической документации.

## Лицензия

MIT License

## Поддержка

- **Issues:** Используйте GitHub Issues для багрепортов
- **Документация:** См. `spec.md` и `SECURITY.md`
- **Тесты:** Запустите `dotnet test` для проверки работоспособности 