# CalSync - Техническая спецификация

## Обзор

CalSync - утилита командной строки для синхронизации календарных событий между .ics файлами (загружаемыми по HTTP/HTTPS) и Microsoft Exchange Server 2019 через EWS API.

## Функциональные требования

### FR-001: Загрузка .ics файлов
- Загрузка .ics файлов по HTTP/HTTPS URL
- Поддержка базовой и digest аутентификации
- Обработка редиректов
- Timeout и retry механизмы
- Валидация SSL сертификатов

### FR-002: Парсинг .ics файлов
- Парсинг стандарта RFC 5545 (iCalendar)
- Поддержка событий (VEVENT)
- Обработка повторяющихся событий (RRULE)
- Поддержка временных зон
- Обработка различных кодировок

### FR-003: Подключение к Exchange Server
- Подключение через EWS API
- Поддержка различных версий Exchange (2013, 2016, 2019)
- Windows Authentication и Basic Authentication
- Autodiscover для автоматического определения EWS URL
- SSL/TLS соединения

### FR-004: Синхронизация событий
- Создание новых событий в Exchange
- Обновление существующих событий
- Удаление событий, отсутствующих в .ics файле
- Сопоставление событий по UID
- Обработка конфликтов временных зон
- Batch операции для повышения производительности

### FR-005: Управление календарями
- Работа с указанным календарем пользователя
- Создание календаря, если он не существует
- Поддержка публичных и приватных календарей

## Нефункциональные требования

### NFR-001: Производительность
- Обработка до 1000 событий за один запуск
- Время синхронизации не более 30 секунд для 100 событий
- Использование памяти не более 100 МБ

### NFR-002: Надежность
- Graceful handling ошибок сети
- Retry механизмы для временных сбоев
- Rollback при критических ошибках
- Детальное логирование всех операций

### NFR-003: Безопасность
- Безопасное хранение учетных данных
- Валидация входных данных
- Защита от injection атак
- Audit trail всех изменений

### NFR-004: Совместимость
- .NET 9.0+
- Windows 10/11, Linux, macOS
- Exchange Server 2013/2016/2019
- PowerShell совместимость

## Архитектура

### Компоненты системы

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   CLI Interface │    │  Configuration  │    │    Logging      │
└─────────────────┘    └─────────────────┘    └─────────────────┘
         │                       │                       │
         └───────────────────────┼───────────────────────┘
                                 │
                    ┌─────────────────┐
                    │  CalSync Core   │
                    └─────────────────┘
                             │
        ┌────────────────────┼────────────────────┐
        │                    │                    │
┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
│  ICS Downloader │  │   ICS Parser    │  │ Exchange Client │
└─────────────────┘  └─────────────────┘  └─────────────────┘
        │                    │                    │
┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
│   HTTP Client   │  │ Calendar Model  │  │   EWS Client    │
└─────────────────┘  └─────────────────┘  └─────────────────┘
```

### Модули

#### 1. CalSync.Core
- `CalSyncService` - основной сервис синхронизации
- `SyncEngine` - движок синхронизации
- `EventMatcher` - сопоставление событий

#### 2. CalSync.Ics
- `IcsDownloader` - загрузка .ics файлов
- `IcsParser` - парсинг iCalendar
- `CalendarEvent` - модель события

#### 3. CalSync.Exchange
- `ExchangeClient` - клиент для работы с EWS
- `ExchangeCalendarService` - операции с календарем
- `ExchangeEventMapper` - маппинг между моделями

#### 4. CalSync.CLI
- `Program` - точка входа
- `CommandLineOptions` - парсинг аргументов
- `ConfigurationManager` - управление конфигурацией

## План разработки

### Этап 1: Базовая инфраструктура (1-2 недели)
- [ ] Настройка проекта и зависимостей
- [ ] Создание базовых моделей данных
- [ ] Настройка логирования и конфигурации
- [ ] CLI интерфейс с базовыми командами
- [ ] Unit тесты для базовых компонентов

**Зависимости:**
- Microsoft.Extensions.Hosting
- Microsoft.Extensions.Configuration
- Microsoft.Extensions.Logging
- System.CommandLine
- Serilog

### Этап 2: ICS модуль (1-2 недели)
- [ ] HTTP клиент для загрузки .ics файлов
- [ ] Парсер iCalendar (RFC 5545)
- [ ] Модели календарных событий
- [ ] Обработка временных зон
- [ ] Unit и интеграционные тесты

**Зависимости:**
- Ical.Net
- System.Net.Http

### Этап 3: Exchange модуль (2-3 недели)
- [ ] EWS клиент
- [ ] Аутентификация (Windows Auth, Basic Auth)
- [ ] CRUD операции для событий
- [ ] Управление календарями
- [ ] Batch операции
- [ ] Тесты с mock Exchange сервером

**Зависимости:**
- Microsoft.Exchange.WebServices

### Этап 4: Синхронизация (1-2 недели)
- [ ] Алгоритм синхронизации
- [ ] Сопоставление событий по UID
- [ ] Обработка конфликтов
- [ ] Dry-run режим
- [ ] Интеграционные тесты

### Этап 5: Расширенные функции (1-2 недели)
- [ ] Повторяющиеся события (RRULE)
- [ ] Обработка исключений в повторяющихся событиях
- [ ] Поддержка вложений
- [ ] Напоминания и уведомления
- [ ] Performance тесты

### Этап 6: Финализация (1 неделя)
- [ ] Документация
- [ ] Packaging и deployment
- [ ] End-to-end тесты
- [ ] Performance оптимизация
- [ ] Security audit

## Модель данных

### CalendarEvent
```csharp
public class CalendarEvent
{
    public string Uid { get; set; }
    public string Summary { get; set; }
    public string Description { get; set; }
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    public string Location { get; set; }
    public string Organizer { get; set; }
    public List<string> Attendees { get; set; }
    public RecurrencePattern Recurrence { get; set; }
    public DateTime LastModified { get; set; }
    public EventStatus Status { get; set; }
}
```

### SyncResult
```csharp
public class SyncResult
{
    public int EventsCreated { get; set; }
    public int EventsUpdated { get; set; }
    public int EventsDeleted { get; set; }
    public List<SyncError> Errors { get; set; }
    public TimeSpan Duration { get; set; }
}
```

## Обработка ошибок

### Категории ошибок
1. **Сетевые ошибки** - timeout, connection refused, DNS errors
2. **Аутентификация** - неверные учетные данные, истекшие токены
3. **Парсинг** - некорректный .ics формат, unsupported features
4. **Exchange ошибки** - EWS errors, календарь не найден, insufficient permissions
5. **Синхронизация** - конфликты данных, duplicate events

### Стратегии обработки
- Retry с exponential backoff для временных ошибок
- Graceful degradation для некритичных ошибок
- Подробное логирование всех ошибок
- User-friendly сообщения об ошибках

## Конфигурация

### appsettings.json
```json
{
  "CalSync": {
    "LogLevel": "Information",
    "DefaultSyncInterval": 3600,
    "Exchange": {
      "Version": "Exchange2016_SP1",
      "Domain": "",
      "UseAutodiscover": true,
      "RequestTimeout": 30000,
      "MaxBatchSize": 100
    },
    "Ics": {
      "RequestTimeout": 30000,
      "MaxRetryAttempts": 3,
      "RetryDelay": 1000,
      "ValidateSslCertificate": true
    },
    "Sync": {
      "DryRun": false,
      "PreserveDuplicates": false,
      "SyncDeletedEvents": true
    }
  }
}
```

## Тестирование

### Unit тесты
- Все публичные методы покрыты тестами
- Mock объекты для внешних зависимостей
- Code coverage > 80%

### Интеграционные тесты
- Тесты с реальными .ics файлами
- Тесты с mock Exchange сервером
- End-to-end сценарии

### Performance тесты
- Benchmark для различных объемов данных
- Memory usage тесты
- Stress тесты для concurrent операций

## Deployment

### Packaging
- Self-contained executable для каждой платформы
- NuGet package для использования как библиотеки
- Docker image для контейнеризации

### CI/CD
- GitHub Actions для автоматической сборки
- Automated тестирование на multiple platforms
- Автоматический release при тегировании 