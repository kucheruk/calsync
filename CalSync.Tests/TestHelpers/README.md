# Test Helpers - Вспомогательные классы для тестирования CalSync

Этот каталог содержит переиспользуемые компоненты для тестирования различных сценариев работы с .ics файлами и HTTP серверами.

## IcsTestDataGenerator

Статический класс для генерации тестовых ICS данных в различных форматах.

### Основные методы:

#### `GenerateSimpleCalendar()`
```csharp
var calendar = IcsTestDataGenerator.GenerateSimpleCalendar(
    "Встреча с командой",
    "Обсуждение планов на спринт", 
    "Конференц-зал А",
    DateTime.UtcNow.AddDays(1),
    DateTime.UtcNow.AddDays(1).AddHours(1)
);
```

#### `GenerateCalendarWithMultipleEvents(int eventCount)`
```csharp
// Создает календарь с 5 событиями
var calendar = IcsTestDataGenerator.GenerateCalendarWithMultipleEvents(5);
```

#### `GenerateRecurringEventCalendar()`
```csharp
// Создает календарь с повторяющимся событием (еженедельно)
var calendar = IcsTestDataGenerator.GenerateRecurringEventCalendar();
```

#### `GenerateCalendarWithTimezone(string timezoneName)`
```csharp
// Создает календарь с событием в указанной временной зоне
var calendar = IcsTestDataGenerator.GenerateCalendarWithTimezone("Europe/Moscow");
```

#### `GenerateInvalidCalendar()`
```csharp
// Создает невалидный ICS файл для тестирования обработки ошибок
var invalidData = IcsTestDataGenerator.GenerateInvalidCalendar();
```

## TestHttpServer

Параметризуемый HTTP сервер для тестирования различных сценариев загрузки .ics файлов.

### Быстрое создание серверов:

#### Простой сервер
```csharp
using var server = TestHttpServer.CreateSimple(8080);
await server.StartAsync();

var downloader = new IcsDownloader();
var content = await downloader.DownloadAsync(server.BaseUrl);
```

#### Сервер с авторизацией
```csharp
using var server = TestHttpServer.CreateWithAuth("username", "password", 8081);
await server.StartAsync();

using var httpClient = new HttpClient();
var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes("username:password"));
httpClient.DefaultRequestHeaders.Authorization = 
    new AuthenticationHeaderValue("Basic", credentials);

var downloader = new IcsDownloader(httpClient);
var content = await downloader.DownloadAsync(server.BaseUrl + "protected");
```

#### Сервер с различными HTTP статусами
```csharp
var statusCodes = new Dictionary<string, HttpStatusCode>
{
    { "/ok", HttpStatusCode.OK },
    { "/notfound", HttpStatusCode.NotFound },
    { "/error", HttpStatusCode.InternalServerError }
};

using var server = TestHttpServer.CreateWithStatusCodes(statusCodes, 8082);
await server.StartAsync();
```

### Продвинутая конфигурация:

```csharp
var config = new TestHttpServerConfig
{
    Port = 8083,
    ResponseDelay = TimeSpan.FromMilliseconds(500), // Искусственная задержка
    Routes = new Dictionary<string, string>
    {
        { "/simple", IcsTestDataGenerator.GenerateSimpleCalendar() },
        { "/multiple", IcsTestDataGenerator.GenerateCalendarWithMultipleEvents(3) },
        { "/recurring", IcsTestDataGenerator.GenerateRecurringEventCalendar() }
    },
    Headers = new Dictionary<string, string>
    {
        { "X-Custom-Header", "Test-Value" }
    }
};

using var server = new TestHttpServer(config);
await server.StartAsync();

// Динамическое добавление маршрутов
server.AddRoute("/new-route", IcsTestDataGenerator.GenerateSimpleCalendar("New Event"));
```

## Использование в тестах

### Пример базового теста:
```csharp
[Fact]
public async Task DownloadIcsFile_ShouldReturnValidContent()
{
    // Arrange
    using var server = TestHttpServer.CreateSimple(8765);
    server.AddRoute("/calendar.ics", IcsTestDataGenerator.GenerateSimpleCalendar("Test Event"));
    await server.StartAsync();
    
    var downloader = new IcsDownloader();

    // Act
    var content = await downloader.DownloadAsync(server.BaseUrl + "calendar.ics");

    // Assert
    Assert.Contains("BEGIN:VCALENDAR", content);
    Assert.Contains("Test Event", content);
}
```

### Пример теста с авторизацией:
```csharp
[Fact]
public async Task DownloadIcsFile_WithAuth_ShouldAuthenticate()
{
    // Arrange
    using var server = TestHttpServer.CreateWithAuth("user", "pass", 8766);
    await server.StartAsync();
    
    using var httpClient = new HttpClient();
    var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes("user:pass"));
    httpClient.DefaultRequestHeaders.Authorization = 
        new AuthenticationHeaderValue("Basic", credentials);
    
    var downloader = new IcsDownloader(httpClient);

    // Act
    var content = await downloader.DownloadAsync(server.BaseUrl + "protected");

    // Assert
    Assert.Contains("BEGIN:VCALENDAR", content);
}
```

### Пример теста с множественными событиями:
```csharp
[Fact]
public async Task DownloadIcsFile_MultipleEvents_ShouldParseAllEvents()
{
    // Arrange
    using var server = TestHttpServer.CreateSimple(8767);
    server.AddRoute("/multi.ics", IcsTestDataGenerator.GenerateCalendarWithMultipleEvents(5));
    await server.StartAsync();
    
    var downloader = new IcsDownloader();

    // Act
    var content = await downloader.DownloadAsync(server.BaseUrl + "multi.ics");

    // Assert
    var eventCount = content.Split("BEGIN:VEVENT").Length - 1;
    Assert.Equal(5, eventCount);
}
```

## Преимущества использования

1. **Переиспользуемость** - компоненты можно использовать в разных тестах
2. **Параметризация** - легкая настройка различных сценариев
3. **Изоляция** - каждый тест использует свой экземпляр сервера
4. **Надежность** - автоматическое управление ресурсами через IDisposable
5. **Гибкость** - поддержка различных HTTP статусов, авторизации, задержек

## Примеры из TestHttpServerExamples.cs

В файле `TestHttpServerExamples.cs` содержится полный набор примеров использования всех возможностей вспомогательных классов, включая:

- Простые HTTP серверы
- Серверы с авторизацией  
- Серверы с различными HTTP статусами
- Серверы с искусственными задержками
- Генерация различных типов ICS файлов

Эти примеры можно использовать как справочник при написании собственных тестов. 