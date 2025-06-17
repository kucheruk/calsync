using System.Net;
using CalSync.Services;
using Xunit;

namespace CalSync.Tests.TestHelpers;

/// <summary>
/// Примеры использования TestHttpServer и IcsTestDataGenerator
/// </summary>
public class TestHttpServerExamples : IDisposable
{
    private readonly List<TestHttpServer> _servers = new();

    [Fact]
    public async Task TestHttpServer_CreateSimple_ShouldServeIcsContent()
    {
        // Arrange
        using var server = TestHttpServer.CreateSimple(8771);
        _servers.Add(server);

        await server.StartAsync();
        var downloader = new IcsDownloader();

        // Act
        var content = await downloader.DownloadAsync(server.BaseUrl);

        // Assert
        Assert.NotNull(content);
        Assert.Contains("BEGIN:VCALENDAR", content);
        Assert.Contains("Test Event", content);
    }

    [Fact]
    public async Task TestHttpServer_WithCustomRoutes_ShouldServeCustomContent()
    {
        // Arrange
        var config = new TestHttpServerConfig
        {
            Port = 8772,
            Routes = new Dictionary<string, string>
            {
                { "/simple", IcsTestDataGenerator.GenerateSimpleCalendar("Simple Event") },
                { "/multiple", IcsTestDataGenerator.GenerateCalendarWithMultipleEvents(3) },
                { "/recurring", IcsTestDataGenerator.GenerateRecurringEventCalendar() },
                { "/timezone", IcsTestDataGenerator.GenerateCalendarWithTimezone("Europe/London") }
            }
        };

        using var server = new TestHttpServer(config);
        _servers.Add(server);

        await server.StartAsync();
        var downloader = new IcsDownloader();

        // Act & Assert
        var simpleContent = await downloader.DownloadAsync(server.BaseUrl + "simple");
        Assert.Contains("Simple Event", simpleContent);

        var multipleContent = await downloader.DownloadAsync(server.BaseUrl + "multiple");
        var eventCount = multipleContent.Split("BEGIN:VEVENT").Length - 1;
        Assert.Equal(3, eventCount);

        var recurringContent = await downloader.DownloadAsync(server.BaseUrl + "recurring");
        Assert.Contains("RRULE:FREQ=WEEKLY", recurringContent);

        var timezoneContent = await downloader.DownloadAsync(server.BaseUrl + "timezone");
        Assert.Contains("Europe/London", timezoneContent);
    }

    [Fact]
    public async Task TestHttpServer_WithAuth_ShouldRequireAuthentication()
    {
        // Arrange
        using var server = TestHttpServer.CreateWithAuth("testuser", "testpass", 8773);
        _servers.Add(server);

        await server.StartAsync();

        // Act & Assert - без авторизации должно быть 401
        using var httpClient = new HttpClient();
        var response = await httpClient.GetAsync(server.BaseUrl + "protected");
        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);

        // Act & Assert - с авторизацией должно быть 200
        var credentials = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("testuser:testpass"));
        httpClient.DefaultRequestHeaders.Authorization =
            new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", credentials);

        var downloader = new IcsDownloader(httpClient);
        var content = await downloader.DownloadAsync(server.BaseUrl + "protected");
        Assert.Contains("Protected Event", content);
    }

    [Fact]
    public async Task TestHttpServer_WithStatusCodes_ShouldReturnCorrectStatuses()
    {
        // Arrange
        var statusCodes = new Dictionary<string, HttpStatusCode>
        {
            { "/ok", HttpStatusCode.OK },
            { "/notfound", HttpStatusCode.NotFound },
            { "/forbidden", HttpStatusCode.Forbidden },
            { "/error", HttpStatusCode.InternalServerError }
        };

        using var server = TestHttpServer.CreateWithStatusCodes(statusCodes, 8774);
        _servers.Add(server);

        await server.StartAsync();

        // Act & Assert
        using var httpClient = new HttpClient();

        var okResponse = await httpClient.GetAsync(server.BaseUrl + "ok");
        Assert.Equal(HttpStatusCode.OK, okResponse.StatusCode);

        var notFoundResponse = await httpClient.GetAsync(server.BaseUrl + "notfound");
        Assert.Equal(HttpStatusCode.NotFound, notFoundResponse.StatusCode);

        var forbiddenResponse = await httpClient.GetAsync(server.BaseUrl + "forbidden");
        Assert.Equal(HttpStatusCode.Forbidden, forbiddenResponse.StatusCode);

        var errorResponse = await httpClient.GetAsync(server.BaseUrl + "error");
        Assert.Equal(HttpStatusCode.InternalServerError, errorResponse.StatusCode);
    }

    [Fact]
    public async Task TestHttpServer_WithDelay_ShouldRespectDelay()
    {
        // Arrange
        var config = new TestHttpServerConfig
        {
            Port = 8775,
            ResponseDelay = TimeSpan.FromMilliseconds(500),
            Routes = new Dictionary<string, string>
            {
                { "/slow", IcsTestDataGenerator.GenerateSimpleCalendar("Slow Event") }
            }
        };

        using var server = new TestHttpServer(config);
        _servers.Add(server);

        await server.StartAsync();
        var downloader = new IcsDownloader();

        // Act
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        var content = await downloader.DownloadAsync(server.BaseUrl + "slow");
        stopwatch.Stop();

        // Assert
        Assert.True(stopwatch.ElapsedMilliseconds >= 400,
            $"Ожидалась задержка минимум 400мс, получено: {stopwatch.ElapsedMilliseconds}мс");
        Assert.Contains("Slow Event", content);
    }

    [Fact]
    public void IcsTestDataGenerator_GenerateSimpleCalendar_ShouldCreateValidIcs()
    {
        // Act
        var calendar = IcsTestDataGenerator.GenerateSimpleCalendar(
            "Custom Event",
            "Custom Description",
            "Custom Location",
            DateTime.UtcNow.AddDays(2),
            DateTime.UtcNow.AddDays(2).AddHours(2)
        );

        // Assert
        Assert.Contains("BEGIN:VCALENDAR", calendar);
        Assert.Contains("END:VCALENDAR", calendar);
        Assert.Contains("Custom Event", calendar);
        Assert.Contains("Custom Description", calendar);
        Assert.Contains("Custom Location", calendar);
    }

    [Fact]
    public void IcsTestDataGenerator_GenerateCalendarWithMultipleEvents_ShouldCreateMultipleEvents()
    {
        // Act
        var calendar = IcsTestDataGenerator.GenerateCalendarWithMultipleEvents(5);

        // Assert
        var eventCount = calendar.Split("BEGIN:VEVENT").Length - 1;
        Assert.Equal(5, eventCount);
        Assert.Contains("Test Event 1", calendar);
        Assert.Contains("Test Event 5", calendar);
    }

    public void Dispose()
    {
        foreach (var server in _servers)
        {
            server.Dispose();
        }
    }
}