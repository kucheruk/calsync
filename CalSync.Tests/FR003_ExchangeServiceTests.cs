using CalSync.Models;
using Microsoft.Extensions.Configuration;
using Xunit;
using ExchangeService = CalSync.Services.ExchangeService;

namespace CalSync.Tests;

/// <summary>
/// Тесты для ExchangeService (FR-003: Подключение к Exchange Server)
/// </summary>
public class FR003_ExchangeServiceTests : IDisposable
{
    private readonly MockEwsServer _mockServer;
    private readonly IConfiguration _configuration;
    private readonly ExchangeService _exchangeService;

    public FR003_ExchangeServiceTests()
    {
        // Запускаем мок-сервер на свободном порту
        _mockServer = new MockEwsServer(8081);
        _mockServer.Start();

        // Даем серверу время запуститься
        Thread.Sleep(500);

        // Создаем конфигурацию для тестов
        var configData = new Dictionary<string, string>
        {
            ["Exchange:ServiceUrl"] = _mockServer.ServiceUrl,
            ["Exchange:Domain"] = "test.local",
            ["Exchange:Username"] = "testuser",
            ["Exchange:Password"] = "testpass",
            ["Exchange:Version"] = "Exchange2013_SP1",
            ["Exchange:UseAutodiscover"] = "false",
            ["Exchange:RequestTimeout"] = "30000",
            ["Exchange:ValidateSslCertificate"] = "false"
        };

        _configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(configData!)
            .Build();

        _exchangeService = new ExchangeService(_configuration);
    }

    [Fact]
    public async Task TestConnectionAsync_WithMockServer_ShouldReturnTrue()
    {
        // Act
        var result = await _exchangeService.TestConnectionAsync();

        // Assert
        Assert.True(result, "Подключение к мок-серверу должно быть успешным");
    }

    [Fact]
    public async Task GetCalendarEventsAsync_WithMockServer_ShouldReturnEvents()
    {
        // Arrange
        var startDate = new DateTime(2025, 6, 17);
        var endDate = new DateTime(2025, 6, 18);

        // Act
        var events = await _exchangeService.GetCalendarEventsAsync(startDate, endDate);

        // Assert
        Assert.NotNull(events);
        Assert.Equal(3, events.Count);

        // Проверяем первое тестовое событие
        var testEvent1 = events.FirstOrDefault(e => e.ExchangeId == "TEST001");
        Assert.NotNull(testEvent1);
        Assert.Equal("[TEST] CalSync тестовое событие 1", testEvent1.Summary);
        Assert.Equal("Тестовая локация 1", testEvent1.Location);
        Assert.Equal(new DateTime(2025, 6, 17, 10, 0, 0, DateTimeKind.Utc), testEvent1.Start);
        Assert.Equal(new DateTime(2025, 6, 17, 11, 0, 0, DateTimeKind.Utc), testEvent1.End);
        Assert.Equal(EventStatus.Confirmed, testEvent1.Status);

        // Проверяем обычное событие
        var normalEvent = events.FirstOrDefault(e => e.ExchangeId == "TEST002");
        Assert.NotNull(normalEvent);
        Assert.Equal("Обычное событие", normalEvent.Summary);
        Assert.Equal(string.Empty, normalEvent.Location);

        // Проверяем второе тестовое событие
        var testEvent2 = events.FirstOrDefault(e => e.ExchangeId == "TEST003");
        Assert.NotNull(testEvent2);
        Assert.Equal("[TEST] CalSync тестовое событие 2", testEvent2.Summary);
        Assert.Equal("Тестовая локация 2", testEvent2.Location);
        Assert.Equal(EventStatus.Tentative, testEvent2.Status);
    }

    [Fact]
    public async Task GetCalendarEventsAsync_WithDefaultDates_ShouldReturnEvents()
    {
        // Act
        var events = await _exchangeService.GetCalendarEventsAsync();

        // Assert
        Assert.NotNull(events);
        Assert.Equal(3, events.Count);
    }

    [Fact]
    public async Task CreateCalendarEventAsync_WithValidEvent_ShouldReturnEventId()
    {
        // Arrange
        var testEvent = new CalendarEvent
        {
            Summary = "[TEST] Новое тестовое событие",
            Description = "Описание тестового события",
            Start = DateTime.Now.AddHours(1),
            End = DateTime.Now.AddHours(2),
            Location = "Тестовая локация"
        };

        // Act
        var eventId = await _exchangeService.CreateCalendarEventAsync(testEvent);

        // Assert
        Assert.NotNull(eventId);
        Assert.NotEmpty(eventId);
        Assert.StartsWith("TEST", eventId);
    }

    [Fact]
    public async Task DeleteCalendarEventAsync_WithTestEvent_ShouldReturnTrue()
    {
        // Act
        var result = await _exchangeService.DeleteCalendarEventAsync("TEST001");

        // Assert
        Assert.True(result, "Удаление тестового события должно быть успешным");
    }

    [Fact]
    public async Task DeleteAllTestEventsAsync_WithMockServer_ShouldReturnCount()
    {
        // Act
        var deletedCount = await _exchangeService.DeleteAllTestEventsAsync();

        // Assert
        // Мок-сервер возвращает 3 события, из которых 2 тестовых
        Assert.Equal(2, deletedCount);
    }

    [Fact]
    public void Constructor_WithValidConfiguration_ShouldInitializeCorrectly()
    {
        // Act & Assert - если конструктор не выбросил исключение, то инициализация прошла успешно
        Assert.NotNull(_exchangeService);
    }

    [Fact]
    public async Task GetCalendarEventsAsync_EventMapping_ShouldMapAllProperties()
    {
        // Act
        var events = await _exchangeService.GetCalendarEventsAsync();
        var testEvent = events.FirstOrDefault(e => e.ExchangeId == "TEST001");

        // Assert
        Assert.NotNull(testEvent);
        Assert.Equal("TEST001", testEvent.ExchangeId);
        Assert.Equal("TEST001", testEvent.Uid); // UID должен совпадать с ExchangeId
        Assert.Equal("[TEST] CalSync тестовое событие 1", testEvent.Summary);
        Assert.Equal("Тестовая локация 1", testEvent.Location);
        Assert.Equal(EventStatus.Confirmed, testEvent.Status);

        // Проверяем, что время корректно распарсилось
        Assert.Equal(DateTimeKind.Utc, testEvent.Start.Kind);
        Assert.Equal(DateTimeKind.Utc, testEvent.End.Kind);
    }

    [Theory]
    [InlineData("TEST001", true)]  // Тестовое событие
    [InlineData("TEST003", true)]  // Тестовое событие
    [InlineData("TEST002", false)] // Обычное событие
    public async Task Event_TestEventIdentification_ShouldIdentifyCorrectly(string eventId, bool shouldBeTestEvent)
    {
        // Arrange
        var events = await _exchangeService.GetCalendarEventsAsync();
        var targetEvent = events.FirstOrDefault(e => e.ExchangeId == eventId);

        // Assert
        Assert.NotNull(targetEvent);

        var isTestEvent = targetEvent.Summary.StartsWith("[TEST]") ||
                         targetEvent.Description.Contains("[CalSync-Test-Event-");

        Assert.Equal(shouldBeTestEvent, isTestEvent);
    }

    public void Dispose()
    {
        _exchangeService?.Dispose();
        _mockServer?.Dispose();
    }
}