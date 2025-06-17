using CalSync.Models;
using Microsoft.Extensions.Configuration;
using Xunit;
using ExchangeService = CalSync.Services.ExchangeService;

namespace CalSync.Tests;

/// <summary>
/// Интеграционные тесты с реальным Exchange сервером
/// Требует наличия appsettings.Local.json с настройками Exchange
/// </summary>
[Collection("Exchange Integration")]
public class ExchangeIntegrationTests : IDisposable
{
    private readonly ExchangeService? _exchangeService;
    private readonly bool _isConfigured;

    public ExchangeIntegrationTests()
    {
        try
        {
            // Пытаемся загрузить конфигурацию
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true)
                .AddJsonFile("appsettings.Local.json", optional: true)
                .Build();

            var exchangeConfig = configuration.GetSection("Exchange");
            var serviceUrl = exchangeConfig["ServiceUrl"];
            var username = exchangeConfig["Username"];
            var password = exchangeConfig["Password"];

            if (!string.IsNullOrEmpty(serviceUrl) && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
            {
                _exchangeService = new ExchangeService(configuration);
                _isConfigured = true;
            }
            else
            {
                _isConfigured = false;
            }
        }
        catch
        {
            _isConfigured = false;
        }
    }

    [Fact]
    public async Task TestConnection_WithRealExchange_ShouldConnect()
    {
        // Skip if not configured
        if (!_isConfigured || _exchangeService == null)
        {
            Assert.True(true, "Тест пропущен - Exchange не настроен");
            return;
        }

        // Act
        var result = await _exchangeService.TestConnectionAsync();

        // Assert
        Assert.True(result, "Подключение к реальному Exchange должно быть успешным");
    }

    [Fact]
    public async Task GetCalendarEvents_WithRealExchange_ShouldReturnEvents()
    {
        // Skip if not configured
        if (!_isConfigured || _exchangeService == null)
        {
            Assert.True(true, "Тест пропущен - Exchange не настроен");
            return;
        }

        // Act
        var events = await _exchangeService.GetCalendarEventsAsync();

        // Assert
        Assert.NotNull(events);

        // Логируем результаты для анализа
        Console.WriteLine($"📅 Получено событий: {events.Count}");

        if (events.Count > 0)
        {
            var firstEvent = events[0];
            Console.WriteLine($"📋 Первое событие:");
            Console.WriteLine($"   Название: {firstEvent.Summary}");
            Console.WriteLine($"   Время: {firstEvent.Start:yyyy-MM-dd HH:mm} - {firstEvent.End:yyyy-MM-dd HH:mm}");
            Console.WriteLine($"   Место: {firstEvent.Location}");
            Console.WriteLine($"   ID: {firstEvent.ExchangeId}");
            Console.WriteLine($"   Статус: {firstEvent.Status}");
        }

        // Базовые проверки
        foreach (var evt in events.Take(5)) // Проверяем первые 5 событий
        {
            Assert.NotNull(evt.ExchangeId);
            Assert.NotEmpty(evt.ExchangeId);
            Assert.NotNull(evt.Summary);
            Assert.True(evt.Start != DateTime.MinValue);
            Assert.True(evt.End != DateTime.MinValue);
            Assert.True(evt.End >= evt.Start);
        }
    }

    [Fact]
    public async Task CreateAndDeleteTestEvent_WithRealExchange_ShouldWork()
    {
        // Skip if not configured
        if (!_isConfigured || _exchangeService == null)
        {
            Assert.True(true, "Тест пропущен - Exchange не настроен");
            return;
        }

        // Arrange
        var testEvent = new CalendarEvent
        {
            Summary = $"[TEST] CalSync Integration Test - {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
            Description = "Автоматический тест интеграции CalSync с Exchange",
            Start = DateTime.Now.AddHours(1),
            End = DateTime.Now.AddHours(2),
            Location = "Тестовая локация для интеграционного теста"
        };

        string? createdEventId = null;

        try
        {
            // Act - создаем событие
            createdEventId = await _exchangeService.CreateCalendarEventAsync(testEvent);

            // Assert - проверяем создание
            Assert.NotNull(createdEventId);
            Assert.NotEmpty(createdEventId);
            Console.WriteLine($"✅ Создано тестовое событие с ID: {createdEventId}");

            // Даем время на синхронизацию
            await Task.Delay(2000);

            // Проверяем, что событие появилось в календаре
            var events = await _exchangeService.GetCalendarEventsAsync(
                DateTime.Today,
                DateTime.Today.AddDays(2));

            var createdEvent = events.FirstOrDefault(e => e.Summary.Contains("CalSync Integration Test"));
            Assert.NotNull(createdEvent);
            Console.WriteLine($"✅ Событие найдено в календаре: {createdEvent.Summary}");
        }
        finally
        {
            // Cleanup - удаляем созданное событие
            if (!string.IsNullOrEmpty(createdEventId))
            {
                try
                {
                    var deleteResult = await _exchangeService.DeleteCalendarEventAsync(createdEventId);
                    Console.WriteLine($"🗑️  Результат удаления: {deleteResult}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️  Не удалось удалить тестовое событие: {ex.Message}");
                }
            }
        }
    }

    [Fact]
    public async Task CleanupTestEvents_WithRealExchange_ShouldRemoveTestEvents()
    {
        // Skip if not configured
        if (!_isConfigured || _exchangeService == null)
        {
            Assert.True(true, "Тест пропущен - Exchange не настроен");
            return;
        }

        // Act
        var deletedCount = await _exchangeService.DeleteAllTestEventsAsync();

        // Assert
        Console.WriteLine($"🧹 Удалено тестовых событий: {deletedCount}");
        Assert.True(deletedCount >= 0, "Количество удаленных событий должно быть >= 0");
    }

    [Fact]
    public async Task GetCalendarEvents_PerformanceTest_ShouldCompleteInReasonableTime()
    {
        // Skip if not configured
        if (!_isConfigured || _exchangeService == null)
        {
            Assert.True(true, "Тест пропущен - Exchange не настроен");
            return;
        }

        // Arrange
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        // Act
        var events = await _exchangeService.GetCalendarEventsAsync();

        // Assert
        stopwatch.Stop();
        Console.WriteLine($"⏱️  Время получения {events.Count} событий: {stopwatch.ElapsedMilliseconds} мс");

        // Проверяем, что операция завершилась за разумное время (менее 30 секунд)
        Assert.True(stopwatch.ElapsedMilliseconds < 30000,
            $"Получение событий заняло слишком много времени: {stopwatch.ElapsedMilliseconds} мс");
    }

    public void Dispose()
    {
        _exchangeService?.Dispose();
    }
}