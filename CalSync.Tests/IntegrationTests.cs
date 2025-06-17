using CalSync.Services;
using Xunit;
using Xunit.Abstractions;

namespace CalSync.Tests;

/// <summary>
/// Интеграционные тесты для проверки реальных календарей
/// </summary>
public class IntegrationTests : TestBase
{
    private readonly ITestOutputHelper _output;

    public IntegrationTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public async Task DownloadAndParseRealICloudCalendar_ShouldFindTestEvent()
    {
        // Arrange
        var downloader = new IcsDownloader();
        var parser = new IcsParser();
        var url = GetTestUrl("RealICloudCalendar");

        _output.WriteLine($"Загружаем календарь с URL: {url}");

        try
        {
            // Act - Скачиваем календарь
            var icsContent = await downloader.DownloadAsync(url);
            _output.WriteLine($"Загружено {icsContent.Length} символов");
            _output.WriteLine("Первые 500 символов:");
            _output.WriteLine(icsContent.Substring(0, Math.Min(500, icsContent.Length)));

            // Act - Парсим события
            var events = parser.Parse(icsContent);
            _output.WriteLine($"Найдено событий: {events.Count}");

            // Выводим все события для отладки
            foreach (var evt in events)
            {
                _output.WriteLine($"Событие: '{evt.Summary}' - {evt.Start:yyyy-MM-dd HH:mm} (TimeZone: {evt.TimeZone})");
            }

            // Assert - Проверяем, что есть события
            Assert.NotEmpty(events);

            // Assert - Ищем событие "test" 19 июня 2025 года в 10:15 MSK
            var testEvent = events.FirstOrDefault(e =>
                e.Summary.Equals("test", StringComparison.OrdinalIgnoreCase) &&
                e.Start.Year == 2025 &&
                e.Start.Month == 6 &&
                e.Start.Day == 19);

            Assert.NotNull(testEvent);
            _output.WriteLine($"Найдено тестовое событие: '{testEvent.Summary}' - {testEvent.Start:yyyy-MM-dd HH:mm:ss} (TimeZone: {testEvent.TimeZone})");

            // Проверяем время (может быть в UTC или с указанием временной зоны)
            // 10:15 MSK = 07:15 UTC (MSK = UTC+3)
            if (testEvent.Start.Kind == DateTimeKind.Utc)
            {
                // Если время в UTC, проверяем 07:15
                Assert.Equal(7, testEvent.Start.Hour);
                Assert.Equal(15, testEvent.Start.Minute);
            }
            else if (!string.IsNullOrEmpty(testEvent.TimeZone))
            {
                // Если указана временная зона, проверяем локальное время
                _output.WriteLine($"Временная зона события: {testEvent.TimeZone}");
                // Время может быть в разных форматах, проверим основные варианты
                var expectedTimes = new[] {
                    new { Hour = 10, Minute = 15 }, // MSK время
                    new { Hour = 7, Minute = 15 },  // UTC время
                    new { Hour = 13, Minute = 15 }  // Если календарь в другой зоне
                };

                var timeMatches = expectedTimes.Any(t => t.Hour == testEvent.Start.Hour && t.Minute == testEvent.Start.Minute);
                Assert.True(timeMatches, $"Время события {testEvent.Start.Hour}:{testEvent.Start.Minute:D2} не соответствует ожидаемому");
            }
            else
            {
                // Локальное время без указания зоны
                Assert.Equal(10, testEvent.Start.Hour);
                Assert.Equal(15, testEvent.Start.Minute);
            }
        }
        finally
        {
            downloader.Dispose();
        }
    }

    [Fact]
    public async Task DownloadICloudCalendar_ShouldHandleWebcalProtocol()
    {
        // Arrange
        var downloader = new IcsDownloader();
        var webcalUrl = GetTestUrl("RealICloudCalendar");
        var httpsUrl = GetTestUrl("HttpsICloudCalendar");

        try
        {
            // Act & Assert - Оба URL должны работать
            var webcalContent = await downloader.DownloadAsync(webcalUrl);
            var httpsContent = await downloader.DownloadAsync(httpsUrl);

            Assert.NotEmpty(webcalContent);
            Assert.NotEmpty(httpsContent);

            // Содержимое должно быть одинаковым
            Assert.Equal(webcalContent, httpsContent);

            _output.WriteLine($"Webcal URL успешно конвертирован и загружен");
            _output.WriteLine($"Размер контента: {webcalContent.Length} символов");
        }
        finally
        {
            downloader.Dispose();
        }
    }

    [Fact]
    public async Task ParseICloudCalendar_ShouldExtractAllEventProperties()
    {
        // Arrange
        var downloader = new IcsDownloader();
        var parser = new IcsParser();
        var url = GetTestUrl("RealICloudCalendar");

        try
        {
            // Act
            var icsContent = await downloader.DownloadAsync(url);
            var events = parser.Parse(icsContent);

            // Assert
            Assert.NotEmpty(events);

            // Проверяем, что у событий есть основные свойства
            foreach (var evt in events.Take(5)) // Проверяем первые 5 событий
            {
                _output.WriteLine($"Событие: {evt.Summary}");
                _output.WriteLine($"  UID: {evt.Uid}");
                _output.WriteLine($"  Начало: {evt.Start:yyyy-MM-dd HH:mm:ss} (Kind: {evt.Start.Kind})");
                _output.WriteLine($"  Конец: {evt.End:yyyy-MM-dd HH:mm:ss}");
                _output.WriteLine($"  Временная зона: {evt.TimeZone}");
                _output.WriteLine($"  Весь день: {evt.IsAllDay}");
                _output.WriteLine($"  Статус: {evt.Status}");
                _output.WriteLine("");

                Assert.NotEmpty(evt.Uid);
                Assert.NotEmpty(evt.Summary);
                Assert.NotEqual(DateTime.MinValue, evt.Start);
            }
        }
        finally
        {
            downloader.Dispose();
        }
    }
}