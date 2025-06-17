using System.Text;
using Xunit;
using CalSync.Services;
using CalSync.Models;
using CalSync.Tests.TestHelpers;

namespace CalSync.Tests;

/// <summary>
/// Тесты для FR-002: Парсинг .ics файлов
/// </summary>
public class FR002_IcsParserTests : TestBase, IDisposable
{
    private readonly IcsParser _parser;
    private readonly TestHttpServer? _testServer;

    public FR002_IcsParserTests()
    {
        _parser = new IcsParser();

        // Создаем тестовый сервер для интеграционных тестов
        try
        {
            _testServer = TestHttpServer.CreateSimple(8778);
            _testServer.AddRoute("/simple.ics", IcsTestDataGenerator.GenerateSimpleCalendar("Server Event"));
            _testServer.AddRoute("/multiple.ics", IcsTestDataGenerator.GenerateCalendarWithMultipleEvents(2));
            _testServer.AddRoute("/recurring.ics", IcsTestDataGenerator.GenerateRecurringEventCalendar());
            _testServer.StartAsync().Wait();
        }
        catch
        {
            // Игнорируем ошибки запуска сервера в тестах
        }
    }

    [Fact]
    public void Parse_ValidRFC5545Format_ShouldParseSuccessfully()
    {
        // Arrange
        var icsContent = IcsTestDataGenerator.GenerateSimpleCalendar(
            "Test Event",
            "Test Description",
            "Test Location",
            new DateTime(2024, 12, 1, 9, 0, 0, DateTimeKind.Utc),
            new DateTime(2024, 12, 1, 10, 0, 0, DateTimeKind.Utc)
        );

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.NotNull(events);
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.NotNull(calendarEvent.Uid);
        Assert.Equal("Test Event", calendarEvent.Summary);
        Assert.Equal("Test Description", calendarEvent.Description);
        Assert.Equal("Test Location", calendarEvent.Location);
        Assert.Equal(EventStatus.Confirmed, calendarEvent.Status);
    }

    [Fact]
    public void Parse_WithVEvent_ShouldExtractEventProperties()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test//EN
BEGIN:VEVENT
UID:event1@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Meeting with Team
DESCRIPTION:Weekly team meeting
LOCATION:Conference Room A
ORGANIZER:MAILTO:organizer@example.com
ATTENDEE:MAILTO:attendee1@example.com
ATTENDEE:MAILTO:attendee2@example.com
STATUS:CONFIRMED
LAST-MODIFIED:20241201T080000Z
END:VEVENT
END:VCALENDAR";

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.Equal("event1@example.com", calendarEvent.Uid);
        Assert.Equal("Meeting with Team", calendarEvent.Summary);
        Assert.Equal("Weekly team meeting", calendarEvent.Description);
        Assert.Equal("Conference Room A", calendarEvent.Location);
        Assert.Equal("organizer@example.com", calendarEvent.Organizer);
        Assert.Equal(2, calendarEvent.Attendees.Count);
        Assert.Contains("attendee1@example.com", calendarEvent.Attendees);
        Assert.Contains("attendee2@example.com", calendarEvent.Attendees);
        Assert.Equal(EventStatus.Confirmed, calendarEvent.Status);
        Assert.Equal(new DateTime(2024, 12, 1, 9, 0, 0, DateTimeKind.Utc), calendarEvent.Start);
        Assert.Equal(new DateTime(2024, 12, 1, 10, 0, 0, DateTimeKind.Utc), calendarEvent.End);
    }

    [Fact]
    public void Parse_WithRecurrenceRule_ShouldParseRRule()
    {
        // Arrange
        var icsContent = IcsTestDataGenerator.GenerateRecurringEventCalendar();

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.NotNull(calendarEvent.RecurrenceRule);
        Assert.Contains("FREQ=WEEKLY", calendarEvent.RecurrenceRule);
        Assert.Contains("BYDAY=MO", calendarEvent.RecurrenceRule);
    }

    [Theory]
    [InlineData("FREQ=DAILY")]
    [InlineData("FREQ=WEEKLY;BYDAY=MO,WE,FR")]
    [InlineData("FREQ=MONTHLY;BYMONTHDAY=15")]
    [InlineData("FREQ=YEARLY;BYMONTH=12;BYMONTHDAY=25")]
    public void Parse_WithVariousRecurrencePatterns_ShouldParseCorrectly(string rrule)
    {
        // Arrange
        var icsContent = $@"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test//EN
BEGIN:VEVENT
UID:test@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Test Event
RRULE:{rrule}
END:VEVENT
END:VCALENDAR";

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);
        Assert.Equal(rrule, events[0].RecurrenceRule);
    }

    [Fact]
    public void Parse_WithTimezone_ShouldHandleTimezones()
    {
        // Arrange
        var icsContent = IcsTestDataGenerator.GenerateCalendarWithTimezone("Europe/Moscow");

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.Equal("Europe/Moscow", calendarEvent.TimeZone);
        Assert.NotEqual(DateTime.MinValue, calendarEvent.Start);
        Assert.NotEqual(DateTime.MinValue, calendarEvent.End);
    }

    [Theory]
    [InlineData("utf-8")]
    [InlineData("iso-8859-1")]
    public void Parse_WithDifferentEncodings_ShouldHandleCorrectly(string encoding)
    {
        // Arrange
        var summary = "Test with special chars: äöü";
        var icsContent = IcsTestDataGenerator.GenerateSimpleCalendar(summary, "Description", "Location");

        // Act
        var encodingObj = Encoding.GetEncoding(encoding);
        var bytes = encodingObj.GetBytes(icsContent);
        var decodedContent = encodingObj.GetString(bytes);

        var events = _parser.Parse(decodedContent);

        // Assert
        Assert.Single(events);
        // Для UTF-8 текст должен сохраниться, для других кодировок может измениться
        Assert.NotEmpty(events[0].Summary);
    }

    [Fact]
    public void Parse_WithMultipleEvents_ShouldParseAllEvents()
    {
        // Arrange
        var icsContent = IcsTestDataGenerator.GenerateCalendarWithMultipleEvents(3);

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Equal(3, events.Count);

        // Проверяем что все события имеют уникальные UID
        var uids = events.Select(e => e.Uid).ToList();
        Assert.Equal(3, uids.Distinct().Count());

        // Проверяем что все события имеют заголовки
        foreach (var calendarEvent in events)
        {
            Assert.NotEmpty(calendarEvent.Summary);
            Assert.Contains("Test Event", calendarEvent.Summary);
        }
    }

    [Fact]
    public void Parse_WithAllDayEvent_ShouldHandleCorrectly()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test//EN
BEGIN:VEVENT
UID:allday@example.com
DTSTART;VALUE=DATE:20241225
DTEND;VALUE=DATE:20241226
SUMMARY:Christmas Day
END:VEVENT
END:VCALENDAR";

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.True(calendarEvent.IsAllDay);
        Assert.Equal("Christmas Day", calendarEvent.Summary);
        Assert.Equal(new DateTime(2024, 12, 25), calendarEvent.Start.Date);
    }

    [Fact]
    public void Parse_WithEscapedText_ShouldUnescapeCorrectly()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test//EN
BEGIN:VEVENT
UID:escaped@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Text with\nline break and\, comma
DESCRIPTION:Text with semicolon\; and backslash\\
END:VEVENT
END:VCALENDAR";

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.Contains("\n", calendarEvent.Summary);
        Assert.Contains(",", calendarEvent.Summary);
        Assert.Contains(";", calendarEvent.Description);
        Assert.Contains("\\", calendarEvent.Description);
    }

    [Theory]
    [InlineData("CONFIRMED", EventStatus.Confirmed)]
    [InlineData("CANCELLED", EventStatus.Cancelled)]
    [InlineData("TENTATIVE", EventStatus.Tentative)]
    [InlineData("UNKNOWN", EventStatus.Tentative)] // Default fallback
    public void Parse_WithDifferentStatuses_ShouldParseCorrectly(string status, EventStatus expectedStatus)
    {
        // Arrange
        var icsContent = $@"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test//EN
BEGIN:VEVENT
UID:status-test@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Status Test Event
STATUS:{status}
END:VEVENT
END:VCALENDAR";

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);
        Assert.Equal(expectedStatus, events[0].Status);
    }

    [Fact]
    public void Parse_WithMultilineValues_ShouldConcatenateProperly()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test//EN
BEGIN:VEVENT
UID:multiline@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Multiline
 Summary Test
DESCRIPTION:This is a very long description that spans
 multiple lines and should be concatenated properly
 when parsed
END:VEVENT
END:VCALENDAR";

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.Equal("MultilineSummary Test", calendarEvent.Summary);
        Assert.Contains("very long description", calendarEvent.Description);
        Assert.Contains("concatenated properly", calendarEvent.Description);
        Assert.DoesNotContain("\n", calendarEvent.Description); // Should not contain line breaks
    }

    [Fact]
    public void Parse_EmptyContent_ShouldThrowArgumentException()
    {
        // Act & Assert
        Assert.Throws<ArgumentException>(() => _parser.Parse(""));
        Assert.Throws<ArgumentException>(() => _parser.Parse("   "));
        Assert.Throws<ArgumentException>(() => _parser.Parse(null!));
    }

    [Fact]
    public void Parse_InvalidContent_ShouldReturnEmptyList()
    {
        // Arrange
        var invalidContent = IcsTestDataGenerator.GenerateInvalidCalendar();

        // Act
        var events = _parser.Parse(invalidContent);

        // Assert
        Assert.NotNull(events);
        Assert.Empty(events);
    }

    [Fact]
    public void Parse_EventWithoutUID_ShouldSkipEvent()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test//EN
BEGIN:VEVENT
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Event without UID
END:VEVENT
BEGIN:VEVENT
UID:valid@example.com
DTSTART:20241201T100000Z
DTEND:20241201T110000Z
SUMMARY:Valid Event
END:VEVENT
END:VCALENDAR";

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events); // Only the valid event should be parsed
        Assert.Equal("valid@example.com", events[0].Uid);
        Assert.Equal("Valid Event", events[0].Summary);
    }

    [Fact]
    public void Parse_ComplexRealWorldExample_ShouldParseSuccessfully()
    {
        // Arrange - используем сложный пример из IcsTestDataGenerator
        var icsContent = IcsTestDataGenerator.GenerateCalendarWithTimezone("Europe/Moscow");

        // Act
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.NotEmpty(calendarEvent.Uid);
        Assert.NotEmpty(calendarEvent.Summary);
        Assert.Equal("Europe/Moscow", calendarEvent.TimeZone);
        Assert.NotEqual(DateTime.MinValue, calendarEvent.Start);
        Assert.NotEqual(DateTime.MinValue, calendarEvent.End);
        Assert.NotEqual(DateTime.MinValue, calendarEvent.LastModified);
    }

    [Fact]
    public async Task Parse_IntegrationWithDownloader_ShouldWorkTogether()
    {
        // Arrange
        if (_testServer?.IsRunning != true)
        {
            // Пропускаем тест если сервер не запущен
            return;
        }

        var downloader = new IcsDownloader();
        var testUrl = _testServer.BaseUrl + "simple.ics";

        // Act
        var icsContent = await downloader.DownloadAsync(testUrl);
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.NotNull(events);
        Assert.Single(events);

        var calendarEvent = events[0];
        Assert.Equal("Server Event", calendarEvent.Summary);
        Assert.NotEmpty(calendarEvent.Uid);
        Assert.Equal(EventStatus.Confirmed, calendarEvent.Status);
    }

    [Fact]
    public async Task Parse_IntegrationWithMultipleEvents_ShouldParseAll()
    {
        // Arrange
        if (_testServer?.IsRunning != true)
        {
            return;
        }

        var downloader = new IcsDownloader();
        var testUrl = _testServer.BaseUrl + "multiple.ics";

        // Act
        var icsContent = await downloader.DownloadAsync(testUrl);
        var events = _parser.Parse(icsContent);

        // Assert
        Assert.Equal(2, events.Count);

        foreach (var calendarEvent in events)
        {
            Assert.NotEmpty(calendarEvent.Uid);
            Assert.NotEmpty(calendarEvent.Summary);
            Assert.NotEqual(DateTime.MinValue, calendarEvent.Start);
            Assert.NotEqual(DateTime.MinValue, calendarEvent.End);
        }
    }

    public void Dispose()
    {
        _testServer?.Dispose();
    }
}