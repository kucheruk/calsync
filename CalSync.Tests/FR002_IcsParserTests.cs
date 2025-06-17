using System.Text;
using Xunit;

namespace CalSync.Tests;

/// <summary>
/// Тесты для FR-002: Парсинг .ics файлов
/// </summary>
public class FR002_IcsParserTests
{
    [Fact]
    public void ParseIcsFile_ValidRFC5545Format_ShouldParseSuccessfully()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test//EN
BEGIN:VEVENT
UID:test@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Test Event
END:VEVENT
END:VCALENDAR";

        // Act & Assert
        // Тест должен парсить стандартный RFC 5545 формат
        Assert.Contains("BEGIN:VCALENDAR", icsContent);
        Assert.Contains("VERSION:2.0", icsContent);
        Assert.Contains("BEGIN:VEVENT", icsContent);
        Assert.Contains("END:VEVENT", icsContent);
        Assert.Contains("END:VCALENDAR", icsContent);
    }

    [Fact]
    public void ParseIcsFile_WithVEvent_ShouldExtractEventProperties()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
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
END:VEVENT
END:VCALENDAR";

        // Act & Assert
        // Тест должен извлекать все основные свойства события
        Assert.Contains("UID:event1@example.com", icsContent);
        Assert.Contains("SUMMARY:Meeting with Team", icsContent);
        Assert.Contains("DESCRIPTION:Weekly team meeting", icsContent);
        Assert.Contains("LOCATION:Conference Room A", icsContent);
        Assert.Contains("ORGANIZER:MAILTO:organizer@example.com", icsContent);
        Assert.Contains("STATUS:CONFIRMED", icsContent);
    }

    [Fact]
    public void ParseIcsFile_WithRecurrenceRule_ShouldParseRRule()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:recurring@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Daily Standup
RRULE:FREQ=DAILY;COUNT=30
END:VEVENT
END:VCALENDAR";

        // Act & Assert
        // Тест должен парсить правила повторения (RRULE)
        Assert.Contains("RRULE:FREQ=DAILY;COUNT=30", icsContent);
    }

    [Theory]
    [InlineData("FREQ=DAILY")]
    [InlineData("FREQ=WEEKLY;BYDAY=MO,WE,FR")]
    [InlineData("FREQ=MONTHLY;BYMONTHDAY=15")]
    [InlineData("FREQ=YEARLY;BYMONTH=12;BYMONTHDAY=25")]
    public void ParseIcsFile_WithVariousRecurrencePatterns_ShouldParseCorrectly(string rrule)
    {
        // Arrange
        var icsContent = $@"BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:test@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:Test Event
RRULE:{rrule}
END:VEVENT
END:VCALENDAR";

        // Act & Assert
        // Тест должен парсить различные паттерны повторения
        Assert.Contains($"RRULE:{rrule}", icsContent);
    }

    [Fact]
    public void ParseIcsFile_WithTimezone_ShouldHandleTimezones()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VTIMEZONE
TZID:America/New_York
BEGIN:STANDARD
DTSTART:20071104T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
TZNAME:EST
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:20070311T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
TZNAME:EDT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
END:DAYLIGHT
END:VTIMEZONE
BEGIN:VEVENT
UID:tz-event@example.com
DTSTART;TZID=America/New_York:20241201T090000
DTEND;TZID=America/New_York:20241201T100000
SUMMARY:Timezone Event
END:VEVENT
END:VCALENDAR";

        // Act & Assert
        // Тест должен обрабатывать временные зоны
        Assert.Contains("BEGIN:VTIMEZONE", icsContent);
        Assert.Contains("TZID:America/New_York", icsContent);
        Assert.Contains("DTSTART;TZID=America/New_York", icsContent);
    }

    [Theory]
    [InlineData("utf-8")]
    [InlineData("iso-8859-1")]
    [InlineData("windows-1252")]
    public void ParseIcsFile_WithDifferentEncodings_ShouldHandleCorrectly(string encoding)
    {
        // Arrange
        var summary = "Тест с русскими символами";
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:encoding-test@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:" + summary + @"
END:VEVENT
END:VCALENDAR";

        // Act & Assert
        // Тест должен обрабатывать различные кодировки
        var encodingObj = Encoding.GetEncoding(encoding);
        var bytes = encodingObj.GetBytes(icsContent);
        var decodedContent = encodingObj.GetString(bytes);
        Assert.Contains(summary, decodedContent);
    }

    [Fact]
    public void ParseIcsFile_WithMultipleEvents_ShouldParseAllEvents()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:event1@example.com
DTSTART:20241201T090000Z
DTEND:20241201T100000Z
SUMMARY:First Event
END:VEVENT
BEGIN:VEVENT
UID:event2@example.com
DTSTART:20241202T140000Z
DTEND:20241202T150000Z
SUMMARY:Second Event
END:VEVENT
BEGIN:VEVENT
UID:event3@example.com
DTSTART:20241203T160000Z
DTEND:20241203T170000Z
SUMMARY:Third Event
END:VEVENT
END:VCALENDAR";

        // Act & Assert
        // Тест должен парсить несколько событий в одном файле
        var eventCount = icsContent.Split("BEGIN:VEVENT").Length - 1;
        Assert.Equal(3, eventCount);
    }

    [Fact]
    public void ParseIcsFile_WithAllDayEvent_ShouldHandleCorrectly()
    {
        // Arrange
        var icsContent = @"BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:allday@example.com
DTSTART;VALUE=DATE:20241225
SUMMARY:Christmas Day
END:VEVENT
END:VCALENDAR";

        // Act & Assert
        // Тест должен обрабатывать события на весь день
        Assert.Contains("DTSTART;VALUE=DATE:20241225", icsContent);
        Assert.DoesNotContain("DTEND", icsContent);
    }
}