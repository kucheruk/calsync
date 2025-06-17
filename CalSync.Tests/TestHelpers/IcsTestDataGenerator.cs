using System.Globalization;

namespace CalSync.Tests.TestHelpers;

/// <summary>
/// Генератор тестовых ICS данных для различных сценариев тестирования
/// </summary>
public static class IcsTestDataGenerator
{
    /// <summary>
    /// Генерирует простой тестовый календарь с одним событием
    /// </summary>
    public static string GenerateSimpleCalendar(
        string eventTitle = "Test Event",
        string eventDescription = "This is a test event for HTTP download testing",
        string location = "Test Location",
        DateTime? startTime = null,
        DateTime? endTime = null)
    {
        var start = startTime ?? DateTime.UtcNow.AddDays(1);
        var end = endTime ?? start.AddHours(1);

        var startStr = start.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
        var endStr = end.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
        var createdStr = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);

        return $@"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test Calendar//EN
CALSCALE:GREGORIAN
METHOD:PUBLISH
BEGIN:VEVENT
UID:test-event-{Guid.NewGuid()}@example.com
DTSTART:{startStr}
DTEND:{endStr}
SUMMARY:{eventTitle}
DESCRIPTION:{eventDescription}
LOCATION:{location}
STATUS:CONFIRMED
CREATED:{createdStr}
LAST-MODIFIED:{createdStr}
END:VEVENT
END:VCALENDAR";
    }

    /// <summary>
    /// Генерирует календарь с множественными событиями
    /// </summary>
    public static string GenerateCalendarWithMultipleEvents(int eventCount = 3)
    {
        var calendar = @"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test Calendar//EN
CALSCALE:GREGORIAN
METHOD:PUBLISH";

        for (int i = 0; i < eventCount; i++)
        {
            var start = DateTime.UtcNow.AddDays(i + 1);
            var end = start.AddHours(1);
            var startStr = start.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
            var endStr = end.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
            var createdStr = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);

            calendar += $@"
BEGIN:VEVENT
UID:test-event-{i}-{Guid.NewGuid()}@example.com
DTSTART:{startStr}
DTEND:{endStr}
SUMMARY:Test Event {i + 1}
DESCRIPTION:This is test event number {i + 1}
LOCATION:Test Location {i + 1}
STATUS:CONFIRMED
CREATED:{createdStr}
LAST-MODIFIED:{createdStr}
END:VEVENT";
        }

        calendar += @"
END:VCALENDAR";
        return calendar;
    }

    /// <summary>
    /// Генерирует календарь с повторяющимся событием
    /// </summary>
    public static string GenerateRecurringEventCalendar()
    {
        var start = DateTime.UtcNow.AddDays(1);
        var end = start.AddHours(1);
        var startStr = start.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
        var endStr = end.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
        var createdStr = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);

        return $@"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test Calendar//EN
CALSCALE:GREGORIAN
METHOD:PUBLISH
BEGIN:VEVENT
UID:recurring-event-{Guid.NewGuid()}@example.com
DTSTART:{startStr}
DTEND:{endStr}
SUMMARY:Weekly Recurring Event
DESCRIPTION:This event repeats every week
LOCATION:Recurring Location
STATUS:CONFIRMED
RRULE:FREQ=WEEKLY;BYDAY=MO
CREATED:{createdStr}
LAST-MODIFIED:{createdStr}
END:VEVENT
END:VCALENDAR";
    }

    /// <summary>
    /// Генерирует невалидный ICS файл для тестирования ошибок
    /// </summary>
    public static string GenerateInvalidCalendar()
    {
        return @"INVALID CALENDAR DATA
This is not a valid ICS file
Missing BEGIN:VCALENDAR";
    }

    /// <summary>
    /// Генерирует календарь с событием в определенной временной зоне
    /// </summary>
    public static string GenerateCalendarWithTimezone(string timezoneName = "Europe/Moscow")
    {
        var start = DateTime.UtcNow.AddDays(1);
        var end = start.AddHours(1);
        var startStr = start.ToString("yyyyMMddTHHmmss", CultureInfo.InvariantCulture);
        var endStr = end.ToString("yyyyMMddTHHmmss", CultureInfo.InvariantCulture);
        var createdStr = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);

        return $@"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Test//Test Calendar//EN
CALSCALE:GREGORIAN
METHOD:PUBLISH
BEGIN:VTIMEZONE
TZID:{timezoneName}
BEGIN:STANDARD
DTSTART:20071028T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
TZNAME:MSK
TZOFFSETFROM:+0400  
TZOFFSETTO:+0300
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:20070325T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
TZNAME:MSD
TZOFFSETFROM:+0300
TZOFFSETTO:+0400
END:DAYLIGHT
END:VTIMEZONE
BEGIN:VEVENT
UID:timezone-event-{Guid.NewGuid()}@example.com
DTSTART;TZID={timezoneName}:{startStr}
DTEND;TZID={timezoneName}:{endStr}
SUMMARY:Event with Timezone
DESCRIPTION:This event has timezone information
LOCATION:Timezone Location
STATUS:CONFIRMED
CREATED:{createdStr}
LAST-MODIFIED:{createdStr}
END:VEVENT
END:VCALENDAR";
    }
}