using CalSync.Models;
using System.Globalization;
using System.Text.RegularExpressions;

namespace CalSync.Services;

/// <summary>
/// Парсер ICS файлов
/// </summary>
public class IcsParser
{
    /// <summary>
    /// Парсить ICS содержимое в список событий
    /// </summary>
    /// <param name="icsContent">Содержимое ICS файла</param>
    /// <returns>Список календарных событий</returns>
    public List<CalendarEvent> Parse(string icsContent)
    {
        if (string.IsNullOrWhiteSpace(icsContent))
            throw new ArgumentException("ICS содержимое не может быть пустым", nameof(icsContent));

        var events = new List<CalendarEvent>();
        var lines = SplitIntoLines(icsContent);

        for (int i = 0; i < lines.Count; i++)
        {
            if (lines[i].StartsWith("BEGIN:VEVENT", StringComparison.OrdinalIgnoreCase))
            {
                var eventEndIndex = FindEventEnd(lines, i);
                if (eventEndIndex > i)
                {
                    var eventLines = lines.GetRange(i, eventEndIndex - i + 1);
                    var calendarEvent = ParseEvent(eventLines);
                    if (calendarEvent != null)
                    {
                        events.Add(calendarEvent);
                    }
                    i = eventEndIndex;
                }
            }
        }

        return events;
    }

    /// <summary>
    /// Разделить содержимое на строки с учетом многострочных значений
    /// </summary>
    private List<string> SplitIntoLines(string content)
    {
        var lines = new List<string>();
        var rawLines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

        for (int i = 0; i < rawLines.Length; i++)
        {
            var line = rawLines[i];

            // Объединяем строки, которые начинаются с пробела или табуляции (RFC 5545)
            while (i + 1 < rawLines.Length &&
                   (rawLines[i + 1].StartsWith(" ") || rawLines[i + 1].StartsWith("\t")))
            {
                i++;
                line += rawLines[i].Substring(1); // Убираем первый символ (пробел/таб)
            }

            if (!string.IsNullOrWhiteSpace(line))
            {
                lines.Add(line.Trim());
            }
        }

        return lines;
    }

    /// <summary>
    /// Найти индекс строки с END:VEVENT
    /// </summary>
    private int FindEventEnd(List<string> lines, int startIndex)
    {
        for (int i = startIndex + 1; i < lines.Count; i++)
        {
            if (lines[i].StartsWith("END:VEVENT", StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }
        return -1;
    }

    /// <summary>
    /// Парсить отдельное событие
    /// </summary>
    private CalendarEvent? ParseEvent(List<string> eventLines)
    {
        var calendarEvent = new CalendarEvent();

        foreach (var line in eventLines)
        {
            var colonIndex = line.IndexOf(':');
            if (colonIndex == -1) continue;

            var propertyPart = line.Substring(0, colonIndex);
            var value = line.Substring(colonIndex + 1);

            // Разделяем свойство и параметры
            var semicolonIndex = propertyPart.IndexOf(';');
            var property = semicolonIndex == -1 ? propertyPart : propertyPart.Substring(0, semicolonIndex);
            var parameters = semicolonIndex == -1 ? "" : propertyPart.Substring(semicolonIndex + 1);

            switch (property.ToUpperInvariant())
            {
                case "UID":
                    calendarEvent.Uid = value;
                    break;
                case "SUMMARY":
                    calendarEvent.Summary = UnescapeText(value);
                    break;
                case "DESCRIPTION":
                    calendarEvent.Description = UnescapeText(value);
                    break;
                case "LOCATION":
                    calendarEvent.Location = UnescapeText(value);
                    break;
                case "URL":
                    calendarEvent.Url = value;
                    break;
                case "ORGANIZER":
                    calendarEvent.Organizer = ExtractEmail(value);
                    break;
                case "ATTENDEE":
                    calendarEvent.Attendees.Add(ExtractEmail(value));
                    break;
                case "DTSTART":
                    calendarEvent.Start = ParseDateTime(value, parameters);
                    calendarEvent.IsAllDay = parameters.Contains("VALUE=DATE");
                    if (parameters.Contains("TZID="))
                    {
                        calendarEvent.TimeZone = ExtractTimeZone(parameters);
                    }
                    break;
                case "DTEND":
                    calendarEvent.End = ParseDateTime(value, parameters);
                    break;
                case "LAST-MODIFIED":
                    calendarEvent.LastModified = ParseDateTime(value, parameters);
                    break;
                case "STATUS":
                    calendarEvent.Status = ParseStatus(value);
                    break;
                case "RRULE":
                    calendarEvent.RecurrenceRule = value;
                    break;
            }
        }

        return string.IsNullOrEmpty(calendarEvent.Uid) ? null : calendarEvent;
    }

    /// <summary>
    /// Парсить дату и время
    /// </summary>
    private DateTime ParseDateTime(string value, string parameters)
    {
        // Убираем Z в конце для UTC времени
        var cleanValue = value.TrimEnd('Z');

        // Парсим дату в формате VALUE=DATE (YYYYMMDD)
        if (parameters.Contains("VALUE=DATE") && cleanValue.Length == 8)
        {
            if (DateTime.TryParseExact(cleanValue, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateOnly))
            {
                return dateOnly;
            }
        }

        // Парсим дату и время в формате YYYYMMDDTHHMMSS
        if (cleanValue.Length >= 15 && cleanValue.Contains('T'))
        {
            if (DateTime.TryParseExact(cleanValue, "yyyyMMddTHHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTime))
            {
                // Если оригинальное значение заканчивалось на Z, это UTC время
                if (value.EndsWith('Z'))
                {
                    return DateTime.SpecifyKind(dateTime, DateTimeKind.Utc);
                }
                return dateTime;
            }
        }

        return DateTime.MinValue;
    }

    /// <summary>
    /// Извлечь временную зону из параметров
    /// </summary>
    private string ExtractTimeZone(string parameters)
    {
        var match = Regex.Match(parameters, @"TZID=([^;]+)", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : "";
    }

    /// <summary>
    /// Извлечь email из строки вида "MAILTO:email@example.com" или "CN=Name:MAILTO:email@example.com"
    /// </summary>
    private string ExtractEmail(string value)
    {
        var match = Regex.Match(value, @"MAILTO:([^;]+)", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : value;
    }

    /// <summary>
    /// Парсить статус события
    /// </summary>
    private EventStatus ParseStatus(string value)
    {
        return value.ToUpperInvariant() switch
        {
            "CONFIRMED" => EventStatus.Confirmed,
            "CANCELLED" => EventStatus.Cancelled,
            _ => EventStatus.Tentative
        };
    }

    /// <summary>
    /// Убрать экранирование из текста (RFC 5545)
    /// </summary>
    private string UnescapeText(string text)
    {
        return text
            .Replace("\\n", "\n")
            .Replace("\\N", "\n")
            .Replace("\\r", "\r")
            .Replace("\\t", "\t")
            .Replace("\\\\", "\\")
            .Replace("\\;", ";")
            .Replace("\\,", ",");
    }
}