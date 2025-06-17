namespace CalSync.Models;

/// <summary>
/// Модель календарного события
/// </summary>
public class CalendarEvent
{
    /// <summary>
    /// Уникальный идентификатор события
    /// </summary>
    public string Uid { get; set; } = string.Empty;

    /// <summary>
    /// Заголовок события
    /// </summary>
    public string Summary { get; set; } = string.Empty;

    /// <summary>
    /// Описание события
    /// </summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// Дата и время начала события
    /// </summary>
    public DateTime Start { get; set; }

    /// <summary>
    /// Дата и время окончания события
    /// </summary>
    public DateTime End { get; set; }

    /// <summary>
    /// Местоположение события
    /// </summary>
    public string Location { get; set; } = string.Empty;

    /// <summary>
    /// Организатор события
    /// </summary>
    public string Organizer { get; set; } = string.Empty;

    /// <summary>
    /// Список участников
    /// </summary>
    public List<string> Attendees { get; set; } = new();

    /// <summary>
    /// Статус события
    /// </summary>
    public EventStatus Status { get; set; } = EventStatus.Tentative;

    /// <summary>
    /// Дата последнего изменения
    /// </summary>
    public DateTime LastModified { get; set; }

    /// <summary>
    /// Событие на весь день
    /// </summary>
    public bool IsAllDay { get; set; }

    /// <summary>
    /// Временная зона события
    /// </summary>
    public string TimeZone { get; set; } = string.Empty;

    /// <summary>
    /// Правило повторения (RRULE)
    /// </summary>
    public string RecurrenceRule { get; set; } = string.Empty;

    /// <summary>
    /// Exchange ID события (для работы с EWS)
    /// </summary>
    public string ExchangeId { get; set; } = string.Empty;

    /// <summary>
    /// Exchange ChangeKey события (для обновлений и удаления)
    /// </summary>
    public string ExchangeChangeKey { get; set; } = string.Empty;
}

/// <summary>
/// Статус события
/// </summary>
public enum EventStatus
{
    Tentative,
    Confirmed,
    Cancelled
}