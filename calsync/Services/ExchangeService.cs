using CalSync.Models;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;

namespace CalSync.Services;

/// <summary>
/// Сервис для работы с Exchange Web Services
/// </summary>
public class ExchangeService : IDisposable
{
    private readonly Microsoft.Exchange.WebServices.Data.ExchangeService _service;
    private readonly IConfiguration _configuration;
    private bool _disposed = false;

    public ExchangeService(IConfiguration configuration)
    {
        _configuration = configuration;
        var exchangeConfig = _configuration.GetSection("Exchange");

        // Создаем сервис с нужной версией
        var version = exchangeConfig["Version"];
        var exchangeVersion = version switch
        {
            "Exchange2013" => ExchangeVersion.Exchange2013,
            "Exchange2013_SP1" => ExchangeVersion.Exchange2013_SP1,
            _ => ExchangeVersion.Exchange2013_SP1
        };

        _service = new Microsoft.Exchange.WebServices.Data.ExchangeService(exchangeVersion);
        Initialize();
    }

    /// <summary>
    /// Инициализация сервиса
    /// </summary>
    private void Initialize()
    {
        var exchangeConfig = _configuration.GetSection("Exchange");

        // Настройка аутентификации
        var domain = exchangeConfig["Domain"];
        var username = exchangeConfig["Username"];
        var password = exchangeConfig["Password"];

        if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
        {
            // Попробуем разные варианты аутентификации
            if (!string.IsNullOrEmpty(domain))
            {
                // Вариант 1: domain\username
                _service.Credentials = new WebCredentials($"{domain}\\{username}", password);
                Console.WriteLine($"🔐 Аутентификация: {domain}\\{username}");
            }
            else
            {
                // Вариант 2: просто username
                _service.Credentials = new WebCredentials(username, password);
                Console.WriteLine($"🔐 Аутентификация: {username}");
            }
        }

        // Настройка URL сервиса
        var serviceUrl = exchangeConfig["ServiceUrl"];
        if (!string.IsNullOrEmpty(serviceUrl))
        {
            _service.Url = new Uri(serviceUrl);
            Console.WriteLine($"🌐 EWS URL: {serviceUrl}");
        }

        // Настройка таймаута
        if (int.TryParse(exchangeConfig["RequestTimeout"], out var timeout))
        {
            _service.Timeout = timeout;
        }

        // Настройка валидации SSL
        var validateSsl = exchangeConfig["ValidateSslCertificate"]?.ToLower() != "false";
        if (!validateSsl)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback =
                (sender, certificate, chain, sslPolicyErrors) => true;
            Console.WriteLine("⚠️  SSL валидация отключена");
        }
    }

    /// <summary>
    /// Тестирование подключения к Exchange
    /// </summary>
    public async Task<bool> TestConnectionAsync()
    {
        var exchangeConfig = _configuration.GetSection("Exchange");
        var domain = exchangeConfig["Domain"];
        var username = exchangeConfig["Username"];
        var password = exchangeConfig["Password"];

        // Попробуем разные варианты аутентификации
        var credentialVariants = new List<(string name, WebCredentials creds)>();

        if (!string.IsNullOrEmpty(domain) && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
        {
            credentialVariants.Add(($"{domain}\\{username}", new WebCredentials($"{domain}\\{username}", password)));
            credentialVariants.Add(($"{username}@{domain}", new WebCredentials($"{username}@{domain}", password)));
            credentialVariants.Add((username, new WebCredentials(username, password)));
        }

        foreach (var (name, creds) in credentialVariants)
        {
            try
            {
                Console.WriteLine($"🔍 Тестирование аутентификации: {name}");
                _service.Credentials = creds;

                // Тестируем подключение через получение календаря
                var calendar = Folder.Bind(_service, WellKnownFolderName.Calendar);

                Console.WriteLine($"✅ Подключение успешно! Календарь: {calendar.DisplayName}");
                Console.WriteLine($"✅ Используется аутентификация: {name}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Не удалось с {name}: {ex.Message}");
            }
        }

        Console.WriteLine("❌ Все варианты аутентификации не сработали");
        return false;
    }

    /// <summary>
    /// Получить события календаря
    /// </summary>
    public async Task<List<CalSync.Models.CalendarEvent>> GetCalendarEventsAsync(DateTime? startDate = null, DateTime? endDate = null)
    {
        try
        {
            Console.WriteLine("📅 Получение событий календаря...");

            var start = startDate ?? DateTime.Today.AddDays(-7);
            var end = endDate ?? DateTime.Today.AddDays(30);

            Console.WriteLine($"📅 Период: {start:yyyy-MM-dd} - {end:yyyy-MM-dd}");

            // Получаем календарь пользователя
            var calendar = CalendarFolder.Bind(_service, WellKnownFolderName.Calendar);

            // Создаем представление календаря
            var calendarView = new CalendarView(start, end);
            calendarView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);

            // Получаем события
            var findResults = calendar.FindAppointments(calendarView);

            Console.WriteLine($"✅ Найдено событий: {findResults.Items.Count}");

            var events = new List<CalSync.Models.CalendarEvent>();

            foreach (var appointment in findResults.Items)
            {
                try
                {
                    // Загружаем дополнительные свойства
                    appointment.Load(new PropertySet(BasePropertySet.FirstClassProperties));

                    var calendarEvent = MapToCalendarEvent(appointment);
                    events.Add(calendarEvent);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️  Ошибка при обработке события {appointment.Id}: {ex.Message}");
                }
            }

            return events;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка получения событий: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Создать событие в календаре
    /// </summary>
    public async Task<string> CreateCalendarEventAsync(CalSync.Models.CalendarEvent calendarEvent)
    {
        try
        {
            Console.WriteLine($"➕ Создание события: {calendarEvent.Summary}");

            var appointment = new Appointment(_service);

            // Основные свойства
            appointment.Subject = calendarEvent.Summary;
            appointment.Body = calendarEvent.Description ?? "";
            appointment.Start = calendarEvent.Start;
            appointment.End = calendarEvent.End;

            if (!string.IsNullOrEmpty(calendarEvent.Location))
            {
                appointment.Location = calendarEvent.Location;
            }

            // Добавляем маркер для идентификации наших событий
            appointment.Body = new MessageBody(BodyType.Text,
                $"{calendarEvent.Description}\n\n[CalSync-Test-Event-{DateTime.UtcNow:yyyyMMddHHmmss}]");

            // Сохраняем событие
            appointment.Save(SendInvitationsMode.SendToNone);

            Console.WriteLine($"✅ Событие создано: {appointment.Id}");
            return appointment.Id.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка создания события: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Удалить событие календаря
    /// </summary>
    public async Task<bool> DeleteCalendarEventAsync(string eventId)
    {
        try
        {
            Console.WriteLine($"🗑️  Удаление события: {eventId}");

            var appointment = Appointment.Bind(_service, new ItemId(eventId));

            // Проверяем, что это наше тестовое событие
            if (!IsTestEvent(appointment))
            {
                Console.WriteLine("⚠️  Событие не помечено как тестовое, удаление отменено");
                return false;
            }

            appointment.Delete(DeleteMode.MoveToDeletedItems);

            Console.WriteLine("✅ Событие удалено");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка удаления события: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Удалить все тестовые события
    /// </summary>
    public async Task<int> DeleteAllTestEventsAsync()
    {
        try
        {
            Console.WriteLine("🧹 Удаление всех тестовых событий...");

            var events = await GetCalendarEventsAsync(DateTime.Today.AddDays(-30), DateTime.Today.AddDays(30));
            var deletedCount = 0;

            foreach (var evt in events)
            {
                if (!string.IsNullOrEmpty(evt.ExchangeId))
                {
                    try
                    {
                        var appointment = Appointment.Bind(_service, new ItemId(evt.ExchangeId));

                        if (IsTestEvent(appointment))
                        {
                            appointment.Delete(DeleteMode.MoveToDeletedItems);
                            deletedCount++;
                            Console.WriteLine($"🗑️  Удалено: {evt.Summary}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️  Не удалось удалить {evt.Summary}: {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"✅ Удалено тестовых событий: {deletedCount}");
            return deletedCount;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка при удалении тестовых событий: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Проверить, является ли событие тестовым
    /// </summary>
    private bool IsTestEvent(Appointment appointment)
    {
        var body = appointment.Body?.Text ?? "";
        return body.Contains("[CalSync-Test-Event-") ||
               appointment.Subject.StartsWith("[TEST]", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Преобразовать Exchange Appointment в CalendarEvent
    /// </summary>
    private CalSync.Models.CalendarEvent MapToCalendarEvent(Appointment appointment)
    {
        return new CalSync.Models.CalendarEvent
        {
            ExchangeId = appointment.Id.ToString(),
            Uid = appointment.Id.ToString(),
            Summary = appointment.Subject ?? "",
            Description = appointment.Body?.Text ?? "",
            Start = appointment.Start,
            End = appointment.End,
            Location = appointment.Location ?? "",
            LastModified = appointment.LastModifiedTime,
            Status = appointment.LegacyFreeBusyStatus switch
            {
                LegacyFreeBusyStatus.Free => EventStatus.Tentative,
                LegacyFreeBusyStatus.Busy => EventStatus.Confirmed,
                LegacyFreeBusyStatus.Tentative => EventStatus.Tentative,
                _ => EventStatus.Confirmed
            }
        };
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
        }
    }
}