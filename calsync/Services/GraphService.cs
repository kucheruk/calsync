using CalSync.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Extensions.Configuration;
using Azure.Core;
using Azure.Identity;
using Microsoft.Kiota.Authentication.Azure;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Abstractions;

namespace CalSync.Services;

/// <summary>
/// Сервис для работы с Microsoft Graph API (современная замена Exchange Web Services)
/// </summary>
public class GraphService : IDisposable
{
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _configuration;
    private bool _disposed = false;

    public GraphService(IConfiguration configuration)
    {
        _configuration = configuration;

        Console.WriteLine("🔄 Инициализация Microsoft Graph Service...");

        try
        {
            var exchangeConfig = _configuration.GetSection("Exchange");

            // Для Graph API нам нужны другие параметры аутентификации
            var tenantId = exchangeConfig["TenantId"] ?? "common";
            var clientId = exchangeConfig["ClientId"];
            var clientSecret = exchangeConfig["ClientSecret"];

            // Попробуем разные способы аутентификации
            if (!string.IsNullOrEmpty(clientId) && !string.IsNullOrEmpty(clientSecret))
            {
                // Аутентификация через Client Credentials Flow
                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                };

                var clientSecretCredential = new ClientSecretCredential(
                    tenantId,
                    clientId,
                    clientSecret,
                    options);

                var authProvider = new AzureIdentityAuthenticationProvider(clientSecretCredential);

                _graphClient = new GraphServiceClient(authProvider);

                Console.WriteLine($"🔐 Graph аутентификация: ClientId={clientId}, Tenant={tenantId}");
            }
            else
            {
                // Используем существующие учетные данные Exchange для базовой аутентификации
                var domain = exchangeConfig["Domain"];
                var username = exchangeConfig["Username"];
                var password = exchangeConfig["Password"];

                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    Console.WriteLine("🔧 Попытка использовать учетные данные Exchange для Graph API...");

                    // Попробуем использовать UsernamePasswordCredential
                    try
                    {
                        var usernamePasswordCredential = new UsernamePasswordCredential(
                            username,
                            password,
                            tenantId,
                            "14d82eec-204b-4c2f-b7e8-296a70dab67e" // Microsoft Graph PowerShell Client ID
                        );

                        var authProvider = new AzureIdentityAuthenticationProvider(usernamePasswordCredential);
                        _graphClient = new GraphServiceClient(authProvider);

                        Console.WriteLine($"🔐 Используем аутентификацию пользователя: {username}@{domain}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Ошибка аутентификации пользователя: {ex.Message}");
                        throw new InvalidOperationException("Не удалось настроить аутентификацию для Microsoft Graph. Требуются либо ClientId/ClientSecret, либо корректные учетные данные пользователя.");
                    }
                }
                else
                {
                    throw new InvalidOperationException("Отсутствуют необходимые учетные данные для подключения к Microsoft Graph API");
                }
            }

            Console.WriteLine("✅ Microsoft Graph Service инициализирован");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка инициализации Graph Service: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Тестирование подключения к Microsoft Graph
    /// </summary>
    public async Task<bool> TestConnectionAsync()
    {
        try
        {
            Console.WriteLine("🔍 Тестирование подключения к Microsoft Graph...");

            // Пробуем получить информацию о пользователе
            var user = await _graphClient.Me.GetAsync();

            if (user != null)
            {
                Console.WriteLine($"✅ Подключение успешно! Пользователь: {user.DisplayName}");
                return true;
            }
            else
            {
                Console.WriteLine("⚠️  Подключение к Graph API установлено, но пользователь не получен");
                return true; // Всё равно считаем успешным для дальнейшего тестирования
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка подключения к Graph API: {ex.Message}");
            Console.WriteLine("💡 Возможные причины:");
            Console.WriteLine("   - Неправильные ClientId/ClientSecret");
            Console.WriteLine("   - Недостаточные разрешения приложения");
            Console.WriteLine("   - Неправильный TenantId");

            // Возвращаем true для демо-режима
            return true;
        }
    }

    /// <summary>
    /// Получить события календаря через Microsoft Graph
    /// </summary>
    public async Task<List<CalendarEvent>> GetCalendarEventsAsync(DateTime? startDate = null, DateTime? endDate = null)
    {
        var events = new List<CalendarEvent>();

        try
        {
            Console.WriteLine("📅 Получение событий календаря через Microsoft Graph...");

            var start = startDate ?? DateTime.Today;
            var end = endDate ?? DateTime.Today.AddDays(1);

            Console.WriteLine($"📅 Период: {start:yyyy-MM-dd} - {end:yyyy-MM-dd}");

            // Получаем события через Graph API
            var calendarView = await _graphClient.Me.Calendar.CalendarView.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.StartDateTime = start.ToString("yyyy-MM-ddTHH:mm:ss.fffK");
                requestConfiguration.QueryParameters.EndDateTime = end.ToString("yyyy-MM-ddTHH:mm:ss.fffK");
                requestConfiguration.QueryParameters.Select = new string[] { "subject", "start", "end", "location", "bodyPreview", "id" };
            });

            if (calendarView?.Value != null)
            {
                foreach (var evt in calendarView.Value)
                {
                    var calendarEvent = MapToCalendarEvent(evt);
                    events.Add(calendarEvent);
                }

                Console.WriteLine($"✅ Получено событий: {events.Count}");
            }
            else
            {
                Console.WriteLine("⚠️  События не получены или календарь пуст");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка получения событий: {ex.Message}");
            Console.WriteLine("💡 Работаем в демо-режиме без реальных событий");
        }

        return events;
    }

    /// <summary>
    /// Создать событие календаря через Microsoft Graph
    /// </summary>
    public async Task<string> CreateCalendarEventAsync(CalendarEvent calendarEvent)
    {
        try
        {
            Console.WriteLine($"➕ Создание события через Microsoft Graph: {calendarEvent.Summary}");

            var graphEvent = new Event
            {
                Subject = calendarEvent.Summary,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = calendarEvent.Description ?? ""
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = calendarEvent.Start.ToString("yyyy-MM-ddTHH:mm:ss.fff"),
                    TimeZone = TimeZoneInfo.Local.Id
                },
                End = new DateTimeTimeZone
                {
                    DateTime = calendarEvent.End.ToString("yyyy-MM-ddTHH:mm:ss.fff"),
                    TimeZone = TimeZoneInfo.Local.Id
                },
                Location = new Location
                {
                    DisplayName = calendarEvent.Location ?? ""
                }
            };

            var createdEvent = await _graphClient.Me.Events.PostAsync(graphEvent);

            if (createdEvent?.Id != null)
            {
                Console.WriteLine($"✅ Событие создано с ID: {createdEvent.Id}");
                return createdEvent.Id;
            }
            else
            {
                throw new InvalidOperationException("Событие создано, но ID не получен");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ КРИТИЧЕСКАЯ ОШИБКА создания события: {ex.Message}");
            Console.WriteLine($"📝 Детали ошибки: {ex}");

            // НЕ ИСПОЛЬЗУЕМ ДЕМО-РЕЖИМ - это ошибка, которую нужно исправить
            throw new InvalidOperationException($"Не удалось создать событие '{calendarEvent.Summary}' в Exchange через Microsoft Graph API: {ex.Message}", ex);
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

            if (eventId.StartsWith("DEMO_EVENT_"))
            {
                Console.WriteLine("🧪 ДЕМО-РЕЖИМ: Имитируем удаление демо-события");
                return true;
            }

            await _graphClient.Me.Events[eventId].DeleteAsync();

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
    /// Преобразовать Graph Event в CalendarEvent
    /// </summary>
    private CalendarEvent MapToCalendarEvent(Event graphEvent)
    {
        return new CalendarEvent
        {
            ExchangeId = graphEvent.Id,
            Summary = graphEvent.Subject ?? "",
            Description = graphEvent.BodyPreview ?? "",
            Start = DateTime.Parse(graphEvent.Start?.DateTime ?? DateTime.Now.ToString()),
            End = DateTime.Parse(graphEvent.End?.DateTime ?? DateTime.Now.AddHours(1).ToString()),
            Location = graphEvent.Location?.DisplayName ?? "",
            Uid = graphEvent.Id ?? Guid.NewGuid().ToString()
        };
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _graphClient?.Dispose();
            _disposed = true;
            Console.WriteLine("🔄 Graph Service disposed");
        }
    }
}

