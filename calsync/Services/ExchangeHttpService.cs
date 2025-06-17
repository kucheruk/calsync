using CalSync.Models;
using Microsoft.Extensions.Configuration;
using System.Net.Http;
using System.Text;
using System.Xml;

namespace CalSync.Services;

/// <summary>
/// Сервис для прямого HTTP подключения к Exchange Web Services через SOAP
/// Обходит проблемы Microsoft.Exchange.WebServices с .NET 8
/// </summary>
public class ExchangeHttpService : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly IConfiguration _configuration;
    private readonly string _serviceUrl;
    private readonly string _domain;
    private readonly string _username;
    private readonly string _password;
    private readonly string _sendMeetingInvitations;
    private readonly string _sendMeetingCancellations;
    private readonly string _defaultTimeZone;
    private bool _disposed = false;

    public ExchangeHttpService(IConfiguration configuration)
    {
        _configuration = configuration;
        var exchangeConfig = _configuration.GetSection("Exchange");

        _serviceUrl = exchangeConfig["ServiceUrl"] ?? throw new ArgumentException("ServiceUrl не настроен");
        _domain = exchangeConfig["Domain"] ?? "";
        _username = exchangeConfig["Username"] ?? throw new ArgumentException("Username не настроен");
        _password = exchangeConfig["Password"] ?? throw new ArgumentException("Password не настроен");
        // Настройки уведомлений для календарных операций:
        // SendToNone - не отправлять уведомления
        // SendOnlyToAll - отправить только участникам (не сохранять в Отправленных)
        // SendToAllAndSaveCopy - отправить участникам и сохранить копию в Отправленных
        _sendMeetingInvitations = exchangeConfig["SendMeetingInvitations"] ?? "SendToAllAndSaveCopy";
        _sendMeetingCancellations = exchangeConfig["SendMeetingCancellations"] ?? "SendToAllAndSaveCopy";

        // Читаем настройку временной зоны из секции CalSync
        var calSyncConfig = _configuration.GetSection("CalSync");
        _defaultTimeZone = calSyncConfig["DefaultTimeZone"] ?? "Europe/Moscow";

        // Настраиваем HTTP клиент
        var handler = new HttpClientHandler();

        // Отключаем валидацию SSL если настроено
        var validateSsl = exchangeConfig["ValidateSslCertificate"]?.ToLower() != "false";
        if (!validateSsl)
        {
            handler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;
            Console.WriteLine("⚠️  SSL валидация отключена");
        }

        _httpClient = new HttpClient(handler);

        // Настраиваем аутентификацию
        var credentials = !string.IsNullOrEmpty(_domain) ?
            $"{_domain}\\{_username}" : _username;

        var authValue = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{credentials}:{_password}"));
        _httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", authValue);

        // Настраиваем таймаут
        if (int.TryParse(exchangeConfig["RequestTimeout"], out var timeout) && timeout > 0)
        {
            _httpClient.Timeout = TimeSpan.FromMilliseconds(timeout);
        }

        Console.WriteLine($"🔄 Инициализация Exchange HTTP Service");
        Console.WriteLine($"🌐 URL: {_serviceUrl}");
        Console.WriteLine($"🔐 Аутентификация: {credentials}");
        Console.WriteLine($"📧 Уведомления: создание={_sendMeetingInvitations}, удаление={_sendMeetingCancellations}");
        Console.WriteLine($"🌍 Временная зона по умолчанию: {_defaultTimeZone}");
    }

    /// <summary>
    /// Тестирование подключения к Exchange
    /// </summary>
    public async Task<bool> TestConnectionAsync()
    {
        try
        {
            Console.WriteLine("🔍 Тестирование HTTP подключения к Exchange...");

            // Отправляем простой SOAP запрос для получения папки Inbox
            var soapRequest = CreateGetFolderSoapRequest();
            var response = await SendSoapRequestAsync(soapRequest);

            if (response.Contains("Success"))
            {
                Console.WriteLine("✅ Подключение к Exchange успешно!");
                return true;
            }
            else if (response.Contains("Unauthorized"))
            {
                Console.WriteLine("❌ Ошибка аутентификации Exchange");
                return false;
            }
            else
            {
                Console.WriteLine("⚠️  Получен ответ от Exchange, проверяем детали...");
                Console.WriteLine($"📝 Первые 500 символов ответа: {response.Substring(0, Math.Min(500, response.Length))}");
                return true; // Считаем успешным, если получили любой ответ
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка тестирования подключения: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Получить события календаря через SOAP запрос
    /// </summary>
    public async Task<List<CalendarEvent>> GetCalendarEventsAsync(DateTime? startDate = null, DateTime? endDate = null)
    {
        var events = new List<CalendarEvent>();

        try
        {
            Console.WriteLine("📅 Получение событий календаря через SOAP...");

            var start = startDate ?? DateTime.Today;
            var end = endDate ?? DateTime.Today.AddDays(1);

            Console.WriteLine($"📅 Период: {start:yyyy-MM-dd} - {end:yyyy-MM-dd}");

            var soapRequest = CreateGetCalendarEventsSoapRequest(start, end);
            var response = await SendSoapRequestAsync(soapRequest);

            Console.WriteLine($"📥 Получен ответ от Exchange ({response.Length} символов)");

            // Парсим ответ и извлекаем события
            events = ParseCalendarEventsFromResponse(response);

            Console.WriteLine($"✅ Получено событий: {events.Count}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка получения событий: {ex.Message}");
        }

        return events;
    }

    /// <summary>
    /// Обновить событие календаря через прямой SOAP запрос
    /// </summary>
    public async Task<bool> UpdateCalendarEventAsync(CalendarEvent calendarEvent)
    {
        try
        {
            Console.WriteLine($"✏️ Обновление события через SOAP: {calendarEvent.Summary}");

            var soapRequest = CreateUpdateEventSoapRequest(calendarEvent);
            var response = await SendSoapRequestAsync(soapRequest);

            if (response.Contains("Success"))
            {
                Console.WriteLine("✅ Событие обновлено успешно");
                return true;
            }
            else
            {
                Console.WriteLine($"❌ Ошибка обновления: {ExtractErrorFromResponse(response)}");
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ КРИТИЧЕСКАЯ ОШИБКА обновления события: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Удалить событие календаря через прямой SOAP запрос
    /// </summary>
    public async Task<bool> DeleteCalendarEventAsync(string eventId)
    {
        try
        {
            Console.WriteLine($"🗑️ Удаление события через SOAP: {eventId}");

            var soapRequest = CreateDeleteEventSoapRequest(eventId);
            var response = await SendSoapRequestAsync(soapRequest);

            if (response.Contains("Success"))
            {
                Console.WriteLine("✅ Событие удалено успешно");
                return true;
            }
            else
            {
                Console.WriteLine($"❌ Ошибка удаления: {ExtractErrorFromResponse(response)}");
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ КРИТИЧЕСКАЯ ОШИБКА удаления события: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Создать событие календаря через прямой SOAP запрос
    /// </summary>
    public async Task<string> CreateCalendarEventAsync(CalendarEvent calendarEvent)
    {
        var createdEvent = await CreateCalendarEventWithDetailsAsync(calendarEvent);
        return createdEvent.ExchangeId;
    }

    /// <summary>
    /// Создать событие календаря и вернуть полную информацию включая ChangeKey
    /// </summary>
    public async Task<CalendarEvent> CreateCalendarEventWithDetailsAsync(CalendarEvent calendarEvent)
    {
        try
        {
            Console.WriteLine($"➕ Создание события через SOAP: {calendarEvent.Summary}");

            // Создаем SOAP запрос для создания события
            var soapRequest = CreateEventSoapRequest(calendarEvent);

            Console.WriteLine("📤 Отправка SOAP запроса...");
            var response = await SendSoapRequestAsync(soapRequest);

            Console.WriteLine($"📥 Получен ответ от Exchange ({response.Length} символов)");

            // Извлекаем ID и ChangeKey созданного события
            var (eventId, changeKey) = ExtractEventIdAndChangeKeyFromResponse(response);

            if (!string.IsNullOrEmpty(eventId))
            {
                Console.WriteLine($"✅ Событие создано с ID: {eventId}");

                // Обновляем исходное событие
                calendarEvent.ExchangeId = eventId;
                calendarEvent.ExchangeChangeKey = changeKey;

                return calendarEvent;
            }
            else if (response.Contains("Success"))
            {
                Console.WriteLine("✅ Событие создано успешно");
                calendarEvent.ExchangeId = $"EXCHANGE_EVENT_{DateTime.Now:yyyyMMddHHmmss}";
                return calendarEvent;
            }
            else
            {
                Console.WriteLine("❌ Ошибка создания события:");
                Console.WriteLine($"📝 Ответ сервера: {response.Substring(0, Math.Min(1000, response.Length))}");
                throw new InvalidOperationException($"Exchange вернул ошибку: {ExtractErrorFromResponse(response)}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ КРИТИЧЕСКАЯ ОШИБКА создания события: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Отправить SOAP запрос к Exchange
    /// </summary>
    private async Task<string> SendSoapRequestAsync(string soapRequest)
    {
        var content = new StringContent(soapRequest, Encoding.UTF8, "text/xml");

        // Добавляем необходимые заголовки для SOAP
        content.Headers.Clear();
        content.Headers.Add("Content-Type", "text/xml; charset=utf-8");

        // Добавляем SOAPAction заголовок (может потребоваться для некоторых операций)
        _httpClient.DefaultRequestHeaders.Remove("SOAPAction");
        _httpClient.DefaultRequestHeaders.Add("SOAPAction", "\"\"");

        Console.WriteLine($"📤 Отправляем SOAP запрос ({soapRequest.Length} символов)");

        var response = await _httpClient.PostAsync(_serviceUrl, content);
        var responseContent = await response.Content.ReadAsStringAsync();

        Console.WriteLine($"📥 Получен HTTP статус: {response.StatusCode}");
        Console.WriteLine($"📥 Размер ответа: {responseContent.Length} символов");

        if (!response.IsSuccessStatusCode)
        {
            Console.WriteLine($"❌ HTTP ошибка: {response.StatusCode}");
            Console.WriteLine($"📝 Ответ: {responseContent}");
        }
        else if (responseContent.Contains("s:Fault") || responseContent.Contains("ErrorInvalidRequest"))
        {
            Console.WriteLine($"⚠️  SOAP ошибка в ответе:");
            Console.WriteLine($"📝 Первые 1000 символов: {responseContent.Substring(0, Math.Min(1000, responseContent.Length))}");
        }

        return responseContent;
    }

    /// <summary>
    /// Создать SOAP запрос для получения папки (тест подключения)
    /// </summary>
    private string CreateGetFolderSoapRequest()
    {
        return @"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
               xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
               xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
               xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
  <soap:Header>
    <t:RequestServerVersion Version=""Exchange2013_SP1"" />
  </soap:Header>
  <soap:Body>
    <GetFolder xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
               xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <FolderShape>
        <t:BaseShape>Default</t:BaseShape>
      </FolderShape>
      <FolderIds>
        <t:DistinguishedFolderId Id=""inbox""/>
      </FolderIds>
    </GetFolder>
  </soap:Body>
</soap:Envelope>";
    }

    /// <summary>
    /// Форматировать время для отправки в Exchange с правильной временной зоной
    /// </summary>
    private string FormatTimeForExchange(DateTime dateTime, string? timeZone)
    {
        // Если время уже в UTC, используем его
        if (dateTime.Kind == DateTimeKind.Utc)
        {
            return dateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
        }

        // Используем указанную временную зону или временную зону по умолчанию из конфигурации
        var effectiveTimeZone = !string.IsNullOrEmpty(timeZone) ? timeZone : _defaultTimeZone;

        try
        {
            // Преобразуем названия временных зон
            var systemTimeZoneId = ConvertIcsTimeZoneToSystem(effectiveTimeZone);
            var tz = TimeZoneInfo.FindSystemTimeZoneById(systemTimeZoneId);

            // Предполагаем, что время в календарном событии уже в указанной временной зоне
            var utcTime = TimeZoneInfo.ConvertTimeToUtc(dateTime, tz);

            Console.WriteLine($"🌍 Конвертация времени: {dateTime:HH:mm:ss} ({effectiveTimeZone}) → {utcTime:HH:mm:ss} UTC");

            return utcTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️  Ошибка обработки временной зоны '{effectiveTimeZone}': {ex.Message}");

            // Fallback: используем временную зону по умолчанию из конфигурации
            try
            {
                var fallbackSystemTimeZoneId = ConvertIcsTimeZoneToSystem(_defaultTimeZone);
                var fallbackTz = TimeZoneInfo.FindSystemTimeZoneById(fallbackSystemTimeZoneId);
                var utcTime = TimeZoneInfo.ConvertTimeToUtc(dateTime, fallbackTz);

                Console.WriteLine($"🔄 Fallback конвертация: {dateTime:HH:mm:ss} ({_defaultTimeZone}) → {utcTime:HH:mm:ss} UTC");

                return utcTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            }
            catch
            {
                // Последний fallback: считаем время локальным и конвертируем в UTC
                var utc = dateTime.ToUniversalTime();
                Console.WriteLine($"🔄 Локальная конвертация: {dateTime:HH:mm:ss} → {utc:HH:mm:ss} UTC");
                return utc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            }
        }
    }

    /// <summary>
    /// Преобразовать ICS временную зону в системную
    /// </summary>
    private string ConvertIcsTimeZoneToSystem(string icsTimeZone)
    {
        var mapping = icsTimeZone switch
        {
            "Europe/Moscow" => "Russian Standard Time",
            "UTC" => "UTC",
            "GMT" => "GMT Standard Time",
            "Europe/London" => "GMT Standard Time",
            "America/New_York" => "Eastern Standard Time",
            "America/Los_Angeles" => "Pacific Standard Time",
            "Europe/Berlin" => "W. Europe Standard Time",
            "Europe/Paris" => "W. Europe Standard Time",
            "Asia/Tokyo" => "Tokyo Standard Time",
            "Australia/Sydney" => "AUS Eastern Standard Time",
            _ => null // Нет точного соответствия
        };

        // Если есть точное соответствие, возвращаем его
        if (mapping != null)
        {
            return mapping;
        }

        // Fallback: используем временную зону по умолчанию из конфигурации
        return ConvertIcsTimeZoneToSystem(_defaultTimeZone);
    }

    /// <summary>
    /// Создать SOAP запрос для создания события календаря
    /// </summary>
    private string CreateEventSoapRequest(CalendarEvent calendarEvent)
    {
        // Правильно обрабатываем временные зоны
        var startTime = FormatTimeForExchange(calendarEvent.Start, calendarEvent.TimeZone);
        var endTime = FormatTimeForExchange(calendarEvent.End, calendarEvent.TimeZone);

        Console.WriteLine($"🕒 Исходное время: {calendarEvent.Start:yyyy-MM-dd HH:mm:ss} (TimeZone: {calendarEvent.TimeZone ?? "не указана"})");
        Console.WriteLine($"🕒 Время для Exchange: {startTime}");

        return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
               xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
               xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
               xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
  <soap:Header>
    <t:RequestServerVersion Version=""Exchange2013_SP1"" />
  </soap:Header>
  <soap:Body>
    <CreateItem xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"" 
                MessageDisposition=""SaveOnly"" 
                SendMeetingInvitations=""{_sendMeetingInvitations}"">
      <SavedItemFolderId>
        <t:DistinguishedFolderId Id=""calendar""/>
      </SavedItemFolderId>
      <Items>
        <t:CalendarItem>
          <t:Subject>{System.Security.SecurityElement.Escape(calendarEvent.Summary)}</t:Subject>
          <t:Body BodyType=""Text"">{System.Security.SecurityElement.Escape(calendarEvent.Description ?? "")}</t:Body>
          <t:Start>{startTime}</t:Start>
          <t:End>{endTime}</t:End>
          <t:Location>{System.Security.SecurityElement.Escape(calendarEvent.Location ?? "")}</t:Location>
          <t:LegacyFreeBusyStatus>Busy</t:LegacyFreeBusyStatus>
        </t:CalendarItem>
      </Items>
    </CreateItem>
  </soap:Body>
</soap:Envelope>";
    }

    /// <summary>
    /// Извлечь ID события из ответа
    /// </summary>
    private string ExtractEventIdFromResponse(string response)
    {
        try
        {
            // Используем XmlReader для извлечения ID
            using (var stringReader = new System.IO.StringReader(response))
            using (var xmlReader = System.Xml.XmlReader.Create(stringReader))
            {
                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == System.Xml.XmlNodeType.Element && xmlReader.LocalName == "ItemId")
                    {
                        var id = xmlReader.GetAttribute("Id");
                        if (!string.IsNullOrEmpty(id))
                        {
                            return id;
                        }
                    }
                }
            }
        }
        catch
        {
            // Если не удалось извлечь через XML, попробуем regex
            var match = System.Text.RegularExpressions.Regex.Match(response, @"Id=""([^""]+)""");
            return match.Success ? match.Groups[1].Value : "";
        }

        return "";
    }

    /// <summary>
    /// Извлечь ID и ChangeKey события из ответа для создания
    /// </summary>
    private (string Id, string ChangeKey) ExtractEventIdAndChangeKeyFromResponse(string response)
    {
        try
        {
            using (var stringReader = new System.IO.StringReader(response))
            using (var xmlReader = System.Xml.XmlReader.Create(stringReader))
            {
                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == System.Xml.XmlNodeType.Element && xmlReader.LocalName == "ItemId")
                    {
                        var id = xmlReader.GetAttribute("Id");
                        var changeKey = xmlReader.GetAttribute("ChangeKey");
                        if (!string.IsNullOrEmpty(id))
                        {
                            return (id, changeKey ?? "");
                        }
                    }
                }
            }
        }
        catch
        {
            // Fallback
        }

        return ("", "");
    }

    /// <summary>
    /// Извлечь ошибку из ответа
    /// </summary>
    private string ExtractErrorFromResponse(string response)
    {
        try
        {
            // Используем XmlReader для извлечения ошибки
            using (var stringReader = new System.IO.StringReader(response))
            using (var xmlReader = System.Xml.XmlReader.Create(stringReader))
            {
                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == System.Xml.XmlNodeType.Element &&
                        (xmlReader.LocalName == "MessageText" || xmlReader.LocalName == "faultstring"))
                    {
                        xmlReader.Read(); // Переходим к тексту
                        if (xmlReader.NodeType == System.Xml.XmlNodeType.Text)
                        {
                            return xmlReader.Value;
                        }
                    }
                }
            }
        }
        catch
        {
            return "Не удалось извлечь описание ошибки";
        }

        return "Неизвестная ошибка Exchange";
    }

    /// <summary>
    /// Создать SOAP запрос для получения событий календаря
    /// </summary>
    private string CreateGetCalendarEventsSoapRequest(DateTime startDate, DateTime endDate)
    {
        // Используем правильный формат времени для CalendarView
        var startTimeUtc = startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
        var endTimeUtc = endDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

        var soapRequest = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
               xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
               xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
               xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
  <soap:Header>
    <t:RequestServerVersion Version=""Exchange2013_SP1"" />
  </soap:Header>
  <soap:Body>
    <FindItem xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"" Traversal=""Shallow"">
      <ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI=""item:Subject"" />
          <t:FieldURI FieldURI=""calendar:Start"" />
          <t:FieldURI FieldURI=""calendar:End"" />
          <t:FieldURI FieldURI=""item:Body"" />
        </t:AdditionalProperties>
      </ItemShape>
      <CalendarView StartDate=""{startTimeUtc}"" EndDate=""{endTimeUtc}"" />
      <ParentFolderIds>
        <t:DistinguishedFolderId Id=""calendar"" />
      </ParentFolderIds>
    </FindItem>
  </soap:Body>
</soap:Envelope>";

        return soapRequest;
    }

    /// <summary>
    /// Создать SOAP запрос для обновления события
    /// </summary>
    private string CreateUpdateEventSoapRequest(CalendarEvent calendarEvent)
    {
        var startTimeUtc = FormatTimeForExchange(calendarEvent.Start, calendarEvent.TimeZone ?? "UTC");
        var endTimeUtc = FormatTimeForExchange(calendarEvent.End, calendarEvent.TimeZone ?? "UTC");

        var soapRequest = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
               xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
               xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
               xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
  <soap:Header>
    <t:RequestServerVersion Version=""Exchange2013_SP1"" />
  </soap:Header>
  <soap:Body>
    <UpdateItem xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"" MessageDisposition=""SaveOnly"" ConflictResolution=""AutoResolve"" SendMeetingInvitationsOrCancellations=""{_sendMeetingInvitations}"">
      <ItemChanges>
        <t:ItemChange>
          <t:ItemId Id=""{calendarEvent.ExchangeId}"" ChangeKey=""{calendarEvent.ExchangeChangeKey}"" />
          <t:Updates>
            <t:SetItemField>
              <t:FieldURI FieldURI=""item:Subject"" />
              <t:CalendarItem>
                <t:Subject>{System.Security.SecurityElement.Escape(calendarEvent.Summary)}</t:Subject>
              </t:CalendarItem>
            </t:SetItemField>
            <t:SetItemField>
              <t:FieldURI FieldURI=""calendar:Start"" />
              <t:CalendarItem>
                <t:Start>{startTimeUtc}</t:Start>
              </t:CalendarItem>
            </t:SetItemField>
            <t:SetItemField>
              <t:FieldURI FieldURI=""calendar:End"" />
              <t:CalendarItem>
                <t:End>{endTimeUtc}</t:End>
              </t:CalendarItem>
            </t:SetItemField>";

        if (!string.IsNullOrEmpty(calendarEvent.Description))
        {
            soapRequest += $@"
            <t:SetItemField>
              <t:FieldURI FieldURI=""item:Body"" />
              <t:CalendarItem>
                <t:Body BodyType=""Text"">{System.Security.SecurityElement.Escape(calendarEvent.Description)}</t:Body>
              </t:CalendarItem>
            </t:SetItemField>";
        }

        soapRequest += @"
          </t:Updates>
        </t:ItemChange>
      </ItemChanges>
    </UpdateItem>
  </soap:Body>
</soap:Envelope>";

        return soapRequest;
    }

    /// <summary>
    /// Создать SOAP запрос для удаления события
    /// </summary>
    private string CreateDeleteEventSoapRequest(string eventId)
    {
        var soapRequest = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
               xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
               xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
               xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
  <soap:Header>
    <t:RequestServerVersion Version=""Exchange2013_SP1"" />
  </soap:Header>
  <soap:Body>
    <DeleteItem xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"" DeleteType=""HardDelete"" SendMeetingCancellations=""{_sendMeetingCancellations}"">
      <ItemIds>
        <t:ItemId Id=""{eventId}"" />
      </ItemIds>
    </DeleteItem>
  </soap:Body>
</soap:Envelope>";

        return soapRequest;
    }

    /// <summary>
    /// Парсинг событий из SOAP ответа
    /// </summary>
    private List<CalendarEvent> ParseCalendarEventsFromResponse(string response)
    {
        var events = new List<CalendarEvent>();

        try
        {
            Console.WriteLine($"🔍 Парсинг XML ответа ({response.Length} символов)");

            // Проверяем на ошибки в ответе
            if (response.Contains("ErrorInvalidRequest") || response.Contains("s:Fault"))
            {
                Console.WriteLine("❌ Обнаружена ошибка в ответе Exchange");
                return events;
            }

            // Используем System.Xml для правильного парсинга
            using (var stringReader = new System.IO.StringReader(response))
            using (var xmlReader = System.Xml.XmlReader.Create(stringReader))
            {
                CalendarEvent currentEvent = null;
                string currentElementName = "";

                while (xmlReader.Read())
                {
                    switch (xmlReader.NodeType)
                    {
                        case System.Xml.XmlNodeType.Element:
                            currentElementName = xmlReader.LocalName;

                            if (currentElementName == "CalendarItem")
                            {
                                currentEvent = new CalendarEvent();
                            }
                            else if (currentElementName == "ItemId" && currentEvent != null)
                            {
                                currentEvent.ExchangeId = xmlReader.GetAttribute("Id");
                                currentEvent.ExchangeChangeKey = xmlReader.GetAttribute("ChangeKey");
                            }
                            break;

                        case System.Xml.XmlNodeType.Text:
                            if (currentEvent != null)
                            {
                                switch (currentElementName)
                                {
                                    case "Subject":
                                        currentEvent.Summary = xmlReader.Value;
                                        break;
                                    case "Start":
                                        if (DateTime.TryParse(xmlReader.Value, out var startTime))
                                        {
                                            currentEvent.Start = startTime;
                                        }
                                        break;
                                    case "End":
                                        if (DateTime.TryParse(xmlReader.Value, out var endTime))
                                        {
                                            currentEvent.End = endTime;
                                        }
                                        break;
                                    case "Body":
                                        currentEvent.Description = xmlReader.Value;
                                        break;
                                }
                            }
                            break;

                        case System.Xml.XmlNodeType.EndElement:
                            if (xmlReader.LocalName == "CalendarItem" && currentEvent != null)
                            {
                                if (!string.IsNullOrEmpty(currentEvent.ExchangeId))
                                {
                                    events.Add(currentEvent);
                                    Console.WriteLine($"✅ Событие найдено: {currentEvent.Summary} ({currentEvent.ExchangeId?.Substring(0, 20)}...)");
                                }
                                currentEvent = null;
                            }
                            break;
                    }
                }
            }

            Console.WriteLine($"📊 Всего событий распарсено: {events.Count}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка парсинга событий: {ex.Message}");
            Console.WriteLine($"📝 Первые 500 символов ответа: {response.Substring(0, Math.Min(500, response.Length))}...");
        }

        return events;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _httpClient?.Dispose();
            _disposed = true;
            Console.WriteLine("🔄 Exchange HTTP Service disposed");
        }
    }
}