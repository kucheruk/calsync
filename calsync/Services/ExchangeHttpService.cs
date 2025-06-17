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
    private bool _disposed = false;

    public ExchangeHttpService(IConfiguration configuration)
    {
        _configuration = configuration;
        var exchangeConfig = _configuration.GetSection("Exchange");

        _serviceUrl = exchangeConfig["ServiceUrl"] ?? throw new ArgumentException("ServiceUrl не настроен");
        _domain = exchangeConfig["Domain"] ?? "";
        _username = exchangeConfig["Username"] ?? throw new ArgumentException("Username не настроен");
        _password = exchangeConfig["Password"] ?? throw new ArgumentException("Password не настроен");

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
        try
        {
            Console.WriteLine($"➕ Создание события через SOAP: {calendarEvent.Summary}");

            // Создаем SOAP запрос для создания события
            var soapRequest = CreateEventSoapRequest(calendarEvent);

            Console.WriteLine("📤 Отправка SOAP запроса...");
            var response = await SendSoapRequestAsync(soapRequest);

            Console.WriteLine($"📥 Получен ответ от Exchange ({response.Length} символов)");

            // Пытаемся извлечь ID созданного события
            var eventId = ExtractEventIdFromResponse(response);

            if (!string.IsNullOrEmpty(eventId))
            {
                Console.WriteLine($"✅ Событие создано с ID: {eventId}");
                return eventId;
            }
            else if (response.Contains("Success"))
            {
                Console.WriteLine("✅ Событие создано успешно");
                return $"EXCHANGE_EVENT_{DateTime.Now:yyyyMMddHHmmss}";
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

        var response = await _httpClient.PostAsync(_serviceUrl, content);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            Console.WriteLine($"❌ HTTP ошибка: {response.StatusCode}");
            Console.WriteLine($"📝 Ответ: {responseContent}");
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

        // Если указана временная зона, обрабатываем ее
        if (!string.IsNullOrEmpty(timeZone))
        {
            try
            {
                // Преобразуем названия временных зон
                var systemTimeZoneId = ConvertIcsTimeZoneToSystem(timeZone);
                var tz = TimeZoneInfo.FindSystemTimeZoneById(systemTimeZoneId);

                // Предполагаем, что время в календарном событии уже в указанной временной зоне
                var utcTime = TimeZoneInfo.ConvertTimeToUtc(dateTime, tz);

                Console.WriteLine($"🌍 Конвертация времени: {dateTime:HH:mm:ss} ({timeZone}) → {utcTime:HH:mm:ss} UTC");

                return utcTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️  Ошибка обработки временной зоны '{timeZone}': {ex.Message}");
                // Fallback: добавляем московское время по умолчанию (UTC+3)
                var moscowTime = dateTime.AddHours(-3);
                return moscowTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            }
        }

        // По умолчанию считаем время локальным и конвертируем в UTC
        var utc = dateTime.ToUniversalTime();
        return utc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
    }

    /// <summary>
    /// Преобразовать ICS временную зону в системную
    /// </summary>
    private string ConvertIcsTimeZoneToSystem(string icsTimeZone)
    {
        return icsTimeZone switch
        {
            "Europe/Moscow" => "Russian Standard Time",
            "UTC" => "UTC",
            "GMT" => "GMT Standard Time",
            "Europe/London" => "GMT Standard Time",
            "America/New_York" => "Eastern Standard Time",
            "America/Los_Angeles" => "Pacific Standard Time",
            _ => "Russian Standard Time" // Fallback для нашего случая
        };
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
                SendMeetingInvitations=""SendToNone"">
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
            var doc = new XmlDocument();
            doc.LoadXml(response);

            var namespaceManager = new XmlNamespaceManager(doc.NameTable);
            namespaceManager.AddNamespace("t", "http://schemas.microsoft.com/exchange/services/2006/types");
            namespaceManager.AddNamespace("m", "http://schemas.microsoft.com/exchange/services/2006/messages");

            var itemIdNode = doc.SelectSingleNode("//t:ItemId", namespaceManager);

            return itemIdNode?.Attributes?["Id"]?.Value ?? "";
        }
        catch
        {
            // Если не удалось извлечь через XML, попробуем regex
            var match = System.Text.RegularExpressions.Regex.Match(response, @"Id=""([^""]+)""");
            return match.Success ? match.Groups[1].Value : "";
        }
    }

    /// <summary>
    /// Извлечь ошибку из ответа
    /// </summary>
    private string ExtractErrorFromResponse(string response)
    {
        try
        {
            var doc = new XmlDocument();
            doc.LoadXml(response);

            var namespaceManager = new XmlNamespaceManager(doc.NameTable);
            namespaceManager.AddNamespace("m", "http://schemas.microsoft.com/exchange/services/2006/messages");

            var errorNode = doc.SelectSingleNode("//m:MessageText", namespaceManager);

            return errorNode?.InnerText ?? "Неизвестная ошибка Exchange";
        }
        catch
        {
            return "Не удалось извлечь описание ошибки";
        }
    }

    /// <summary>
    /// Создать SOAP запрос для получения событий календаря
    /// </summary>
    private string CreateGetCalendarEventsSoapRequest(DateTime startDate, DateTime endDate)
    {
        var startTimeUtc = FormatTimeForExchange(startDate, "UTC");
        var endTimeUtc = FormatTimeForExchange(endDate, "UTC");

        var soapRequest = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
               xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
               xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
               xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
  <soap:Header>
    <t:RequestServerVersion Version=""Exchange2016_SP1"" />
  </soap:Header>
  <soap:Body>
    <FindItem Traversal=""Shallow"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"">
      <ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
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
    <t:RequestServerVersion Version=""Exchange2016_SP1"" />
  </soap:Header>
  <soap:Body>
    <UpdateItem MessageDisposition=""SaveOnly"" ConflictResolution=""AutoResolve"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"">
      <ItemChanges>
        <t:ItemChange>
          <t:ItemId Id=""{calendarEvent.ExchangeId}"" />
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
    <t:RequestServerVersion Version=""Exchange2016_SP1"" />
  </soap:Header>
  <soap:Body>
    <DeleteItem DeleteType=""HardDelete"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"">
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
            // Простой парсинг XML ответа без XmlDocument для избежания проблем
            var lines = response.Split('\n');
            CalendarEvent currentEvent = null;

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();

                if (trimmedLine.Contains("<t:CalendarItem>") || trimmedLine.Contains("<CalendarItem"))
                {
                    currentEvent = new CalendarEvent();
                }
                else if (currentEvent != null && (trimmedLine.Contains("</t:CalendarItem>") || trimmedLine.Contains("</CalendarItem")))
                {
                    if (currentEvent != null && !string.IsNullOrEmpty(currentEvent.ExchangeId))
                    {
                        events.Add(currentEvent);
                    }
                    currentEvent = null;
                }
                else if (currentEvent != null)
                {
                    // Парсим свойства события
                    if (trimmedLine.Contains("<t:ItemId") && trimmedLine.Contains("Id=\""))
                    {
                        var idStart = trimmedLine.IndexOf("Id=\"") + 4;
                        var idEnd = trimmedLine.IndexOf("\"", idStart);
                        if (idEnd > idStart)
                        {
                            currentEvent.ExchangeId = trimmedLine.Substring(idStart, idEnd - idStart);
                        }
                    }
                    else if (trimmedLine.StartsWith("<t:Subject>") && trimmedLine.EndsWith("</t:Subject>"))
                    {
                        currentEvent.Summary = trimmedLine.Replace("<t:Subject>", "").Replace("</t:Subject>", "").Trim();
                    }
                    else if (trimmedLine.StartsWith("<t:Start>") && trimmedLine.EndsWith("</t:Start>"))
                    {
                        var timeStr = trimmedLine.Replace("<t:Start>", "").Replace("</t:Start>", "").Trim();
                        if (DateTime.TryParse(timeStr, out var startTime))
                        {
                            currentEvent.Start = startTime;
                        }
                    }
                    else if (trimmedLine.StartsWith("<t:End>") && trimmedLine.EndsWith("</t:End>"))
                    {
                        var timeStr = trimmedLine.Replace("<t:End>", "").Replace("</t:End>", "").Trim();
                        if (DateTime.TryParse(timeStr, out var endTime))
                        {
                            currentEvent.End = endTime;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка парсинга событий: {ex.Message}");
            Console.WriteLine($"📝 Ответ: {response.Substring(0, Math.Min(500, response.Length))}...");
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