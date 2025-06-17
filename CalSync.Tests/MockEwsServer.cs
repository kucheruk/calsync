using System.Net;
using System.Text;
using System.Xml.Linq;

namespace CalSync.Tests;

/// <summary>
/// Мок-сервер для имитации Exchange Web Services
/// </summary>
public class MockEwsServer : IDisposable
{
    private readonly HttpListener _listener;
    private readonly string _baseUrl;
    private readonly CancellationTokenSource _cancellationTokenSource;
    private Task? _serverTask;
    private readonly Dictionary<string, Func<XDocument, string>> _requestHandlers;

    public MockEwsServer(int port = 8080)
    {
        _baseUrl = $"http://localhost:{port}/EWS/Exchange.asmx";
        _listener = new HttpListener();
        _listener.Prefixes.Add($"http://localhost:{port}/");
        _cancellationTokenSource = new CancellationTokenSource();
        _requestHandlers = new Dictionary<string, Func<XDocument, string>>();

        SetupDefaultHandlers();
    }

    public string ServiceUrl => _baseUrl;

    /// <summary>
    /// Запустить мок-сервер
    /// </summary>
    public void Start()
    {
        _listener.Start();
        _serverTask = Task.Run(async () => await HandleRequestsAsync(_cancellationTokenSource.Token));
        Console.WriteLine($"🚀 Mock EWS Server запущен на {_baseUrl}");
    }

    /// <summary>
    /// Остановить мок-сервер
    /// </summary>
    public void Stop()
    {
        _cancellationTokenSource.Cancel();
        _listener.Stop();
        _serverTask?.Wait(TimeSpan.FromSeconds(5));
        Console.WriteLine("🛑 Mock EWS Server остановлен");
    }

    /// <summary>
    /// Настройка обработчиков по умолчанию
    /// </summary>
    private void SetupDefaultHandlers()
    {
        // Обработчик для получения календаря (Folder.Bind)
        _requestHandlers["GetFolder"] = HandleGetFolder;

        // Обработчик для поиска событий (FindAppointments)
        _requestHandlers["FindItem"] = HandleFindItem;

        // Обработчик для загрузки деталей события (Load)
        _requestHandlers["GetItem"] = HandleGetItem;

        // Обработчик для создания события
        _requestHandlers["CreateItem"] = HandleCreateItem;

        // Обработчик для удаления события
        _requestHandlers["DeleteItem"] = HandleDeleteItem;
    }

    /// <summary>
    /// Обработка HTTP запросов
    /// </summary>
    private async Task HandleRequestsAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                var context = await _listener.GetContextAsync();
                _ = Task.Run(() => ProcessRequest(context), cancellationToken);
            }
            catch (HttpListenerException)
            {
                // Сервер остановлен
                break;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Ошибка в Mock EWS Server: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Обработка отдельного запроса
    /// </summary>
    private void ProcessRequest(HttpListenerContext context)
    {
        try
        {
            var request = context.Request;
            var response = context.Response;

            // Читаем тело запроса
            string requestBody;
            using (var reader = new StreamReader(request.InputStream, Encoding.UTF8))
            {
                requestBody = reader.ReadToEnd();
            }

            Console.WriteLine($"📨 EWS Request: {request.HttpMethod} {request.Url}");
            Console.WriteLine($"📄 Request Body: {requestBody}");

            // Парсим SOAP запрос
            var soapDoc = XDocument.Parse(requestBody);
            var soapAction = GetSoapAction(soapDoc);

            // Получаем обработчик для действия
            var responseXml = _requestHandlers.ContainsKey(soapAction)
                ? _requestHandlers[soapAction](soapDoc)
                : CreateErrorResponse("Неподдерживаемое действие");

            Console.WriteLine($"📤 EWS Response: {responseXml}");

            // Отправляем ответ
            var responseBytes = Encoding.UTF8.GetBytes(responseXml);
            response.ContentType = "text/xml; charset=utf-8";
            response.ContentLength64 = responseBytes.Length;
            response.StatusCode = 200;

            response.OutputStream.Write(responseBytes, 0, responseBytes.Length);
            response.OutputStream.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка обработки запроса: {ex.Message}");
            try
            {
                context.Response.StatusCode = 500;
                context.Response.Close();
            }
            catch { }
        }
    }

    /// <summary>
    /// Извлечь SOAP действие из запроса
    /// </summary>
    private string GetSoapAction(XDocument soapDoc)
    {
        var ns = XNamespace.Get("http://schemas.microsoft.com/exchange/services/2006/messages");

        if (soapDoc.Descendants(ns + "GetFolder").Any()) return "GetFolder";
        if (soapDoc.Descendants(ns + "FindItem").Any()) return "FindItem";
        if (soapDoc.Descendants(ns + "GetItem").Any()) return "GetItem";
        if (soapDoc.Descendants(ns + "CreateItem").Any()) return "CreateItem";
        if (soapDoc.Descendants(ns + "DeleteItem").Any()) return "DeleteItem";

        return "Unknown";
    }

    /// <summary>
    /// Обработчик получения календаря
    /// </summary>
    private string HandleGetFolder(XDocument request)
    {
        return @"<?xml version=""1.0"" encoding=""utf-8""?>
<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
    <s:Header>
        <h:ServerVersionInfo MajorVersion=""15"" MinorVersion=""0"" MajorBuildNumber=""1497"" MinorBuildNumber=""2"" Version=""V2016_SP1"" xmlns:h=""http://schemas.microsoft.com/exchange/services/2006/types"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/types"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>
    </s:Header>
    <s:Body xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
        <m:GetFolderResponse xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"" xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
            <m:ResponseMessages>
                <m:GetFolderResponseMessage ResponseClass=""Success"">
                    <m:ResponseCode>NoError</m:ResponseCode>
                    <m:Folders>
                        <t:CalendarFolder>
                            <t:FolderId Id=""MOCK_CALENDAR_ID_123456789ABCDEF"" ChangeKey=""MOCK_CHANGE_KEY_123""/>
                            <t:DisplayName>Календарь</t:DisplayName>
                            <t:TotalCount>3</t:TotalCount>
                            <t:ChildFolderCount>0</t:ChildFolderCount>
                        </t:CalendarFolder>
                    </m:Folders>
                </m:GetFolderResponseMessage>
            </m:ResponseMessages>
        </m:GetFolderResponse>
    </s:Body>
</s:Envelope>";
    }

    /// <summary>
    /// Обработчик поиска событий
    /// </summary>
    private string HandleFindItem(XDocument request)
    {
        return @"<?xml version=""1.0"" encoding=""utf-8""?>
<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
    <s:Header>
        <h:ServerVersionInfo MajorVersion=""15"" MinorVersion=""0"" MajorBuildNumber=""1497"" MinorBuildNumber=""2"" Version=""V2016_SP1"" xmlns:h=""http://schemas.microsoft.com/exchange/services/2006/types""/>
    </s:Header>
    <s:Body>
        <m:FindItemResponse xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"" xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
            <m:ResponseMessages>
                <m:FindItemResponseMessage ResponseClass=""Success"">
                    <m:ResponseCode>NoError</m:ResponseCode>
                    <m:RootFolder TotalItemsInView=""3"" IncludesLastItemInRange=""true"">
                        <t:Items>
                            <t:CalendarItem>
                                <t:ItemId Id=""TEST001"" ChangeKey=""DwAAABYAAAA""/>
                                <t:Subject>[TEST] CalSync тестовое событие 1</t:Subject>
                                <t:Start>2025-06-17T10:00:00Z</t:Start>
                                <t:End>2025-06-17T11:00:00Z</t:End>
                                <t:Location>Тестовая локация 1</t:Location>
                                <t:LegacyFreeBusyStatus>Busy</t:LegacyFreeBusyStatus>
                            </t:CalendarItem>
                            <t:CalendarItem>
                                <t:ItemId Id=""TEST002"" ChangeKey=""DwAAABYAAAA""/>
                                <t:Subject>Обычное событие</t:Subject>
                                <t:Start>2025-06-17T14:00:00Z</t:Start>
                                <t:End>2025-06-17T15:00:00Z</t:End>
                                <t:LegacyFreeBusyStatus>Busy</t:LegacyFreeBusyStatus>
                            </t:CalendarItem>
                            <t:CalendarItem>
                                <t:ItemId Id=""TEST003"" ChangeKey=""DwAAABYAAAA""/>
                                <t:Subject>[TEST] CalSync тестовое событие 2</t:Subject>
                                <t:Start>2025-06-18T09:00:00Z</t:Start>
                                <t:End>2025-06-18T10:00:00Z</t:End>
                                <t:Location>Тестовая локация 2</t:Location>
                                <t:LegacyFreeBusyStatus>Tentative</t:LegacyFreeBusyStatus>
                            </t:CalendarItem>
                        </t:Items>
                    </m:RootFolder>
                </m:FindItemResponseMessage>
            </m:ResponseMessages>
        </m:FindItemResponse>
    </s:Body>
</s:Envelope>";
    }

    /// <summary>
    /// Обработчик получения деталей события
    /// </summary>
    private string HandleGetItem(XDocument request)
    {
        return @"<?xml version=""1.0"" encoding=""utf-8""?>
<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
    <s:Header>
        <h:ServerVersionInfo MajorVersion=""15"" MinorVersion=""0"" MajorBuildNumber=""1497"" MinorBuildNumber=""2"" Version=""V2016_SP1"" xmlns:h=""http://schemas.microsoft.com/exchange/services/2006/types""/>
    </s:Header>
    <s:Body>
        <m:GetItemResponse xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"" xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
            <m:ResponseMessages>
                <m:GetItemResponseMessage ResponseClass=""Success"">
                    <m:ResponseCode>NoError</m:ResponseCode>
                    <m:Items>
                        <t:CalendarItem>
                            <t:ItemId Id=""TEST001"" ChangeKey=""DwAAABYAAAA""/>
                            <t:Subject>[TEST] CalSync тестовое событие 1</t:Subject>
                            <t:Body BodyType=""Text"">Тестовое событие для проверки CalSync

[CalSync-Test-Event-MOCK]</t:Body>
                            <t:Start>2025-06-17T10:00:00Z</t:Start>
                            <t:End>2025-06-17T11:00:00Z</t:End>
                            <t:Location>Тестовая локация 1</t:Location>
                            <t:LegacyFreeBusyStatus>Busy</t:LegacyFreeBusyStatus>
                            <t:LastModifiedTime>2025-06-17T13:15:00Z</t:LastModifiedTime>
                        </t:CalendarItem>
                    </m:Items>
                </m:GetItemResponseMessage>
            </m:ResponseMessages>
        </m:GetItemResponse>
    </s:Body>
</s:Envelope>";
    }

    /// <summary>
    /// Обработчик создания события
    /// </summary>
    private string HandleCreateItem(XDocument request)
    {
        var newId = "TEST" + DateTime.Now.Ticks.ToString().Substring(10);
        return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
    <s:Header>
        <h:ServerVersionInfo MajorVersion=""15"" MinorVersion=""0"" MajorBuildNumber=""1497"" MinorBuildNumber=""2"" Version=""V2016_SP1"" xmlns:h=""http://schemas.microsoft.com/exchange/services/2006/types""/>
    </s:Header>
    <s:Body>
        <m:CreateItemResponse xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"" xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
            <m:ResponseMessages>
                <m:CreateItemResponseMessage ResponseClass=""Success"">
                    <m:ResponseCode>NoError</m:ResponseCode>
                    <m:Items>
                        <t:CalendarItem>
                            <t:ItemId Id=""{newId}"" ChangeKey=""DwAAABYAAAA""/>
                        </t:CalendarItem>
                    </m:Items>
                </m:CreateItemResponseMessage>
            </m:ResponseMessages>
        </m:CreateItemResponse>
    </s:Body>
</s:Envelope>";
    }

    /// <summary>
    /// Обработчик удаления события
    /// </summary>
    private string HandleDeleteItem(XDocument request)
    {
        return @"<?xml version=""1.0"" encoding=""utf-8""?>
<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
    <s:Header>
        <h:ServerVersionInfo MajorVersion=""15"" MinorVersion=""0"" MajorBuildNumber=""1497"" MinorBuildNumber=""2"" Version=""V2016_SP1"" xmlns:h=""http://schemas.microsoft.com/exchange/services/2006/types""/>
    </s:Header>
    <s:Body>
        <m:DeleteItemResponse xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"" xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
            <m:ResponseMessages>
                <m:DeleteItemResponseMessage ResponseClass=""Success"">
                    <m:ResponseCode>NoError</m:ResponseCode>
                </m:DeleteItemResponseMessage>
            </m:ResponseMessages>
        </m:DeleteItemResponse>
    </s:Body>
</s:Envelope>";
    }

    /// <summary>
    /// Создать ответ об ошибке
    /// </summary>
    private string CreateErrorResponse(string errorMessage)
    {
        return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
    <s:Body>
        <s:Fault>
            <faultcode>s:Server</faultcode>
            <faultstring>{errorMessage}</faultstring>
        </s:Fault>
    </s:Body>
</s:Envelope>";
    }

    public void Dispose()
    {
        Stop();
        _listener?.Close();
        _cancellationTokenSource?.Dispose();
    }
}