using System.Net;
using System.Text;

namespace CalSync.Tests.TestHelpers;

/// <summary>
/// Конфигурация для тестового HTTP сервера
/// </summary>
public class TestHttpServerConfig
{
    public int Port { get; set; } = 8765;
    public string Host { get; set; } = "localhost";
    public Dictionary<string, string> Routes { get; set; } = new();
    public Dictionary<string, HttpStatusCode> StatusCodes { get; set; } = new();
    public Dictionary<string, string> Headers { get; set; } = new();
    public TimeSpan? ResponseDelay { get; set; }
    public bool RequireAuth { get; set; } = false;
    public string? Username { get; set; }
    public string? Password { get; set; }
}

/// <summary>
/// Параметризуемый HTTP сервер для тестирования различных сценариев
/// </summary>
public class TestHttpServer : IDisposable
{
    private readonly HttpListener? _httpListener;
    private readonly TestHttpServerConfig _config;
    private readonly CancellationTokenSource _cancellationTokenSource;
    private Task? _serverTask;

    public string BaseUrl { get; }
    public bool IsRunning => _httpListener?.IsListening ?? false;

    public TestHttpServer(TestHttpServerConfig config)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _cancellationTokenSource = new CancellationTokenSource();
        BaseUrl = $"http://{_config.Host}:{_config.Port}/";

        try
        {
            _httpListener = new HttpListener();
            _httpListener.Prefixes.Add(BaseUrl);
        }
        catch (PlatformNotSupportedException)
        {
            // HttpListener не поддерживается на этой платформе
            _httpListener = null;
        }
    }

    /// <summary>
    /// Создает сервер с простой конфигурацией
    /// </summary>
    public static TestHttpServer CreateSimple(int port = 8765, string defaultContent = "")
    {
        var config = new TestHttpServerConfig
        {
            Port = port,
            Routes = new Dictionary<string, string>
            {
                { "/", defaultContent.IsNullOrEmpty() ? IcsTestDataGenerator.GenerateSimpleCalendar() : defaultContent }
            }
        };
        return new TestHttpServer(config);
    }

    /// <summary>
    /// Создает сервер для тестирования аутентификации
    /// </summary>
    public static TestHttpServer CreateWithAuth(string username, string password, int port = 8766)
    {
        var config = new TestHttpServerConfig
        {
            Port = port,
            RequireAuth = true,
            Username = username,
            Password = password,
            Routes = new Dictionary<string, string>
            {
                { "/protected", IcsTestDataGenerator.GenerateSimpleCalendar("Protected Event") }
            }
        };
        return new TestHttpServer(config);
    }

    /// <summary>
    /// Создает сервер для тестирования различных HTTP статусов
    /// </summary>
    public static TestHttpServer CreateWithStatusCodes(Dictionary<string, HttpStatusCode> statusCodes, int port = 8767)
    {
        var config = new TestHttpServerConfig
        {
            Port = port,
            StatusCodes = statusCodes,
            Routes = statusCodes.ToDictionary(
                kvp => kvp.Key,
                kvp => kvp.Value == HttpStatusCode.OK ? IcsTestDataGenerator.GenerateSimpleCalendar() : "Error"
            )
        };
        return new TestHttpServer(config);
    }

    /// <summary>
    /// Запускает HTTP сервер
    /// </summary>
    public async Task StartAsync()
    {
        if (_httpListener == null)
            throw new PlatformNotSupportedException("HttpListener не поддерживается на этой платформе");

        try
        {
            _httpListener.Start();
            _serverTask = Task.Run(HandleRequests, _cancellationTokenSource.Token);

            // Ждем немного чтобы сервер успел запуститься
            await Task.Delay(100);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Не удалось запустить HTTP сервер на {BaseUrl}: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Останавливает HTTP сервер
    /// </summary>
    public void Stop()
    {
        if (!_cancellationTokenSource.IsCancellationRequested)
        {
            _cancellationTokenSource.Cancel();
        }
        _httpListener?.Stop();
        _serverTask?.Wait(TimeSpan.FromSeconds(5));
    }

    /// <summary>
    /// Добавляет маршрут с контентом
    /// </summary>
    public void AddRoute(string path, string content, HttpStatusCode statusCode = HttpStatusCode.OK)
    {
        _config.Routes[path] = content;
        _config.StatusCodes[path] = statusCode;
    }

    /// <summary>
    /// Обрабатывает HTTP запросы
    /// </summary>
    private async Task HandleRequests()
    {
        if (_httpListener == null) return;

        while (_httpListener.IsListening && !_cancellationTokenSource.Token.IsCancellationRequested)
        {
            try
            {
                var context = await _httpListener.GetContextAsync();
                _ = Task.Run(() => ProcessRequest(context), _cancellationTokenSource.Token);
            }
            catch (HttpListenerException)
            {
                // Ожидается при остановке сервера
                break;
            }
            catch (ObjectDisposedException)
            {
                // Ожидается при освобождении ресурсов  
                break;
            }
        }
    }

    /// <summary>
    /// Обрабатывает отдельный HTTP запрос
    /// </summary>
    private async Task ProcessRequest(HttpListenerContext context)
    {
        try
        {
            var request = context.Request;
            var response = context.Response;
            var path = request.Url?.AbsolutePath ?? "/";

            // Задержка если настроена
            if (_config.ResponseDelay.HasValue)
            {
                await Task.Delay(_config.ResponseDelay.Value);
            }

            // Проверка аутентификации
            if (_config.RequireAuth && !IsAuthorized(request))
            {
                response.StatusCode = (int)HttpStatusCode.Unauthorized;
                response.Headers.Add("WWW-Authenticate", "Basic realm=\"Test\"");
                response.Close();
                return;
            }

            // Определение контента и статуса
            var content = GetContentForPath(path);
            var statusCode = _config.StatusCodes.GetValueOrDefault(path, HttpStatusCode.OK);

            // Установка заголовков
            foreach (var header in _config.Headers)
            {
                response.Headers.Add(header.Key, header.Value);
            }

            // Отправка ответа
            response.StatusCode = (int)statusCode;
            response.ContentType = "text/calendar; charset=utf-8";

            var buffer = Encoding.UTF8.GetBytes(content);
            response.ContentLength64 = buffer.Length;

            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.OutputStream.Close();
        }
        catch (Exception)
        {
            // Игнорируем ошибки обработки отдельных запросов
            try { context.Response.Close(); } catch { }
        }
    }

    /// <summary>
    /// Проверяет авторизацию Basic Auth
    /// </summary>
    private bool IsAuthorized(HttpListenerRequest request)
    {
        var authHeader = request.Headers["Authorization"];
        if (string.IsNullOrEmpty(authHeader) || !authHeader.StartsWith("Basic "))
            return false;

        try
        {
            var encoded = authHeader.Substring("Basic ".Length);
            var decoded = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
            var parts = decoded.Split(':');

            return parts.Length == 2 &&
                   parts[0] == _config.Username &&
                   parts[1] == _config.Password;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Получает контент для указанного пути
    /// </summary>
    private string GetContentForPath(string path)
    {
        // Точное совпадение
        if (_config.Routes.TryGetValue(path, out var content))
            return content;

        // Ищем частичное совпадение или wildcard
        var matchingRoute = _config.Routes.Keys
            .Where(route => path.StartsWith(route) || route == "*")
            .OrderByDescending(route => route.Length)
            .FirstOrDefault();

        if (matchingRoute != null)
            return _config.Routes[matchingRoute];

        // По умолчанию возвращаем простой календарь
        return IcsTestDataGenerator.GenerateSimpleCalendar();
    }

    public void Dispose()
    {
        try
        {
            Stop();
        }
        catch (ObjectDisposedException)
        {
            // Ignore already disposed
        }

        try
        {
            _cancellationTokenSource?.Dispose();
        }
        catch (ObjectDisposedException)
        {
            // Ignore already disposed
        }

        _httpListener?.Close();
    }
}

/// <summary>
/// Расширения для работы со строками
/// </summary>
internal static class StringExtensions
{
    public static bool IsNullOrEmpty(this string? value) => string.IsNullOrEmpty(value);
}