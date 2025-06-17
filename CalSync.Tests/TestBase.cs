using Microsoft.Extensions.Configuration;

namespace CalSync.Tests;

/// <summary>
/// Базовый класс для тестов с поддержкой конфигурации
/// </summary>
public abstract class TestBase
{
    protected IConfiguration Configuration { get; }

    protected TestBase()
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .AddJsonFile("appsettings.Development.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables();

        Configuration = builder.Build();
    }

    /// <summary>
    /// Получить URL из конфигурации
    /// </summary>
    protected string GetTestUrl(string key)
    {
        return Configuration[$"TestConfiguration:IcsUrls:{key}"] ??
               throw new InvalidOperationException($"URL с ключом '{key}' не найден в конфигурации");
    }

    /// <summary>
    /// Получить настройки аутентификации
    /// </summary>
    protected (string username, string password) GetTestCredentials()
    {
        var username = Configuration["TestConfiguration:Authentication:Username"] ?? "testuser";
        var password = Configuration["TestConfiguration:Authentication:Password"] ?? "testpass";
        return (username, password);
    }

    /// <summary>
    /// Получить настройки таймаутов
    /// </summary>
    protected (int timeoutMs, int maxRetries, int baseDelayMs) GetTestTimeouts()
    {
        var timeoutMs = Configuration.GetValue<int>("TestConfiguration:Timeouts:DefaultTimeoutMs", 5000);
        var maxRetries = Configuration.GetValue<int>("TestConfiguration:Timeouts:MaxRetries", 3);
        var baseDelayMs = Configuration.GetValue<int>("TestConfiguration:Timeouts:BaseDelayMs", 1000);
        return (timeoutMs, maxRetries, baseDelayMs);
    }
}