using System.Text;

namespace CalSync.Services;

/// <summary>
/// Сервис для скачивания ICS файлов
/// </summary>
public class IcsDownloader
{
    private readonly HttpClient _httpClient;

    public IcsDownloader(HttpClient? httpClient = null)
    {
        _httpClient = httpClient ?? new HttpClient();
        _httpClient.DefaultRequestHeaders.Add("User-Agent", "CalSync/1.0");
    }

    /// <summary>
    /// Скачать ICS файл по URL
    /// </summary>
    /// <param name="url">URL календаря</param>
    /// <param name="cancellationToken">Токен отмены</param>
    /// <returns>Содержимое ICS файла</returns>
    public async Task<string> DownloadAsync(string url, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(url))
            throw new ArgumentException("URL не может быть пустым", nameof(url));

        // Конвертируем webcal:// в https://
        var httpUrl = ConvertWebcalToHttps(url);

        try
        {
            using var response = await _httpClient.GetAsync(httpUrl, cancellationToken);
            response.EnsureSuccessStatusCode();

            var content = await response.Content.ReadAsStringAsync(cancellationToken);

            // Проверяем, что это действительно ICS файл
            if (!content.TrimStart().StartsWith("BEGIN:VCALENDAR", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Загруженный файл не является валидным ICS календарем");
            }

            return content;
        }
        catch (HttpRequestException ex)
        {
            throw new InvalidOperationException($"Ошибка при загрузке календаря: {ex.Message}", ex);
        }
        catch (TaskCanceledException ex)
        {
            throw new TimeoutException($"Превышено время ожидания при загрузке календаря: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Конвертировать webcal:// URL в https://
    /// </summary>
    private static string ConvertWebcalToHttps(string url)
    {
        if (url.StartsWith("webcal://", StringComparison.OrdinalIgnoreCase))
        {
            return "https://" + url.Substring(9);
        }
        return url;
    }

    public void Dispose()
    {
        _httpClient?.Dispose();
    }
}