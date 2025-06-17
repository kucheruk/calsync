using System.Text.Json;
using System.Xml.Linq;

namespace CalSync.Tests;

/// <summary>
/// Класс для записи и воспроизведения EWS запросов и ответов
/// </summary>
public class EwsRequestRecorder
{
    private readonly List<EwsRequestResponse> _recordings = new();
    private readonly string _recordingsPath;

    public EwsRequestRecorder(string recordingsPath = "ews_recordings.json")
    {
        _recordingsPath = recordingsPath;
        LoadRecordings();
    }

    /// <summary>
    /// Записать EWS запрос и ответ
    /// </summary>
    public void RecordRequest(string action, string requestXml, string responseXml, string? description = null)
    {
        var recording = new EwsRequestResponse
        {
            Action = action,
            RequestXml = FormatXml(requestXml),
            ResponseXml = FormatXml(responseXml),
            Description = description ?? $"EWS {action} request",
            Timestamp = DateTime.UtcNow
        };

        _recordings.Add(recording);
        SaveRecordings();

        Console.WriteLine($"📼 Записан EWS запрос: {action}");
    }

    /// <summary>
    /// Получить все записи для действия
    /// </summary>
    public List<EwsRequestResponse> GetRecordings(string action)
    {
        return _recordings.Where(r => r.Action.Equals(action, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    /// <summary>
    /// Получить последнюю запись для действия
    /// </summary>
    public EwsRequestResponse? GetLatestRecording(string action)
    {
        return _recordings
            .Where(r => r.Action.Equals(action, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(r => r.Timestamp)
            .FirstOrDefault();
    }

    /// <summary>
    /// Получить все записи
    /// </summary>
    public List<EwsRequestResponse> GetAllRecordings()
    {
        return _recordings.ToList();
    }

    /// <summary>
    /// Очистить все записи
    /// </summary>
    public void ClearRecordings()
    {
        _recordings.Clear();
        SaveRecordings();
        Console.WriteLine("🗑️  Все записи EWS запросов очищены");
    }

    /// <summary>
    /// Сгенерировать тесты на основе записей
    /// </summary>
    public string GenerateTests(string testClassName = "GeneratedEwsTests")
    {
        var testCode = $@"using Xunit;
using System.Xml.Linq;

namespace CalSync.Tests;

/// <summary>
/// Автоматически сгенерированные тесты на основе записанных EWS запросов
/// Сгенерировано: {DateTime.Now:yyyy-MM-dd HH:mm:ss}
/// </summary>
public class {testClassName}
{{";

        var actionGroups = _recordings.GroupBy(r => r.Action);
        var testIndex = 1;

        foreach (var group in actionGroups)
        {
            var action = group.Key;
            var recordings = group.ToList();

            for (int i = 0; i < recordings.Count; i++)
            {
                var recording = recordings[i];
                var testName = recordings.Count > 1
                    ? $"Test{action}_{i + 1:D2}"
                    : $"Test{action}";

                testCode += $@"

    [Fact]
    public void {testName}()
    {{
        // Arrange - {recording.Description}
        var expectedRequest = @""{EscapeString(recording.RequestXml)}"";
        var expectedResponse = @""{EscapeString(recording.ResponseXml)}"";

        // Act - парсим XML для проверки корректности
        var requestDoc = XDocument.Parse(expectedRequest);
        var responseDoc = XDocument.Parse(expectedResponse);

        // Assert - проверяем структуру
        Assert.NotNull(requestDoc.Root);
        Assert.NotNull(responseDoc.Root);
        
        // Проверяем, что это SOAP сообщения
        Assert.Equal(""Envelope"", requestDoc.Root.Name.LocalName);
        Assert.Equal(""Envelope"", responseDoc.Root.Name.LocalName);
        
        // Проверяем наличие Body
        var requestBody = requestDoc.Descendants().FirstOrDefault(x => x.Name.LocalName == ""Body"");
        var responseBody = responseDoc.Descendants().FirstOrDefault(x => x.Name.LocalName == ""Body"");
        
        Assert.NotNull(requestBody);
        Assert.NotNull(responseBody);
        
        // Проверяем наличие EWS действия в запросе
        var ewsAction = requestBody.Descendants().FirstOrDefault(x => x.Name.LocalName == ""{action}"");
        Assert.NotNull(ewsAction);
        
        Console.WriteLine(""✅ Тест {testName} прошел успешно"");
    }}";
                testIndex++;
            }
        }

        testCode += @"
}";

        return testCode;
    }

    /// <summary>
    /// Сохранить сгенерированные тесты в файл
    /// </summary>
    public void SaveGeneratedTests(string filePath, string testClassName = "GeneratedEwsTests")
    {
        var testCode = GenerateTests(testClassName);
        File.WriteAllText(filePath, testCode);
        Console.WriteLine($"💾 Сгенерированные тесты сохранены в {filePath}");
    }

    /// <summary>
    /// Загрузить записи из файла
    /// </summary>
    private void LoadRecordings()
    {
        if (File.Exists(_recordingsPath))
        {
            try
            {
                var json = File.ReadAllText(_recordingsPath);
                var recordings = JsonSerializer.Deserialize<List<EwsRequestResponse>>(json);
                if (recordings != null)
                {
                    _recordings.AddRange(recordings);
                    Console.WriteLine($"📂 Загружено {recordings.Count} записей EWS запросов");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️  Ошибка загрузки записей: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Сохранить записи в файл
    /// </summary>
    private void SaveRecordings()
    {
        try
        {
            var json = JsonSerializer.Serialize(_recordings, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            File.WriteAllText(_recordingsPath, json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️  Ошибка сохранения записей: {ex.Message}");
        }
    }

    /// <summary>
    /// Форматировать XML для лучшей читаемости
    /// </summary>
    private string FormatXml(string xml)
    {
        try
        {
            var doc = XDocument.Parse(xml);
            return doc.ToString();
        }
        catch
        {
            return xml; // Возвращаем как есть, если не удалось распарсить
        }
    }

    /// <summary>
    /// Экранировать строку для использования в коде
    /// </summary>
    private string EscapeString(string str)
    {
        return str.Replace("\"", "\"\"");
    }
}

/// <summary>
/// Модель для хранения EWS запроса и ответа
/// </summary>
public class EwsRequestResponse
{
    public string Action { get; set; } = string.Empty;
    public string RequestXml { get; set; } = string.Empty;
    public string ResponseXml { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public DateTime Timestamp { get; set; }
}