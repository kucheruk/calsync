using Xunit;

namespace CalSync.Tests;

/// <summary>
/// Пример использования EwsRequestRecorder для записи и генерации тестов
/// </summary>
public class EwsRecordingExample
{
    [Fact]
    public void EwsRequestRecorder_Example_ShouldRecordAndGenerateTests()
    {
        // Arrange
        var recorder = new EwsRequestRecorder("test_recordings.json");

        // Очищаем предыдущие записи для чистого теста
        recorder.ClearRecordings();

        // Пример записи GetFolder запроса
        var getFolderRequest = @"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">
  <soap:Body>
    <m:GetFolder xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"">
      <m:FolderShape>
        <t:BaseShape xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">AllProperties</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"" Id=""calendar"" />
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>";

        var getFolderResponse = @"<?xml version=""1.0"" encoding=""utf-8""?>
<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
  <s:Body>
    <m:GetFolderResponse xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"">
      <m:ResponseMessages>
        <m:GetFolderResponseMessage ResponseClass=""Success"">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:Folders>
            <t:CalendarFolder xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
              <t:FolderId Id=""CALENDAR123"" ChangeKey=""CHANGE123""/>
              <t:DisplayName>Календарь</t:DisplayName>
              <t:TotalCount>5</t:TotalCount>
            </t:CalendarFolder>
          </m:Folders>
        </m:GetFolderResponseMessage>
      </m:ResponseMessages>
    </m:GetFolderResponse>
  </s:Body>
</s:Envelope>";

        // Act - записываем запрос
        recorder.RecordRequest("GetFolder", getFolderRequest, getFolderResponse,
            "Получение календаря пользователя");

        // Пример записи FindItem запроса
        var findItemRequest = @"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">
  <soap:Body>
    <m:FindItem xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"" Traversal=""Shallow"">
      <m:ItemShape>
        <t:BaseShape xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">AllProperties</t:BaseShape>
      </m:ItemShape>
      <m:CalendarView StartDate=""2025-06-17T00:00:00.000Z"" EndDate=""2025-06-18T00:00:00.000Z"" />
      <m:ParentFolderIds>
        <t:FolderId xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"" Id=""CALENDAR123"" />
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>";

        var findItemResponse = @"<?xml version=""1.0"" encoding=""utf-8""?>
<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
  <s:Body>
    <m:FindItemResponse xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"">
      <m:ResponseMessages>
        <m:FindItemResponseMessage ResponseClass=""Success"">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:RootFolder TotalItemsInView=""2"" IncludesLastItemInRange=""true"">
            <t:Items xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
              <t:CalendarItem>
                <t:ItemId Id=""EVENT001"" ChangeKey=""CHANGE001""/>
                <t:Subject>Тестовое событие 1</t:Subject>
                <t:Start>2025-06-17T10:00:00Z</t:Start>
                <t:End>2025-06-17T11:00:00Z</t:End>
              </t:CalendarItem>
              <t:CalendarItem>
                <t:ItemId Id=""EVENT002"" ChangeKey=""CHANGE002""/>
                <t:Subject>Тестовое событие 2</t:Subject>
                <t:Start>2025-06-17T14:00:00Z</t:Start>
                <t:End>2025-06-17T15:00:00Z</t:End>
              </t:CalendarItem>
            </t:Items>
          </m:RootFolder>
        </m:FindItemResponseMessage>
      </m:ResponseMessages>
    </m:FindItemResponse>
  </s:Body>
</s:Envelope>";

        recorder.RecordRequest("FindItem", findItemRequest, findItemResponse,
            "Поиск событий в календаре за период");

        // Assert - проверяем записи
        var allRecordings = recorder.GetAllRecordings();
        Assert.Equal(2, allRecordings.Count);

        var getFolderRecording = recorder.GetLatestRecording("GetFolder");
        Assert.NotNull(getFolderRecording);
        Assert.Equal("GetFolder", getFolderRecording.Action);
        Assert.Contains("GetFolder", getFolderRecording.RequestXml);
        Assert.Contains("GetFolderResponse", getFolderRecording.ResponseXml);

        var findItemRecording = recorder.GetLatestRecording("FindItem");
        Assert.NotNull(findItemRecording);
        Assert.Equal("FindItem", findItemRecording.Action);
        Assert.Contains("FindItem", findItemRecording.RequestXml);
        Assert.Contains("FindItemResponse", findItemRecording.ResponseXml);

        // Генерируем тесты на основе записей
        var generatedTests = recorder.GenerateTests("ExampleGeneratedTests");

        // Проверяем, что тесты сгенерированы
        Assert.NotNull(generatedTests);
        Assert.Contains("TestGetFolder", generatedTests);
        Assert.Contains("TestFindItem", generatedTests);
        Assert.Contains("ExampleGeneratedTests", generatedTests);

        // Сохраняем сгенерированные тесты в файл для демонстрации
        recorder.SaveGeneratedTests("GeneratedEwsTests_Example.cs", "ExampleGeneratedTests");

        Console.WriteLine("✅ Пример записи и генерации EWS тестов выполнен успешно");
        Console.WriteLine($"📊 Записано запросов: {allRecordings.Count}");
        Console.WriteLine("📄 Сгенерированные тесты сохранены в GeneratedEwsTests_Example.cs");
    }

    [Fact]
    public void EwsRequestRecorder_LoadExistingRecordings_ShouldWork()
    {
        // Arrange - создаем recorder с тем же файлом
        var recorder1 = new EwsRequestRecorder("test_recordings.json");

        // Записываем тестовые данные
        recorder1.RecordRequest("CreateItem", "<request>test</request>", "<response>success</response>");

        // Act - создаем новый recorder с тем же файлом
        var recorder2 = new EwsRequestRecorder("test_recordings.json");

        // Assert - проверяем, что данные загрузились
        var recordings = recorder2.GetAllRecordings();
        Assert.True(recordings.Count > 0);

        var createItemRecording = recorder2.GetLatestRecording("CreateItem");
        Assert.NotNull(createItemRecording);
        Assert.Equal("CreateItem", createItemRecording.Action);

        Console.WriteLine("✅ Загрузка существующих записей работает корректно");
    }
}