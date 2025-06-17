using CalSync.Services;
using CalSync.Models;
using Microsoft.Extensions.Configuration;

namespace CalSync;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("CalSync - Синхронизация календарей ICS ↔ Exchange");
        Console.WriteLine("================================================");

        try
        {
            // Загружаем конфигурацию
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true)
                .AddJsonFile("appsettings.Local.json", optional: true)
                .Build();

            // Получаем URL календаря из конфигурации или аргументов
            var icsUrl = configuration["IcsUrl"] ?? args.FirstOrDefault();
            if (string.IsNullOrEmpty(icsUrl))
            {
                Console.WriteLine("❌ Не указан URL календаря ICS.");
                Console.WriteLine("Добавьте \"IcsUrl\": \"your-calendar-url\" в appsettings.Local.json");
                Console.WriteLine("или передайте URL как аргумент командной строки.");
                return;
            }

            Console.WriteLine($"📅 ICS календарь: {icsUrl}");

            // Определяем период синхронизации (расширенный период с 1 июня 2025)
            var startDate = new DateTime(2025, 6, 1);
            var endDate = new DateTime(2025, 6, 30); // Весь июнь для тестирования

            // Создаем сервисы
            using var icsDownloader = new IcsDownloader();
            var icsParser = new IcsParser();
            using var exchangeHttpService = new ExchangeHttpService(configuration);

            // Тестируем подключение к Exchange через HTTP
            Console.WriteLine("\n🔍 Проверка подключения к Exchange...");
            var connectionResult = await exchangeHttpService.TestConnectionAsync();

            if (!connectionResult)
            {
                Console.WriteLine("❌ Не удалось подключиться к Exchange.");
                Console.WriteLine("🛑 КРИТИЧЕСКАЯ ОШИБКА: Программа не может работать без подключения к Exchange");
                return;
            }

            Console.WriteLine($"\n📊 Период синхронизации: {startDate:yyyy-MM-dd} - {endDate:yyyy-MM-dd}");

            // Выполняем полный цикл синхронизации
            await PerformFullSyncCycleAsync(icsDownloader, icsParser, exchangeHttpService, icsUrl, startDate, endDate);

            Console.WriteLine("\n✅ Синхронизация завершена!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Критическая ошибка: {ex.Message}");
            Console.WriteLine($"📝 Детали: {ex}");
        }
    }

    /// <summary>
    /// Выполнить полный цикл синхронизации: добавление, обновление, удаление событий
    /// </summary>
    private static async Task PerformFullSyncCycleAsync(
        IcsDownloader icsDownloader,
        IcsParser icsParser,
        ExchangeHttpService exchangeHttpService,
        string icsUrl,
        DateTime startDate,
        DateTime endDate)
    {
        try
        {
            Console.WriteLine("\n🔄 Начинаем полный цикл синхронизации...");
            Console.WriteLine("".PadRight(50, '='));

            // Шаг 1: Загружаем и парсим ICS календарь
            Console.WriteLine("\n📥 Загрузка ICS календаря...");
            var icsContent = await icsDownloader.DownloadAsync(icsUrl);
            var icsEvents = icsParser.Parse(icsContent);

            // Фильтруем события по периоду (включая конечную дату)
            var filteredIcsEvents = icsEvents.Where(e =>
                e.Start.Date >= startDate.Date && e.Start.Date <= endDate.Date).ToList();

            Console.WriteLine($"✅ Загружено ICS событий: {icsEvents.Count}");
            Console.WriteLine($"📅 В указанном периоде: {filteredIcsEvents.Count}");

            // Показываем найденные ICS события
            if (filteredIcsEvents.Any())
            {
                Console.WriteLine("\n📋 События из ICS календаря:");
                foreach (var icsEvent in filteredIcsEvents)
                {
                    Console.WriteLine($"  • {icsEvent.Summary} ({icsEvent.Start:yyyy-MM-dd HH:mm}) - {icsEvent.TimeZone}");

                    // Дополнительная информация для отладки
                    if (!string.IsNullOrEmpty(icsEvent.Location))
                        Console.WriteLine($"    📍 Место: {icsEvent.Location}");
                    if (!string.IsNullOrEmpty(icsEvent.Organizer))
                        Console.WriteLine($"    👤 Организатор: {icsEvent.Organizer}");
                    if (icsEvent.Attendees.Any())
                        Console.WriteLine($"    👥 Участники: {string.Join(", ", icsEvent.Attendees)}");
                    if (!string.IsNullOrEmpty(icsEvent.Url))
                        Console.WriteLine($"    🔗 URL: {icsEvent.Url}");
                    if (!string.IsNullOrEmpty(icsEvent.Description))
                        Console.WriteLine($"    📝 Описание: {icsEvent.Description}");
                }
            }

            // Шаг 2: Получаем события из Exchange
            Console.WriteLine("\n📥 Получение событий из Exchange...");
            var exchangeEvents = await exchangeHttpService.GetCalendarEventsAsync(startDate, endDate);

            Console.WriteLine($"✅ Получено Exchange событий: {exchangeEvents.Count}");

            // Показываем найденные Exchange события
            if (exchangeEvents.Any())
            {
                Console.WriteLine("\n📋 События из Exchange:");
                foreach (var evt in exchangeEvents)
                {
                    Console.WriteLine($"  • {evt.Summary} ({evt.Start:yyyy-MM-dd HH:mm}) - ID: {evt.ExchangeId?.Substring(0, 20)}...");
                }
            }

            // Шаг 3: Выполняем синхронизацию
            Console.WriteLine("\n🔄 Анализ различий и синхронизация...");
            await SynchronizeEventsAsync(exchangeHttpService, filteredIcsEvents, exchangeEvents);

            // Шаг 4: Демонстрируем полный цикл (ОТКЛЮЧЕНО - чтобы не удалять созданные события)
            // await DemonstrateFullSyncCycleAsync(exchangeHttpService, filteredIcsEvents);

        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка синхронизации: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Синхронизировать события между ICS и Exchange
    /// </summary>
    private static async Task SynchronizeEventsAsync(
        ExchangeHttpService exchangeHttpService,
        List<CalendarEvent> icsEvents,
        List<CalendarEvent> exchangeEvents)
    {
        var stats = new SyncStats();

        try
        {
            // Сопоставляем события по UID (правильный подход)
            var eventPairs = new List<(CalendarEvent ics, CalendarEvent exchange)>();
            var unmatchedIcsEvents = new List<CalendarEvent>();
            var unmatchedExchangeEvents = new List<CalendarEvent>(exchangeEvents);

            foreach (var icsEvent in icsEvents)
            {
                CalendarEvent? matchingExchangeEvent = null;

                // Приоритет 1: Поиск по UID (правильный подход для календарей)
                if (!string.IsNullOrEmpty(icsEvent.Uid))
                {
                    matchingExchangeEvent = exchangeEvents.FirstOrDefault(e =>
                        !string.IsNullOrEmpty(e.Uid) &&
                        string.Equals(e.Uid, icsEvent.Uid, StringComparison.OrdinalIgnoreCase));

                    if (matchingExchangeEvent != null)
                    {
                        Console.WriteLine($"🔗 Найдено соответствие по UID: {icsEvent.Summary} ↔ {matchingExchangeEvent.Summary}");
                    }
                }

                // Приоритет 2: Fallback - поиск по названию и времени (для событий без UID)
                if (matchingExchangeEvent == null)
                {
                    matchingExchangeEvent = exchangeEvents.FirstOrDefault(e =>
                        string.Equals(e.Summary?.Trim(), icsEvent.Summary?.Trim(), StringComparison.OrdinalIgnoreCase) &&
                        Math.Abs((e.Start - icsEvent.Start).TotalMinutes) < 30); // 30 минут tolerance

                    if (matchingExchangeEvent != null)
                    {
                        Console.WriteLine($"🔗 Найдено соответствие по названию+времени: {icsEvent.Summary} ↔ {matchingExchangeEvent.Summary}");
                    }
                }

                if (matchingExchangeEvent != null)
                {
                    eventPairs.Add((icsEvent, matchingExchangeEvent));
                    unmatchedExchangeEvents.Remove(matchingExchangeEvent);
                }
                else
                {
                    unmatchedIcsEvents.Add(icsEvent);
                }
            }

            Console.WriteLine($"\n📊 Анализ синхронизации:");
            Console.WriteLine($"  🔗 Сопоставленных событий: {eventPairs.Count}");
            Console.WriteLine($"  ➕ Новых для создания в Exchange: {unmatchedIcsEvents.Count}");
            Console.WriteLine($"  🗑️  Для удаления из Exchange: {unmatchedExchangeEvents.Count}");

            // Создаем новые события в Exchange
            foreach (var icsEvent in unmatchedIcsEvents)
            {
                try
                {
                    Console.WriteLine($"\n➕ Создание события: {icsEvent.Summary}");
                    var eventId = await exchangeHttpService.CreateCalendarEventAsync(icsEvent);

                    if (!string.IsNullOrEmpty(eventId))
                    {
                        icsEvent.ExchangeId = eventId;
                        stats.EventsCreated++;
                        Console.WriteLine($"✅ Событие создано с ID: {eventId.Substring(0, 20)}...");
                    }
                    else
                    {
                        stats.Errors++;
                        Console.WriteLine($"❌ Не удалось создать событие: {icsEvent.Summary}");
                    }
                }
                catch (Exception ex)
                {
                    stats.Errors++;
                    Console.WriteLine($"❌ Ошибка создания события {icsEvent.Summary}: {ex.Message}");
                }
            }

            // Обновляем существующие события
            foreach (var (icsEvent, exchangeEvent) in eventPairs)
            {
                try
                {
                    // Проверяем, нужно ли обновление
                    var needsUpdate = !string.Equals(icsEvent.Summary, exchangeEvent.Summary, StringComparison.OrdinalIgnoreCase) ||
                                     Math.Abs((icsEvent.Start - exchangeEvent.Start).TotalMinutes) > 5 ||
                                     Math.Abs((icsEvent.End - exchangeEvent.End).TotalMinutes) > 5;

                    if (needsUpdate)
                    {
                        Console.WriteLine($"\n✏️ Обновление события: {icsEvent.Summary}");
                        icsEvent.ExchangeId = exchangeEvent.ExchangeId;
                        icsEvent.ExchangeChangeKey = exchangeEvent.ExchangeChangeKey;
                        var success = await exchangeHttpService.UpdateCalendarEventAsync(icsEvent);

                        if (success)
                        {
                            stats.EventsUpdated++;
                            Console.WriteLine("✅ Событие обновлено успешно");
                        }
                        else
                        {
                            stats.Errors++;
                            Console.WriteLine("❌ Не удалось обновить событие");
                        }
                    }
                    else
                    {
                        stats.EventsUpToDate++;
                        Console.WriteLine($"✅ Событие актуально: {icsEvent.Summary}");
                    }
                }
                catch (Exception ex)
                {
                    stats.Errors++;
                    Console.WriteLine($"❌ Ошибка обновления события {icsEvent.Summary}: {ex.Message}");
                }
            }

            // Удаляем лишние события из Exchange (только события с меткой CalSync!)
            var eventsToDelete = unmatchedExchangeEvents.Where(e => e.IsCalSyncEvent).ToList();
            var eventsToSkip = unmatchedExchangeEvents.Where(e => !e.IsCalSyncEvent).ToList();

            if (eventsToDelete.Any())
            {
                Console.WriteLine($"\n🗑️  Удаление {eventsToDelete.Count} событий CalSync из Exchange (их нет в ICS календаре):");

                foreach (var evt in eventsToDelete)
                {
                    try
                    {
                        Console.WriteLine($"  🗑️  Удаление: {evt.Summary} ({evt.Start:yyyy-MM-dd HH:mm})");
                        var success = await exchangeHttpService.DeleteCalendarEventAsync(evt.ExchangeId);

                        if (success)
                        {
                            stats.EventsDeleted++;
                            Console.WriteLine($"    ✅ Удалено успешно");
                        }
                        else
                        {
                            stats.Errors++;
                            Console.WriteLine($"    ❌ Не удалось удалить");
                        }
                    }
                    catch (Exception ex)
                    {
                        stats.Errors++;
                        Console.WriteLine($"    ❌ Ошибка удаления: {ex.Message}");
                    }
                }
            }

            if (eventsToSkip.Any())
            {
                Console.WriteLine($"\n⏭️  Пропущено {eventsToSkip.Count} событий без метки CalSync (не наши события):");
                foreach (var evt in eventsToSkip)
                {
                    Console.WriteLine($"  ⏭️  {evt.Summary} ({evt.Start:yyyy-MM-dd HH:mm})");
                    stats.EventsSkipped++;
                }
            }

            // Показываем статистику
            Console.WriteLine($"\n📊 Результаты синхронизации:");
            Console.WriteLine($"  ✅ Создано: {stats.EventsCreated}");
            Console.WriteLine($"  ✏️ Обновлено: {stats.EventsUpdated}");
            Console.WriteLine($"  🗑️ Удалено: {stats.EventsDeleted}");
            Console.WriteLine($"  ✅ Актуально: {stats.EventsUpToDate}");
            Console.WriteLine($"  ⏭️ Пропущено: {stats.EventsSkipped}");
            Console.WriteLine($"  ❌ Ошибок: {stats.Errors}");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка синхронизации: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Демонстрация полного цикла синхронизации
    /// </summary>
    private static async Task DemonstrateFullSyncCycleAsync(
        ExchangeHttpService exchangeHttpService,
        List<CalendarEvent> icsEvents)
    {
        Console.WriteLine("\n🧪 Демонстрация полного цикла синхронизации...");
        Console.WriteLine("".PadRight(50, '='));

        if (!icsEvents.Any())
        {
            Console.WriteLine("⚠️  Нет событий для демонстрации");
            return;
        }

        var testEvent = icsEvents.First();
        var originalSummary = testEvent.Summary;
        var originalStart = testEvent.Start;

        try
        {
            // 1. Создание события (если еще не создано)
            if (string.IsNullOrEmpty(testEvent.ExchangeId))
            {
                Console.WriteLine("\n1️⃣ Создание тестового события...");
                var createdEvent = await exchangeHttpService.CreateCalendarEventWithDetailsAsync(testEvent);
                testEvent.ExchangeId = createdEvent.ExchangeId;
                testEvent.ExchangeChangeKey = createdEvent.ExchangeChangeKey;
                Console.WriteLine($"✅ Событие создано: {testEvent.ExchangeId?.Substring(0, 20)}...");
            }
            else
            {
                Console.WriteLine($"\n1️⃣ Событие уже существует: {testEvent.ExchangeId?.Substring(0, 20)}...");
            }

            // 2. Обновление события
            Console.WriteLine("\n2️⃣ Обновление события...");
            testEvent.Summary = $"{originalSummary} [ОБНОВЛЕНО]";
            testEvent.Start = originalStart.AddMinutes(15);
            testEvent.End = testEvent.End.AddMinutes(15);

            var updateSuccess = await exchangeHttpService.UpdateCalendarEventAsync(testEvent);
            if (updateSuccess)
            {
                Console.WriteLine("✅ Событие успешно обновлено");
            }
            else
            {
                Console.WriteLine("❌ Не удалось обновить событие");
            }

            // 3. Получение обновленного события
            Console.WriteLine("\n3️⃣ Проверка обновления...");
            var updatedEvents = await exchangeHttpService.GetCalendarEventsAsync(
                testEvent.Start.Date, testEvent.Start.Date.AddDays(1));

            var updatedEvent = updatedEvents.FirstOrDefault(e => e.ExchangeId == testEvent.ExchangeId);
            if (updatedEvent != null)
            {
                Console.WriteLine($"✅ Событие найдено: {updatedEvent.Summary}");
                Console.WriteLine($"📅 Время: {updatedEvent.Start:yyyy-MM-dd HH:mm}");
            }

            // 4. Удаление события (опционально)
            Console.WriteLine("\n4️⃣ Удаление тестового события...");
            var deleteSuccess = await exchangeHttpService.DeleteCalendarEventAsync(testEvent.ExchangeId);
            if (deleteSuccess)
            {
                Console.WriteLine("✅ Событие успешно удалено");
            }
            else
            {
                Console.WriteLine("❌ Не удалось удалить событие");
            }

            // 5. Проверка удаления
            Console.WriteLine("\n5️⃣ Проверка удаления...");
            var finalEvents = await exchangeHttpService.GetCalendarEventsAsync(
                testEvent.Start.Date, testEvent.Start.Date.AddDays(1));

            var deletedEvent = finalEvents.FirstOrDefault(e => e.ExchangeId == testEvent.ExchangeId);
            if (deletedEvent == null)
            {
                Console.WriteLine("✅ Событие успешно удалено из Exchange");
            }
            else
            {
                Console.WriteLine("⚠️  Событие все еще существует в Exchange");
            }

            Console.WriteLine("\n🎉 Полный цикл синхронизации завершен!");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка в демонстрации: {ex.Message}");
        }
    }

    private class SyncStats
    {
        public int EventsCreated { get; set; }
        public int EventsUpdated { get; set; }
        public int EventsDeleted { get; set; }
        public int EventsUpToDate { get; set; }
        public int EventsSkipped { get; set; }
        public int Errors { get; set; }
    }
}
