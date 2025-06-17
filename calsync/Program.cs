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

            // Определяем период синхронизации (фокус на 19 июня 2025)
            var targetDate = new DateTime(2025, 6, 19);
            var startDate = targetDate.Date;
            var endDate = targetDate.Date.AddDays(1);

            // Создаем сервисы
            using var exchangeService = new ExchangeService(configuration);
            using var icsDownloader = new IcsDownloader();
            var icsParser = new IcsParser();

            // Тестируем подключение к Exchange
            Console.WriteLine("\n🔍 Проверка подключения к Exchange...");
            var connectionResult = await exchangeService.TestConnectionAsync();

            if (!connectionResult)
            {
                Console.WriteLine("❌ Не удалось подключиться к Exchange.");
                Console.WriteLine("🧪 Запускаем тестовый режим для проверки timezone исправлений...");

                // Тестовый режим: проверяем только создание событий без Exchange
                await TestTimezoneFixesAsync(icsDownloader, icsParser, exchangeService, icsUrl, startDate, endDate);
                return;
            }

            Console.WriteLine($"\n📊 Период синхронизации: {startDate:yyyy-MM-dd} - {endDate:yyyy-MM-dd}");

            // Выполняем синхронизацию
            await PerformSyncAsync(icsDownloader, icsParser, exchangeService, icsUrl, startDate, endDate);

            Console.WriteLine("\n✅ Синхронизация завершена!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Критическая ошибка: {ex.Message}");
            Console.WriteLine($"📝 Детали: {ex}");
        }

    }

    /// <summary>
    /// Выполнить синхронизацию между ICS календарем и Exchange
    /// </summary>
    private static async Task PerformSyncAsync(
        IcsDownloader icsDownloader,
        IcsParser icsParser,
        ExchangeService exchangeService,
        string icsUrl,
        DateTime startDate,
        DateTime endDate)
    {
        try
        {
            Console.WriteLine("\n🔄 Начинаем синхронизацию...");
            Console.WriteLine("".PadRight(50, '='));

            // Шаг 1: Загружаем и парсим ICS календарь
            Console.WriteLine("\n📥 Загрузка ICS календаря...");
            var icsContent = await icsDownloader.DownloadAsync(icsUrl);
            var icsEvents = icsParser.Parse(icsContent);

            // Фильтруем события по периоду
            var filteredIcsEvents = icsEvents.Where(e =>
                e.Start.Date >= startDate.Date && e.Start.Date < endDate.Date).ToList();

            Console.WriteLine($"✅ Загружено ICS событий: {icsEvents.Count}");
            Console.WriteLine($"📅 В указанном периоде: {filteredIcsEvents.Count}");

            // Показываем найденные ICS события
            if (filteredIcsEvents.Any())
            {
                Console.WriteLine("\n📋 События из ICS календаря:");
                foreach (var evt in filteredIcsEvents)
                {
                    Console.WriteLine($"  • {evt.Summary} ({evt.Start:yyyy-MM-dd HH:mm})");
                }
            }

            // Шаг 2: Получаем события из Exchange
            Console.WriteLine("\n📥 Получение событий из Exchange...");
            var exchangeEvents = await exchangeService.GetCalendarEventsAsync(startDate, endDate);

            // Исключаем тестовые события из синхронизации
            var nonTestExchangeEvents = exchangeEvents.Where(e =>
                !e.Summary.StartsWith("[TEST]", StringComparison.OrdinalIgnoreCase) &&
                !e.Description.Contains("[CalSync-Test-Event-")).ToList();

            Console.WriteLine($"✅ Получено Exchange событий: {exchangeEvents.Count}");
            Console.WriteLine($"📅 Без тестовых событий: {nonTestExchangeEvents.Count}");

            // Показываем найденные Exchange события
            if (nonTestExchangeEvents.Any())
            {
                Console.WriteLine("\n📋 События из Exchange:");
                foreach (var evt in nonTestExchangeEvents)
                {
                    Console.WriteLine($"  • {evt.Summary} ({evt.Start:yyyy-MM-dd HH:mm})");
                }
            }

            // Шаг 3: Выполняем синхронизацию
            Console.WriteLine("\n🔄 Анализ различий и синхронизация...");
            await SynchronizeEventsAsync(exchangeService, filteredIcsEvents, nonTestExchangeEvents);

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
        ExchangeService exchangeService,
        List<CalendarEvent> icsEvents,
        List<CalendarEvent> exchangeEvents)
    {
        var stats = new SyncStats();

        try
        {
            // Сопоставляем события по заголовку и времени (так как UID может отличаться)
            var eventPairs = new List<(CalendarEvent ics, CalendarEvent exchange)>();
            var unmatchedIcsEvents = new List<CalendarEvent>();
            var unmatchedExchangeEvents = new List<CalendarEvent>(exchangeEvents);

            foreach (var icsEvent in icsEvents)
            {
                var matchingExchangeEvent = exchangeEvents.FirstOrDefault(e =>
                    string.Equals(e.Summary?.Trim(), icsEvent.Summary?.Trim(), StringComparison.OrdinalIgnoreCase) &&
                    Math.Abs((e.Start - icsEvent.Start).TotalMinutes) < 5); // 5 минут tolerance

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

                    // Помечаем событие как синхронизированное из ICS
                    var eventToCreate = new CalendarEvent
                    {
                        Summary = icsEvent.Summary,
                        Description = $"{icsEvent.Description}\n\n[CalSync-Synced-{DateTime.UtcNow:yyyyMMddHHmmss}]",
                        Start = icsEvent.Start,
                        End = icsEvent.End,
                        Location = icsEvent.Location,
                        Uid = icsEvent.Uid
                    };

                    var createdId = await exchangeService.CreateCalendarEventAsync(eventToCreate);
                    stats.EventsCreated++;
                    Console.WriteLine($"✅ Создано: {createdId}");
                }
                catch (Exception ex)
                {
                    stats.Errors++;
                    Console.WriteLine($"❌ Ошибка создания '{icsEvent.Summary}': {ex.Message}");
                }
            }

            // Обновляем существующие события (если нужно)
            foreach (var (icsEvent, exchangeEvent) in eventPairs)
            {
                try
                {
                    var needsUpdate = false;
                    var changes = new List<string>();

                    // Проверяем различия
                    if (!string.Equals(icsEvent.Description?.Trim(), exchangeEvent.Description?.Trim(), StringComparison.Ordinal))
                    {
                        needsUpdate = true;
                        changes.Add("описание");
                    }

                    if (!string.Equals(icsEvent.Location?.Trim(), exchangeEvent.Location?.Trim(), StringComparison.Ordinal))
                    {
                        needsUpdate = true;
                        changes.Add("место");
                    }

                    if (Math.Abs((icsEvent.Start - exchangeEvent.Start).TotalMinutes) > 1)
                    {
                        needsUpdate = true;
                        changes.Add("время начала");
                    }

                    if (Math.Abs((icsEvent.End - exchangeEvent.End).TotalMinutes) > 1)
                    {
                        needsUpdate = true;
                        changes.Add("время окончания");
                    }

                    if (needsUpdate)
                    {
                        Console.WriteLine($"\n🔄 Обновление события: {icsEvent.Summary}");
                        Console.WriteLine($"   Изменения: {string.Join(", ", changes)}");

                        // TODO: Реализовать UpdateCalendarEventAsync в ExchangeService
                        Console.WriteLine($"⚠️  Обновление пока не реализовано");
                        stats.EventsSkipped++;
                    }
                    else
                    {
                        Console.WriteLine($"✓ Событие актуально: {icsEvent.Summary}");
                        stats.EventsUpToDate++;
                    }
                }
                catch (Exception ex)
                {
                    stats.Errors++;
                    Console.WriteLine($"❌ Ошибка при проверке '{icsEvent.Summary}': {ex.Message}");
                }
            }

            // Удаляем события, которых нет в ICS (только синхронизированные нами)
            foreach (var exchangeEvent in unmatchedExchangeEvents)
            {
                try
                {
                    // Удаляем только события, которые были синхронизированы нами
                    if (exchangeEvent.Description.Contains("[CalSync-Synced-"))
                    {
                        Console.WriteLine($"\n🗑️  Удаление события: {exchangeEvent.Summary}");
                        var deleted = await exchangeService.DeleteCalendarEventAsync(exchangeEvent.ExchangeId);
                        if (deleted)
                        {
                            stats.EventsDeleted++;
                            Console.WriteLine($"✅ Удалено");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"⚠️  Пропускаем удаление (не синхронизировано нами): {exchangeEvent.Summary}");
                        stats.EventsSkipped++;
                    }
                }
                catch (Exception ex)
                {
                    stats.Errors++;
                    Console.WriteLine($"❌ Ошибка удаления '{exchangeEvent.Summary}': {ex.Message}");
                }
            }

            // Выводим статистику
            Console.WriteLine("\n📊 Результаты синхронизации:");
            Console.WriteLine("".PadRight(40, '='));
            Console.WriteLine($"✅ Создано событий: {stats.EventsCreated}");
            Console.WriteLine($"🔄 Обновлено событий: {stats.EventsUpdated}");
            Console.WriteLine($"🗑️  Удалено событий: {stats.EventsDeleted}");
            Console.WriteLine($"✓ Актуальных событий: {stats.EventsUpToDate}");
            Console.WriteLine($"⚠️  Пропущено событий: {stats.EventsSkipped}");
            Console.WriteLine($"❌ Ошибок: {stats.Errors}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Критическая ошибка синхронизации: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Тестовый режим для проверки timezone исправлений
    /// </summary>
    private static async Task TestTimezoneFixesAsync(
        IcsDownloader icsDownloader,
        IcsParser icsParser,
        ExchangeService exchangeService,
        string icsUrl,
        DateTime startDate,
        DateTime endDate)
    {
        try
        {
            Console.WriteLine("\n🧪 ТЕСТОВЫЙ РЕЖИМ: Проверка timezone исправлений");
            Console.WriteLine("".PadRight(60, '='));

            // Шаг 1: Загружаем и парсим ICS календарь
            Console.WriteLine("\n📥 Загрузка ICS календаря...");
            var icsContent = await icsDownloader.DownloadAsync(icsUrl);
            var icsEvents = icsParser.Parse(icsContent);

            // Фильтруем события по периоду
            var filteredIcsEvents = icsEvents.Where(e =>
                e.Start.Date >= startDate.Date && e.Start.Date < endDate.Date).ToList();

            Console.WriteLine($"✅ Загружено ICS событий: {icsEvents.Count}");
            Console.WriteLine($"📅 В указанном периоде: {filteredIcsEvents.Count}");

            // Показываем найденные ICS события
            if (filteredIcsEvents.Any())
            {
                Console.WriteLine("\n📋 События из ICS календаря:");
                foreach (var evt in filteredIcsEvents)
                {
                    Console.WriteLine($"  • {evt.Summary} ({evt.Start:yyyy-MM-dd HH:mm})");
                }

                // Тестируем создание каждого события через все наши timezone исправления
                Console.WriteLine("\n🔧 Тестирование timezone исправлений...");

                foreach (var evt in filteredIcsEvents)
                {
                    if (evt.Summary.ToLower().Contains("test"))
                    {
                        Console.WriteLine($"\n🎯 Найдено тестовое событие: {evt.Summary}");
                        Console.WriteLine($"   Время: {evt.Start:yyyy-MM-dd HH:mm:ss} - {evt.End:yyyy-MM-dd HH:mm:ss}");
                        Console.WriteLine($"   Описание: {evt.Description}");
                        Console.WriteLine($"   Место: {evt.Location}");

                        // Тестируем создание события
                        await TestEventCreationAsync(exchangeService, evt);
                        break; // Тестируем только первое найденное событие "test"
                    }
                }
            }
            else
            {
                Console.WriteLine("❌ Событие 'test' на 19 июня 2025 не найдено в календаре");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка тестового режима: {ex.Message}");
        }
    }

    /// <summary>
    /// Тестировать создание события с timezone исправлениями
    /// </summary>
    private static async Task TestEventCreationAsync(ExchangeService exchangeService, CalendarEvent calendarEvent)
    {
        Console.WriteLine("\n🔧 Тестирование создания события с timezone исправлениями...");

        try
        {
            var eventId = await exchangeService.CreateCalendarEventAsync(calendarEvent);
            Console.WriteLine($"🎉 УСПЕХ! Событие создано с ID: {eventId}");
            Console.WriteLine("✅ Timezone исправления работают!");

            // Попробуем удалить созданное тестовое событие
            try
            {
                await exchangeService.DeleteCalendarEventAsync(eventId);
                Console.WriteLine("🧹 Тестовое событие удалено");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️  Не удалось удалить тестовое событие: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Создание события не удалось: {ex.Message}");
            Console.WriteLine("⚠️  Timezone конфликт все еще присутствует");

            // Показываем детали ошибки для анализа
            if (ex.Message.Contains("Dlt/1880") || ex.Message.Contains("timezone") || ex.Message.Contains("same key"))
            {
                Console.WriteLine("🔍 Обнаружен известный .NET 9 timezone конфликт");
                Console.WriteLine("💡 Попробуйте запустить в .NET 8 или ждите исправления Microsoft");
            }
        }
    }

    /// <summary>
    /// Статистика синхронизации
    /// </summary>
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
