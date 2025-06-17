using CalSync.Services;
using CalSync.Models;
using Microsoft.Extensions.Configuration;

namespace CalSync;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("CalSync - Тестирование Exchange Web Services");
        Console.WriteLine("===========================================");

        try
        {
            // Загружаем конфигурацию
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true)
                .AddJsonFile("appsettings.Local.json", optional: true)
                .Build();

            // Создаем Exchange сервис
            using var exchangeService = new CalSync.Services.ExchangeService(configuration);

            // Тестируем подключение
            Console.WriteLine("\n🔍 Тестирование подключения к Exchange...");
            var connectionResult = await exchangeService.TestConnectionAsync();

            if (!connectionResult)
            {
                Console.WriteLine("❌ Не удалось подключиться к Exchange. Проверьте настройки.");
                return;
            }

            // Показываем меню
            while (true)
            {
                Console.WriteLine("\n📋 Выберите действие:");
                Console.WriteLine("1. Показать события календаря");
                Console.WriteLine("2. Создать тестовое событие");
                Console.WriteLine("3. Удалить все тестовые события");
                Console.WriteLine("4. Выход");
                Console.Write("Ваш выбор: ");

                var choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        await ShowCalendarEvents(exchangeService);
                        break;
                    case "2":
                        await CreateTestEvent(exchangeService);
                        break;
                    case "3":
                        await DeleteTestEvents(exchangeService);
                        break;
                    case "4":
                        Console.WriteLine("👋 До свидания!");
                        return;
                    default:
                        Console.WriteLine("❌ Неверный выбор. Попробуйте снова.");
                        break;
                }

                Console.WriteLine("\nНажмите любую клавишу для продолжения...");
                Console.ReadKey();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Критическая ошибка: {ex.Message}");
            Console.WriteLine($"📝 Детали: {ex}");
        }
    }

    /// <summary>
    /// Показать события календаря
    /// </summary>
    private static async Task ShowCalendarEvents(CalSync.Services.ExchangeService exchangeService)
    {
        try
        {
            Console.WriteLine("\n📅 Получение событий календаря...");

            var events = await exchangeService.GetCalendarEventsAsync();

            if (events.Count == 0)
            {
                Console.WriteLine("📭 События не найдены.");
                return;
            }

            Console.WriteLine($"\n📋 Найдено событий: {events.Count}");
            Console.WriteLine("".PadRight(80, '='));

            for (int i = 0; i < Math.Min(events.Count, 10); i++) // Показываем максимум 10 событий
            {
                var evt = events[i];
                Console.WriteLine($"{i + 1:D2}. {evt.Summary}");
                Console.WriteLine($"    📅 Дата: {evt.Start:yyyy-MM-dd HH:mm} - {evt.End:yyyy-MM-dd HH:mm}");
                if (!string.IsNullOrEmpty(evt.Location))
                    Console.WriteLine($"    📍 Место: {evt.Location}");
                Console.WriteLine($"    🆔 ID: {evt.ExchangeId}");

                // Проверяем, тестовое ли это событие
                if (evt.Description.Contains("[CalSync-Test-Event-") || evt.Summary.StartsWith("[TEST]"))
                {
                    Console.WriteLine($"    🧪 ТЕСТОВОЕ СОБЫТИЕ");
                }

                Console.WriteLine();
            }

            if (events.Count > 10)
            {
                Console.WriteLine($"... и еще {events.Count - 10} событий");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка получения событий: {ex.Message}");
        }
    }

    /// <summary>
    /// Создать тестовое событие
    /// </summary>
    private static async Task CreateTestEvent(CalSync.Services.ExchangeService exchangeService)
    {
        try
        {
            Console.WriteLine("\n➕ Создание тестового события...");

            var testEvent = new CalendarEvent
            {
                Summary = $"[TEST] CalSync тест - {DateTime.Now:yyyy-MM-dd HH:mm}",
                Description = "Тестовое событие, созданное CalSync для проверки работы с Exchange",
                Start = DateTime.Now.AddHours(1),
                End = DateTime.Now.AddHours(2),
                Location = "Тестовая локация"
            };

            var eventId = await exchangeService.CreateCalendarEventAsync(testEvent);
            Console.WriteLine($"✅ Тестовое событие создано с ID: {eventId}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка создания события: {ex.Message}");
        }
    }

    /// <summary>
    /// Удалить все тестовые события
    /// </summary>
    private static async Task DeleteTestEvents(CalSync.Services.ExchangeService exchangeService)
    {
        try
        {
            Console.WriteLine("\n🗑️  Удаление тестовых событий...");
            Console.Write("Вы уверены, что хотите удалить все тестовые события? (y/N): ");

            var confirmation = Console.ReadLine();
            if (confirmation?.ToLower() != "y" && confirmation?.ToLower() != "yes")
            {
                Console.WriteLine("❌ Удаление отменено.");
                return;
            }

            var deletedCount = await exchangeService.DeleteAllTestEventsAsync();

            if (deletedCount > 0)
            {
                Console.WriteLine($"✅ Удалено тестовых событий: {deletedCount}");
            }
            else
            {
                Console.WriteLine("📭 Тестовые события не найдены.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка удаления событий: {ex.Message}");
        }
    }
}
