using CalSync.Services;
using Microsoft.Extensions.Configuration;

namespace CalSync;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("CalSync - Тестирование загрузки и парсинга календаря");
        Console.WriteLine("=================================================");

        try
        {
            // Загружаем конфигурацию
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true)
                .AddJsonFile("appsettings.Development.json", optional: true)
                .Build();

            // Получаем URL календаря из конфигурации
            var calendarUrl = configuration["TestUrls:RealICloudCalendar"];
            if (string.IsNullOrEmpty(calendarUrl))
            {
                Console.WriteLine("❌ URL календаря не найден в конфигурации");
                return;
            }

            Console.WriteLine($"📅 Загружаем календарь: {calendarUrl}");

            // Создаем сервисы
            var downloader = new IcsDownloader();
            var parser = new IcsParser();

            // Загружаем календарь
            Console.WriteLine("⬇️  Скачиваем календарь...");
            var icsContent = await downloader.DownloadAsync(calendarUrl);
            Console.WriteLine($"✅ Загружено {icsContent.Length} символов");

            // Парсим события
            Console.WriteLine("🔍 Парсим события...");
            var events = parser.Parse(icsContent);
            Console.WriteLine($"✅ Найдено событий: {events.Count}");

            if (events.Count == 0)
            {
                Console.WriteLine("⚠️  События не найдены");
                return;
            }

            // Выводим все события
            Console.WriteLine("\n📋 Список всех событий:");
            Console.WriteLine("".PadRight(80, '='));

            for (int i = 0; i < events.Count; i++)
            {
                var evt = events[i];
                Console.WriteLine($"{i + 1:D2}. {evt.Summary}");
                Console.WriteLine($"    📅 Дата: {evt.Start:yyyy-MM-dd HH:mm:ss} - {evt.End:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"    🌍 Временная зона: {(string.IsNullOrEmpty(evt.TimeZone) ? "не указана" : evt.TimeZone)}");
                Console.WriteLine($"    🆔 UID: {evt.Uid}");
                if (!string.IsNullOrEmpty(evt.Location))
                    Console.WriteLine($"    📍 Место: {evt.Location}");
                if (!string.IsNullOrEmpty(evt.Description))
                    Console.WriteLine($"    📝 Описание: {evt.Description}");
                Console.WriteLine($"    📊 Статус: {evt.Status}");
                Console.WriteLine();
            }

            // Ищем тестовое событие
            Console.WriteLine("🔍 Поиск тестового события 'test' 19 июня 2025 года...");
            var testEvent = events.FirstOrDefault(e =>
                e.Summary.Equals("test", StringComparison.OrdinalIgnoreCase) &&
                e.Start.Year == 2025 &&
                e.Start.Month == 6 &&
                e.Start.Day == 19);

            if (testEvent != null)
            {
                Console.WriteLine("🎉 ТЕСТОВОЕ СОБЫТИЕ НАЙДЕНО!");
                Console.WriteLine($"   📅 Название: {testEvent.Summary}");
                Console.WriteLine($"   🕐 Время: {testEvent.Start:yyyy-MM-dd HH:mm:ss} (Kind: {testEvent.Start.Kind})");
                Console.WriteLine($"   🌍 Временная зона: {testEvent.TimeZone}");

                // Проверяем время
                Console.WriteLine($"   ⏰ Час: {testEvent.Start.Hour}, Минута: {testEvent.Start.Minute}");

                // Анализируем время
                if (testEvent.Start.Hour == 10 && testEvent.Start.Minute == 15)
                {
                    Console.WriteLine("   ✅ Время соответствует ожидаемому: 10:15 (MSK время)");
                }
                else if (testEvent.Start.Hour == 7 && testEvent.Start.Minute == 15)
                {
                    Console.WriteLine("   ✅ Время соответствует ожидаемому: 07:15 (UTC время, что равно 10:15 MSK)");
                }
                else
                {
                    Console.WriteLine($"   ⚠️  Время не соответствует ожидаемому. Ожидалось: 10:15 MSK или 07:15 UTC");
                }
            }
            else
            {
                Console.WriteLine("❌ Тестовое событие не найдено!");
                Console.WriteLine("   Проверьте, что в календаре есть событие:");
                Console.WriteLine("   - Название: 'test'");
                Console.WriteLine("   - Дата: 19 июня 2025 года");
                Console.WriteLine("   - Время: 10:15 MSK");
            }

            // Очистка ресурсов
            downloader.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка: {ex.Message}");
            Console.WriteLine($"📝 Детали: {ex}");
        }

        Console.WriteLine("\n🏁 Нажмите любую клавишу для выхода...");
        Console.ReadKey();
    }
}
