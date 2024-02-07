using GeocodingAppConsole.Services;

namespace GeocodingAppConsole;

internal sealed class View
{
    public static void Start()
    {
        var exit = false;
        while (!exit)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("GeocodingAppConsole");
            Console.ResetColor();

            Console.WriteLine("*****************************************");
            Console.WriteLine("1. Геокодирования адресов из файла Excel");
            Console.WriteLine("2. FAQ");
            Console.WriteLine("0. Exit");
            Console.WriteLine("*****************************************");

            var option = Console.ReadLine();

            Console.Clear();

            switch (option)
            {
                case "1": Geocode(); break;
                case "2": FAQ(); break;
                case "0": exit = true; break;
                default: Console.WriteLine("Неверная опция."); break;
            }

            Console.ReadKey();
            Console.Clear();
        }
    }

    private static async void Geocode()
    {
        Console.WriteLine("Введите путь к файлу Excel: ");

        Console.SetCursorPosition(0, 2);
        Console.WriteLine("Введите ваш ключ API: ");

        Console.SetCursorPosition(28, 0);
        var filePath = Console.ReadLine();

        Console.SetCursorPosition(22, 2);
        var apiKey = Console.ReadLine();

        var geocoder = new Geocoder(apiKey);
        var addressGeocoder = new ExcelAddressGeocoder(geocoder);
        await addressGeocoder.AddressHandler(filePath);
    }

    private static void FAQ()
    {

    }
}
