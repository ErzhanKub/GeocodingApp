using GeocodingAppConsole.Abstraction;
using Syncfusion.XlsIO;

namespace GeocodingAppConsole.Services;

internal sealed class ExcelAddressGeocoder
{
    private readonly IGeocoder _geocoder;

    public ExcelAddressGeocoder(IGeocoder geocoder)
    {
        _geocoder = geocoder;
    }

    public async Task AddressHandler(string path)
    {
        // Использование ExcelEngine для работы с Excel файлом
        using (var excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            IWorkbook workbook;

            // Открытие Excel файла по указанному пути
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                workbook = application.Workbooks.Open(stream);

                // Выбор первого листа в рабочей книге Excel
                IWorksheet worksheet = workbook.Worksheets[0];

                // Получение количества строк в листе
                int rowCount = worksheet.UsedRange.Rows.Count();

                // Обработка адресов начиная со второй строки (пропуская заголовки)
                for (int i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        // Получение адреса из текущей строки первой ячейки
                        var address = worksheet.Range[i, 2].Value;

                        // Геокодирование адреса и получение широты и долготы
                        var (lat, lng) = await _geocoder.GeocodeAsync(address);

                        // Запись широты и долготы в текущую строку
                        worksheet.Range[i, 3].Value = lat.ToString();
                        worksheet.Range[i, 4].Value = lng.ToString();

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Адрес: {address} успешно обработан. Широта: {lat} | Долгота: {lng}");

                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.SetCursorPosition(0, 5);
                        int progress = (i - 1) * 100 / (rowCount - 1);
                        Console.Write($"\rОбработано адресов: {progress}%");
                        Console.SetCursorPosition(0,0);
                        Console.ResetColor();
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Ошибка при обработке адреса: {worksheet.Range[i, 2].Value} || {ex.Message}");
                        Console.ResetColor();
                    }
                }
            }

            // Сохранение и закрытие рабочей книги
            using (var saveStream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.SaveAs(saveStream);
            }
            workbook.Close();
        }
        Console.SetCursorPosition(0,2);
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("##### - Обработка всех адресов завершена. Файл Excel сохранен - ####");
        Console.ResetColor();
    }
}
