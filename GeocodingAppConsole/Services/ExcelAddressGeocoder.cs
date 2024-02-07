using GeocodingAppConsole.Abstraction;
using Excel = Microsoft.Office.Interop.Excel;

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
        // Запуск нового экземпляра Excel в памяти
        var xlApp = new Excel.Application();
        // Открытие файла Excel, по пути
        var xlWorkbook = xlApp.Workbooks.Open(path);
        // Выбор первого листа в рабочей книге Excel
        var xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
        // Диапазон ячеек, которые содержат данные
        var xlRange = xlWorksheet.UsedRange;

        // Количесво строк
        int rowCount = xlRange.Rows.Count;

        for (int i = 1; i <= rowCount; i++)
        {
            try
            {
                // Получение адреса из текущей строки первой ячейки
                var address = (string)(xlRange.Cells[i, 1] as Excel.Range).Value2;

                // Геокодирование адреса и получние широты и долготы в виде кортежа
                var (lat, lng) = await _geocoder.GeocodeAsync(address);

                // Запись широты и долготы в текущую строку. Во вторую и третью соотвественно
                (xlRange.Cells[i, 2] as Excel.Range).Value2 = lat;
                (xlRange.Cells[i, 3] as Excel.Range).Value2 = lng;

                Console.WriteLine($"Адрес: {address} успешно обработан. Широта: {lat} | Долгота: {lng}");

                Console.Write($"\rОбработано адресов: {i} из {rowCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при обработке адреса: {xlRange.Cells[i, 1] as Excel.Range} || {ex.Message}");
            }
        }

        // Сохранение и закрытие рабочей книги / выход из приложения Excel
        xlWorkbook.Save();
        xlWorkbook.Close();
        xlApp.Quit();

        Console.WriteLine("Обработка всех адресов завершена. Файл Excel сохранен.");
    }
}
