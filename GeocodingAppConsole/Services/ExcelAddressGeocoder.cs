using GeocodingAppConsole.Abstraction;
using OfficeOpenXml;


namespace GeocodingAppConsole.Services
{
    internal sealed class ExcelAddressGeocoder
    {
        private readonly IGeocoder _geocoder;

        public ExcelAddressGeocoder(IGeocoder geocoder)
        {
            _geocoder = geocoder;
        }

        public async Task AddressHandler(string path)
        {
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int i = 2; i <= 50; i++)
                {
                    try
                    {
                        var address = worksheet.Cells[i, 2].Value.ToString();
                        var (lat, lng) = await _geocoder.GeocodeAsync(address);

                        worksheet.Cells[i, 3].Value = lat;
                        worksheet.Cells[i, 4].Value = lng;

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Адрес: {address} успешно обработан. Широта: {lat} | Долгота: {lng}");

                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.SetCursorPosition(0, 5);
                        int progress = (i - 1) * 100 / (50 - 1);
                        Console.Write($"\rОбработано адресов: {progress}%");
                        Console.SetCursorPosition(0, 0);
                        Console.ResetColor();
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Ошибка при обработке адреса: {worksheet.Cells[i, 2].Value} || {ex.Message}");
                        Console.ResetColor();
                    }
                }

                package.Save();
            }

            Console.SetCursorPosition(0, 2);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("##### - Обработка всех адресов завершена. Файл Excel сохранен - ####");
            Console.ResetColor();
        }
    }
}
