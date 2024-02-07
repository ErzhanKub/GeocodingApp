using GeocodingAppConsole.Abstraction;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GeocodingAppConsole.Services;

internal sealed class Geocoder : IGeocoder
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;
    private const string googleMapsUrl = "https://maps.googleapis.com/maps/api/geocode/json?address=";

    public Geocoder(string apiKey)
    {
        _httpClient = new HttpClient();
        _apiKey = apiKey;
    }

    public async Task<(double lat, double lng)> GeocodeAsync(string adress)
    {
        try
        {
            // Форматирование строки(адрес) для URL (заменяет пробел, заятую и т.д в процентное кодирование)
            var formattedAdress = Uri.EscapeDataString(adress);

            var url = $"{googleMapsUrl}{formattedAdress}&key={_apiKey}";

            var response = await _httpClient.GetAsync(url);

            // Проверка на то что запрос был успешен, иначе сгенерирует исключение (HttpRequestException)
            response.EnsureSuccessStatusCode();

            // Получение тело ответа и преоброзование строки обратно в JSON
            var responseBody = await response.Content.ReadAsStringAsync();
            var data = JObject.Parse(responseBody);

            // Получение широты и долготы из JSON
            var lat = data["results"]?[0]["geometry"]["location"]["lat"]?.ToObject<double>() ?? 0.0;
            var lng = data["results"]?[0]["geametry"]["location"]["lng"]?.ToObject<double>() ?? 0.0;

            return (lat, lng);
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"Ошибка при отправке запроса: {ex.Message}");
            throw;
        }
        catch (JsonReaderException ex)
        {
            Console.WriteLine($"Ошибка при чтении JSON: {ex.Message}");
            throw;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Неизвестная ошибка: {ex.Message}");
            throw;
        }
    }
}
