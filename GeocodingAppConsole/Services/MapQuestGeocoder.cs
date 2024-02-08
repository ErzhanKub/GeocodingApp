using GeocodingAppConsole.Abstraction;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GeocodingAppConsole.Services;

internal sealed class MapQuestGeocoder : IGeocoder
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;
    private const string mapQuestBaseUrl = "http://www.mapquestapi.com/geocoding/v1/address?key=";

    public MapQuestGeocoder(string apiKey)
    {
        _httpClient = new HttpClient();
        _apiKey = apiKey;
    }

    public async Task<(double lat, double lng)> GeocodeAsync(string address)
    {
        try
        {
            var formattedAddress = Uri.EscapeDataString(address);

            var url = $"{mapQuestBaseUrl}{_apiKey}&location={formattedAddress}";

            var response = await _httpClient.GetAsync(url);

            response.EnsureSuccessStatusCode();

            var responseBody = await response.Content.ReadAsStringAsync();
            var data = JObject.Parse(responseBody);

            var lat = data["results"]?[0]["locations"]?[0]["latLng"]["lat"]?.ToObject<double>() ?? 0.0;
            var lng = data["results"]?[0]["locations"]?[0]["latLng"]["lng"]?.ToObject<double>() ?? 0.0;

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
