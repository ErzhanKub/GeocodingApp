using GeocodingAppConsole.Abstraction;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GeocodingAppConsole.Services.Geocoders;

internal sealed class TwoGisGeocoder : IGeocoder
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;
    private const string twoGisBaseUrl = "https://catalog.api.2gis.com/3.0/items/geocode";

    public TwoGisGeocoder(string apiKey)
    {
        _httpClient = new HttpClient();
        _apiKey = apiKey;
    }

    public async Task<(double lat, double lng)> GeocodeAsync(string address)
    {
        try
        {
            var formattedAddress = Uri.EscapeDataString(address);

            var url = $"{twoGisBaseUrl}?q={formattedAddress}&fields=items.point,items.geometry.centroid&key={_apiKey}";

            var response = await _httpClient.GetAsync(url);

            response.EnsureSuccessStatusCode();

            var responseBody = await response.Content.ReadAsStringAsync();
            var data = JObject.Parse(responseBody);

            var lat = data["result"]?["items"]?[0]["point"]["lat"]?.ToObject<double>() ?? 0.0;
            var lng = data["result"]?["items"]?[0]["point"]["lon"]?.ToObject<double>() ?? 0.0;

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
