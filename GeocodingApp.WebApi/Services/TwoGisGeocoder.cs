using GeocodingApp.WebApi.Abstraction;
using Newtonsoft.Json.Linq;

namespace GeocodingApp.WebApi.Services
{
    internal sealed class TwoGisGeocoder : IGeocoder
    {
        private readonly HttpClient _httpClient;
        private const string twoGisBaseUrl = "https://catalog.api.2gis.com/3.0/items/geocode";

        public TwoGisGeocoder()
        {
            _httpClient = new HttpClient();
        }

        public async Task<(double lat, double lng)> GeocodeAsync(string address, string apiKey)
        {
            var formattedAddress = Uri.EscapeDataString(address);

            var url = $"{twoGisBaseUrl}?q={formattedAddress}&fields=items.point,items.geometry.centroid&key={apiKey}";

            var response = await _httpClient.GetAsync(url);

            response.EnsureSuccessStatusCode();

            var responseBody = await response.Content.ReadAsStringAsync();
            var data = JObject.Parse(responseBody);

            var lat = data["result"]?["items"]?[0]["point"]["lat"]?.ToObject<double>() ?? 0.0;
            var lng = data["result"]?["items"]?[0]["point"]["lon"]?.ToObject<double>() ?? 0.0;

            return (lat, lng);
        }
    }
}