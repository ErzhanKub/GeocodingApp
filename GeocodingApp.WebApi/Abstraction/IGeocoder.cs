namespace GeocodingApp.WebApi.Abstraction
{
    public interface IGeocoder
    {
        Task<(double lat, double lng)> GeocodeAsync(string address, string apiKey);
    }
}