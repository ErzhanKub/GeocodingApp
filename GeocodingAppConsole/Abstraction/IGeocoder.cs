namespace GeocodingAppConsole.Abstraction;

internal interface IGeocoder
{
    Task<(double lat, double lng)> GeocodeAsync(string address);
}
