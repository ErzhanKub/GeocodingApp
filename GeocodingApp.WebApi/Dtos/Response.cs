namespace GeocodingApp.WebApi.Dtos
{
    internal sealed record Response
    {
        public required string responseCode { get; init; }
        public required string responseMessage { get; init; }
        public required string responseBody { get; init; }
    }
}
