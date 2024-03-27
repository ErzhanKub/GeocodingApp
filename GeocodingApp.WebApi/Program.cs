using GeocodingApp.WebApi.Abstraction;
using GeocodingApp.WebApi.Dtos;
using GeocodingApp.WebApi.Services;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddScoped<IGeocoder, TwoGisGeocoder>();

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI();

app.UseHttpsRedirection();

app.MapPost("/geocode", async (string path, string apiKey) =>
{
    try
    {
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(apiKey))
            return Results.BadRequest(new Response
            {
                responseCode = "1",
                responseMessage = "������ �������� ������",
                responseBody = $"����������, ��������� ������������ ���������� ���� � ����� � API-�����. " +
                $"���������, ��� ��� ������ ������� ��������� � ��������� �������. " +
                $"Path: {path} | Api Key: {apiKey}"
            });

        if (!File.Exists(path))
            return Results.BadRequest(new Response
            {
                responseCode = "1",
                responseMessage = "������: ���� �� ������",
                responseBody = $"��������� ���� � ����� �� ����������. ����������, ��������� ������������ ���� � ���������, ��� ���� ��������. " +
                $"Path: {path}"
            });

        var geocoder = new TwoGisGeocoder();

        using var package = new ExcelPackage(new FileInfo(path));

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var worksheet = package.Workbook.Worksheets[0];
        int rowCount = worksheet.Dimension.Rows;

        for (int i = 2; i <= rowCount + 1; i++)
        {
            try
            {
                if (worksheet.Cells[i, 2].Value is null)
                    continue;

                var address = worksheet.Cells[i, 2].Value.ToString();

                if (string.IsNullOrWhiteSpace(address))
                    continue;

                var (lat, lng) = await geocoder.GeocodeAsync(address, apiKey);

                worksheet.Cells[i, 3].Value = lat;
                worksheet.Cells[i, 4].Value = lng;
            }
            catch (Exception ex)
            {
                return Results.BadRequest(new Response
                {
                    responseCode = "1",
                    responseMessage = $"������ ��� ��������� ������: {worksheet.Cells[i, 2]}",
                    responseBody = $"{ex.Message}"
                });
            }
        }

        package.Save();

        return Results.Ok(new Response
        {
            responseCode = "0",
            responseMessage = "�������� ������� ���������",
            responseBody = $"���������� ����������� �������: {rowCount}"
        });
    }
    catch (Exception ex)
    {
        return Results.Ok(new Response
        {
            responseCode = "2",
            responseMessage = "����������� ������ ������",
            responseBody = ex.Message
        });
    }
})
.WithName("GeocodeAddresses")
.Produces<Response>(StatusCodes.Status200OK)
.Produces<Response>(StatusCodes.Status400BadRequest)
.Produces<Response>(StatusCodes.Status500InternalServerError)
.WithOpenApi(operation => new(operation)
{
    Summary = "�������������� ������� � Excal",
    Description = "�������� ��������� ���� � ����� Excal � API-����. ����� ��������� ������, ���������� �������������� �������, ������������ � �������."
});

app.Run();

