using WordToPdf.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new() { Title = "Word to PDF Converter API", Version = "v1" });
});

// Register our conversion service
builder.Services.AddSingleton<WordToPdfService>();

// Configure CORS (allow all for development)
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

var app = builder.Build();

// Configure the HTTP request pipeline
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors();
app.UseAuthorization();
app.MapControllers();

// Configure port for cloud hosting (Fly.io, Railway, etc.)
var port = Environment.GetEnvironmentVariable("PORT") ?? "8080";
app.Urls.Add($"http://0.0.0.0:{port}");

// Print startup info
var urls = app.Urls.Any() ? string.Join(", ", app.Urls) : "http://localhost:5000";
Console.WriteLine("======================================");
Console.WriteLine("  Word to PDF Converter API");
Console.WriteLine("======================================");
Console.WriteLine($"  Listening on: {urls}");
Console.WriteLine();
Console.WriteLine("  Endpoints:");
Console.WriteLine("    POST /api/convert - Convert Word to PDF");
Console.WriteLine("    GET  /api/convert/health - Health check");
Console.WriteLine();
Console.WriteLine("  Swagger UI: /swagger");
Console.WriteLine("======================================");

app.Run();
