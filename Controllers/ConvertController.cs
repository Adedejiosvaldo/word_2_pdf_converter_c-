using Microsoft.AspNetCore.Mvc;
using WordToPdf.Services;

namespace WordToPdf.Controllers;

/// <summary>
/// API Controller for Word to PDF conversion
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class ConvertController : ControllerBase
{
    private readonly WordToPdfService _conversionService;
    private readonly ILogger<ConvertController> _logger;

    public ConvertController(WordToPdfService conversionService, ILogger<ConvertController> logger)
    {
        _conversionService = conversionService;
        _logger = logger;
    }

    /// <summary>
    /// Convert a Word document to PDF
    /// </summary>
    /// <param name="file">The Word document file (.docx)</param>
    /// <param name="outputPath">The destination path where the PDF should be saved</param>
    /// <returns>Conversion result with the PDF path</returns>
    [HttpPost]
    [RequestSizeLimit(50_000_000)] // 50MB limit
    public async Task<IActionResult> Convert(
        [FromForm] IFormFile file,
        [FromForm] string outputPath)
    {
        // Validate file
        if (file == null || file.Length == 0)
        {
            return BadRequest(new ConversionResult
            {
                Success = false,
                Message = "No file uploaded"
            });
        }

        // Validate file extension
        var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
        if (extension != ".docx" && extension != ".doc")
        {
            return BadRequest(new ConversionResult
            {
                Success = false,
                Message = "Invalid file type. Only .docx and .doc files are supported"
            });
        }

        // Validate output path
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return BadRequest(new ConversionResult
            {
                Success = false,
                Message = "Output path is required"
            });
        }

        _logger.LogInformation("Converting {FileName} to PDF at {OutputPath}", file.FileName, outputPath);

        try
        {
            // Convert the file
            using var stream = file.OpenReadStream();
            var result = _conversionService.ConvertToPdf(stream, outputPath);

            if (result.Success)
            {
                _logger.LogInformation("Successfully converted to {PdfPath}", result.PdfPath);
                return Ok(result);
            }
            else
            {
                _logger.LogWarning("Conversion failed: {Message}", result.Message);
                return BadRequest(result);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error converting document");
            return StatusCode(500, new ConversionResult
            {
                Success = false,
                Message = "Internal server error during conversion",
                Error = ex.Message
            });
        }
    }

    /// <summary>
    /// Health check endpoint
    /// </summary>
    [HttpGet("health")]
    public IActionResult Health()
    {
        return Ok(new { status = "healthy", service = "Word to PDF Converter" });
    }
}
