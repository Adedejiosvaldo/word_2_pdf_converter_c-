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
    /// <param name="outputPath">The destination directory where the PDF should be saved. If not provided, uses current directory. Filename is auto-generated as: {original}_{policy}_{timestamp}.pdf</param>
    /// <returns>Conversion result with the PDF path</returns>
    [HttpPost]
    [RequestSizeLimit(50_000_000)] // 50MB limit
    public async Task<IActionResult> Convert(
        [FromForm] IFormFile file,
        [FromForm] string? outputPath = null)
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

        // Validate and prepare output directory
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            outputPath = Directory.GetCurrentDirectory();
        }

        // If outputPath ends with .pdf, treat it as a directory path (remove .pdf)
        if (outputPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            outputPath = Path.GetDirectoryName(outputPath) ?? outputPath;
        }

        // Ensure directory exists
        if (!Directory.Exists(outputPath))
        {
            try
            {
                Directory.CreateDirectory(outputPath);
            }
            catch (Exception ex)
            {
                return BadRequest(new ConversionResult
                {
                    Success = false,
                    Message = $"Cannot create output directory: {ex.Message}"
                });
            }
        }

        // Get original filename without extension
        var originalFileName = Path.GetFileNameWithoutExtension(file.FileName);

        _logger.LogInformation("Converting {FileName} to PDF in directory {OutputDirectory}", file.FileName, outputPath);

        try
        {
            // Convert the file - document loaded ONCE inside the service
            using var stream = file.OpenReadStream();
            var result = _conversionService.ConvertToPdf(stream, outputPath, originalFileName);

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
