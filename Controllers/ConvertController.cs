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
    /// Convert a Word document to PDF and download it directly
    /// </summary>
    /// <param name="file">The Word document file (.docx)</param>
    /// <returns>The PDF file download</returns>
    [HttpPost("Download")]
    [RequestSizeLimit(50_000_000)] // 50MB limit
    public async Task<IActionResult> Download([FromForm] IFormFile file)
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

        // Get original filename without extension
        var originalFileName = Path.GetFileNameWithoutExtension(file.FileName);

        _logger.LogInformation("Converting {FileName} to PDF for download", file.FileName);

        try
        {
            using var stream = file.OpenReadStream();
            // Call the stream-based conversion
            var (result, pdfStream) = _conversionService.ConvertToPdfStream(stream, originalFileName);

            if (result.Success && pdfStream != null)
            {
                _logger.LogInformation("Successfully converted to stream");

                // Construct a filename for download
                var downloadName = $"{originalFileName}_{result.PolicyNumber ?? "converted"}.pdf";

                // Return file stream (FileContentResult would require reading all bytes, FileStreamResult is better for memory)
                // Note: The stream must be open. But 'pdfStream' is a MemoryStream created in service.
                // We need to ensure it's not disposed prematurely. The service returns it, transferring ownership to controller.
                return File(pdfStream, "application/pdf", downloadName);
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
    /// Convert multiple Word documents to PDF and save to disk (single LibreOffice call = faster)
    /// </summary>
    [HttpPost("batch")]
    [RequestSizeLimit(200_000_000)] // 200MB for batch
    public async Task<IActionResult> ConvertBatch(
        [FromForm] List<IFormFile> files,
        [FromForm] string? outputPath = null)
    {
        if (files == null || files.Count == 0)
        {
            return BadRequest(new BatchConversionResult
            {
                TotalFiles = 0,
                Failed = 0,
                Successful = 0,
                Results = new() { new ConversionResult { Success = false, Message = "No files uploaded" } }
            });
        }

        // Validate all file extensions
        foreach (var file in files)
        {
            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".docx" && ext != ".doc")
            {
                return BadRequest(new BatchConversionResult
                {
                    TotalFiles = files.Count,
                    Failed = files.Count,
                    Successful = 0,
                    Results = new() { new ConversionResult
                    {
                        Success = false,
                        Message = $"Invalid file type: {file.FileName}. Only .docx and .doc files are supported"
                    }}
                });
            }
        }

        // Prepare output directory
        if (string.IsNullOrWhiteSpace(outputPath))
            outputPath = Directory.GetCurrentDirectory();

        if (outputPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            outputPath = Path.GetDirectoryName(outputPath) ?? outputPath;

        if (!Directory.Exists(outputPath))
        {
            try { Directory.CreateDirectory(outputPath); }
            catch (Exception ex)
            {
                return BadRequest(new BatchConversionResult
                {
                    TotalFiles = files.Count,
                    Failed = files.Count,
                    Successful = 0,
                    Results = new() { new ConversionResult
                    {
                        Success = false,
                        Message = $"Cannot create output directory: {ex.Message}"
                    }}
                });
            }
        }

        _logger.LogInformation("Batch converting {Count} files to PDF", files.Count);

        var fileInputs = new List<(Stream stream, string originalFileName)>();
        try
        {
            foreach (var file in files)
                fileInputs.Add((file.OpenReadStream(), Path.GetFileNameWithoutExtension(file.FileName)));

            var result = _conversionService.ConvertBatchToPdf(fileInputs, outputPath);

            _logger.LogInformation("Batch conversion complete: {Successful}/{Total} succeeded", result.Successful, result.TotalFiles);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during batch conversion");
            return StatusCode(500, new BatchConversionResult
            {
                TotalFiles = files.Count,
                Failed = files.Count,
                Successful = 0,
                Results = new() { new ConversionResult
                {
                    Success = false,
                    Message = "Internal server error during batch conversion",
                    Error = ex.Message
                }}
            });
        }
        finally
        {
            foreach (var (stream, _) in fileInputs)
                stream.Dispose();
        }
    }

    /// <summary>
    /// Convert multiple Word documents to PDF and download as a ZIP file
    /// </summary>
    [HttpPost("batch/download")]
    [RequestSizeLimit(200_000_000)] // 200MB for batch
    public async Task<IActionResult> DownloadBatch([FromForm] List<IFormFile> files)
    {
        if (files == null || files.Count == 0)
        {
            return BadRequest(new BatchConversionResult
            {
                TotalFiles = 0,
                Results = new() { new ConversionResult { Success = false, Message = "No files uploaded" } }
            });
        }

        foreach (var file in files)
        {
            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".docx" && ext != ".doc")
            {
                return BadRequest(new BatchConversionResult
                {
                    TotalFiles = files.Count,
                    Failed = files.Count,
                    Results = new() { new ConversionResult
                    {
                        Success = false,
                        Message = $"Invalid file type: {file.FileName}. Only .docx and .doc files are supported"
                    }}
                });
            }
        }

        _logger.LogInformation("Batch converting {Count} files to PDF for download", files.Count);

        var fileInputs = new List<(Stream stream, string originalFileName)>();
        try
        {
            foreach (var file in files)
                fileInputs.Add((file.OpenReadStream(), Path.GetFileNameWithoutExtension(file.FileName)));

            var (result, zipStream) = _conversionService.ConvertBatchToPdfStream(fileInputs);

            if (zipStream != null && result.Successful > 0)
            {
                _logger.LogInformation("Batch download: {Successful}/{Total} converted", result.Successful, result.TotalFiles);
                return File(zipStream, "application/zip", $"converted_{DateTime.Now:yyyyMMdd_HHmmss}.zip");
            }

            _logger.LogWarning("Batch conversion failed: {Message}", result.Results.FirstOrDefault()?.Message);
            return BadRequest(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during batch download");
            return StatusCode(500, new BatchConversionResult
            {
                TotalFiles = files.Count,
                Failed = files.Count,
                Results = new() { new ConversionResult
                {
                    Success = false,
                    Message = "Internal server error during batch conversion",
                    Error = ex.Message
                }}
            });
        }
        finally
        {
            foreach (var (stream, _) in fileInputs)
                stream.Dispose();
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
