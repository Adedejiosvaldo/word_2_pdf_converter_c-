using System.Diagnostics;

namespace WordToPdf.Services;

/// <summary>
/// Service that converts Word documents to PDF using LibreOffice headless
/// </summary>
public class WordToPdfService
{
    private readonly string _libreOfficePath;
    private readonly ILogger<WordToPdfService> _logger;
    private static readonly SemaphoreSlim _semaphore = new(4); // Limit concurrent conversions

    public WordToPdfService(ILogger<WordToPdfService> logger)
    {
        _logger = logger;
        _libreOfficePath = FindLibreOffice();
        _logger.LogInformation("LibreOffice path: {Path}", _libreOfficePath);
    }

    /// <summary>
    /// Convert a Word document stream to PDF and save to disk
    /// </summary>
    public ConversionResult ConvertToPdf(Stream inputStream, string outputDirectory, string originalFileName)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"word2pdf_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            // Save input stream to temp file
            var safeFileName = SanitizeFileName(originalFileName);
            var tempDocx = Path.Combine(tempDir, $"{safeFileName}.docx");
            using (var fileStream = new FileStream(tempDocx, FileMode.Create))
            {
                inputStream.CopyTo(fileStream);
            }

            // Convert using LibreOffice
            var (success, error) = RunLibreOffice(tempDocx, tempDir);
            if (!success)
            {
                return new ConversionResult
                {
                    Success = false,
                    Message = $"LibreOffice conversion failed: {error}",
                    Error = error
                };
            }

            // Find the generated PDF
            var tempPdf = Path.Combine(tempDir, $"{safeFileName}.pdf");
            if (!File.Exists(tempPdf))
            {
                return new ConversionResult
                {
                    Success = false,
                    Message = "Conversion produced no output PDF"
                };
            }

            // Move to output directory
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var pdfFileName = $"{originalFileName}_{timestamp}.pdf";
            var pdfPath = Path.Combine(outputDirectory, pdfFileName);
            File.Copy(tempPdf, pdfPath, overwrite: true);

            return new ConversionResult
            {
                Success = true,
                Message = "Conversion successful",
                PdfPath = pdfPath,
                FileName = pdfFileName
            };
        }
        catch (Exception ex)
        {
            return new ConversionResult
            {
                Success = false,
                Message = $"Conversion failed: {ex.Message}",
                Error = ex.ToString()
            };
        }
        finally
        {
            // Cleanup temp directory
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }

    /// <summary>
    /// Convert a Word document stream to PDF and return as a stream (for download)
    /// </summary>
    public (ConversionResult result, MemoryStream? pdfStream) ConvertToPdfStream(Stream inputStream, string originalFileName)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"word2pdf_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            // Save input stream to temp file
            var safeFileName = SanitizeFileName(originalFileName);
            var tempDocx = Path.Combine(tempDir, $"{safeFileName}.docx");
            using (var fileStream = new FileStream(tempDocx, FileMode.Create))
            {
                inputStream.CopyTo(fileStream);
            }

            // Convert using LibreOffice
            var (success, error) = RunLibreOffice(tempDocx, tempDir);
            if (!success)
            {
                var failResult = new ConversionResult
                {
                    Success = false,
                    Message = $"LibreOffice conversion failed: {error}",
                    Error = error
                };
                return (failResult, null);
            }

            // Find the generated PDF
            var tempPdf = Path.Combine(tempDir, $"{safeFileName}.pdf");
            if (!File.Exists(tempPdf))
            {
                var failResult = new ConversionResult
                {
                    Success = false,
                    Message = "Conversion produced no output PDF"
                };
                return (failResult, null);
            }

            // Read into memory stream
            var outputStream = new MemoryStream();
            using (var fs = new FileStream(tempPdf, FileMode.Open, FileAccess.Read))
            {
                fs.CopyTo(outputStream);
            }
            outputStream.Position = 0;

            var result = new ConversionResult
            {
                Success = true,
                Message = "Conversion successful",
                FileName = $"{originalFileName}.pdf"
            };

            return (result, outputStream);
        }
        catch (Exception ex)
        {
            var result = new ConversionResult
            {
                Success = false,
                Message = $"Conversion failed: {ex.Message}",
                Error = ex.ToString()
            };
            return (result, null);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }

    private (bool success, string? error) RunLibreOffice(string inputFile, string outputDir)
    {
        _semaphore.Wait(); // Throttle concurrent conversions

        // Each conversion gets its own user profile so LibreOffice instances don't lock each other
        var userProfile = Path.Combine(Path.GetTempPath(), $"lo_profile_{Guid.NewGuid():N}");
        Directory.CreateDirectory(userProfile);

        try
        {
            var userProfileUri = OperatingSystem.IsWindows()
                ? $"file:///{userProfile.Replace('\\', '/')}"
                : $"file://{userProfile}";

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = _libreOfficePath,
                    Arguments = string.Join(" ",
                        $"-env:UserInstallation=\"{userProfileUri}\"",
                        "--headless",
                        "--norestore",
                        // "--infilter=\"Microsoft Word 2007-2013 XML\"",
                        "--convert-to \"pdf:writer_pdf_Export:{" +
                            "\\\"UseLosslessCompression\\\":{\\\"type\\\":\\\"boolean\\\",\\\"value\\\":\\\"true\\\"}," +
                            "\\\"Quality\\\":{\\\"type\\\":\\\"long\\\",\\\"value\\\":\\\"100\\\"}," +
                            "\\\"EmbedStandardFonts\\\":{\\\"type\\\":\\\"boolean\\\",\\\"value\\\":\\\"true\\\"}," +
                            "\\\"UseTaggedPDF\\\":{\\\"type\\\":\\\"boolean\\\",\\\"value\\\":\\\"true\\\"}," +
                            "\\\"ExportBookmarks\\\":{\\\"type\\\":\\\"boolean\\\",\\\"value\\\":\\\"true\\\"}," +
                            "\\\"IsSkipEmptyPages\\\":{\\\"type\\\":\\\"boolean\\\",\\\"value\\\":\\\"true\\\"}," +
                            "\\\"SelectPdfVersion\\\":{\\\"type\\\":\\\"long\\\",\\\"value\\\":\\\"0\\\"}" +
                        "}\"",
                        $"--outdir \"{outputDir}\"",
                        $"\"{inputFile}\""
                    ),
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            _logger.LogInformation("Running: {FileName} {Arguments}", process.StartInfo.FileName, process.StartInfo.Arguments);

            process.Start();
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit(120_000); // 120 second timeout for large documents

            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                return (false, "LibreOffice timed out after 120 seconds");
            }

            _logger.LogInformation("LibreOffice stdout: {Output}", stdout);
            if (!string.IsNullOrEmpty(stderr))
                _logger.LogWarning("LibreOffice stderr: {Error}", stderr);

            if (process.ExitCode != 0)
            {
                return (false, $"Exit code {process.ExitCode}: {stderr}");
            }

            return (true, null);
        }
        finally
        {
            _semaphore.Release();
            // Cleanup the temporary user profile
            try { Directory.Delete(userProfile, recursive: true); } catch { }
        }
    }

    /// <summary>
    /// Sanitize filename to remove characters that could cause issues
    /// </summary>
    private static string SanitizeFileName(string fileName)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sanitized = string.Join("_", fileName.Split(invalid, StringSplitOptions.RemoveEmptyEntries));
        // Also replace spaces to avoid shell quoting issues
        return sanitized.Trim();
    }

    private static string FindLibreOffice()
    {
        // Windows paths
        var windowsPaths = new[]
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        };

        // Linux paths
        var linuxPaths = new[]
        {
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            "/usr/local/bin/libreoffice",
            "/usr/local/bin/soffice",
            "/snap/bin/libreoffice"
        };

        var paths = OperatingSystem.IsWindows() ? windowsPaths : linuxPaths;

        foreach (var path in paths)
        {
            if (File.Exists(path))
                return path;
        }

        // Fallback: assume it's on PATH
        return OperatingSystem.IsWindows() ? "soffice.exe" : "libreoffice";
    }
}

/// <summary>
/// Result of a conversion operation
/// </summary>
public class ConversionResult
{
    public bool Success { get; set; }
    public string? Message { get; set; }
    public string? PdfPath { get; set; }
    public string? FileName { get; set; }
    public string? PolicyNumber { get; set; }
    public string? Error { get; set; }
}
