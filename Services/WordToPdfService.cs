using System.Diagnostics;
using System.IO.Compression;

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

    /// <summary>
    /// Convert multiple Word documents to PDF and save to disk (single LibreOffice invocation = faster)
    /// </summary>
    public BatchConversionResult ConvertBatchToPdf(List<(Stream stream, string originalFileName)> files, string outputDirectory)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"word2pdf_batch_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            // Save all input streams to temp files
            var tempFiles = new List<(string tempDocx, string safeFileName, string originalFileName)>();
            foreach (var (stream, originalFileName) in files)
            {
                var safeFileName = SanitizeFileName(originalFileName);
                // Append a short guid to avoid collisions if two files have the same name
                var uniqueSafe = $"{safeFileName}_{Guid.NewGuid().ToString("N")[..6]}";
                var tempDocx = Path.Combine(tempDir, $"{uniqueSafe}.docx");
                using (var fileStream = new FileStream(tempDocx, FileMode.Create))
                {
                    stream.CopyTo(fileStream);
                }
                tempFiles.Add((tempDocx, uniqueSafe, originalFileName));
            }

            // Single LibreOffice invocation for ALL files
            var inputPaths = tempFiles.Select(f => f.tempDocx).ToList();
            var (success, error) = RunLibreOfficeBatch(inputPaths, tempDir);

            var results = new List<ConversionResult>();
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            foreach (var (tempDocx, safeFileName, originalFileName) in tempFiles)
            {
                var tempPdf = Path.Combine(tempDir, $"{safeFileName}.pdf");
                if (success && File.Exists(tempPdf))
                {
                    var pdfFileName = $"{originalFileName}_{timestamp}.pdf";
                    var pdfPath = Path.Combine(outputDirectory, pdfFileName);
                    File.Copy(tempPdf, pdfPath, overwrite: true);
                    results.Add(new ConversionResult
                    {
                        Success = true,
                        Message = "Conversion successful",
                        PdfPath = pdfPath,
                        FileName = pdfFileName
                    });
                }
                else
                {
                    results.Add(new ConversionResult
                    {
                        Success = false,
                        Message = error ?? "Conversion produced no output PDF",
                        Error = error,
                        FileName = $"{originalFileName}.pdf"
                    });
                }
            }

            return new BatchConversionResult
            {
                TotalFiles = files.Count,
                Successful = results.Count(r => r.Success),
                Failed = results.Count(r => !r.Success),
                Results = results
            };
        }
        catch (Exception ex)
        {
            return new BatchConversionResult
            {
                TotalFiles = files.Count,
                Successful = 0,
                Failed = files.Count,
                Results = files.Select(f => new ConversionResult
                {
                    Success = false,
                    Message = $"Batch conversion failed: {ex.Message}",
                    Error = ex.ToString(),
                    FileName = $"{f.originalFileName}.pdf"
                }).ToList()
            };
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }

    /// <summary>
    /// Convert multiple Word documents to PDF and return as a ZIP stream (for download)
    /// </summary>
    public (BatchConversionResult result, MemoryStream? zipStream) ConvertBatchToPdfStream(List<(Stream stream, string originalFileName)> files)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"word2pdf_batch_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            var tempFiles = new List<(string tempDocx, string safeFileName, string originalFileName)>();
            foreach (var (stream, originalFileName) in files)
            {
                var safeFileName = SanitizeFileName(originalFileName);
                var uniqueSafe = $"{safeFileName}_{Guid.NewGuid().ToString("N")[..6]}";
                var tempDocx = Path.Combine(tempDir, $"{uniqueSafe}.docx");
                using (var fileStream = new FileStream(tempDocx, FileMode.Create))
                {
                    stream.CopyTo(fileStream);
                }
                tempFiles.Add((tempDocx, uniqueSafe, originalFileName));
            }

            var inputPaths = tempFiles.Select(f => f.tempDocx).ToList();
            var (success, error) = RunLibreOfficeBatch(inputPaths, tempDir);

            var results = new List<ConversionResult>();
            var zipStream = new MemoryStream();

            using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var (tempDocx, safeFileName, originalFileName) in tempFiles)
                {
                    var tempPdf = Path.Combine(tempDir, $"{safeFileName}.pdf");
                    if (success && File.Exists(tempPdf))
                    {
                        // Ensure unique names in ZIP
                        var pdfName = $"{originalFileName}.pdf";
                        var counter = 1;
                        while (!usedNames.Add(pdfName))
                        {
                            pdfName = $"{originalFileName}_{counter++}.pdf";
                        }

                        var entry = archive.CreateEntry(pdfName);
                        using var entryStream = entry.Open();
                        using var fs = new FileStream(tempPdf, FileMode.Open, FileAccess.Read);
                        fs.CopyTo(entryStream);

                        results.Add(new ConversionResult
                        {
                            Success = true,
                            Message = "Conversion successful",
                            FileName = pdfName
                        });
                    }
                    else
                    {
                        results.Add(new ConversionResult
                        {
                            Success = false,
                            Message = error ?? "Conversion produced no output PDF",
                            Error = error,
                            FileName = $"{originalFileName}.pdf"
                        });
                    }
                }
            }

            zipStream.Position = 0;

            var batchResult = new BatchConversionResult
            {
                TotalFiles = files.Count,
                Successful = results.Count(r => r.Success),
                Failed = results.Count(r => !r.Success),
                Results = results
            };

            return (batchResult, batchResult.Successful > 0 ? zipStream : null);
        }
        catch (Exception ex)
        {
            var batchResult = new BatchConversionResult
            {
                TotalFiles = files.Count,
                Successful = 0,
                Failed = files.Count,
                Results = files.Select(f => new ConversionResult
                {
                    Success = false,
                    Message = $"Batch conversion failed: {ex.Message}",
                    Error = ex.ToString(),
                    FileName = $"{f.originalFileName}.pdf"
                }).ToList()
            };
            return (batchResult, null);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }

    /// <summary>
    /// Run LibreOffice with multiple input files in a single invocation (faster than one-per-file)
    /// </summary>
    private (bool success, string? error) RunLibreOfficeBatch(List<string> inputFiles, string outputDir)
    {
        _semaphore.Wait();

        var userProfile = Path.Combine(Path.GetTempPath(), $"lo_profile_{Guid.NewGuid():N}");
        Directory.CreateDirectory(userProfile);

        try
        {
            var userProfileUri = OperatingSystem.IsWindows()
                ? $"file:///{userProfile.Replace('\\', '/')}"
                : $"file://{userProfile}";

            var fileArgs = string.Join(" ", inputFiles.Select(f => $"\"{f}\""));

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = _libreOfficePath,
                    Arguments = string.Join(" ",
                        $"-env:UserInstallation=\"{userProfileUri}\"",
                        "--headless",
                        "--norestore",
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
                        fileArgs
                    ),
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            _logger.LogInformation("Running batch conversion of {Count} files", inputFiles.Count);

            process.Start();
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();

            // Scale timeout with number of files: 120s base + 30s per additional file
            var timeoutMs = 120_000 + (inputFiles.Count - 1) * 30_000;
            process.WaitForExit(timeoutMs);

            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                return (false, $"LibreOffice timed out after {timeoutMs / 1000} seconds");
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
            try { Directory.Delete(userProfile, recursive: true); } catch { }
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

public class BatchConversionResult
{
    public int TotalFiles { get; set; }
    public int Successful { get; set; }
    public int Failed { get; set; }
    public List<ConversionResult> Results { get; set; } = new();
}
