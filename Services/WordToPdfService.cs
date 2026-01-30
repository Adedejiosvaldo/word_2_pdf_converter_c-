using QuestPDF.Infrastructure;
using Xceed.Words.NET;
using Xceed.Document.NET;
using QuestPDF.Helpers;
using QuestPDF.Fluent;
using System.IO.Compression;
using System.Xml.Linq;

namespace WordToPdf.Services;

/// <summary>
/// Service that converts Word documents to PDF using QuestPDF
/// </summary>
public class WordToPdfService
{
    private const float TWIPS_TO_POINTS = 1f / 20f;
    private const float EMUS_TO_POINTS = 1f / 12700f;
    // Sections that must start on new pages
    private static readonly string[] NewPageSections = {
        "POLICY SCHEDULE",
        "MEMORANDA ATTACHING TO AND FORMING PART OF"
    };

    // Text patterns that are NOT headings
    private static readonly string[] ExcludePatterns = {
        "IT IS AGREED",
        "PROVIDED THAT",
        "PROVIDED ALSO",
        "WHEREAS",
        "NOW THEREFORE",
        "IN WITNESS",
        "OZUMBA MBADIWE",
        "VICTORIA ISLAND",
        "LAGOS",
        "Followed by",
        "13TH",
        "14TH",
        "CIVIC TOWERS"
    };

    // Valid section headings
    private static readonly string[] ValidHeadings = {
        "PRODUCT LIABILITY INSURANCE",
        "HEALTHCARE PROFESSIONAL INDEMNITY",
        "IMPORTANT",
        "PLEASE NOTE",
        "EXCEPTIONS",
        "CONDITIONS",
        "POLICY CONDITIONS",
        "THE SCHEDULE",
        "MEMO ",
        "MEMORANDUM",
        "CLAIMS COMPLAINT",
        "Important Notice"
    };

    // Insurance type headings (centered)
    private static readonly string[] InsuranceTypes = {
        "PRODUCT LIABILITY INSURANCE",
        "HEALTHCARE PROFESSIONAL INDEMNITY INSURANCE",
        "PROFESSIONAL INDEMNITY INSURANCE",
        "PUBLIC LIABILITY INSURANCE",
        "MOTOR INSURANCE",
        "FIRE INSURANCE"
    };

    public WordToPdfService()
    {
        QuestPDF.Settings.License = LicenseType.Community;
    }

    /// <summary>
    /// Convert a Word document stream to PDF and save to the specified directory
    /// </summary>
    /// <param name="docxStream">The Word document stream</param>
    /// <param name="outputDirectory">The directory where the PDF should be saved</param>
    /// <param name="originalFileName">The original filename (without extension) for generating output filename</param>
    /// <summary>
    /// Convert a Word document stream to PDF and save to the specified directory
    /// </summary>
    /// <param name="docxStream">The Word document stream</param>
    /// <param name="outputDirectory">The directory where the PDF should be saved</param>
    /// <param name="originalFileName">The original filename (without extension) for generating output filename</param>
    public ConversionResult ConvertToPdf(Stream docxStream, string outputDirectory, string originalFileName)
    {
        var result = new ConversionResult();
        try
        {
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Generate the PDF document object
            var (pdfDocument, metadata, docxMs) = GenerateDocumentModel(docxStream);

            // We need docxMs to stay open during generation if it's used?
            // Actually GenerateDocumentModel consumes docxStream and returns a new MemoryStream if needed,
            // but QuestPDF generation happens here.

            // Let's refine: GenerateDocumentModel should handle the DocX loading and QuestPDF definition.
            // But QuestPDF generation needs the DocX object references.

            // To properly refactor without breaking scoping:
            // I'll implement a full "ConvertToPdfStream" separately reusing a private "PrepareDocument" method.

            using (docxMs) // Ensure disposal of the memory stream created in helper
            {
                result.InsuredName = metadata.InsuredName;
                result.PolicyNumber = metadata.PolicyNumber;

                // Generate filename
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var sanitizedPolicyNumber = string.IsNullOrEmpty(result.PolicyNumber)
                    ? "NOPOLICY"
                    : string.Join("_", result.PolicyNumber.Split(Path.GetInvalidFileNameChars()));

                var generatedFileName = $"{originalFileName}_{sanitizedPolicyNumber}_{timestamp}.pdf";
                var outputPath = Path.Combine(outputDirectory, generatedFileName);

                pdfDocument.GeneratePdf(outputPath);

                result.Success = true;
                result.PdfPath = Path.GetFullPath(outputPath);
                result.Message = "PDF created successfully";
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = $"Conversion failed: {ex.Message}";
            result.Error = ex.ToString();
        }
        return result;
    }

    /// <summary>
    /// Convert a Word document stream to PDF and return the PDF file stream
    /// </summary>
    public (ConversionResult Result, MemoryStream PdfStream) ConvertToPdfStream(Stream docxStream, string originalFileName)
    {
        var result = new ConversionResult();
        MemoryStream pdfStream = new MemoryStream();

        try
        {
             var (pdfDocument, metadata, docxMs) = GenerateDocumentModel(docxStream);
             using (docxMs)
             {
                 result.InsuredName = metadata.InsuredName;
                 result.PolicyNumber = metadata.PolicyNumber;

                 pdfDocument.GeneratePdf(pdfStream);
                 pdfStream.Position = 0;

                 result.Success = true;
                 result.Message = "PDF generated in memory";
                 result.PdfPath = "MEMORY_STREAM";
             }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = $"Conversion failed: {ex.Message}";
            result.Error = ex.ToString();
        }

        return (result, pdfStream);
    }

    private (QuestPDF.Infrastructure.IDocument pdfDocument, ConversionResult metadata, MemoryStream docxMs) GenerateDocumentModel(Stream inputDocxStream)
    {
        var ms = new MemoryStream();
        inputDocxStream.CopyTo(ms);
        ms.Position = 0;

        // Sanitize
        Console.WriteLine("Starting document sanitization...");
        try
        {
            SanitizeDocument(ms);
            ms.Position = 0;
            Console.WriteLine("Document sanitization completed.");
        }
        catch (Exception ex)
        {
             Console.WriteLine($"Sanitization FAILED: {ex.Message}");
             ms.Position = 0;
        }

        // Load Numbering & Sections
        ms.Position = 0;
        LoadNumberingDefinitions(ms);
        ms.Position = 0;
        LoadSectionProperties(ms);
        ms.Position = 0;

        // Load DocX
        var wordDocument = DocX.Load(ms);

        // Metadata
        var metadata = new ConversionResult();
        ExtractDocumentMetadata(wordDocument, metadata);

        // Track tables
        var allTables = wordDocument.Tables.ToList();
        var paragraphsInTables = new HashSet<Paragraph>();
        foreach (var table in allTables)
        {
            foreach (var p in table.Paragraphs)
            {
                paragraphsInTables.Add(p);
            }
        }

        // Create QuestPDF Document
        var pdf = QuestPDF.Fluent.Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(2, Unit.Centimetre);
                page.PageColor(Colors.White);
                page.DefaultTextStyle(x => x.FontSize(11).FontFamily("Arial"));

                // Header
                page.Header().Column(headerCol =>
                {
                    // First Page Header
                    headerCol.Item().ShowIf(ctx => ctx.PageNumber == 1 && _sectionProps.TitlePg && !string.IsNullOrEmpty(_sectionProps.FirstPageHeaderId))
                        .Text(_headerFooterContent.GetValueOrDefault(_sectionProps.FirstPageHeaderId ?? "", ""))
                        .FontSize(9).AlignRight();

                    // Default Header (Page 2+)
                    headerCol.Item().ShowIf(ctx => ctx.PageNumber > 1 || (!_sectionProps.TitlePg))
                        .Row(row =>
                        {
                            row.RelativeItem().Text(text =>
                            {
                                if (_sectionProps.HeaderId != null && _headerFooterContent.TryGetValue(_sectionProps.HeaderId, out var hText))
                                {
                                    text.Span(hText).FontSize(9);
                                }
                            });
                            row.RelativeItem().AlignRight().Text($"POLICY NO: {metadata.PolicyNumber}").FontSize(9);
                        });
                });

                // Footer
                page.Footer().Column(footerColumn =>
                {
                    footerColumn.Item().PaddingBottom(4).LineHorizontal(0.5f).LineColor(Colors.Black);

                    // Dynamic Word Footer Content
                    footerColumn.Item().Row(row =>
                    {
                        row.RelativeItem().AlignLeft().Column(c =>
                        {
                            c.Item().ShowIf(ctx => ctx.PageNumber == 1 && _sectionProps.TitlePg)
                                .Text(_headerFooterContent.GetValueOrDefault(_sectionProps.FirstPageFooterId ?? "", "ZENITH GENERAL INSURANCE CO. LTD")).FontSize(8);

                            c.Item().ShowIf(ctx => ctx.PageNumber > 1 || (!_sectionProps.TitlePg))
                                .Text(_headerFooterContent.GetValueOrDefault(_sectionProps.FooterId ?? "", "ZENITH GENERAL INSURANCE CO. LTD")).FontSize(8);
                        });

                        row.RelativeItem().AlignCenter().Text(text =>
                        {
                            text.CurrentPageNumber().FontSize(8);
                        });

                        row.RelativeItem().AlignRight().Text(metadata.InsuredName).FontSize(8);
                    });

                    footerColumn.Item().PaddingTop(4).AlignCenter()
                        .Text("Authorised and regulated by the National Insurance Commission [RIC-048]")
                        .FontSize(7);
                });

                // Content
                page.Content().Column(column =>
                {
                    RenderContent(column, wordDocument, paragraphsInTables, allTables);
                });
            });
        });

        return (pdf, metadata, ms);
    }

    private void ExtractDocumentMetadata(DocX wordDocument, ConversionResult result)
    {
        foreach (var p in wordDocument.Paragraphs)
        {
            var text = p.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(text)) continue;

            if (string.IsNullOrEmpty(result.InsuredName) && text.StartsWith("INSURED:", StringComparison.OrdinalIgnoreCase))
            {
                var parts = text.Split(':', 2);
                if (parts.Length == 2) result.InsuredName = parts[1].Trim();
            }

            if (string.IsNullOrEmpty(result.PolicyNumber))
            {
                if (text.StartsWith("POLICY NO:", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith("POLICY NUMBER:", StringComparison.OrdinalIgnoreCase))
                {
                    var parts = text.Split(':', 2);
                    if (parts.Length == 2) result.PolicyNumber = parts[1].Trim();
                }

            }

            if (!string.IsNullOrEmpty(result.InsuredName) && !string.IsNullOrEmpty(result.PolicyNumber))
                break;
        }
    }

    private void SanitizeDocument(Stream stream)
    {
        // We need to edit the zip archive.
        // Important: We must keep the stream open and seekable.
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, true))
        {
            // 1. Identify altChunks to remove by looking at document.xml
            var documentEntry = archive.GetEntry("word/document.xml");
            var altChunkIds = new HashSet<string>();

            if (documentEntry != null)
            {
                XDocument doc;
                using (var entryStream = documentEntry.Open())
                {
                    doc = XDocument.Load(entryStream);
                }

                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                var altChunkElements = doc.Descendants(w + "altChunk").ToList();

                foreach (var ac in altChunkElements)
                {
                     var rId = ac.Attribute(r + "id")?.Value;
                     if (!string.IsNullOrEmpty(rId))
                     {
                         altChunkIds.Add(rId);
                     }
                }

                if (altChunkElements.Any())
                {
                    Console.WriteLine($"Found {altChunkElements.Count} altChunks in document.xml. Removing...");
                    altChunkElements.Remove();

                    using (var entryStream = documentEntry.Open())
                    {
                        entryStream.SetLength(0);
                        doc.Save(entryStream);
                    }
                }
            }

            // 2. Remove relationships and catch targets
            var targetsToRemove = new HashSet<string>();
            var relsEntry = archive.GetEntry("word/_rels/document.xml.rels");
            if (relsEntry != null)
            {
                XDocument relsDoc;
                using (var entryStream = relsEntry.Open())
                {
                    relsDoc = XDocument.Load(entryStream);
                }

                XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

                // Find relationships that match our altChunk Ids OR have the aFChunk type
                var badRels = relsDoc.Descendants(rel + "Relationship")
                    .Where(r =>
                        (r.Attribute("Id")?.Value != null && altChunkIds.Contains(r.Attribute("Id")!.Value)) ||
                        r.Attribute("Type")?.Value.Contains("aFChunk") == true
                     )
                    .ToList();

                foreach (var br in badRels)
                {
                    var target = br.Attribute("Target")?.Value;
                    if (!string.IsNullOrEmpty(target))
                    {
                        // Targets are usually relative, e.g., "afchunk2.docx" or "/word/afchunk2.docx"
                        // We need to normalize to find the zip entry.
                        // Assuming they are in 'word/' folder if they don't start with /
                        targetsToRemove.Add(target);
                    }
                }

                if (badRels.Any())
                {
                    Console.WriteLine($"Removing {badRels.Count} relationships.");
                    badRels.Remove();

                    using (var entryStream = relsEntry.Open())
                    {
                        entryStream.SetLength(0);
                        relsDoc.Save(entryStream);
                    }
                }
            }

            // 3. Delete the actual chunk files from the archive
            foreach (var target in targetsToRemove)
            {
                 // Handle path variations
                 string entryName = target.TrimStart('/');
                 if (!entryName.StartsWith("word/") && !entryName.Contains("/"))
                 {
                     entryName = "word/" + entryName;
                 }

                 var chunkEntry = archive.GetEntry(entryName);
                 if (chunkEntry != null)
                 {
                     Console.WriteLine($"Deleting entry: {entryName}");
                     chunkEntry.Delete();
                 }
            }

            // 4. Remove from [Content_Types].xml
            var contentTypesEntry = archive.GetEntry("[Content_Types].xml");
            if (contentTypesEntry != null)
            {
                XDocument contentTypesDoc;
                using (var entryStream = contentTypesEntry.Open())
                {
                    contentTypesDoc = XDocument.Load(entryStream);
                }

                XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
                var badTypes = contentTypesDoc.Descendants(ns + "Override")
                    .Where(t => t.Attribute("PartName")?.Value.Contains("afchunk") == true)
                    .ToList();

                 badTypes.AddRange(contentTypesDoc.Descendants(ns + "Default")
                    .Where(t => t.Attribute("Extension")?.Value.Contains("afchunk") == true));

                if (badTypes.Any())
                {
                     Console.WriteLine($"Removing {badTypes.Count} content types.");
                     badTypes.Remove();
                     using (var entryStream = contentTypesEntry.Open())
                     {
                         entryStream.SetLength(0);
                         contentTypesDoc.Save(entryStream);
                     }
                }
            }
        }
    }



    private void RenderContent(ColumnDescriptor column, DocX wordDocument, HashSet<Paragraph> paragraphsInTables, List<Table> allTables)
    {
        _numCounters.Clear();
        bool insideImportantSection = false;
        List<Paragraph> importantParagraphs = new();

        foreach (var paragraph in wordDocument.Paragraphs)
        {
            if (paragraphsInTables.Contains(paragraph))
                continue;

            var paragraphText = paragraph.Text?.Trim() ?? string.Empty;

            // Forced page breaks for major sections
            foreach (var section in NewPageSections)
            {
                if (paragraphText.StartsWith(section, StringComparison.OrdinalIgnoreCase))
                {
                    if (insideImportantSection && importantParagraphs.Count > 0)
                    {
                        RenderImportantBox(column, importantParagraphs);
                        importantParagraphs.Clear();
                        insideImportantSection = false;
                    }
                    column.Item().PageBreak();
                    break;
                }
            }

            // Page breaks (form feed characters)
            if (paragraph.Text?.Contains("\f") == true || paragraph.Text?.Contains("\x0C") == true)
            {
                if (insideImportantSection && importantParagraphs.Count > 0)
                {
                    RenderImportantBox(column, importantParagraphs);
                    importantParagraphs.Clear();
                    insideImportantSection = false;
                }

                column.Item().PageBreak();
                if (string.IsNullOrWhiteSpace(paragraph.Text?.Replace("\f", "").Replace("\x0C", "")))
                    continue;
            }

            // Images
            IList<Picture> pictures = null;
            bool manualExtractionNeeded = false;
            try
            {
               // This property getter can throw OverflowException for corrupt images
               pictures = paragraph.Pictures;
            }
            catch
            {
                manualExtractionNeeded = true;
            }

            // Normal processing
            if (!manualExtractionNeeded && pictures != null && pictures.Count > 0)
            {
                foreach (var picture in pictures)
                {
                    try
                    {
                        var imageId = picture.Id;
                        ProcessImage(column, wordDocument, imageId, null, null);
                    }
                    catch { /* Skip failed images */ }
                }

                if (string.IsNullOrWhiteSpace(paragraphText))
                    continue;
            }
            // Fallback: Manual Extraction from XML
            else if (manualExtractionNeeded)
            {
                 try
                 {
                     var manualImages = ManuallyExtractImages(paragraph);
                     if (manualImages.Count > 0)
                     {
                         foreach (var img in manualImages)
                         {
                             ProcessImage(column, wordDocument, img.Id, img.Width, img.Height);
                         }

                         if (string.IsNullOrWhiteSpace(paragraphText))
                            continue;
                     }
                 }
                 catch (Exception ex)
                 {
                     Console.WriteLine($"Manual image extraction failed: {ex.Message}");
                 }
            }

            // Empty paragraphs
            if (string.IsNullOrWhiteSpace(paragraphText))
            {
                float spacingAfter = GetParagraphSpacingAfter(paragraph);
                column.Item().Height(spacingAfter > 0 ? spacingAfter : 6);
                continue;
            }

            // Detect IMPORTANT section
            if (paragraphText.Equals("IMPORTANT", StringComparison.OrdinalIgnoreCase))
            {
                insideImportantSection = true;
                importantParagraphs.Add(paragraph);
                continue;
            }

            // Collect IMPORTANT section content
            if (insideImportantSection)
            {
                bool isEndOfImportant =
                    paragraphText.StartsWith("Followed by", StringComparison.OrdinalIgnoreCase) ||
                    paragraphText.StartsWith("PLEASE NOTE", StringComparison.OrdinalIgnoreCase) ||
                    IsInsuranceTypeHeading(paragraphText);

                if (isEndOfImportant)
                {
                    RenderImportantBox(column, importantParagraphs);
                    importantParagraphs.Clear();
                    insideImportantSection = false;
                }
                else
                {
                    importantParagraphs.Add(paragraph);
                    continue;
                }
            }

            // Label-value pairs
            if (IsLabelValuePair(paragraphText))
            {
                RenderLabelValuePair(column, paragraph);
                continue;
            }

            // Period of insurance handling
            if (paragraphText.StartsWith("PERIOD OF", StringComparison.OrdinalIgnoreCase))
            {
                column.Item().PaddingTop(8).PaddingBottom(4).Row(row =>
                {
                    row.ConstantItem(140).Text(text => text.Span("PERIOD OF").Bold());
                    row.RelativeItem();
                });
                continue;
            }

            if (paragraphText.StartsWith("INSURANCE:", StringComparison.OrdinalIgnoreCase))
            {
                var value = paragraphText.Substring(10).Trim();
                column.Item().PaddingBottom(4).Row(row =>
                {
                    row.ConstantItem(140).Text(text => text.Span("INSURANCE:").Bold());
                    row.RelativeItem().Text(value);
                });
                continue;
            }

            if (paragraphText.StartsWith("FROM:", StringComparison.OrdinalIgnoreCase))
            {
                var value = paragraphText.Substring(5).Trim();
                column.Item().PaddingBottom(2).PaddingLeft(140).Text($"FROM: {value}");
                continue;
            }

            if (paragraphText.StartsWith("TO:", StringComparison.OrdinalIgnoreCase) && paragraphText.Length < 50)
            {
                var value = paragraphText.Substring(3).Trim();
                column.Item().PaddingBottom(4).PaddingLeft(140).Text($"TO: {value}");
                continue;
            }

            // Heading detection
            bool isStyleHeading = paragraph.StyleId?.Contains("Heading") == true;
            bool isValidHeading = IsValidHeading(paragraphText, paragraph, isStyleHeading);

            if (isValidHeading)
            {

                var alignment = paragraph.Alignment;
                float fontSize = 14;
                float paddingTop = GetParagraphSpacingBefore(paragraph);
                float paddingBottom = GetParagraphSpacingAfter(paragraph);
                if (paddingTop <= 0) paddingTop = 10;
                if (paddingBottom <= 0) paddingBottom = 6;

                if (isStyleHeading && paragraph.StyleId != null)
                {
                    if (paragraph.StyleId.Contains("1")) fontSize = 16;
                    else if (paragraph.StyleId.Contains("2")) fontSize = 14;
                    else if (paragraph.StyleId.Contains("3")) fontSize = 13;
                    else fontSize = 12;
                }
                else if (IsInsuranceTypeHeading(paragraphText))
                {
                    fontSize = 16;
                    if (paddingTop < 16) paddingTop = 16;
                    if (paddingBottom < 12) paddingBottom = 12;
                    alignment = Alignment.center;
                }
                else if (paragraphText.StartsWith("MEMO", StringComparison.OrdinalIgnoreCase))
                {
                    fontSize = 13;
                    paddingTop = 12;
                    paddingBottom = 6;
                }

                column.Item().PaddingTop(paddingTop).PaddingBottom(paddingBottom).Text(text =>
                {
                    text.Span(paragraphText).Bold().FontSize(fontSize);
                    ApplyAlignment(text, alignment);
                });
                continue;
            }

            // Section headings (bold, short)
            bool isBold = paragraph.MagicText.Count > 0 &&
                          paragraph.MagicText.All(r => r.formatting?.Bold == true);
            bool isBoldShort = paragraphText.Length < 60 && isBold &&
                               !IsExcludedFromHeading(paragraphText) &&
                               !paragraphText.EndsWith(":-") &&
                               !paragraphText.EndsWith(":") &&
                               paragraph.MagicText.Count <= 3;

            if (isBoldShort && !paragraph.IsListItem && !IsAllCaps(paragraphText))
            {
                RenderParagraph(column, paragraph);
                continue;
            }

            // List items
            if (paragraph.IsListItem)
            {
                RenderListItem(column, paragraph);
                continue;
            }

            // Regular paragraphs
            RenderParagraph(column, paragraph);
        }

        // Render any remaining IMPORTANT section
        if (insideImportantSection && importantParagraphs.Count > 0)
        {
            RenderImportantBox(column, importantParagraphs);
        }

        // Tables
        foreach (var table in allTables)
        {
            RenderTable(column, table);
        }
    }

    // Helper method
    private static bool IsAllCaps(string text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        return text.All(c => char.IsUpper(c) || char.IsWhiteSpace(c) || char.IsPunctuation(c) || char.IsDigit(c));
    }

    private static bool IsInsuranceTypeHeading(string text)
    {
        foreach (var type in InsuranceTypes)
        {
            if (text.Equals(type, StringComparison.OrdinalIgnoreCase) ||
                text.StartsWith(type, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static bool IsExcludedFromHeading(string text)
    {
        foreach (var pattern in ExcludePatterns)
        {
            if (text.StartsWith(pattern, StringComparison.OrdinalIgnoreCase) ||
                text.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        if (text.EndsWith(":-") || text.EndsWith(": -"))
            return true;

        if (text.Length > 80)
            return true;

        return false;
    }

    private static bool IsValidHeading(string text, Paragraph paragraph, bool isStyleHeading)
    {
        if (isStyleHeading)
            return true;

        foreach (var heading in ValidHeadings)
        {
            if (text.StartsWith(heading, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        bool isAllCaps = IsAllCaps(text);
        bool isBold = paragraph.MagicText.Count > 0 &&
                      paragraph.MagicText.All(r => r.formatting?.Bold == true);
        bool isShort = text.Length < 50;
        bool isExcluded = IsExcludedFromHeading(text);

        if (isAllCaps && isBold && isShort && !isExcluded && !text.Contains(","))
            return true;

        return false;
    }

    private float GetParagraphSpacingBefore(Paragraph paragraph)
    {
        return ExtractParagraphStyle(paragraph).SpacingBefore ?? 0;
    }

    private float GetParagraphSpacingAfter(Paragraph paragraph)
    {
        return ExtractParagraphStyle(paragraph).SpacingAfter ?? 0;
    }

    private static bool IsLabelValuePair(string text)
    {
        string[] labels = { "INSURED:", "POLICY NUMBER:", "POLICY NO:" };
        foreach (var label in labels)
        {
            if (text.StartsWith(label, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private void RenderLabelValuePair(ColumnDescriptor column, Paragraph paragraph)
    {
        var text = paragraph.Text?.Trim() ?? string.Empty;
        var colonIdx = text.IndexOf(':');
        if (colonIdx < 0) return;

        var label = text.Substring(0, colonIdx + 1);
        var value = text.Substring(colonIdx + 1).Trim();

        column.Item().PaddingTop(6).PaddingBottom(4).Row(row =>
        {
            row.ConstantItem(140).Text(t => t.Span(label).Bold());
            row.RelativeItem().Text(t => t.Span(value).Bold());
        });
    }

    private struct ManualImageInfo
    {
        public string Id;
        public float? Width; // in points
        public float? Height; // in points
    }

    private List<ManualImageInfo> ManuallyExtractImages(Paragraph paragraph)
    {
        var results = new List<ManualImageInfo>();
        var xml = paragraph.Xml; // Get raw XML XElement if exposed, or we might need to rely on other properties.
        // Xceed Paragraph doesn't expose public XElement usually, but if it does:
        // Checking documentation: DocX elements usually have an Xml property (XElement).

        if (xml == null) return results;

        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

        // Find blips
        var blips = xml.Descendants(a + "blip");
        foreach (var blip in blips)
        {
            var embed = blip.Attribute(r + "embed")?.Value;
            if (!string.IsNullOrEmpty(embed))
            {
                float? width = null;
                float? height = null;

                // Try to find dimensions in parent 'xfrm' -> 'ext'
                // hierarchy: drawing -> inline -> graphic -> graphicData -> pic -> spPr -> xfrm -> ext
                var xfrm = blip.Ancestors(a + "graphic").FirstOrDefault()?.Descendants(a + "xfrm").FirstOrDefault();
                if (xfrm != null)
                {
                    var ext = xfrm.Element(a + "ext");
                    if (ext != null)
                    {
                         var cx = ext.Attribute("cx")?.Value;
                         var cy = ext.Attribute("cy")?.Value;

                         // EMU to Points: 1 pt = 12700 EMUs
                         if (long.TryParse(cx, out long cxVal)) width = cxVal * EMUS_TO_POINTS;
                         if (long.TryParse(cy, out long cyVal)) height = cyVal * EMUS_TO_POINTS;
                    }
                }

                results.Add(new ManualImageInfo { Id = embed, Width = width, Height = height });
            }
        }

        return results;
    }

    private struct ParagraphStyle
    {
        public Alignment? Alignment;
        public float? SpacingBefore;
        public float? SpacingAfter;
        public float? LineSpacing;
        public float LeftIndent;
        public float RightIndent;
        public float HangingIndent;
        public float FirstLineIndent;
    }

    private ParagraphStyle ExtractParagraphStyle(Paragraph paragraph)
    {
        var style = new ParagraphStyle
        {
            Alignment = paragraph.Alignment,
            LeftIndent = paragraph.IndentationBefore * TWIPS_TO_POINTS,
            RightIndent = paragraph.IndentationAfter * TWIPS_TO_POINTS,
            SpacingBefore = (float)paragraph.LineSpacingBefore * TWIPS_TO_POINTS,
            SpacingAfter = (float)paragraph.LineSpacingAfter * TWIPS_TO_POINTS
        };

        // Deep XML parsing for exact values (e.g. Line Spacing which library simplifies)
        var pPr = paragraph.Xml?.Element(XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main") + "pPr");
        if (pPr != null)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Indentation
            var ind = pPr.Element(w + "ind");
            if (ind != null)
            {
                if (float.TryParse(ind.Attribute(w + "left")?.Value, out float left)) style.LeftIndent = left * TWIPS_TO_POINTS;
                if (float.TryParse(ind.Attribute(w + "right")?.Value, out float right)) style.RightIndent = right * TWIPS_TO_POINTS;
                if (float.TryParse(ind.Attribute(w + "hanging")?.Value, out float hanging)) style.HangingIndent = hanging * TWIPS_TO_POINTS;
                if (float.TryParse(ind.Attribute(w + "firstLine")?.Value, out float first)) style.FirstLineIndent = first * TWIPS_TO_POINTS;
            }

            // Spacing
            var spacing = pPr.Element(w + "spacing");
            if (spacing != null)
            {
                if (float.TryParse(spacing.Attribute(w + "before")?.Value, out float before)) style.SpacingBefore = before * TWIPS_TO_POINTS;
                if (float.TryParse(spacing.Attribute(w + "after")?.Value, out float after)) style.SpacingAfter = after * TWIPS_TO_POINTS;

                var line = spacing.Attribute(w + "line")?.Value;
                var lineRule = spacing.Attribute(w + "lineRule")?.Value;
                if (float.TryParse(line, out float lineVal))
                {
                    // Interpretation depends on lineRule
                    if (lineRule == "atLeast" || lineRule == "exact") style.LineSpacing = lineVal * TWIPS_TO_POINTS;
                    else style.LineSpacing = lineVal / 240f * 12f; // Auto/Multiple
                }
            }

            // Justification
            var jc = pPr.Element(w + "jc");
            if (jc != null)
            {
                var val = jc.Attribute(w + "val")?.Value;
                if (val == "both") style.Alignment = Alignment.both;
            }
        }

        return style;
    }

    private struct RunStyle
    {
        public string FontName;
        public float? FontSize;
        public string Color;
        public bool Bold;
        public bool Italic;
        public bool Underline;
    }

    private RunStyle ExtractRunStyle(dynamic run)
    {
        var style = new RunStyle();

        try { style.FontName = run.FontFamily?.Name; } catch { }
        try { style.FontSize = (float?)run.FontSize; } catch { }
        try { style.Bold = run.Bold; } catch { }
        try { style.Italic = run.Italic; } catch { }
        try { style.Underline = run.UnderlineStyle != UnderlineStyle.none; } catch { }

        XElement? rPr = null;
        try { rPr = run.Xml?.Element(XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main") + "rPr"); } catch { }

        if (rPr != null)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Font
            var rFonts = rPr.Element(w + "rFonts");
            if (rFonts != null)
            {
                var font = rFonts.Attribute(w + "ascii")?.Value ?? rFonts.Attribute(w + "hAnsi")?.Value;
                if (!string.IsNullOrEmpty(font)) style.FontName = font;
            }

            // Size
            var sz = rPr.Element(w + "sz");
            if (sz != null)
            {
                if (float.TryParse(sz.Attribute(w + "val")?.Value, out float halfPts))
                    style.FontSize = halfPts / 2f;
            }

            // Color
            var color = rPr.Element(w + "color");
            if (color != null)
            {
                style.Color = color.Attribute(w + "val")?.Value;
            }

            // Styles
            if (rPr.Element(w + "b") != null) style.Bold = true;
            if (rPr.Element(w + "i") != null) style.Italic = true;
            if (rPr.Element(w + "u") != null) style.Underline = true;
        }

        return style;
    }

    private struct CellStyle
    {
        public int GridSpan;
        public bool IsVerticalMergeRestart;
        public bool IsVerticalMergeContinue;
        public string TopBorderColor;
        public float TopBorderSize;
        public string BottomBorderColor;
        public float BottomBorderSize;
        public string LeftBorderColor;
        public float LeftBorderSize;
        public string RightBorderColor;
        public float RightBorderSize;
    }

    private CellStyle ExtractCellStyle(Cell cell)
    {
        var style = new CellStyle
        {
            GridSpan = 1,
            TopBorderSize = 0.5f,
            BottomBorderSize = 0.5f,
            LeftBorderSize = 0.5f,
            RightBorderSize = 0.5f,
            TopBorderColor = "000000",
            BottomBorderColor = "000000",
            LeftBorderColor = "000000",
            RightBorderColor = "000000"
        };

        var tcPr = cell.Xml?.Element(XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main") + "tcPr");
        if (tcPr != null)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Grid Span (ColSpan)
            var gridSpan = tcPr.Element(w + "gridSpan");
            if (gridSpan != null && int.TryParse(gridSpan.Attribute(w + "val")?.Value, out int gs))
            {
                style.GridSpan = gs;
            }

            // Vertical Merge (RowSpan)
            var vMerge = tcPr.Element(w + "vMerge");
            if (vMerge != null)
            {
                var val = vMerge.Attribute(w + "val")?.Value;
                if (val == "restart") style.IsVerticalMergeRestart = true;
                else style.IsVerticalMergeContinue = true;
            }

            // Borders
            var tcBorders = tcPr.Element(w + "tcBorders");
            if (tcBorders != null)
            {
                ProcessBorder(tcBorders.Element(w + "top"), out style.TopBorderSize, out style.TopBorderColor);
                ProcessBorder(tcBorders.Element(w + "bottom"), out style.BottomBorderSize, out style.BottomBorderColor);
                ProcessBorder(tcBorders.Element(w + "left"), out style.LeftBorderSize, out style.LeftBorderColor);
                ProcessBorder(tcBorders.Element(w + "right"), out style.RightBorderSize, out style.RightBorderColor);
            }
        }

        return style;
    }

    private class NumberingLevel
    {
        public int LevelIndex;
        public string Start;
        public string NumberFormat;
        public string LevelText;
        public float Indent;
        public float Hanging;
    }

    private class NumberingDefinition
    {
        public int AbstractNumId;
        public Dictionary<int, NumberingLevel> Levels = new();
    }

    private Dictionary<int, int> _numIdToAbstractId = new();
    private Dictionary<int, NumberingDefinition> _abstractNumbering = new();

    private void LoadNumberingDefinitions(Stream docxStream)
    {
        _numIdToAbstractId.Clear();
        _abstractNumbering.Clear();

        try
        {
            using (var archive = new ZipArchive(docxStream, ZipArchiveMode.Read, true))
            {
                var entry = archive.GetEntry("word/numbering.xml");
                if (entry == null) return;

                using (var stream = entry.Open())
                {
                    var doc = XDocument.Load(stream);
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                    // Parse abstractNum
                    foreach (var abstractNum in doc.Descendants(w + "abstractNum"))
                    {
                        var idVal = abstractNum.Attribute(w + "abstractNumId")?.Value;
                        if (!int.TryParse(idVal, out int id)) continue;

                        var def = new NumberingDefinition { AbstractNumId = id };
                        foreach (var lvl in abstractNum.Descendants(w + "lvl"))
                        {
                            var ilvlVal = lvl.Attribute(w + "ilvl")?.Value;
                            if (!int.TryParse(ilvlVal, out int ilvl)) continue;

                            var level = new NumberingLevel
                            {
                                LevelIndex = ilvl,
                                Start = lvl.Element(w + "start")?.Attribute(w + "val")?.Value ?? "1",
                                NumberFormat = lvl.Element(w + "numFmt")?.Attribute(w + "val")?.Value ?? "decimal",
                                LevelText = lvl.Element(w + "lvlText")?.Attribute(w + "val")?.Value ?? ""
                            };

                            var pPr = lvl.Element(w + "pPr");
                            if (pPr != null)
                            {
                                var ind = pPr.Element(w + "ind");
                                if (ind != null)
                                {
                                    if (float.TryParse(ind.Attribute(w + "left")?.Value, out float left)) level.Indent = left * TWIPS_TO_POINTS;
                                    if (float.TryParse(ind.Attribute(w + "hanging")?.Value, out float hanging)) level.Hanging = hanging * TWIPS_TO_POINTS;
                                }
                            }
                            def.Levels[ilvl] = level;
                        }
                        _abstractNumbering[id] = def;
                    }

                    // Parse num (maps numId to abstractNumId)
                    foreach (var num in doc.Descendants(w + "num"))
                    {
                        var idVal = num.Attribute(w + "numId")?.Value;
                        var abstractIdVal = num.Element(w + "abstractNumId")?.Attribute(w + "val")?.Value;

                        if (int.TryParse(idVal, out int id) && int.TryParse(abstractIdVal, out int abstractId))
                        {
                            _numIdToAbstractId[id] = abstractId;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load numbering definitions: {ex.Message}");
        }
    }

    private Dictionary<int, Dictionary<int, int>> _numCounters = new();

    private struct SectionProperties
    {
        public bool TitlePg; // First page different
        public string? HeaderId;
        public string? FirstPageHeaderId;
        public string? FooterId;
        public string? FirstPageFooterId;
    }

    private SectionProperties _sectionProps;
    private Dictionary<string, string> _headerFooterContent = new();

    private void LoadSectionProperties(Stream docxStream)
    {
        _sectionProps = new SectionProperties();
        _headerFooterContent.Clear();

        try
        {
            using (var archive = new ZipArchive(docxStream, ZipArchiveMode.Read, true))
            {
                var documentEntry = archive.GetEntry("word/document.xml");
                if (documentEntry == null) return;

                using (var stream = documentEntry.Open())
                {
                    var doc = XDocument.Load(stream);
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                    var sectPr = doc.Descendants(w + "sectPr").LastOrDefault();
                    if (sectPr != null)
                    {
                        _sectionProps.TitlePg = sectPr.Element(w + "titlePg") != null;

                        foreach (var hdr in sectPr.Elements(w + "headerReference"))
                        {
                            var type = hdr.Attribute(w + "type")?.Value;
                            var id = hdr.Attribute(r + "id")?.Value;
                            if (type == "default") _sectionProps.HeaderId = id;
                            else if (type == "first") _sectionProps.FirstPageHeaderId = id;
                        }

                        foreach (var ftr in sectPr.Elements(w + "footerReference"))
                        {
                            var type = ftr.Attribute(w + "type")?.Value;
                            var id = ftr.Attribute(r + "id")?.Value;
                            if (type == "default") _sectionProps.FooterId = id;
                            else if (type == "first") _sectionProps.FirstPageFooterId = id;
                        }
                    }
                }

                // Load Relationship mapping to find filenames
                var relsEntry = archive.GetEntry("word/_rels/document.xml.rels");
                if (relsEntry != null)
                {
                    using (var relsStream = relsEntry.Open())
                    {
                        var relsDoc = XDocument.Load(relsStream);
                        XNamespace relsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

                        var relMap = relsDoc.Descendants(relsNs + "Relationship")
                            .ToDictionary(x => x.Attribute("Id")?.Value ?? "", x => x.Attribute("Target")?.Value ?? "");

                        // Load actual content
                        LoadHdrFtrContent(archive, relMap, _sectionProps.HeaderId);
                        LoadHdrFtrContent(archive, relMap, _sectionProps.FirstPageHeaderId);
                        LoadHdrFtrContent(archive, relMap, _sectionProps.FooterId);
                        LoadHdrFtrContent(archive, relMap, _sectionProps.FirstPageFooterId);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load section properties: {ex.Message}");
        }
    }

    private void LoadHdrFtrContent(ZipArchive archive, Dictionary<string, string> relMap, string? id)
    {
        if (string.IsNullOrEmpty(id) || !relMap.TryGetValue(id, out var target)) return;

        var path = target.StartsWith("/") ? target.Substring(1) : "word/" + target;
        var entry = archive.GetEntry(path);
        if (entry == null) return;

        using (var stream = entry.Open())
        {
            var doc = XDocument.Load(stream);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Extract text content from the header/footer
            var text = string.Join(" ", doc.Descendants(w + "t").Select(t => t.Value));
            if (!string.IsNullOrWhiteSpace(text))
            {
                _headerFooterContent[id] = text;
            }
        }
    }

    private void RenderListItem(ColumnDescriptor column, Paragraph paragraph)
    {
        int numId = paragraph.Xml?.Descendants(XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main") + "numPr")
                             .Elements(XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main") + "numId")
                             .Select(e => int.TryParse(e.Attribute(XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main") + "val")?.Value, out int id) ? id : 0)
                             .FirstOrDefault() ?? 0;

        int ilvl = paragraph.IndentLevel ?? 0;

        NumberingLevel? lvlDef = null;
        if (numId > 0 && _numIdToAbstractId.TryGetValue(numId, out int abstractId))
        {
            if (_abstractNumbering.TryGetValue(abstractId, out var def) && def.Levels.TryGetValue(ilvl, out var level))
            {
                lvlDef = level;
            }
        }

        // Increment counter
        int counter = 1;
        if (numId > 0)
        {
            if (!_numCounters.ContainsKey(numId)) _numCounters[numId] = new();
            if (!_numCounters[numId].ContainsKey(ilvl)) _numCounters[numId][ilvl] = 0;
            _numCounters[numId][ilvl]++;
            counter = _numCounters[numId][ilvl];

            // Reset sub-levels
            for (int i = ilvl + 1; i < 10; i++)
            {
                if (_numCounters[numId].ContainsKey(i)) _numCounters[numId][i] = 0;
            }
        }

        string marker = GetListMarker(lvlDef, counter);
        float indent = lvlDef?.Indent ?? (ilvl * 20f);
        float hanging = lvlDef?.Hanging ?? 15f;

        column.Item().PaddingTop(3).PaddingBottom(3).PaddingLeft(indent).Row(row =>
        {
            row.ConstantItem(hanging).Text(marker).FontSize(11);
            row.RelativeItem().Text(text =>
            {
                ProcessTextRuns(text, paragraph);
                ApplyAlignment(text, paragraph.Alignment);
            });
        });
    }

    private string GetListMarker(NumberingLevel? lvl, int counter)
    {
        if (lvl == null) return "•";

        var format = lvl.NumberFormat;
        var text = lvl.LevelText;

        if (format == "bullet")
        {
            if (text == "o") return "○";
            if (text == "·") return "•";
            return text.Length > 0 ? text[0].ToString() : "•";
        }

        string value = counter.ToString();
        if (format == "lowerLetter") value = ((char)('a' + (counter - 1))).ToString();
        else if (format == "upperLetter") value = ((char)('A' + (counter - 1))).ToString();
        else if (format == "lowerRoman") value = ToRoman(counter).ToLower();
        else if (format == "upperRoman") value = ToRoman(counter);

        if (!string.IsNullOrEmpty(text))
        {
            return text.Replace($"%{lvl.LevelIndex + 1}", value);
        }

        return value + ".";
    }

    private string ToRoman(int number)
    {
        if (number <= 0) return number.ToString();
        if (number >= 1000) return "M" + ToRoman(number - 1000);
        if (number >= 900) return "CM" + ToRoman(number - 900);
        if (number >= 500) return "D" + ToRoman(number - 500);
        if (number >= 400) return "CD" + ToRoman(number - 400);
        if (number >= 100) return "C" + ToRoman(number - 100);
        if (number >= 90) return "XC" + ToRoman(number - 90);
        if (number >= 50) return "L" + ToRoman(number - 50);
        if (number >= 40) return "XL" + ToRoman(number - 40);
        if (number >= 10) return "X" + ToRoman(number - 10);
        if (number >= 9) return "IX" + ToRoman(number - 9);
        if (number >= 5) return "V" + ToRoman(number - 5);
        if (number >= 4) return "IV" + ToRoman(number - 4);
        if (number >= 1) return "I" + ToRoman(number - 1);
        return string.Empty;
    }

    private void RenderTable(ColumnDescriptor parentColumn, Table table)
    {
        int totalGridColumns = 0;
        if (table.Rows.Count > 0)
        {
            totalGridColumns = table.Rows[0].Cells.Sum(c => ExtractCellStyle(c).GridSpan);
        }

        if (totalGridColumns <= 0) return;

        parentColumn.Item().PaddingVertical(8).Table(tableElement =>
        {
            tableElement.ColumnsDefinition(columns =>
            {
                for (int i = 0; i < totalGridColumns; i++)
                {
                    columns.RelativeColumn();
                }
            });

            // Track vertical merges: grid column index -> (rowSpanCount, cellElement)
            // Since QuestPDF RowSpan needs to know the count upfront, we have to look ahead.
            var rowOccupancy = new bool[table.Rows.Count + 1, totalGridColumns + 1];

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var row = table.Rows[rowIndex];
                int currentGridCol = 1;

                for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                {
                    var cell = row.Cells[cellIndex];
                    var style = ExtractCellStyle(cell);

                    // Find next available grid column
                    while (currentGridCol <= totalGridColumns && rowOccupancy[rowIndex + 1, currentGridCol])
                    {
                        currentGridCol++;
                    }
                    if (currentGridCol > totalGridColumns) break;

                    if (style.IsVerticalMergeRestart)
                    {
                        // Look ahead to find RowSpan
                        int rowSpan = 1;
                        for (int nextRowIdx = rowIndex + 1; nextRowIdx < table.Rows.Count; nextRowIdx++)
                        {
                            // We need to check the cell at the same horizontal position
                            // DocX doesn't make this easy. We'll check the cell and its XML
                            var nextRow = table.Rows[nextRowIdx];
                            // This is still complex because cell indices shift.
                            // For now, let's stick to the grid occupancy we can detect.
                        }

                        // Marking occupancy (simplification: if it's restart, we mark it.
                        // If it's continue, we mark it.)
                    }

                    // Mark occupancy for this cell's grid area
                    for (int i = 0; i < style.GridSpan; i++)
                    {
                        if (currentGridCol + i <= totalGridColumns)
                            rowOccupancy[rowIndex + 1, currentGridCol + i] = true;
                    }

                    if (style.IsVerticalMergeContinue)
                    {
                        currentGridCol += style.GridSpan;
                        continue;
                    }

                    var cellElement = tableElement.Cell()
                        .Row((uint)(rowIndex + 1))
                        .Column((uint)currentGridCol)
                        .ColumnSpan((uint)style.GridSpan);

                    QuestPDF.Infrastructure.IContainer container = cellElement;

                    // Apply Borders
                    if (style.TopBorderSize > 0) container = container.BorderTop(style.TopBorderSize).BorderColor("#" + style.TopBorderColor);
                    if (style.BottomBorderSize > 0) container = container.BorderBottom(style.BottomBorderSize).BorderColor("#" + style.BottomBorderColor);
                    if (style.LeftBorderSize > 0) container = container.BorderLeft(style.LeftBorderSize).BorderColor("#" + style.LeftBorderColor);
                    if (style.RightBorderSize > 0) container = container.BorderRight(style.RightBorderSize).BorderColor("#" + style.RightBorderColor);

                    container.Padding(5).Column(cellColumn =>
                    {
                        foreach (var cellParagraph in cell.Paragraphs)
                        {
                            RenderParagraph(cellColumn, cellParagraph);
                        }
                    });

                    currentGridCol += style.GridSpan;
                }
            }
        });
    }

    private void ProcessBorder(XElement? border, out float size, out string color)
    {
        size = 0.5f;
        color = "000000";
        if (border == null)
        {
            size = 0; // No border
            return;
        }

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var sz = border.Attribute(w + "sz")?.Value;
        if (float.TryParse(sz, out float szVal))
        {
            // Border size in eighths of a point
            size = szVal / 8f;
        }

        var col = border.Attribute(w + "color")?.Value;
        if (!string.IsNullOrEmpty(col) && col != "auto")
        {
            color = col;
        }
    }

    private void RenderParagraph(ColumnDescriptor column, Paragraph paragraph)
    {
        var style = ExtractParagraphStyle(paragraph);

        var item = column.Item();

        // Spacing
        float paddingTop = style.SpacingBefore ?? 3;
        float paddingBottom = style.SpacingAfter ?? 3;

        // Ensure some minimum spacing if it likely needs it (headings etc) handled by caller or style?
        // Let's trust the style for now.

        item.PaddingTop(paddingTop)
            .PaddingBottom(paddingBottom)
            .PaddingLeft(style.LeftIndent)
            .PaddingRight(style.RightIndent)
            .Text(text =>
            {
                // Handle FirstLineIndent/Hanging if possible
                // QuestPDF doesn't have an 'indent' property on text, so we might need a workaround for hanging.
                // For now, focus on standard flow.
                ProcessTextRuns(text, paragraph);
                ApplyAlignment(text, style.Alignment);
            });
    }

    private void ProcessImage(ColumnDescriptor column, DocX wordDocument, string imageId, float? width, float? height)
    {
         var imagePart = wordDocument.Images.FirstOrDefault(img => img.Id == imageId);
         if (imagePart != null)
         {
             using var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
             using var memoryStream = new MemoryStream();
             stream.CopyTo(memoryStream);
             var imageBytes = memoryStream.ToArray();

             if (imageBytes.Length > 0)
             {
                 var imgContainer = column.Item()
                     .PaddingVertical(8)
                     .AlignCenter();

                 // Apply dimensions if manually extracted
                 if (width.HasValue && width.Value > 0) imgContainer = imgContainer.Width(width.Value);
                 // if (height.HasValue && height.Value > 0) imgContainer = imgContainer.Height(height.Value); // QuestPDF usually handles aspect ratio if only width is set, or max width.
                 else imgContainer = imgContainer.MaxWidth(150);

                 imgContainer.Image(imageBytes);
             }
         }
    }

    private void RenderImportantBox(ColumnDescriptor column, List<Paragraph> paragraphs)
    {
        column.Item()
            .PaddingVertical(16)
            .Border(1)
            .BorderColor(Colors.Grey.Darken2)
            .Padding(20)
            .Column(boxColumn =>
            {
                foreach (var p in paragraphs)
                {
                    var text = p.Text?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(text)) continue;

                    if (text.Equals("IMPORTANT", StringComparison.OrdinalIgnoreCase))
                    {
                        boxColumn.Item().PaddingBottom(12).AlignCenter().Text(t =>
                        {
                            t.Span("IMPORTANT").Bold().FontSize(14);
                        });
                    }
                    else if (text.Contains("ZENITH GENERAL INSURANCE", StringComparison.OrdinalIgnoreCase))
                    {
                        boxColumn.Item().PaddingTop(8).AlignCenter().Text(t =>
                        {
                            t.Span(text).Bold().FontSize(11);
                        });
                    }
                    else if (text.StartsWith("OFFICE:", StringComparison.OrdinalIgnoreCase) ||
                             text.Contains("CIVIC TOWERS", StringComparison.OrdinalIgnoreCase) ||
                             text.Contains("OZUMBA MBADIWE", StringComparison.OrdinalIgnoreCase) ||
                             text.Contains("VICTORIA ISLAND", StringComparison.OrdinalIgnoreCase) ||
                             text.Contains("LAGOS", StringComparison.OrdinalIgnoreCase))
                    {
                        boxColumn.Item().AlignCenter().Text(t => t.Span(text).FontSize(9));
                    }
                    else if (text.StartsWith("Tel:", StringComparison.OrdinalIgnoreCase) ||
                             text.Contains("Fax:", StringComparison.OrdinalIgnoreCase))
                    {
                        boxColumn.Item().AlignCenter().Text(t => t.Span(text).FontSize(9));
                    }
                    else
                    {
                        boxColumn.Item().PaddingBottom(4).AlignCenter().Text(t =>
                        {
                            ProcessTextRunsForBox(t, p);
                        });
                    }
                }
            });
    }

    private void ProcessTextRunsForBox(TextDescriptor text, Paragraph paragraph)
    {
        foreach (var run in paragraph.MagicText)
        {
            if (string.IsNullOrEmpty(run.text)) continue;

            var style = ExtractRunStyle(run);
            var span = text.Span(run.text);

            if (style.Bold) span.Bold();
            if (style.Italic) span.Italic();
            if (style.Underline) span.Underline();

            if (style.FontSize.HasValue) span.FontSize(style.FontSize.Value);
            if (!string.IsNullOrEmpty(style.FontName)) span.FontFamily(style.FontName);
            if (!string.IsNullOrEmpty(style.Color)) span.FontColor("#" + style.Color);
        }
    }

    private void ApplyAlignment(TextDescriptor text, Alignment? alignment)
    {
        switch (alignment)
        {
            case Alignment.left:
                text.AlignLeft();
                break;
            case Alignment.center:
                text.AlignCenter();
                break;
            case Alignment.right:
                text.AlignRight();
                break;
            case Alignment.both:
                text.Justify();
                break;
            default:
                text.AlignLeft();
                break;
        }
    }

    private void ProcessTextRuns(TextDescriptor text, Paragraph paragraph)
    {
        foreach (var run in paragraph.MagicText)
        {
            if (string.IsNullOrEmpty(run.text)) continue;

            var style = ExtractRunStyle(run);
            var textParts = run.text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);

            for (int i = 0; i < textParts.Length; i++)
            {
                if (i > 0) text.Span("\n");

                var part = textParts[i];
                if (string.IsNullOrEmpty(part)) continue;

                var hyperlink = paragraph.Hyperlinks?.FirstOrDefault(h =>
                    !string.IsNullOrEmpty(h.Text) && h.Text.Contains(part));

                if (hyperlink != null && hyperlink.Uri != null)
                {
                    try
                    {
                        var uriString = hyperlink.Uri.ToString();
                        if (!string.IsNullOrEmpty(uriString))
                        {
                            var linkSpan = text.Hyperlink(part, uriString);
                            linkSpan.FontColor(Colors.Blue.Darken1);
                            linkSpan.Underline();

                            if (style.Bold) linkSpan.Bold();
                            if (style.Italic) linkSpan.Italic();
                            if (style.FontSize.HasValue) linkSpan.FontSize(style.FontSize.Value);
                            if (!string.IsNullOrEmpty(style.FontName)) linkSpan.FontFamily(style.FontName);

                            continue;
                        }
                    }
                    catch { }
                }

                var span = text.Span(part);

                if (style.Bold) span.Bold();
                if (style.Italic) span.Italic();
                if (style.Underline) span.Underline();

                if (style.FontSize.HasValue)
                {
                    span.FontSize(style.FontSize.Value);
                }

                if (!string.IsNullOrEmpty(style.FontName))
                {
                    span.FontFamily(style.FontName);
                }

                if (!string.IsNullOrEmpty(style.Color))
                {
                    var hex = style.Color;
                    if (hex.Length == 6 || hex.Length == 8)
                    {
                         span.FontColor("#" + hex);
                    }
                }
            }
        }
    }
}

/// <summary>
/// Result of a Word to PDF conversion
/// </summary>
public class ConversionResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public string? PdfPath { get; set; }
    public string? Error { get; set; }
    public string InsuredName { get; set; } = string.Empty;
    public string PolicyNumber { get; set; } = string.Empty;
}
