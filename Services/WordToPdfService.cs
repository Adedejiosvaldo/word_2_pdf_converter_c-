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
                page.Header()
                    .ShowIf(ctx => ctx.PageNumber > 1)
                    .AlignRight()
                    .Text($"POLICY NO: {metadata.PolicyNumber}")
                    .FontSize(9);

                // Footer
                page.Footer()
                    .ShowIf(ctx => ctx.PageNumber > 0)
                    .Column(footerColumn =>
                    {
                        footerColumn.Item().PaddingBottom(4).LineHorizontal(0.5f).LineColor(Colors.Black);

                        footerColumn.Item().Row(row =>
                        {
                            row.RelativeItem().AlignLeft().Text("ZENITH GENERAL INSURANCE CO. LTD").FontSize(8);
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
        int numberListCounter = 0;
        ListItemType? currentListType = null;
        int currentIndentLevel = -1;
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
                currentListType = null;
                numberListCounter = 0;

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
                float paddingTop = GetParagraphSpacingBefore(paragraph);
                float paddingBottom = GetParagraphSpacingAfter(paragraph);
                if (paddingTop <= 0) paddingTop = 10;
                if (paddingBottom <= 0) paddingBottom = 4;

                column.Item().PaddingTop(paddingTop).PaddingBottom(paddingBottom).Text(text =>
                {
                    ProcessTextRuns(text, paragraph);
                    ApplyAlignment(text, paragraph.Alignment);
                });
                continue;
            }

            // List items
            if (paragraph.IsListItem)
            {
                int indentLevel = paragraph.IndentLevel ?? 0;
                float indent = indentLevel * 20f;

                if (paragraph.ListItemType == ListItemType.Numbered)
                {
                    if (currentListType != ListItemType.Numbered || currentIndentLevel != indentLevel)
                    {
                        numberListCounter = 0;
                    }
                    currentListType = ListItemType.Numbered;
                    currentIndentLevel = indentLevel;
                    numberListCounter++;
                }
                else
                {
                    currentListType = ListItemType.Bulleted;
                    currentIndentLevel = indentLevel;
                }

                string marker = paragraph.ListItemType == ListItemType.Numbered
                    ? $"{numberListCounter})"
                    : "•";

                float listPaddingTop = GetParagraphSpacingBefore(paragraph);
                float listPaddingBottom = GetParagraphSpacingAfter(paragraph);
                if (listPaddingTop <= 0) listPaddingTop = 3;
                if (listPaddingBottom <= 0) listPaddingBottom = 3;

                column.Item().PaddingTop(listPaddingTop).PaddingBottom(listPaddingBottom).PaddingLeft(indent).Row(row =>
                {
                    row.ConstantItem(25).Text(marker).FontSize(11);
                    row.RelativeItem().Text(text =>
                    {
                        ProcessTextRuns(text, paragraph);
                        ApplyAlignment(text, paragraph.Alignment);
                    });
                });
                continue;
            }

            // Reset list counter
            if (currentListType != null)
            {
                currentListType = null;
                numberListCounter = 0;
                currentIndentLevel = -1;
            }

            // Regular paragraphs
            float regPaddingTop = GetParagraphSpacingBefore(paragraph);
            float regPaddingBottom = GetParagraphSpacingAfter(paragraph);
            if (regPaddingTop <= 0) regPaddingTop = 3;
            if (regPaddingBottom <= 0) regPaddingBottom = 3;

            float paragraphIndent = (paragraph.IndentLevel ?? 0) * 20f;

            var item = column.Item().PaddingTop(regPaddingTop).PaddingBottom(regPaddingBottom);
            if (paragraphIndent > 0)
            {
                item = item.PaddingLeft(paragraphIndent);
            }

            item.Text(text =>
            {
                ProcessTextRuns(text, paragraph);
                ApplyAlignment(text, paragraph.Alignment);
            });
        }

        // Render any remaining IMPORTANT section
        if (insideImportantSection && importantParagraphs.Count > 0)
        {
            RenderImportantBox(column, importantParagraphs);
        }

        // Tables
        foreach (var table in allTables)
        {
            int columnCount = table.ColumnCount;
            if (columnCount <= 0) continue;

            column.Item().PaddingVertical(8).Table(tableElement =>
            {
                tableElement.ColumnsDefinition(columns =>
                {
                    for (int i = 0; i < columnCount; i++)
                    {
                        columns.RelativeColumn();
                    }
                });

                foreach (var row in table.Rows)
                {
                    int cellIdx = 0;
                    foreach (var cell in row.Cells)
                    {
                        if (cellIdx >= columnCount) break;

                        var rowIndex = table.Rows.IndexOf(row);
                        tableElement.Cell()
                            .Row((uint)(rowIndex >= 0 ? rowIndex + 1 : 1))
                            .Column((uint)(cellIdx + 1))
                            .Border(0.5f)
                            .BorderColor(Colors.Black)
                            .Padding(5)
                            .Column(cellColumn =>
                            {
                                foreach (var cellParagraph in cell.Paragraphs)
                                {
                                    var cellText = cellParagraph.Text?.Trim() ?? string.Empty;
                                    if (string.IsNullOrEmpty(cellText)) continue;

                                    cellColumn.Item().Text(text =>
                                    {
                                        ProcessTextRuns(text, cellParagraph);
                                        ApplyAlignment(text, cellParagraph.Alignment);
                                    });
                                }
                            });
                        cellIdx++;
                    }
                }
            });
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

    private static float GetParagraphSpacingBefore(Paragraph paragraph)
    {
        try
        {
            var lineSpacing = paragraph.LineSpacingBefore;
            if (lineSpacing > 0) return (float)lineSpacing;
        }
        catch { }
        return 0;
    }

    private static float GetParagraphSpacingAfter(Paragraph paragraph)
    {
        try
        {
            var lineSpacing = paragraph.LineSpacingAfter;
            if (lineSpacing > 0) return (float)lineSpacing;
        }
        catch { }
        return 0;
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

    private static void RenderLabelValuePair(ColumnDescriptor column, Paragraph paragraph)
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
                         if (long.TryParse(cx, out long cxVal)) width = cxVal / 12700f;
                         if (long.TryParse(cy, out long cyVal)) height = cyVal / 12700f;
                    }
                }

                results.Add(new ManualImageInfo { Id = embed, Width = width, Height = height });
            }
        }

        return results;
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

    private static void RenderImportantBox(ColumnDescriptor column, List<Paragraph> paragraphs)
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

    private static void ProcessTextRunsForBox(TextDescriptor text, Paragraph paragraph)
    {
        foreach (var run in paragraph.MagicText)
        {
            if (string.IsNullOrEmpty(run.text)) continue;

            var span = text.Span(run.text);

            if (run.formatting?.Bold == true) span.Bold();
            if (run.formatting?.Italic == true) span.Italic();
            if (run.formatting?.UnderlineStyle == UnderlineStyle.singleLine) span.Underline();

            if (run.formatting?.Size.HasValue == true)
            {
                span.FontSize((float)run.formatting.Size.Value);
            }
        }
    }

    private static void ApplyAlignment(TextDescriptor text, Alignment? alignment)
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

    private static void ProcessTextRuns(TextDescriptor text, Paragraph paragraph)
    {
        foreach (var run in paragraph.MagicText)
        {
            if (string.IsNullOrEmpty(run.text)) continue;

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
                            if (run.formatting?.Bold == true) linkSpan.Bold();
                            if (run.formatting?.Italic == true) linkSpan.Italic();
                            if (run.formatting?.Size.HasValue == true)
                            {
                                linkSpan.FontSize((float)run.formatting.Size.Value);
                            }
                            continue;
                        }
                    }
                    catch { }
                }

                var span = text.Span(part);

                if (run.formatting?.Bold == true) span.Bold();
                if (run.formatting?.Italic == true) span.Italic();
                if (run.formatting?.UnderlineStyle == UnderlineStyle.singleLine) span.Underline();

                if (run.formatting?.FontFamily != null)
                {
                    span.FontFamily(run.formatting.FontFamily.ToString());
                }

                if (run.formatting?.FontColor.HasValue == true)
                {
                    var color = run.formatting.FontColor.Value;
                    span.FontColor(Color.FromARGB(color.A, color.R, color.G, color.B));
                }

                if (run.formatting?.Size.HasValue == true)
                {
                    span.FontSize((float)run.formatting.Size.Value);
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
