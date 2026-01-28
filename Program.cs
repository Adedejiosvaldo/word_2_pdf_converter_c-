using QuestPDF.Infrastructure;
using Xceed.Words.NET;
using Xceed.Document.NET;
using QuestPDF.Helpers;
using QuestPDF.Fluent;

QuestPDF.Settings.License = LicenseType.Community;
// QuestPDF.Settings.EnableDebugging = true;

Console.WriteLine("Word 2 PDF Converter");
Console.WriteLine("----------------------");

Console.WriteLine("Enter your file path");
string? filePath = Console.ReadLine();

if (!File.Exists(filePath))
{
    Console.WriteLine("No File found");
    return;
}

Console.WriteLine("Enter your file path to save");
string? pdfPath = Console.ReadLine();

if (string.IsNullOrEmpty(pdfPath))
{
    Console.WriteLine("Invalid Path");
    return;
}

if (!pdfPath.EndsWith(".pdf"))
{
    pdfPath += ".pdf";
}

Console.WriteLine("Conversion Starting...");
Console.WriteLine();

try
{
    using var wordDocument = DocX.Load(filePath);
    Console.WriteLine("Document Loaded");

    // ========================================
    // STEP 1: Extract dynamic values FIRST
    // ========================================
    string policyNumber = string.Empty;
    string insuredName = string.Empty;
    string periodFrom = string.Empty;
    string periodTo = string.Empty;

    foreach (var p in wordDocument.Paragraphs)
    {
        var text = p.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(text)) continue;

        // Extract INSURED value
        if (string.IsNullOrEmpty(insuredName) && text.StartsWith("INSURED:", StringComparison.OrdinalIgnoreCase))
        {
            var parts = text.Split(':', 2);
            if (parts.Length == 2) insuredName = parts[1].Trim();
        }

        // Extract POLICY NO value
        if (string.IsNullOrEmpty(policyNumber))
        {
            if (text.StartsWith("POLICY NO:", StringComparison.OrdinalIgnoreCase) ||
                text.StartsWith("POLICY NUMBER:", StringComparison.OrdinalIgnoreCase))
            {
                var parts = text.Split(':', 2);
                if (parts.Length == 2) policyNumber = parts[1].Trim();
            }
        }

        // Extract FROM date
        if (string.IsNullOrEmpty(periodFrom) && text.Contains("FROM:", StringComparison.OrdinalIgnoreCase))
        {
            var fromIdx = text.IndexOf("FROM:", StringComparison.OrdinalIgnoreCase);
            if (fromIdx >= 0)
            {
                var afterFrom = text.Substring(fromIdx + 5).Trim();
                // Take until next label or end
                var toIdx = afterFrom.IndexOf("TO:", StringComparison.OrdinalIgnoreCase);
                periodFrom = toIdx > 0 ? afterFrom.Substring(0, toIdx).Trim() : afterFrom.Trim();
            }
        }

        // Extract TO date
        if (string.IsNullOrEmpty(periodTo) && text.Contains("TO:", StringComparison.OrdinalIgnoreCase))
        {
            var toIdx = text.IndexOf("TO:", StringComparison.OrdinalIgnoreCase);
            if (toIdx >= 0)
            {
                periodTo = text.Substring(toIdx + 3).Trim();
            }
        }
    }

    Console.WriteLine($"Extracted - INSURED: {insuredName}");
    Console.WriteLine($"Extracted - POLICY NO: {policyNumber}");
    Console.WriteLine($"Extracted - FROM: {periodFrom}");
    Console.WriteLine($"Extracted - TO: {periodTo}");

    // ========================================
    // STEP 2: Track tables and special sections
    // ========================================
    var allTables = wordDocument.Tables.ToList();
    var paragraphsInTables = new HashSet<Paragraph>();
    foreach (var table in allTables)
    {
        foreach (var p in table.Paragraphs)
        {
            paragraphsInTables.Add(p);
        }
    }

    // ========================================
    // STEP 3: Create PDF Document
    // ========================================
    QuestPDF.Fluent.Document.Create(container =>
    {
        container.Page(page =>
        {
            page.Size(PageSizes.A4);
            page.Margin(2, Unit.Centimetre);
            page.PageColor(Colors.White);
            page.DefaultTextStyle(x => x.FontSize(11).FontFamily("Arial"));

            // HEADER: Policy number (page 2+ only)
            page.Header()
                .ShowIf(ctx => ctx.PageNumber > 1)
                .AlignRight()
                .Text($"POLICY NO: {policyNumber}")
                .FontSize(9);

            // FOOTER: Only show from page 1 (but hide line on page 1 if desired)
            page.Footer()
                .ShowIf(ctx => ctx.PageNumber > 0) // Show on all pages
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
                        row.RelativeItem().AlignRight().Text(insuredName).FontSize(8);
                    });

                    footerColumn.Item().PaddingTop(4).AlignCenter()
                        .Text("Authorised and regulated by the National Insurance Commission [RIC-048]")
                        .FontSize(7);
                });

            // CONTENT
            page.Content().Column(column =>
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

                // ========================================
                // PAGE BREAKS
                // ========================================
                if (paragraph.Text?.Contains("\f") == true || paragraph.Text?.Contains("\x0C") == true)
                {
                    // If we were collecting IMPORTANT paragraphs, render the box first
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

                // ========================================
                // IMAGES
                // ========================================
                if (paragraph.Pictures != null && paragraph.Pictures.Count > 0)
                {
                    foreach (var picture in paragraph.Pictures)
                    {
                        try
                        {
                            var imageId = picture.Id;
                            var imagePart = wordDocument.Images.FirstOrDefault(img => img.Id == imageId);

                            if (imagePart != null)
                            {
                                using var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                                using var memoryStream = new MemoryStream();
                                stream.CopyTo(memoryStream);
                                var imageBytes = memoryStream.ToArray();

                                if (imageBytes.Length > 0)
                                {
                                    column.Item()
                                        .PaddingVertical(8)
                                        .AlignCenter()
                                        .MaxWidth(150) // Logo size
                                        .Image(imageBytes);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not load image - {ex.Message}");
                        }
                    }

                    if (string.IsNullOrWhiteSpace(paragraphText))
                        continue;
                }

                // Skip empty paragraphs
                if (string.IsNullOrWhiteSpace(paragraphText))
                {
                    column.Item().Height(6);
                    continue;
                }

                // ========================================
                // DETECT IMPORTANT SECTION START
                // ========================================
                if (paragraphText.Equals("IMPORTANT", StringComparison.OrdinalIgnoreCase))
                {
                    insideImportantSection = true;
                    importantParagraphs.Add(paragraph);
                    continue;
                }

                // If inside IMPORTANT section, collect paragraphs until page break or specific markers
                if (insideImportantSection)
                {
                    // Check for end markers (like next main section or page break indicator)
                    bool isEndOfImportant = paragraphText.StartsWith("POLICY NO:", StringComparison.OrdinalIgnoreCase) ||
                                            paragraphText.Contains("HEALTHCARE PROFESSIONAL") ||
                                            paragraphText.Contains("PRODUCT LIABILITY") ||
                                            (paragraphText.All(c => char.IsUpper(c) || char.IsWhiteSpace(c)) &&
                                             paragraphText.Length > 20 &&
                                             !paragraphText.Contains("ZENITH"));

                    if (isEndOfImportant)
                    {
                        // Render the IMPORTANT box and continue with this paragraph normally
                        RenderImportantBox(column, importantParagraphs);
                        importantParagraphs.Clear();
                        insideImportantSection = false;
                        // Fall through to process this paragraph
                    }
                    else
                    {
                        importantParagraphs.Add(paragraph);
                        continue;
                    }
                }

                // ========================================
                // LABEL-VALUE PAIRS (INSURED:, POLICY NUMBER:, etc.)
                // ========================================
                if (IsLabelValuePair(paragraphText))
                {
                    RenderLabelValuePair(column, paragraph);
                    continue;
                }

                // ========================================
                // PERIOD OF INSURANCE (multi-line handling)
                // ========================================
                if (paragraphText.StartsWith("PERIOD OF", StringComparison.OrdinalIgnoreCase))
                {
                    column.Item().PaddingTop(8).PaddingBottom(4).Row(row =>
                    {
                        row.ConstantItem(140).Text(text =>
                        {
                            text.Span("PERIOD OF").Bold();
                        });
                        row.RelativeItem(); // Empty for this line
                    });
                    continue;
                }

                if (paragraphText.StartsWith("INSURANCE:", StringComparison.OrdinalIgnoreCase))
                {
                    var value = paragraphText.Substring(10).Trim();
                    column.Item().PaddingBottom(4).Row(row =>
                    {
                        row.ConstantItem(140).Text(text =>
                        {
                            text.Span("INSURANCE:").Bold();
                        });
                        row.RelativeItem().Text(value);
                    });
                    continue;
                }

                // Handle standalone FROM: and TO: lines
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

                // ========================================
                // MAIN TITLES (all caps, centered)
                // ========================================
                bool isStyleHeading = paragraph.StyleId?.Contains("Heading") == true;
                bool isAllCapsTitle = paragraphText.Length > 10 &&
                                      paragraphText.Length < 100 &&
                                      paragraphText.All(c => char.IsUpper(c) || char.IsWhiteSpace(c) || char.IsPunctuation(c)) &&
                                      !paragraphText.Contains(":");

                if (isStyleHeading || isAllCapsTitle)
                {
                    currentListType = null;
                    numberListCounter = 0;

                    var alignment = paragraph.Alignment;
                    float fontSize = 14;
                    float paddingTop = 10;
                    float paddingBottom = 6;

                    if (isStyleHeading && paragraph.StyleId != null)
                    {
                        if (paragraph.StyleId.Contains("1")) fontSize = 16;
                        else if (paragraph.StyleId.Contains("2")) fontSize = 14;
                        else if (paragraph.StyleId.Contains("3")) fontSize = 13;
                        else fontSize = 12;
                    }
                    else if (isAllCapsTitle)
                    {
                        fontSize = 16;
                        paddingTop = 16;
                        paddingBottom = 12;
                        alignment = Alignment.center;
                    }

                    column.Item().PaddingTop(paddingTop).PaddingBottom(paddingBottom).Text(text =>
                    {
                        text.Span(paragraphText).Bold().FontSize(fontSize);
                        ApplyAlignment(text, alignment);
                    });
                    continue;
                }

                // ========================================
                // SECTION HEADINGS (Bold, short)
                // ========================================
                bool isBoldShort = paragraphText.Length < 60 &&
                                   paragraph.MagicText.Count > 0 &&
                                   paragraph.MagicText.All(r => r.formatting?.Bold == true);

                if (isBoldShort && !paragraph.IsListItem)
                {
                    column.Item().PaddingTop(10).PaddingBottom(4).Text(text =>
                    {
                        text.Span(paragraphText).Bold().FontSize(12);
                        ApplyAlignment(text, paragraph.Alignment);
                    });
                    continue;
                }

                // ========================================
                // LIST ITEMS
                // ========================================
                if (paragraph.IsListItem)
                {
                    int indentLevel = paragraph.IndentLevel ?? 0;
                    float indent = indentLevel * 15f;

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

                    column.Item().PaddingTop(3).PaddingBottom(3).PaddingLeft(indent).Row(row =>
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

                // Reset list counter for non-list items
                if (currentListType != null)
                {
                    currentListType = null;
                    numberListCounter = 0;
                    currentIndentLevel = -1;
                }

                // ========================================
                // REGULAR PARAGRAPHS
                // ========================================
                column.Item().PaddingTop(3).PaddingBottom(3).Text(text =>
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

            // ========================================
            // TABLES
            // ========================================
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
                                            foreach (var run in cellParagraph.MagicText)
                                            {
                                                if (string.IsNullOrEmpty(run.text)) continue;
                                                var span = text.Span(run.text);
                                                if (run.formatting?.Bold == true) span.Bold();
                                                if (run.formatting?.Italic == true) span.Italic();
                                                if (run.formatting?.UnderlineStyle == UnderlineStyle.singleLine) span.Underline();
                                            }
                                            ApplyAlignment(text, cellParagraph.Alignment);
                                        });
                                    }
                                });
                            cellIdx++;
                        }
                    }
                });
            }
        });
        });
    }).GeneratePdf(pdfPath);

    string fullPath = Path.GetFullPath(pdfPath);
    Console.WriteLine($"✓ PDF saved to: {fullPath}");
    Console.WriteLine("Conversion completed successfully!");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.ToString()}");
}

Console.WriteLine("Press any key to exit");
try { Console.ReadKey(); } catch (InvalidOperationException) { }

// ========================================
// HELPER FUNCTIONS
// ========================================

static bool IsLabelValuePair(string text)
{
    // Check for common label patterns
    string[] labels = { "INSURED:", "POLICY NUMBER:", "POLICY NO:" };
    foreach (var label in labels)
    {
        if (text.StartsWith(label, StringComparison.OrdinalIgnoreCase))
            return true;
    }
    return false;
}

static void RenderLabelValuePair(ColumnDescriptor column, Paragraph paragraph)
{
    var text = paragraph.Text?.Trim() ?? string.Empty;
    var colonIdx = text.IndexOf(':');
    if (colonIdx < 0) return;

    var label = text.Substring(0, colonIdx + 1);
    var value = text.Substring(colonIdx + 1).Trim();

    column.Item().PaddingTop(6).PaddingBottom(4).Row(row =>
    {
        row.ConstantItem(140).Text(t => t.Span(label).Bold());
        row.RelativeItem().Text(t =>
        {
            t.Span(value).Bold();
        });
    });
}

static void RenderImportantBox(ColumnDescriptor column, List<Paragraph> paragraphs)
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

                // Title "IMPORTANT"
                if (text.Equals("IMPORTANT", StringComparison.OrdinalIgnoreCase))
                {
                    boxColumn.Item().PaddingBottom(12).AlignCenter().Text(t =>
                    {
                        t.Span("IMPORTANT").Bold().FontSize(14);
                    });
                }
                // Company name
                else if (text.Contains("ZENITH GENERAL INSURANCE", StringComparison.OrdinalIgnoreCase))
                {
                    boxColumn.Item().PaddingTop(8).AlignCenter().Text(t =>
                    {
                        t.Span(text).Bold().FontSize(11);
                    });
                }
                // Contact details (smaller, centered)
                else if (text.StartsWith("OFFICE:", StringComparison.OrdinalIgnoreCase) ||
                         text.StartsWith("Tel:", StringComparison.OrdinalIgnoreCase))
                {
                    boxColumn.Item().AlignCenter().Text(t =>
                    {
                        t.Span(text).FontSize(9);
                    });
                }
                // Regular text in box
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

static void ProcessTextRunsForBox(TextDescriptor text, Paragraph paragraph)
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

static void ApplyAlignment(TextDescriptor text, Alignment? alignment)
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

static void ProcessTextRuns(TextDescriptor text, Paragraph paragraph)
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
