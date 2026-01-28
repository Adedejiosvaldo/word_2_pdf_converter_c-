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

// ========================================
// SECTIONS THAT MUST START ON NEW PAGES
// ========================================
string[] newPageSections = {
    "POLICY SCHEDULE",
    "MEMORANDA ATTACHING TO AND FORMING PART OF"
};

// ========================================
// TEXT PATTERNS THAT ARE NOT HEADINGS
// ========================================
string[] notHeadingPatterns = {
    "IT IS AGREED",
    "PROVIDED THAT",
    "PROVIDED ALSO",
    "WHEREAS",
    "NOW THEREFORE",
    "IN WITNESS",
    "OZUMBA MBADIWE",
    "VICTORIA ISLAND",
    "LAGOS",
    "Followed by"
};

// ========================================
// VALID SECTION HEADINGS (exact or starts with)
// ========================================
string[] validHeadings = {
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

        if (string.IsNullOrEmpty(insuredName) && text.StartsWith("INSURED:", StringComparison.OrdinalIgnoreCase))
        {
            var parts = text.Split(':', 2);
            if (parts.Length == 2) insuredName = parts[1].Trim();
        }

        if (string.IsNullOrEmpty(policyNumber))
        {
            if (text.StartsWith("POLICY NO:", StringComparison.OrdinalIgnoreCase) ||
                text.StartsWith("POLICY NUMBER:", StringComparison.OrdinalIgnoreCase))
            {
                var parts = text.Split(':', 2);
                if (parts.Length == 2) policyNumber = parts[1].Trim();
            }
        }

        if (string.IsNullOrEmpty(periodFrom) && text.Contains("FROM:", StringComparison.OrdinalIgnoreCase))
        {
            var fromIdx = text.IndexOf("FROM:", StringComparison.OrdinalIgnoreCase);
            if (fromIdx >= 0)
            {
                var afterFrom = text.Substring(fromIdx + 5).Trim();
                var toIdx = afterFrom.IndexOf("TO:", StringComparison.OrdinalIgnoreCase);
                periodFrom = toIdx > 0 ? afterFrom.Substring(0, toIdx).Trim() : afterFrom.Trim();
            }
        }

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
    // STEP 2: Track tables
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

            // FOOTER: Show on all pages
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
                    // FORCED PAGE BREAKS FOR MAJOR SECTIONS
                    // ========================================
                    foreach (var section in newPageSections)
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

                    // ========================================
                    // PAGE BREAKS (form feed characters)
                    // ========================================
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
                                            .MaxWidth(150)
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

                    // ========================================
                    // EMPTY PARAGRAPHS
                    // ========================================
                    if (string.IsNullOrWhiteSpace(paragraphText))
                    {
                        float spacingAfter = GetParagraphSpacingAfter(paragraph);
                        column.Item().Height(spacingAfter > 0 ? spacingAfter : 6);
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

                    // ========================================
                    // COLLECT IMPORTANT SECTION CONTENT
                    // End when we hit "Followed by" or next major section
                    // ========================================
                    if (insideImportantSection)
                    {
                        // Check for end of IMPORTANT section
                        // End markers: "Followed by such further steps" or insurance type headings
                        bool isEndOfImportant =
                            paragraphText.StartsWith("Followed by", StringComparison.OrdinalIgnoreCase) ||
                            paragraphText.StartsWith("PLEASE NOTE", StringComparison.OrdinalIgnoreCase) ||
                            IsInsuranceTypeHeading(paragraphText);

                        if (isEndOfImportant)
                        {
                            // Render the IMPORTANT box
                            RenderImportantBox(column, importantParagraphs);
                            importantParagraphs.Clear();
                            insideImportantSection = false;
                            // Fall through to process "Followed by" as regular text
                        }
                        else
                        {
                            // Still inside IMPORTANT - collect this paragraph
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
                            row.RelativeItem();
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
                    // HEADING DETECTION (Improved)
                    // ========================================
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

                        // Determine font size based on heading type
                        if (isStyleHeading && paragraph.StyleId != null)
                        {
                            if (paragraph.StyleId.Contains("1")) fontSize = 16;
                            else if (paragraph.StyleId.Contains("2")) fontSize = 14;
                            else if (paragraph.StyleId.Contains("3")) fontSize = 13;
                            else fontSize = 12;
                        }
                        else if (IsInsuranceTypeHeading(paragraphText))
                        {
                            // Main insurance type heading
                            fontSize = 16;
                            if (paddingTop < 16) paddingTop = 16;
                            if (paddingBottom < 12) paddingBottom = 12;
                            alignment = Alignment.center;
                        }
                        else if (paragraphText.StartsWith("MEMO", StringComparison.OrdinalIgnoreCase))
                        {
                            // MEMO headings
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

                    // ========================================
                    // SECTION HEADINGS (Bold, short, not excluded)
                    // ========================================
                    bool isBold = paragraph.MagicText.Count > 0 &&
                                  paragraph.MagicText.All(r => r.formatting?.Bold == true);
                    bool isBoldShort = paragraphText.Length < 60 && isBold &&
                                       !IsExcludedFromHeading(paragraphText) &&
                                       !paragraphText.EndsWith(":-") &&
                                       !paragraphText.EndsWith(":") &&
                                       paragraph.MagicText.Count <= 3; // Short headings have few runs

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

                    // ========================================
                    // LIST ITEMS
                    // ========================================
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

static bool IsAllCaps(string text)
{
    if (string.IsNullOrEmpty(text)) return false;
    return text.All(c => char.IsUpper(c) || char.IsWhiteSpace(c) || char.IsPunctuation(c) || char.IsDigit(c));
}

static bool IsInsuranceTypeHeading(string text)
{
    // These are the main insurance type headings that should be centered
    string[] insuranceTypes = {
        "PRODUCT LIABILITY INSURANCE",
        "HEALTHCARE PROFESSIONAL INDEMNITY INSURANCE",
        "PROFESSIONAL INDEMNITY INSURANCE",
        "PUBLIC LIABILITY INSURANCE",
        "MOTOR INSURANCE",
        "FIRE INSURANCE"
    };

    foreach (var type in insuranceTypes)
    {
        if (text.Equals(type, StringComparison.OrdinalIgnoreCase) ||
            text.StartsWith(type, StringComparison.OrdinalIgnoreCase))
            return true;
    }
    return false;
}

static bool IsExcludedFromHeading(string text)
{
    // These patterns should NOT be treated as headings even if bold/all-caps
    string[] excludePatterns = {
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

    foreach (var pattern in excludePatterns)
    {
        if (text.StartsWith(pattern, StringComparison.OrdinalIgnoreCase) ||
            text.Contains(pattern, StringComparison.OrdinalIgnoreCase))
            return true;
    }

    // Exclude if it ends with ":-" (introduces a list)
    if (text.EndsWith(":-") || text.EndsWith(": -"))
        return true;

    // Exclude if it's too long to be a heading (> 80 chars usually means it's a paragraph)
    if (text.Length > 80)
        return true;

    return false;
}

static bool IsValidHeading(string text, Paragraph paragraph, bool isStyleHeading)
{
    // If Word says it's a heading style, trust it
    if (isStyleHeading)
        return true;

    // Check if it matches known valid headings
    string[] validHeadings = {
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

    foreach (var heading in validHeadings)
    {
        if (text.StartsWith(heading, StringComparison.OrdinalIgnoreCase))
            return true;
    }

    // Check if it's a short all-caps text that's bold and NOT excluded
    bool isAllCaps = IsAllCaps(text);
    bool isBold = paragraph.MagicText.Count > 0 &&
                  paragraph.MagicText.All(r => r.formatting?.Bold == true);
    bool isShort = text.Length < 50;
    bool isExcluded = IsExcludedFromHeading(text);

    if (isAllCaps && isBold && isShort && !isExcluded && !text.Contains(","))
        return true;

    return false;
}

static float GetParagraphSpacingBefore(Paragraph paragraph)
{
    try
    {
        var lineSpacing = paragraph.LineSpacingBefore;
        if (lineSpacing > 0) return (float)lineSpacing;
    }
    catch { }
    return 0;
}

static float GetParagraphSpacingAfter(Paragraph paragraph)
{
    try
    {
        var lineSpacing = paragraph.LineSpacingAfter;
        if (lineSpacing > 0) return (float)lineSpacing;
    }
    catch { }
    return 0;
}

static bool IsLabelValuePair(string text)
{
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
                // Company name (bold, centered)
                else if (text.Contains("ZENITH GENERAL INSURANCE", StringComparison.OrdinalIgnoreCase))
                {
                    boxColumn.Item().PaddingTop(8).AlignCenter().Text(t =>
                    {
                        t.Span(text).Bold().FontSize(11);
                    });
                }
                // Office/Address lines (smaller, centered)
                else if (text.StartsWith("OFFICE:", StringComparison.OrdinalIgnoreCase) ||
                         text.Contains("CIVIC TOWERS", StringComparison.OrdinalIgnoreCase) ||
                         text.Contains("OZUMBA MBADIWE", StringComparison.OrdinalIgnoreCase) ||
                         text.Contains("VICTORIA ISLAND", StringComparison.OrdinalIgnoreCase) ||
                         text.Contains("LAGOS", StringComparison.OrdinalIgnoreCase))
                {
                    boxColumn.Item().AlignCenter().Text(t =>
                    {
                        t.Span(text).FontSize(9);
                    });
                }
                // Tel/Fax line (smaller, centered)
                else if (text.StartsWith("Tel:", StringComparison.OrdinalIgnoreCase) ||
                         text.Contains("Fax:", StringComparison.OrdinalIgnoreCase))
                {
                    boxColumn.Item().AlignCenter().Text(t =>
                    {
                        t.Span(text).FontSize(9);
                    });
                }
                // Regular text in box (centered)
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
