using System.ComponentModel;
using QuestPDF.Infrastructure;
using Xceed.Words.NET;
using Xceed.Document.NET;
using QuestPDF.Helpers;
using QuestPDF.Fluent;

QuestPDF.Settings.License = LicenseType.Community;
QuestPDF.Settings.EnableDebugging = true; // WHY: Enable debugging to find layout constraint issues

// Terminal output
Console.WriteLine("Word 2 PDF Converter");
Console.WriteLine("----------------------");

// Get word file path
Console.WriteLine("Enter your file path");
string? filePath = Console.ReadLine();

// Check if file exists
if (!File.Exists(filePath))
{
    Console.WriteLine("No File found");
    return;
}

// A place to save
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
    using (var wordDocument = DocX.Load(filePath))
    {
        System.Console.WriteLine("Document Loaded");

        // creating a PDF
        QuestPDF.Fluent.Document.Create(container =>
        {
            container.Page(page =>
            {
                // page configurations
                page.Size(PageSizes.A4);
                page.Margin(2, Unit.Centimetre);
                page.PageColor(Colors.White);
                page.DefaultTextStyle(x => x.FontSize(12).FontFamily("Century Gothic"));

                // WHY: Enhanced footer with company information
                page.Footer().Column(footerColumn =>
                {
                    // WHY: Company name on the left
                    footerColumn.Item().AlignLeft().Text("ZENITH GENERAL INSURANCE CO. LTD").FontSize(8);
                    
                    // WHY: Page numbers in the center
                    footerColumn.Item().AlignCenter().Text(text =>
                    {
                        text.CurrentPageNumber().FontSize(8);
                        text.Span(" of ").FontSize(8);
                        text.TotalPages().FontSize(8);
                    });
                    
                    // WHY: Regulatory information at the bottom center
                    footerColumn.Item().AlignCenter().Text("Authorised and regulated by the National Insurance Commission [RIC-048]").FontSize(7);
                });

                // List
                // adding page content
                page.Content().Column(column =>
                {
                    int numberListCounter = 0;
                    ListItemType? currentListType = null;

                    // Tables
                    var allTables = wordDocument.Tables;
                    int tableIndex = 0;




                    foreach (var paragraph in wordDocument.Paragraphs)
                    {
                        // check if the paragraph is in a table
                        bool isInTable = allTables.Any(t => t.Paragraphs.Contains(paragraph));
                        if (isInTable)
                        {
                            continue;
                        }

                        // WHY: Check for page breaks in paragraph text
                        // Page breaks in Word are often represented as form feed characters (\f) or special break characters
                        // Check if paragraph text contains page break markers
                        bool isPageBreak = paragraph.Text?.Contains("\f") == true || // Form feed character (page break)
                                          paragraph.Text?.Contains("\x000c") == true || // Page break character (Unicode)
                                          paragraph.Text?.Contains("\x0C") == true; // Another form of page break

                        if (isPageBreak)
                        {
                            column.Item().PageBreak(); // WHY: Force a new page in QuestPDF
                            // WHY: Remove the break character from text if it exists, or skip the paragraph
                            if (paragraph.Text?.Length <= 1) // If paragraph is just the break character
                            {
                                continue; // Skip this paragraph entirely
                            }
                            // If paragraph has text after break, continue processing but the page break is already added
                        }

                        // WHY: Check for images in paragraphs first
                        // Images can be in paragraphs even if there's no text
                        if (paragraph.Pictures != null && paragraph.Pictures.Count > 0)
                        {
                            // WHY: Process each image in the paragraph
                            foreach (var picture in paragraph.Pictures)
                            {
                                try
                                {
                                    // WHY: Get image bytes from Word document
                                    // In Xceed.Words.NET, images are accessed through the document's image parts
                                    var imageId = picture.Id;
                                    var imagePart = wordDocument.Images.FirstOrDefault(img => img.Id == imageId);

                                    if (imagePart != null)
                                    {
                                        // WHY: Get image bytes from the image part
                                        // GetStream requires FileMode and FileAccess parameters
                                        byte[] imageBytes;
                                        using (var stream = imagePart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read))
                                        {
                                            using (var memoryStream = new MemoryStream())
                                            {
                                                stream.CopyTo(memoryStream);
                                                imageBytes = memoryStream.ToArray();
                                            }
                                        }

                                        // WHY: Add image to PDF with proper spacing, alignment, and size constraints
                                        // Images need size constraints to prevent layout conflicts
                                        // Use FitArea to constrain image to available space
                                        column.Item()
                                            .PaddingTop(6)
                                            .PaddingBottom(6)
                                            .AlignCenter()
                                            .Image(imageBytes)
                                            .FitArea(); // WHY: Constrain image to fit available area, preventing layout conflicts
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Warning: Could not load image - {ex.Message}");
                                }
                            }

                            // WHY: If paragraph only has images (no text), skip text processing
                            if (string.IsNullOrEmpty(paragraph.Text))
                            {
                                continue;
                            }
                            // If paragraph has both images and text, continue to process text below
                        }

                        // WHY: Skip empty paragraphs to reduce layout complexity
                        // Empty paragraphs create unnecessary padding elements that can cause layout conflicts
                        if (string.IsNullOrEmpty(paragraph.Text))
                        {
                            // Skip empty paragraphs entirely - don't add padding
                            continue;
                        }

                        // check heading
                        var startsWithNumber = paragraph.Text?.Length > 0 && char.IsDigit(paragraph.Text[0]);


                        var isNumberedSectionHeading = startsWithNumber && paragraph?.Text?.Length < 60 && (paragraph.Text.IndexOf(". ") > 0 || paragraph.Text.IndexOf(".") == 1) && !paragraph.Text.Contains("–") && !paragraph.Text.Contains("-") && paragraph.Text.Split(' ').Length <= 8;

                        // WHY: Check if paragraph is a section heading (bold, short, all caps like "IMPORTANT", "PLEASE NOTE")
                        var isBoldSectionHeading = paragraph?.Text?.All(char.IsUpper) == true &&
                                                   paragraph.Text.Length < 50 &&
                                                   paragraph.Text.Length > 3 &&
                                                   paragraph.MagicText.Any(r => r.formatting?.Bold == true) &&
                                                   !paragraph.Text.Contains(":"); // Exclude labels like "INSURED:"

                        // WHY: Enhanced heading detection for all-caps titles
                        // Main titles like "PRODUCT LIABILITY INSURANCE POLICY" should be detected
                        var IsHeading = paragraph?.StyleId?.Contains("Heading") == true ||
                        (paragraph?.Text?.All(char.IsUpper) == true && 
                         paragraph?.Text?.Length < 150 && 
                         paragraph?.Text?.Length > 5 && // Increased from 3 to catch longer titles
                         !paragraph.Text.Contains(":") && // Exclude labels like "INSURED:"
                         paragraph.Text.Split(' ').Length <= 10) || // Allow more words for titles
                        isNumberedSectionHeading ||
                        isBoldSectionHeading; // Add bold section heading detection


                        if (IsHeading)
                        {
                            // we need space for above and below the heading
                            var alignment = paragraph?.Alignment;
                            int levelSize = 1;
                            if (paragraph?.StyleId?.Contains("Heading") == true)
                            {
                                var StyleId = paragraph.StyleId;
                                if (StyleId?.Contains("Heading 1") == true)
                                {
                                    levelSize = 1;
                                }
                                else if (StyleId?.Contains("Heading 2") == true)
                                {
                                    levelSize = 2;
                                }
                                else if (StyleId?.Contains("Heading 3") == true)
                                {
                                    levelSize = 3;
                                }
                                else if (StyleId?.Contains("Heading 4") == true)
                                {
                                    levelSize = 4;
                                }
                                else if (StyleId?.Contains("Heading 5") == true)
                                {
                                    levelSize = 5;
                                }
                                else if (StyleId?.Contains("Heading 6") == true)
                                {
                                    levelSize = 6;
                                }
                                else
                                {
                                    levelSize = 1;
                                }
                            }
                            else if (startsWithNumber)
                            {
                                levelSize = 2;
                            }
                            else
                            {
                                levelSize = 1;
                            }

                            // font calculation
                            float fontSize = 16 - (levelSize * 2);
                            column.Item().PaddingTop(6).PaddingBottom(6)
                            .Text(text =>
                            {
                                text.Span(paragraph?.Text).Bold().FontSize(fontSize);
                                if (alignment == Alignment.left) { text.AlignLeft(); }
                                else if (alignment == Alignment.center) { text.AlignCenter(); }
                                else if (alignment == Alignment.right) { text.AlignRight(); }
                                else if (alignment == Alignment.both) { text.Justify(); }
                            });

                            continue;
                        }


                        if (paragraph?.IsListItem == true)
                        {
                            var indent = paragraph.IndentLevel * 6;
                            // var indent = paragraph.IndentLevel * 10;

                            if (paragraph.ListItemType == ListItemType.Numbered)
                            {
                                if (currentListType != ListItemType.Numbered)
                                {
                                    numberListCounter = 0;
                                }
                                currentListType = ListItemType.Numbered;
                                numberListCounter++;
                            }
                            else
                            {
                                currentListType = ListItemType.Bulleted;
                            }

                            column.Item().PaddingBottom(6).PaddingTop(6).Row(row =>
                            {
                                row.ConstantItem(30).Text(paragraph.ListItemType == ListItemType.Numbered ? $"{numberListCounter}." : "•").Bold();
                                row.RelativeItem().PaddingLeft((float)indent).Text(text =>
                                {
                                    var alignment = paragraph.Alignment;
                                    foreach (var run in paragraph.MagicText)
                                    {
                                        // WHY: Handle line breaks in list items
                                        if (string.IsNullOrEmpty(run.text))
                                        {
                                            if (run.GetType().Name.Contains("Break") || run.GetType().Name.Contains("Line"))
                                            {
                                                text.Span("\n"); // WHY: Use newline character for line breaks
                                            }
                                            continue;
                                        }

                                        // WHY: Split text by line breaks
                                        var textParts = run.text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                                        bool isFirstPart = true;

                                        foreach (var textPart in textParts)
                                        {
                                            if (!isFirstPart)
                                            {
                                                text.Span("\n"); // WHY: Use newline character for line breaks
                                            }
                                            isFirstPart = false;

                                            if (string.IsNullOrEmpty(textPart)) continue;

                                            var textSpan = text.Span(textPart);

                                            // applying style
                                            if (run.formatting?.Bold == true) { textSpan.Bold(); }
                                            if (run.formatting?.Italic == true) { textSpan.Italic(); }
                                            if (run.formatting?.UnderlineStyle == UnderlineStyle.singleLine) { textSpan.Underline(); }

                                            if (run.formatting?.FontFamily != null)
                                            {
                                                textSpan.FontFamily(run.formatting.FontFamily.ToString());
                                            }

                                            if (run.formatting?.FontColor.HasValue == true)
                                            {
                                                var color = run.formatting.FontColor.Value;
                                                textSpan.FontColor(Color.FromARGB(color.A, color.R, color.G, color.B));
                                            }
                                        }
                                    }

                                    if (alignment == Alignment.left) { text.AlignLeft(); }
                                    else if (alignment == Alignment.center) { text.AlignCenter(); }
                                    else if (alignment == Alignment.right) { text.AlignRight(); }
                                    else if (alignment == Alignment.both) { text.Justify(); }
                                });
                            });
                            continue;
                            // Skip regular paragraph processing for list items
                        }


                        if (!paragraph.IsListItem && !IsHeading)
                        {
                            currentListType = null;
                            numberListCounter = 0;
                        }

                        // regular paragraph
                        column.Item().PaddingTop(6).PaddingBottom(6).Text(text =>
                        {

                            var alignment = paragraph.Alignment;
                            foreach (var run in paragraph.MagicText)
                            {
                                // WHY: Check for line breaks in the text
                                // Line breaks can be \n, \r\n, or represented as separate break elements
                                if (string.IsNullOrEmpty(run.text))
                                {
                                    // WHY: Empty run might be a line break marker
                                    // Check if this is a break element
                                    if (run.GetType().Name.Contains("Break") || run.GetType().Name.Contains("Line"))
                                    {
                                        text.Span("\n"); // WHY: Add line break in PDF using newline character
                                    }
                                    continue;
                                }

                                // WHY: Check if text contains line break characters
                                // Split text by line breaks and add each part separately
                                var textParts = run.text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                                bool isFirstPart = true;

                                foreach (var textPart in textParts)
                                {
                                    // WHY: Add line break before each part except the first
                                    if (!isFirstPart)
                                    {
                                        text.Span("\n"); // WHY: Use newline character for line breaks in QuestPDF
                                    }
                                    isFirstPart = false;

                                    if (string.IsNullOrEmpty(textPart)) continue;
                                    // starting the text
                                    // var textSpan = text.Span(run.text);

                                    // check if the run is a hyperlink
                                    var hyperLink = paragraph.Hyperlinks.FirstOrDefault(h => h.Text.Contains(textPart));


                                    if (hyperLink != null && !string.IsNullOrEmpty(hyperLink.Uri.ToString()))
                                    {
                                        var linkSpan = text.Hyperlink(hyperLink.Text, hyperLink.Uri.ToString().AsSpan().ToString());

                                        // styling for indicators
                                        linkSpan.FontColor(Colors.Blue.Darken1);
                                        linkSpan.Underline();
                                        // applying style
                                        if (run.formatting?.Bold == true) { linkSpan.Bold(); }

                                        if (run.formatting?.Italic == true) { linkSpan.Italic(); }

                                        if (run.formatting?.UnderlineStyle == UnderlineStyle.singleLine) { linkSpan.Underline(); }

                                        if (run.formatting?.FontFamily != null)
                                        {
                                            linkSpan.FontFamily(run.formatting.FontFamily.ToString());
                                        }

                                        if (run.formatting?.FontColor.HasValue == true)
                                        {
                                            var color = run.formatting.FontColor.Value;
                                            linkSpan.FontColor(Color.FromARGB(color.A, color.R, color.G, color.B));
                                        }
                                    }
                                    else
                                    {
                                        // regular text - use textPart instead of run.text
                                        var regtextSpan = text.Span(textPart);

                                        // applying style
                                        if (run.formatting?.Bold == true) { regtextSpan.Bold(); }

                                        if (run.formatting?.Italic == true) { regtextSpan.Italic(); }

                                        if (run.formatting?.UnderlineStyle == UnderlineStyle.singleLine) { regtextSpan.Underline(); }

                                        if (run.formatting?.FontFamily != null)
                                        {
                                            regtextSpan.FontFamily(run.formatting.FontFamily.ToString());
                                        }

                                        if (run.formatting?.FontColor.HasValue == true)
                                        {
                                            var color = run.formatting.FontColor.Value;
                                            regtextSpan.FontColor(Color.FromARGB(color.A, color.R, color.G, color.B));
                                        }
                                    }
                                }

                            }

                            if (alignment == Alignment.left)
                            {
                                text.AlignLeft();
                            }
                            else if (alignment == Alignment.center)
                            {
                                text.AlignCenter();
                            }
                            else if (alignment == Alignment.right)
                            {
                                text.AlignRight();
                            }
                            else if (alignment == Alignment.both)
                            {
                                text.Justify();
                            }
                        });
                        // string text = paragraph.Text;
                    }

                    // WHY: Process all tables after paragraphs
                    // Tables are separate from paragraphs in Word documents
                    foreach (var table in allTables)
                    {
                        // WHY: Limit table columns to prevent layout issues
                        // Too many columns can cause width constraint conflicts
                        int columnCount = Math.Min(table.ColumnCount, 10); // Max 10 columns to prevent issues

                        column.Item().PaddingTop(6).PaddingBottom(6).Table(tableElement =>
                        {
                            // WHY: Define table columns - equal width for all columns
                            tableElement.ColumnsDefinition(columns =>
                            {
                                for (int i = 0; i < columnCount; i++)
                                {
                                    columns.RelativeColumn(); // Equal width columns
                                }
                            });

                            // WHY: Process each row in the table
                            foreach (var row in table.Rows)
                            {
                                tableElement.Cell().Row(rowElement =>
                                {
                                    // WHY: Process each cell in the row
                                    // In QuestPDF, use RelativeItem() for each cell in the row
                                    // Limit cells to match column count to prevent mismatches
                                    int cellIndex = 0;
                                    foreach (var cell in row.Cells)
                                    {
                                        if (cellIndex >= columnCount) break; // WHY: Prevent too many cells

                                        rowElement.RelativeItem().Padding(5).Column(cellColumn =>
                                        {
                                            // WHY: Process paragraphs in each cell
                                            foreach (var cellParagraph in cell.Paragraphs)
                                            {
                                                if (!string.IsNullOrEmpty(cellParagraph.Text))
                                                {
                                                    cellColumn.Item().Text(text =>
                                                    {
                                                        // WHY: Process text runs in cell
                                                        foreach (var run in cellParagraph.MagicText)
                                                        {
                                                            if (string.IsNullOrEmpty(run.text)) continue;

                                                            var textSpan = text.Span(run.text);

                                                            // applying style
                                                            if (run.formatting?.Bold == true) { textSpan.Bold(); }
                                                            if (run.formatting?.Italic == true) { textSpan.Italic(); }
                                                            if (run.formatting?.UnderlineStyle == UnderlineStyle.singleLine) { textSpan.Underline(); }

                                                            if (run.formatting?.FontFamily != null)
                                                            {
                                                                textSpan.FontFamily(run.formatting.FontFamily.ToString());
                                                            }

                                                            if (run.formatting?.FontColor.HasValue == true)
                                                            {
                                                                var color = run.formatting.FontColor.Value;
                                                                textSpan.FontColor(Color.FromARGB(color.A, color.R, color.G, color.B));
                                                            }
                                                        }

                                                        // WHY: Alignment should be OUTSIDE the run loop
                                                        // Apply alignment once per paragraph, not per run
                                                        if (cellParagraph.Alignment == Alignment.left) { text.AlignLeft(); }
                                                        else if (cellParagraph.Alignment == Alignment.center) { text.AlignCenter(); }
                                                        else if (cellParagraph.Alignment == Alignment.right) { text.AlignRight(); }
                                                        else if (cellParagraph.Alignment == Alignment.both) { text.Justify(); }
                                                    });
                                                }
                                            }
                                        });
                                        cellIndex++;
                                    }
                                });
                            }
                        });
                    }
                });
            });
        }).GeneratePdf(pdfPath);

        string fullPath = Path.GetFullPath(pdfPath);
        Console.WriteLine($"✓ PDF saved to: {fullPath}");
    }

    Console.WriteLine("Conversion completed successfully!");


}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    // throw;
}

Console.WriteLine("Press any key to exit");
// WHY: Only read key if console input is available (not redirected)
try
{
    Console.ReadKey();
}
catch (InvalidOperationException)
{
    // Console input not available (e.g., when running via script)
    // Just exit silently
}
