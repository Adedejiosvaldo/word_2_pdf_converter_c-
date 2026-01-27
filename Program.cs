using System.ComponentModel;
using QuestPDF.Infrastructure;
using Xceed.Words.NET;
using Xceed.Document.NET;
using QuestPDF.Helpers;
using QuestPDF.Fluent;

QuestPDF.Settings.License = LicenseType.Community;

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
                page.DefaultTextStyle(x => x.FontSize(11).FontFamily("Calibri"));

                page.Footer().AlignCenter().Text(text =>
                {
                    text.CurrentPageNumber().FontSize(8);
                    text.Span(" of ").FontSize(8);
                    text.TotalPages().FontSize(8);


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

                        // check if the paragraph is in a list
                        if (string.IsNullOrEmpty(paragraph.Text))
                        {
                            column.Item().PaddingBottom(5);
                            continue;
                        }

                        // check heading
                        var startsWithNumber = paragraph.Text?.Length > 0 && char.IsDigit(paragraph.Text[0]);


                        var isNumberedSectionHeading = startsWithNumber && paragraph?.Text?.Length < 60 && (paragraph.Text.IndexOf(". ") > 0 || paragraph.Text.IndexOf(".") == 1) && !paragraph.Text.Contains("–") && !paragraph.Text.Contains("-") && paragraph.Text.Split(' ').Length <= 8;

                        var IsHeading = paragraph?.StyleId?.Contains("Heading") == true ||
                        (paragraph?.Text?.All(char.IsUpper) == true && paragraph?.Text?.Length < 100 && paragraph?.Text?.Length > 3) ||
                        isNumberedSectionHeading;


                        if (IsHeading)
                        {
                            // we need space for above and below the heading
                            var alignment = paragraph.Alignment;
                            int levelSize = 1;
                            if (paragraph.StyleId.Contains("Heading"))
                            {
                                var StyleId = paragraph.StyleId;
                                if (StyleId.Contains("Heading 1"))
                                {
                                    levelSize = 1;
                                }
                                else if (StyleId.Contains("Heading 2"))
                                {
                                    levelSize = 2;
                                }
                                else if (StyleId.Contains("Heading 3"))
                                {
                                    levelSize = 3;
                                }
                                else if (StyleId.Contains("Heading 4"))
                                {
                                    levelSize = 4;
                                }
                                else if (StyleId.Contains("Heading 5"))
                                {
                                    levelSize = 5;
                                }
                                else if (StyleId.Contains("Heading 6"))
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
                                text.Span(paragraph.Text).Bold().FontSize(fontSize);
                                if (alignment == Alignment.left) { text.AlignLeft(); }
                                else if (alignment == Alignment.center) { text.AlignCenter(); }
                                else if (alignment == Alignment.right) { text.AlignRight(); }
                                else if (alignment == Alignment.both) { text.Justify(); }
                            });

                            continue;
                        }


                        if (paragraph.IsListItem)
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


                                if (string.IsNullOrEmpty(run.text)) continue;
                                // starting the text
                                // var textSpan = text.Span(run.text);

                                // check if the run is a hyperlink
                                var hyperLink = paragraph.Hyperlinks.FirstOrDefault(h => h.Text.Contains(run.text));


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
                                    // regular text
                                    var regtextSpan = text.Span(run.text);// applying style
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


                        // Tables code
                        foreach (var table in allTables)
                        {
                            column.Item().PaddingTop(3).PaddingBottom(3).Table(tableElement =>
                            {
                                tableElement.ColumnsDefinition(column =>
                                {
                                    for (int i = 0; i < table.ColumnCount; i++)
                                    {
                                        column.RelativeColumn();
                                    }
                                });

                                // Process each row
                                foreach (var row in table.Rows)
                                {
                                    tableElement.Cell().Row(rowElement =>
                                    {
                                        foreach (var cell in row.Cells)
                                        {
                                            tableElement.Cell().Padding(5).Column(column =>
                                            {

                                                foreach (var cellParagraph in cell.Paragraphs)
                                                {

                                                    if (!string.IsNullOrEmpty(cellParagraph.Text))
                                                    {
                                                        column.Item().Text(text =>
                                                        {

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

                                                                if (cellParagraph.Alignment == Alignment.left) { text.AlignLeft(); }
                                                                else if (cellParagraph.Alignment == Alignment.center) { text.AlignCenter(); }
                                                                else if (cellParagraph.Alignment == Alignment.right) { text.AlignRight(); }
                                                                else if (cellParagraph.Alignment == Alignment.both) { text.Justify(); }
                                                            }
                                                        });
                                                    }
                                                    {

                                                    }
                                                }

                                            });


                                        }
                                        ;
                                    });
                                }
                            }); ;
                        }
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
Console.ReadKey();
