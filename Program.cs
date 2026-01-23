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


                // List
                // adding page content
                page.Content().Column(column =>
                {
                    int numberListCounter = 0;
                    ListItemType? currentListType = null;
                    foreach (var paragraph in wordDocument.Paragraphs)
                    {
                        if (string.IsNullOrEmpty(paragraph.Text))
                        {
                            column.Item().PaddingBottom(5);
                            continue;
                        }

                        // check heading
                        var startsWithNumber = paragraph.Text?.Length > 0 && char.IsDigit(paragraph.Text[0]);
                        var IsHeading = paragraph.StyleId.Contains("Heading") == true ||
                        (paragraph.Text?.All(char.IsUpper) == true && paragraph.Text.Length < 100 && paragraph.Text.Length > 3) ||
                        (startsWithNumber && paragraph.Text?.Length < 100);


                        if (IsHeading)
                        {
                            column.Item().Text(text =>
                            {
                                text.Span(paragraph.Text).Bold().FontSize(14);
                            });
                            continue;
                        }


                        if (paragraph.IsListItem)
                        {
                            var indent = paragraph.IndentLevel * 10;

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

                            column.Item().Row(row =>
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
                            continue; // Skip regular paragraph processing for list items
                        }


                        if (!paragraph.IsListItem && !IsHeading)
                        {
                            currentListType = null;
                            numberListCounter = 0;
                        }

                        // regular paragraph
                        column.Item().Text(text =>
                        {
                            var alignment = paragraph.Alignment;
                            foreach (var run in paragraph.MagicText)
                            {
                                if (string.IsNullOrEmpty(run.text)) continue;
                                // starting the text
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

Console.WriteLine("Press anykey to exit");
Console.ReadKey();
