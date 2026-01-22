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

                // adding page content
                page.Content().Column(column =>
                {
                    foreach (var paragraph in wordDocument.Paragraphs)
                    {
                        string text = paragraph.Text;

                        if (string.IsNullOrEmpty(text))
                        {
                            column.Item().PaddingBottom(5);
                            continue;
                        }

                        column.Item().Text(text);
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
