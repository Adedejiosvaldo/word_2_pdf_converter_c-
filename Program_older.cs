// using QuestPDF.Infrastructure;
// using Xceed.Words.NET;
// using Xceed.Document.NET;
// using QuestPDF.Helpers;
// using QuestPDF.Fluent;

// QuestPDF.Settings.License = LicenseType.Community;

// Console.WriteLine("Word Document Explorer");
// Console.WriteLine("======================");

// Console.WriteLine("Enter your Word file path (.docx):");
// string? filePath = Console.ReadLine();

// if (!File.Exists(filePath))
// {
//     Console.WriteLine("No File found");
//     return;
// }

// try
// {
//     using (var wordDocument = DocX.Load(filePath))
//     {
//         Console.WriteLine($"Document Loaded: {wordDocument.Paragraphs.Count} paragraphs");
//         Console.WriteLine();
//         Console.WriteLine("Exploring first 3 paragraphs...");
//         Console.WriteLine();

//         int count = 0;
//         foreach (var paragraph in wordDocument.Paragraphs)
//         {
//             if (count >= 3) break;
//             if (string.IsNullOrWhiteSpace(paragraph.Text)) continue;

//             Console.WriteLine($"--- Paragraph {count + 1} ---");
//             Console.WriteLine($"Text: {paragraph.Text}");
//             Console.WriteLine($"Alignment: {paragraph.Alignment}");
//             Console.WriteLine();

//             // Explore MagicText
//             var magicTexts = paragraph.MagicText.ToList();
//             Console.WriteLine($"MagicText items: {magicTexts.Count}");

//             int runIndex = 0;
//             foreach (var magic in magicTexts)
//             {
//                 Console.WriteLine($"  Run {runIndex}:");
//                 Console.WriteLine($"    Type: {magic.GetType().Name}");

//                 // Use reflection to see all properties
//                 var properties = magic.GetType().GetProperties();
//                 foreach (var prop in properties)
//                 {
//                     try
//                     {
//                         var value = prop.GetValue(magic);
//                         if (value != null)
//                         {
//                             Console.WriteLine($"      {prop.Name}: {value}");
//                         }
//                     }
//                     catch
//                     {
//                         // Skip properties we can't read
//                     }
//                 }
//                 Console.WriteLine();
//                 runIndex++;
//             }

//             count++;
//         }
//     }
// }
// catch (Exception ex)
// {
//     Console.WriteLine($"Error: {ex.Message}");
//     Console.WriteLine($"Stack: {ex.StackTrace}");
// }

// Console.WriteLine();
// Console.WriteLine("Press any key to exit");
// Console.ReadKey();
