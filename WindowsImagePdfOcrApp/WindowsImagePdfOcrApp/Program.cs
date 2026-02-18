using Windows.Media.Ocr;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Data.Pdf;
using Windows.Storage.Streams;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace WindowsImagePdfOcrApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // 1. Configure encoding for correct display of Cyrillic characters in the console
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            // 2. Get file path (from arguments or hardcoded for testing)
            // If running from console: OcrTool.exe "C:\Path\To\File.pdf"
            string inputPath = args.Length > 0 ? args[0] : @"C:\Test\scan.pdf";

            Console.WriteLine($"=== Starting OCR Tool ===");
            Console.WriteLine($"Input file: {inputPath}");

            try
            {
                // 3. Initialize engine core (V12 - Invert + Scale 2.0)
                // Explicitly specify Russian language; the engine will pick up English as well
                var engine = new PowerOcrEngine("ru-RU");

                // 4. Initialize processors
                var pdfProcessor = new PdfProcessor(engine);
                var imageProcessor = new ImageProcessor(engine);

                // 5. Determine file type and route processing
                string extension = Path.GetExtension(inputPath)?.ToLowerInvariant();

                if (extension == ".pdf")
                {
                    // Branch for PDFs
                    await ProcessPdfDocument(pdfProcessor, inputPath);
                }
                else
                {
                    // Branch for images (png, jpg, bmp, etc.)
                    await ProcessImageFile(imageProcessor, inputPath);
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"\n[ERROR] File not found: {ex.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[CRITICAL ERROR]: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("\nPress Enter to exit...");
            Console.ReadLine();
        }

        // --- METHOD 1: PDF Processing ---
        private static async Task ProcessPdfDocument(PdfProcessor processor, string filePath)
        {
            Console.WriteLine(">>> PDF document detected. Starting page-by-page rendering...");

            var watch = System.Diagnostics.Stopwatch.StartNew();

            // Call PDF parsing logic
            string extractedText = await processor.ProcessPdfAsync(filePath);

            watch.Stop();
            Console.WriteLine($">>> Processing completed in {watch.Elapsed.TotalSeconds:F2} sec.");

            // Save and output
            SaveResult(filePath, extractedText);
        }

        // --- METHOD 2: Image Processing ---
        private static async Task ProcessImageFile(ImageProcessor processor, string filePath)
        {
            Console.WriteLine(">>> Image detected. Starting preprocessing and OCR...");

            var watch = System.Diagnostics.Stopwatch.StartNew();

            // Call image parsing logic (validation -> bitmap -> OCR)
            string extractedText = await processor.ProcessImageAsync(filePath);

            watch.Stop();
            Console.WriteLine($">>> Processing completed in {watch.Elapsed.TotalSeconds:F2} sec.");

            // Save and output
            SaveResult(filePath, extractedText);
        }

        // --- Helper method to save the result ---
        private static void SaveResult(string originalPath, string text)
        {
            Console.WriteLine("\n--- BEGIN RESULT ---");
            // Print the first 500 characters to the console to avoid clutter when text is very large
            Console.WriteLine(text.Length > 500 ? text.Substring(0, 500) + "\n... [text truncated for console] ..." : text);
            Console.WriteLine("--- END RESULT ---");

            string outputPath = originalPath + ".txt";
            File.WriteAllText(outputPath, text);
            Console.WriteLine($"\n[SUCCESS] Full text saved to file: {outputPath}");
        }


    }
}
