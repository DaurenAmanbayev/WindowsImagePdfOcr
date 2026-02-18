using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace WindowsImagePdfOcrApp
{
    public class PdfProcessor
    {
        private readonly PowerOcrEngine _ocrEngine;

        public PdfProcessor(PowerOcrEngine engine)
        {
            _ocrEngine = engine;
        }

        public async Task<string> ProcessPdfAsync(string pdfPath)
        {
            // 1. Validate path and get the file via Windows Storage API
            // WinRT requires an absolute path
            string fullPath = Path.GetFullPath(pdfPath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("PDF not found", fullPath);

            StorageFile file = await StorageFile.GetFileFromPathAsync(fullPath);

            // 2. Load the PDF document
            PdfDocument pdfDoc = await PdfDocument.LoadFromFileAsync(file);

            StringBuilder fullText = new StringBuilder();
            Console.WriteLine($"Pages found: {pdfDoc.PageCount}");

            // 3. Process each page
            for (uint i = 0; i < pdfDoc.PageCount; i++)
            {
                Console.WriteLine($"Processing page {i + 1} of {pdfDoc.PageCount}...");

                using (var page = pdfDoc.GetPage(i))
                {
                    // Render the page into an in-memory stream
                    using (var stream = new InMemoryRandomAccessStream())
                    {
                        // Rendering settings.
                        // You can increase resolution here, but our OCR engine already performs Scale 2.0.
                        // Keep default — it's fast and produces good results.
                        var options = new PdfPageRenderOptions();

                        await page.RenderToStreamAsync(stream, options);

                        // Convert WinRT stream to .NET Bitmap
                        using (var bitmap = await StreamToBitmap(stream))
                        {
                            // 4. Call our powerful OCR engine (V12)
                            string pageText = await _ocrEngine.ExtractTextAsync(bitmap);

                            fullText.AppendLine($"--- Page {i + 1} ---");
                            fullText.AppendLine(pageText);
                            fullText.AppendLine();
                        }
                    }
                }
            }

            return fullText.ToString();
        }

        // Helper for stream conversion
        private async Task<Bitmap> StreamToBitmap(IRandomAccessStream winRtStream)
        {
            // Copy data from WinRT stream into a .NET MemoryStream
            using (var netStream = new MemoryStream())
            {
                var reader = new DataReader(winRtStream.GetInputStreamAt(0));
                await reader.LoadAsync((uint)winRtStream.Size);
                byte[] buffer = new byte[winRtStream.Size];
                reader.ReadBytes(buffer);

                await netStream.WriteAsync(buffer, 0, buffer.Length);
                netStream.Position = 0;

                return new Bitmap(netStream);
            }
        }
    }
}
