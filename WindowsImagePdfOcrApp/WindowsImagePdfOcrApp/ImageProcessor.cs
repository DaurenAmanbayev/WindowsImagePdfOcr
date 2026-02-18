using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsImagePdfOcrApp
{
    public class ImageProcessor
    {
        private readonly PowerOcrEngine _ocrEngine;

        public ImageProcessor(PowerOcrEngine engine)
        {
            _ocrEngine = engine;
        }

        public async Task<string> ProcessImageAsync(string imagePath)
        {
            // 1. Path validation
            string fullPath = Path.GetFullPath(imagePath);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException($"File not found: {fullPath}", fullPath);

            // 2. Format check (by extension)
            // Basic protection against feeding .txt or .exe as an image
            if (!IsImageExtension(fullPath))
            {
                throw new InvalidOperationException($"File {Path.GetFileName(fullPath)} is not a supported image.");
            }

            Console.WriteLine($"Processing image: {Path.GetFileName(fullPath)}...");

            try
            {
                // 3. Load into Bitmap
                // Use 'using' to avoid leaving the file locked by the system
                using (var bitmap = new Bitmap(fullPath))
                {
                    // 4. Run OCR
                    string text = await _ocrEngine.ExtractTextAsync(bitmap);
                    return text;
                }
            }
            catch (ArgumentException)
            {
                // System.Drawing throws ArgumentException if the file is corrupted or the format is invalid
                throw new InvalidDataException("Failed to open image. The file may be corrupted or the format is not supported by GDI+.");
            }
        }

        private bool IsImageExtension(string path)
        {
            string ext = Path.GetExtension(path)?.ToLowerInvariant();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg" ||
                   ext == ".bmp" || ext == ".tif" || ext == ".tiff" || ext == ".gif";
        }
    }
}
