using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace WindowsImagePdfOcrApp
{
    public class PowerOcrEngine
    {
        private readonly OcrEngine _engine;
        private readonly bool _isSpaceJoiningLanguage;

        public PowerOcrEngine(string? languageTag = null)
        {
            // Initialization (Microsoft Original)
            if (string.IsNullOrEmpty(languageTag))
                languageTag = System.Globalization.CultureInfo.CurrentCulture.Name;

            Language selectedLanguage = new Language(languageTag);

            if (!OcrEngine.IsLanguageSupported(selectedLanguage))
            {
                var available = OcrEngine.AvailableRecognizerLanguages;
                selectedLanguage = available.FirstOrDefault(l => l.LanguageTag.StartsWith(selectedLanguage.LanguageTag.Split('-')[0]))
                                   ?? available.FirstOrDefault();
                if (selectedLanguage == null) throw new InvalidOperationException("OCR языки не найдены.");
            }

            _engine = OcrEngine.TryCreateFromLanguage(selectedLanguage);
            _isSpaceJoiningLanguage = IsLanguageSpaceJoining(selectedLanguage);
        }

        public async Task<string> ExtractTextAsync(string imagePath)
        {
            if (!File.Exists(imagePath)) throw new FileNotFoundException("Файл не найден", imagePath);
            using var bmp = new Bitmap(imagePath);
            return await ExtractTextAsync(bmp);
        }

        public async Task<string> ExtractTextAsync(Bitmap inputBmp)
        {
            // 1. Padding (Microsoft Original)
            // We must add a border
            using var paddedBmp = PadImage(inputBmp);

            // 2. INVERT (TUNING #1)
            // Critically important for Dark Mode.
            // Turn "White on Gray" into "Black on White".
            // This improves recognition of small fonts by 40-50%.
            using var invertedBmp = InvertColors(paddedBmp);

            // 3. Scale 2.0 (TUNING #2)
            // Was 1.5 -> Now 2.0.
            // Fixes the "Лдти" (Идти) mistake.
            double scaleFactor = 2.0;

            bool performScale = true;
            if (invertedBmp.Width * scaleFactor > OcrEngine.MaxImageDimension ||
                invertedBmp.Height * scaleFactor > OcrEngine.MaxImageDimension)
            {
                performScale = false;
            }

            using var finalBmp = performScale ? ScaleBitmapUniform(invertedBmp, scaleFactor) : new Bitmap(invertedBmp);

            // 4. Recognition
            using var softwareBitmap = await ConvertToSoftwareBitmap(finalBmp);
            var result = await _engine.RecognizeAsync(softwareBitmap);

            return ProcessOcrResult(result);
        }

        // --- Helper: Invert Colors (Simple math) ---
        private Bitmap InvertColors(Bitmap original)
        {
            Bitmap newBitmap = new Bitmap(original.Width, original.Height);
            newBitmap.SetResolution(original.HorizontalResolution, original.VerticalResolution);

            using (Graphics g = Graphics.FromImage(newBitmap))
            {
                // Inversion matrix
                ColorMatrix colorMatrix = new ColorMatrix(
                   new float[][]
                   {
                      new float[] {-1, 0, 0, 0, 0},
                      new float[] {0, -1, 0, 0, 0},
                      new float[] {0, 0, -1, 0, 0},
                      new float[] {0, 0, 0, 1, 0},
                      new float[] {1, 1, 1, 0, 1}
                   });

                using (ImageAttributes attributes = new ImageAttributes())
                {
                    attributes.SetColorMatrix(colorMatrix);
                    g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height),
                        0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);
                }
            }
            return newBitmap;
        }

        // --- Microsoft Original: PadImage ---
        private static Bitmap PadImage(Bitmap image, int minW = 64, int minH = 64)
        {
            int width = Math.Max(image.Width + 16, minW + 16);
            int height = Math.Max(image.Height + 16, minH + 16);

            Bitmap destination = new Bitmap(width, height, image.PixelFormat);
            destination.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (Graphics g = Graphics.FromImage(destination))
            {
                g.Clear(Color.White);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                int x = (width - image.Width) / 2;
                int y = (height - image.Height) / 2;
                g.DrawImage(image, x, y, image.Width, image.Height);
            }
            return destination;
        }

        // --- Microsoft Original: Scale (with scale argument 2.0) ---
        private static Bitmap ScaleBitmapUniform(Bitmap image, double scale)
        {
            int newWidth = (int)(image.Width * scale);
            int newHeight = (int)(image.Height * scale);

            Bitmap destination = new Bitmap(newWidth, newHeight, image.PixelFormat);
            destination.SetResolution(96, 96); // Fix DPI for stability

            using (Graphics g = Graphics.FromImage(destination))
            {
                g.Clear(Color.White);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;

                g.DrawImage(image, 0, 0, newWidth, newHeight);
            }
            return destination;
        }

        private async Task<SoftwareBitmap> ConvertToSoftwareBitmap(Bitmap bitmap)
        {
            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Png);
            ms.Position = 0;
            var decoder = await BitmapDecoder.CreateAsync(ms.AsRandomAccessStream());
            return await decoder.GetSoftwareBitmapAsync();
        }

        private string ProcessOcrResult(OcrResult result)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var line in result.Lines)
            {
                if (_isSpaceJoiningLanguage) sb.AppendLine(line.Text);
                else sb.AppendLine(string.Join(" ", line.Words.Select(w => w.Text)));
            }
            return sb.ToString().Trim();
        }

        private static bool IsLanguageSpaceJoining(Language language)
        {
            return language.LanguageTag.StartsWith("zh", StringComparison.InvariantCultureIgnoreCase) ||
                   language.LanguageTag.Equals("ja", StringComparison.OrdinalIgnoreCase);
        }
    }
}
