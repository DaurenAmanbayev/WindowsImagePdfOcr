# WindowsImagePdfOcrApp

A small command-line and library tool that performs OCR on PDF pages and image files using Windows built-in OCR (Windows.Media.Ocr). The application renders PDF pages, preprocesses images (padding, inversion, scaling), and extracts text using the Windows OCR engine. The project is implemented in C# targeting .NET 8 and is designed to run on Windows (desktop) environments that support WinRT OCR APIs.

## Features

- OCR for PDF documents (page-by-page rendering)
- OCR for common image formats: PNG, JPEG, BMP, TIFF, GIF
- Preprocessing tuned for better accuracy: padding, color inversion, uniform scaling
- Reusable OCR engine wrapper with language selection (example: `ru-RU`)
- Saves extracted text to a `.txt` file next to the input file

## Prerequisites

- Windows 10 (version 1809+) or Windows 11 � required for WinRT OCR APIs
- .NET 8 SDK
- Visual Studio 2022 (recommended) or newer with .NET desktop development workload

**Note:** The project uses WinRT APIs (`Windows.Media.Ocr`, `Windows.Data.Pdf`, `Windows.Graphics.Imaging`) and therefore must run on Windows. Running on non-Windows platforms is not supported.

## Build

1. Clone the repository:

    ```bash
    git clone https://github.com/your/repo.git
    cd WindowsImagePdfOcr
    ```

2. Open the solution in Visual Studio 2022.
3. Restore NuGet packages and build the solution (Build > Build Solution).

Alternatively, build from the command line with the .NET SDK:


dotnet build -c Release


## Run

Usage (console):


WindowsImagePdfOcrApp.exe "C:\path\to\file.pdf"


If no path is provided, the app uses a test path configured in `Program.cs`.

### Examples:

- OCR a PDF:

    ```bash
    WindowsImagePdfOcrApp.exe "C:\Scans\document.pdf"
    ```

- OCR an image:

    ```bash
    WindowsImagePdfOcrApp.exe "C:\Images\scan.png"
    ```

The extracted text will be saved to `file.pdf.txt` or `scan.png.txt` next to the source file.

## Configuration

- **Language:** The engine can be initialized with a language tag (e.g. `ru-RU`). If the requested language is not available, the engine falls back to an available recognizer language.
- **Scaling:** The engine applies a scaling factor (default 2.0) before recognition. You can change scaling or pre-processing in `PowerOcrEngine.cs`.

## Project layout

- `Program.cs` CLI entry point, detects file type and orchestrates processing
- `PdfProcessor.cs` Loads PDF, renders pages via `Windows.Data.Pdf`, converts pages to bitmaps
- `ImageProcessor.cs` Validates and loads images
- `PowerOcrEngine.cs` Image preprocessing (pad, invert, scale), converts to `SoftwareBitmap`, calls `OcrEngine`

## Performance and optimization tips

- **Parallelize PDF page processing** when working with multi-page PDFs. Use a limited concurrency (for example `Environment.ProcessorCount`) to avoid saturating CPU and IO.
- **Reduce render DPI** (or scale) for large documents if OCR quality remains acceptable � lower resolution reduces memory usage and recognition time.
- **Avoid unnecessary conversions:** prefer `SoftwareBitmap`/WinRT types where possible to reduce copying between managed and native buffers.
- **Reuse OCR engine instances** where thread-safety permits. If `OcrEngine` is not thread-safe, implement a pool of engines to reduce initialization overhead.
- **Add timing instrumentation** (Stopwatch) around rendering, conversion, and OCR steps to identify bottlenecks.

## Troubleshooting

- **File not found:** Ensure the full absolute path is provided and the app has read permissions.
- **Unsupported image:** Only common image formats are supported. Corrupted images may throw `ArgumentException` during load.
- **Missing OCR languages:** Install the required language packs for Windows if your language is not available to `OcrEngine`.

## Contributing

Contributions are welcome. Please open an issue for discussion before submitting larger changes. Follow the repository's coding conventions and create small, focused pull requests.

### Suggested areas for contribution:

- Parallel PDF processing with throttled concurrency
- Improved preprocessing filters (deskew, denoise, binarization)
- Unit tests for processors
- Cross-platform abstraction (if replacing WinRT OCR with a cross-platform engine)

## License

Specify your license here (for example MIT). If no license is added, the project remains proprietary by default.

---

If you want, I can also add a short `CONTRIBUTING.md` and `.editorconfig` tuned to the project's style.


This revised `README.md` maintains the original structure while enhancing clarity and coherence, ensuring that users can easily understand and navigate the documentation.
