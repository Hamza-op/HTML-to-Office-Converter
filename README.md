# HTML to Office Converter

A modern desktop application to convert HTML files into editable DOCX and PPTX formats. Built with Python and `customtkinter`, featuring a beautiful dark/light mode interface.

## Features

- **HTML to DOCX**: Converts HTML to highly-accurate, editable Word documents. It carefully preserves layout and styling by utilizing a Playwright to PDF to DOCX pipeline.
- **HTML to Editable PPTX**: Parses HTML natively into editable PowerPoint shapes, text frames, and tables using `python-pptx`.
- **Modern UI**: Polished, responsive interface built with CustomTkinter, featuring drag-and-drop, real-time activity logging, and live visual previews.
- **Cross-Platform**: Designed to work gracefully on Windows, macOS, and Linux.

## Requirements

Ensure you have Python 3.8+ installed. The following libraries are required:

- `customtkinter`
- `Pillow`
- `playwright`
- `pdf2docx`
- `python-pptx`
- `PyMuPDF` (imported as `fitz`)

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Hamza-op/HTML-to-Office-Converter.git
   cd HTML-to-Office-Converter
   ```

2. **Install the dependencies:**
   ```bash
   pip install customtkinter Pillow playwright pdf2docx python-pptx PyMuPDF
   ```

3. **Background Browser Setup**:
   The application requires Chromium for the HTML rendering pipeline. The app will attempt to automatically install the Playwright browser on the first run, or you can install it manually:
   ```bash
   python -m playwright install chromium
   ```

## Usage

1. **Run the Application:**
   ```bash
   python main.py
   ```
2. **Select Files**: Drag and drop your `.html` files into the input zone or click "Browse".
3. **Choose Output Format**:
   - **DOCX**: Options for page size, orientation, and margins are available.
   - **PPTX**: Options for slide size and quality are available.
4. **Convert**: Click the **Convert Now** button. The built-in activity log will show real-time progress.
5. **View Preview**: After conversion, you can navigate and zoom through a visual preview of the document right inside the app!

## License

This project is licensed under the MIT License.
