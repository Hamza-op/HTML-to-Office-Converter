# HTML to Office Converter

A modern desktop application to convert HTML files into editable DOCX and PPTX formats. Built with Python and `customtkinter`, featuring a beautiful dark/light mode interface.

## Features

- **HTML to DOCX**: Converts HTML to highly-accurate, editable Word documents. It carefully preserves layout and styling by utilizing a Playwright to PDF to DOCX pipeline.
- **HTML to Editable PPTX**: Parses HTML natively into editable PowerPoint shapes, text frames, and tables using `python-pptx`.
- **Modern UI**: Polished, responsive interface built with CustomTkinter, featuring drag-and-drop, real-time activity logging, and live visual previews.
- **Cross-Platform**: Designed to work gracefully on Windows, macOS, and Linux.
- **Standalone Executable**: Run the application as a single `.exe` file without installing Python or dependencies.
- **System Browser Integration**: Automatically detects and uses installed Google Chrome or Microsoft Edge for HTML parsing, eliminating the need to download massive browser binaries.

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
   The application intelligently detects and uses **Google Chrome** or **Microsoft Edge** if they are already installed on your system. 
   
   If neither is found, the app will attempt to automatically install the Playwright Chromium browser on the first run, or you can install it manually:
   ```bash
   playwright install chromium
   ```

## Usage

1. **Run the Application (Source):**
   ```bash
   python app.py
   ```
   **OR Run the Executable (Windows):**
   Simply double-click the `HTML-to-Office-Converter.exe` file. No Python installation is required.

2. **Select Files**: Drag and drop your `.html` files into the input zone or click "Browse".
3. **Choose Output Format**:
   - **DOCX**: Options for page size, orientation, and margins are available.
   - **PPTX**: Options for slide size and quality are available.
4. **Convert**: Click the **Convert Now** button. The built-in activity log will show real-time progress.
5. **View Preview**: After conversion, you can navigate and zoom through a visual preview of the document right inside the app!
## Building the Executable (Windows)

If you want to build the standalone `.exe` yourself:

1. Ensure all dependencies are installed, including PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Choose a build type:

   - **Fast start (recommended)**: folder-based build (starts much faster than onefile)
     ```bash
     pyinstaller HTML-to-Office-Converter.onedir.spec
     ```

   - **Single EXE**: onefile build (larger and slower to start because it unpacks on every launch)
     ```bash
     pyinstaller HTML-to-Office-Converter.spec
     ```

3. The output will be generated in the `dist` folder.

   - `onedir`: `dist/HTML-to-Office-Converter/HTML-to-Office-Converter.exe`
   - `onefile`: `dist/HTML-to-Office-Converter.exe`

## Notes On EXE Size / Startup

- A ~130 MB Windows onefile EXE is normal for PyInstaller apps that bundle Python + Tk + Pillow and also include Playwright/PyMuPDF for high-fidelity rendering and previews.
- For the fastest startup, prefer the `onedir` build.


## License

This project is licensed under the MIT License.
