# HTML to Office Converter

A modern desktop application to convert HTML files into editable DOCX and PPTX formats. Built with Python and `customtkinter`, featuring a beautiful dark/light mode interface.

## Features

- **HTML/PDF to DOCX**: Converts HTML or PDF to highly-accurate, editable Word documents.
- **HTML/PDF to Editable PPTX**: Parses HTML or PDF natively into editable PowerPoint shapes, text frames, and images. Move, edit, or resize any element.
- **Modern UI**: Polished "Warm Amber" interface built with CustomTkinter, featuring drag-and-drop, real-time activity logging, and live visual previews.
- **Cross-Platform**: Designed to work gracefully on Windows, macOS, and Linux.
- **Standalone Executable**: Run the application as a single `.exe` file without installing Python or dependencies.
- **System Browser Integration**: Automatically detects and uses installed Google Chrome or Microsoft Edge for HTML parsing.

## Requirements

Ensure you have Python 3.9+ installed. The following libraries are required:

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

2. **Select Files**: Click "Browse" and choose `.html`, `.htm`, or `.pdf` files.
3. **Choose Output Format**:
   - **DOCX**: Options for page size, orientation, and margins are available.
   - **PPTX**: Options for slide size and quality are available.
4. **Convert**: Click the **Convert Now** button. The built-in activity log will show real-time progress.
5. **View Preview**: After conversion, you can navigate and zoom through a visual preview of the document right inside the app!

## PDF -> Editable PPTX Quality Gate

To validate production readiness for PDF to editable PPTX conversion, run:

```bash
python quality_gate_pdf_to_pptx.py --pdf input.pdf --convert --out output.pptx
```

The gate prints JSON coverage metrics for pages, text runs, images, vectors, and tables, and exits non-zero when thresholds fail.

To score an already generated PPTX:

```bash
python quality_gate_pdf_to_pptx.py --pdf input.pdf --pptx output.pptx
```
## Building the Executable (Windows)

If you want to build the standalone `.exe` yourself:

1. Ensure all dependencies are installed, including PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Build executable:
   ```bash
   pyinstaller --noconfirm HTML_to_Office_Converter.spec
   ```

3. The output will be generated in the `dist` folder:
   - `dist/HTML_to_Office_Converter.exe`

## Notes On EXE Size / Startup

- A ~130 MB Windows onefile EXE is normal for PyInstaller apps that bundle Python + Tk + Pillow and also include Playwright/PyMuPDF for high-fidelity rendering and previews.
- For the fastest startup, prefer the `onedir` build.


## GitHub Auto-Release

The project includes a GitHub Actions workflow to automatically build and release the Windows executable:

1. Push a version tag:
   ```bash
   git tag v2.1.0
   git push origin v2.1.0
   ```
2. The `.github/workflows/releaser.yml` will trigger, build the EXE on a Windows runner, create a new GitHub Release, and upload the `HTML-to-Office-Converter.exe` asset automatically.

## License

This project is licensed under the MIT License.
