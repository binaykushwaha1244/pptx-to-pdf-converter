# PPTX to PDF Converter

A lightweight Windows desktop application for converting PowerPoint `.pptx` presentations into PDF files. Supports batch conversion, folder selection, output directory selection, and conversion progress/error reporting.

**Your files stay on your computer** — presentations are converted locally using Microsoft PowerPoint or LibreOffice.

## Features

- **Batch conversion** — select multiple `.pptx` files at once.
- **Folder import** — add all `.pptx` files from a selected folder.
- **Choose output folder** — decide where converted PDFs are saved.
- **Progress reporting** — track conversion status and see errors for failed files.
- **Local processing** — no upload to a third-party conversion website.
- **Two conversion backends**:
  - Microsoft PowerPoint (recommended for layout fidelity)
  - LibreOffice (alternative if PowerPoint is not installed)

## Preview

When launched, the app opens a desktop window where you can add presentations, choose an output folder, and start conversion.

## Requirements

- Windows 10 or Windows 11
- Python 3.10 or newer
- **One** of the following conversion options:
  - Microsoft PowerPoint installed, plus `pywin32`
  - LibreOffice installed and discoverable by the app

## Installation

### 1. Download the project

Clone this repository:

```bash
git clone https://github.com/YOUR_USERNAME/pptx-to-pdf-converter.git
cd pptx-to-pdf-converter
```

Or download the repository as a ZIP file from GitHub.

### 2. Install Python dependencies

If you plan to use Microsoft PowerPoint as the conversion backend, install `pywin32`:

```bash
py -m pip install pywin32
```

If you plan to use LibreOffice, install it from the [official LibreOffice website](https://www.libreoffice.org/download/download-libreoffice/). The app checks common Windows installation paths and the system `PATH` for `soffice.exe`.

> You generally do not need both backends. Install the option that fits your computer.

## Run the application

From the project directory, run:

```bash
py pptx_to_pdf_converter.py
```

A desktop window will open.

## How to use

1. Click **Add PPTX files…** to choose individual presentations, or **Add folder…** to add all `.pptx` files from a folder.
2. Click **Browse…** to choose the folder where PDFs should be saved.
3. Click **Convert to PDF**.
4. Wait for conversion to finish. The app reports successes and any failed files.

## How it works

- If `pywin32` and Microsoft PowerPoint are available, the application uses PowerPoint's COM automation to export each presentation as a PDF.
- Otherwise, if LibreOffice is found, it uses LibreOffice in headless mode.
- The conversion runs locally. No presentation content is sent to an external service.

## Project structure

```text
pptx-to-pdf-converter/
├── pptx_to_pdf_converter.py   # Desktop GUI and conversion logic
├── README.md                  # Project documentation
└── .gitignore                 # Git exclusions
```

## Limitations and notes

- PDFs are static documents; PowerPoint animations and transitions are not preserved as interactive effects.
- Rendering can differ between systems if fonts or other presentation assets are missing.
- Microsoft PowerPoint usually gives the closest match to the original slide layout.
- The folder picker currently scans the selected folder's top level, not nested subfolders.
- If different source folders contain `.pptx` files with the same filename, their PDFs may target the same output filename. Rename files or convert them into separate output folders to avoid overwriting.
- Very large presentations or large batches can take longer to process.

## Troubleshooting

### `No supported converter found`

Install Microsoft PowerPoint and run:

```bash
py -m pip install pywin32
```

Or install LibreOffice and ensure `soffice.exe` is available in a standard installation path or the system `PATH`.

### `No module named win32com`

Install the PowerPoint automation dependency:

```bash
py -m pip install pywin32
```

### A presentation fails to convert

- Open the `.pptx` in PowerPoint or LibreOffice to confirm it is valid.
- Check that the output folder is writable.
- Confirm that the presentation is not password-protected or corrupted.
- Review the error shown by the application.

## Contributing

Issues and pull requests are welcome. If you report a bug, include your Windows version, conversion backend, and the error message. Do not upload confidential presentations when reporting issues.

## License

No license has been specified yet. Add a `LICENSE` file before publishing if you want to grant others explicit permission to use, modify, and redistribute this project.
