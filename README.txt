# PPTX to PDF Converter (Windows)

A small desktop GUI for converting one or many `.pptx` PowerPoint files to PDF. It runs locally on your computer; your presentations are not uploaded to a website.

## Option A — Use Microsoft PowerPoint (recommended for best fidelity)

1. Install Microsoft PowerPoint / Microsoft Office.
2. Install Python 3.10+ from https://www.python.org/downloads/ and enable **Add Python to PATH** during setup.
3. Open Command Prompt in this folder and run:
   ```bat
   py -m pip install pywin32
   ```
4. Run the tool:
   ```bat
   py pptx_to_pdf_converter.py
   ```

## Option B — Use LibreOffice

If Microsoft PowerPoint is not installed, install LibreOffice from https://www.libreoffice.org/download/download-libreoffice/. The tool looks for `soffice.exe` in the usual Windows install locations or on PATH.

Then run:
```bat
py pptx_to_pdf_converter.py
```

## Features

- Add individual `.pptx` files or all `.pptx` files in a folder.
- Choose where PDFs should be saved.
- Batch conversion with progress and error reporting.
- Keeps conversion local.

## Notes

- Microsoft PowerPoint is generally the best option for preserving exact fonts, animations-to-static-slide rendering, and layout fidelity.
- PDFs cannot preserve slide animations or transitions.
- The current folder picker scans only the selected folder's top level. (The recursive option is reserved for a future enhancement.)
- If two selected presentations have the same filename, the later output may overwrite the earlier PDF. Choose separate input batches or rename files to avoid collisions.
