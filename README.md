<img width="1000" height="500" alt="AnalYzer Screenshot" src="https://github.com/user-attachments/assets/e979ca68-09d7-48ff-b30e-4801e6529da7" />

# AnalYzer (Image OCR → Excel/Word/PDF + AI)

AnalYzer is a Windows desktop application that extracts text from images (OCR), lets you OCR a selected region (“Lens” mode), improves OCR readability, and exports the result to Excel, Word, or PDF. It also includes optional AI tools for refining OCR text and analyzing/transcribing image content.

## Features

- Full image OCR (extract all text from an image)
- Selection/Lens OCR (drag-select a region and OCR only that part)
- “Increase Accuracy” helpers to clean up OCR output
- Optional AI tools (ask questions about extracted text, vision re-read for blurry images)
- Export to: Excel (`.xlsx`), Word (`.docx`), PDF (`.pdf`)

## Download (Windows)

Go to the **Releases** page of this repository and download the latest installer:

- `AnalYzer_Setup.exe`

Then run it and follow the setup steps.

## Requirements (Run from source)

- Python 3.10+ (recommended)
- Tesseract OCR installed (Windows)
- Groq API key (only if you use AI features)

## Setup (Development)

### 1) Clone the repo

```bash
git clone https://github.com/gjohn-jm/ImageAnalYzer.git
cd ImageAnalYzer
```
### 2) Create & activate a virtual environment (Windows)
```text
python -m venv .venv
.venv\Scripts\activate
```
### 3) Install dependencies
```text
pip install -r requirements.txt
```
### 4) Run the app
```text
python "ImageAnalYzer.py"
```
### Environment variables (AI)
This app uses an environment variable for the Groq API key (only needed for AI features).

### Windows PowerShell
powershell
```
setx GROQ_API_KEY "YOUR_KEY_HERE"
```
Restart the terminal (or reopen the app) after setting it.

Security: Do not commit API keys to GitHub.

### Build (Create the EXE)
1) Build with PyInstaller (onedir)
From your project folder:

```text
py -m PyInstaller --noconfirm --clean --windowed --onedir --name ImageAnalyzer --icon "PATH_TO_Analyzer.ico" "ImageAnalYzer.py"
```
This generates:

dist\ImageAnalyzer\ImageAnalyzer.exe (plus required folders/files)

### 2) Create the installer with Inno Setup
Open AnalYzer.iss in Inno Setup Compiler and compile.

It generates your shareable installer EXE:

AnalYzer_Setup.exe (inside your configured output folder)

Upload this installer to GitHub Releases.

Troubleshooting
Shortcut / “missing shortcut”
Make sure the installed folder contains:

ImageAnalyzer.exe

If your shortcut points to a filename that doesn’t exist, rebuild with PyInstaller and recompile the installer.

Icon not updating
Windows can cache icons. Try renaming the output installer/exe or restarting Windows Explorer, then rebuild.

Screenshots
You can add more screenshots by creating an assets/ folder and referencing images like:

text
undefined
