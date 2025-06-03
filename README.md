# Zotero2Word

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)](https://www.python.org/)
[![License: AGPL v3.0](https://img.shields.io/badge/License-AGPL%20v3.0-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20Windows-lightgrey)](https://wkhtmltopdf.org/)

A tool to export Zotero items and notes to Microsoft Word (.docx) with rich formatting, metadata, and embedded images.

## Features

- Export Zotero items and notes to Word with rich formatting
- Embed metadata (author(s), date, date added, tags, source URL) as small italic text
- Parse and render HTML content, including lists, code blocks, blockquotes, and images
- Insert images from local files or base64-encoded HTML
- Render HTML snapshots as images using html2image
- Add clickable hyperlinks

## Requirements
- Python 3.8+
- pyzotero
- python-docx
- tqdm
- beautifulsoup4
- lxml
- requests
- imgkit (optional)
- html2image
- wkhtmltopdf (system dependency, required for imgkit/html2image)

## Installation
> Make sure you have Python installed

### Windows (Recommended)

1. **Install Python and wkhtmltopdf automatically using winget:**

   ```powershell
   winget install --id=Python.Python.3.11 -e && winget install --id=wkhtmltopdf.wkhtmltox -e
   ```
   - If you already have Python installed, you can skip the first part of the command.
   - `winget` will add both Python and wkhtmltopdf to your PATH automatically.

2. **Clone this repository:**

   ```powershell
   git clone https://github.com/WarreTh/Zotero2Word.git
   cd Zotero2Word
   ```

3. **Install Python dependencies:**

   ```powershell
   pip install pyzotero python-docx tqdm beautifulsoup4 lxml requests imgkit html2image
   ```

   - You do **NOT** need pipx or pipenv on Windows. Just use the above pip command.

### Linux/macOS

1. **Clone this repository:**

   ```bash
   git clone https://github.com/WarreTh/Zotero2Word.git
   cd Zotero2Word
   ```

2. **Install Python dependencies:**

   ```bash
   pip install pyzotero python-docx tqdm beautifulsoup4 lxml requests imgkit html2image
   ```

3. **Install system dependencies:**

   - On Ubuntu/Debian:

     ```bash
     sudo apt install wkhtmltopdf
     ```

   - On macOS:

     ```bash
     brew install wkhtmltopdf
     ```

   - Or download from [wkhtmltopdf.org](https://wkhtmltopdf.org/)

5. **Edit `config.py` file:**
   - Probably not needed

6. **Enable Zotero Local Server API:**
   - Open Zotero.
   - Go to `Edit` > `Preferences` > `Advanced` > `General` > `Advanced Configuration`
   - Tick `Allow other applications on this computer to communicate with Zotero`
   - Restart Zotero if needed.

## Usage

Run the script:

```powershell
python Zotero2Word.py
```

The output Word document will be saved to the path specified in your `config.py`.

# License

This project is licensed under the GNU Affero General Public License v3.0. See the [LICENSE](LICENSE) file for details.

## Common Errors

**Error:** `wkhtmltoimage is not installed or not in PATH`
- **Solution:** Make sure you installed wkhtmltopdf using winget (Windows), `sudo apt install wkhtmltopdf` (Ubuntu/Debian), or `brew install wkhtmltopdf` (macOS). Restart your terminal or computer if needed.

**Error:** `ModuleNotFoundError: No module named 'pyzotero'`
- **Solution:** Run `pip install pyzotero` in your terminal.

**Error:** `Could not connect to local Zotero`
- **Solution:** Ensure Zotero is running and the Local Server API is enabled (see step 6 in Installation).

**Error:** `Image file not found` or `Attachment file does not exist`
- **Solution:** Check your `STORAGE_DIR` path in `config.py` and make sure your Zotero storage is accessible.
