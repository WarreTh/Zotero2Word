# Zotero2Word

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)](https://www.python.org/)
[![License: AGPL v3.0](https://img.shields.io/badge/License-AGPL%20v3.0-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20Windows-lightgrey)](https://wkhtmltopdf.org/)

Create a beautiful Word document from your Zotero library, including all folders, subfolders, notes, images, and HTML snapshots.

## Features

- Exports your entire Zotero library to a well-formatted Word (.docx) document
- Preserves folder and subfolder structure as Word headers
- Includes all notes, formatted for readability
- Embeds images and HTML snapshots from attachments
- Manual clickable Table of Contents

## Installation

1. **Clone this repository:**

   ```fish
   git clone https://github.com/WarreTh/Zotero2Word.git
   cd Zotero2Word
   ```

2. **Install Python dependencies:**

   ```bash
   pip install pyzotero python-docx tqdm beautifulsoup4 lxml requests imgkit
   ```

3. **Install system dependencies (choose your OS):**

   **On Windows:**
   - Download and install `wkhtmltopdf` (includes `wkhtmltoimage`) from [wkhtmltopdf.org](https://wkhtmltopdf.org/downloads.html)
   - Add the installation folder (usually `C:\Program Files\wkhtmltopdf\bin`) to your PATH environment variable.
   - (Optional, recommended) Use pipenv for isolated Python environments:

     ```fish
     pip install pipenv
     pipenv install --dev
     pipenv shell
     ```

   **Other OSes (Linux/macOS):**
   - On Ubuntu/Debian:

     ```fish
     sudo apt install wkhtmltopdf
     ```

   - On macOS:

     ```fish
     brew install wkhtmltopdf
     ```

   - Or download from [wkhtmltopdf.org](https://wkhtmltopdf.org/)
   - (Optional, recommended) Use pipenv for isolated Python environments:

     ```fish
     pip install pipenv
     pipenv install --dev
     pipenv shell
     ```

4. **Edit `config.py` file:**
   - Propably not needed

5. **Enable Zotero Local Server API:**
   - Open Zotero.
   - Go to `Edit` > `Preferences` > `Advanced` > `General` > `Advanced Configuration`
   - Tick `Allow other applications on this computer to communicate with Zotero`
   - Restart Zotero if needed.

## Usage

Run the script:

```fish
python Zotero2Word.py
```

## TODO:
- Fix attachment-path showing as None
  - Thats why attachments arent shown in the doc
The output Word document will be saved to the path specified in your `config.py`.

## License

This project is licensed under the GNU Affero General Public License v3.0. See the [LICENSE](LICENSE) file for details.
