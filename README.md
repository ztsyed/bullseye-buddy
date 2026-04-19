# Bullseye Score Reader

A web app that reads Bullseye Pistol League score sheets from photos and converts them into editable, downloadable Excel spreadsheets.

## Features

- **AI-powered OCR** — Upload a photo of a handwritten score sheet and Claude's vision API extracts all scores automatically
- **Live editing** — Review and correct extracted scores in an interactive table that mirrors the score sheet layout
- **Auto-calculation** — Row subtotals, stage totals, match aggregates, and grand aggregate update in real-time as you edit
- **Validation** — Invalid shot values (outside 0-10, X, M) are highlighted in red
- **Floating image viewer** — Draggable, zoomable overlay of the original score sheet for easy comparison while editing
- **Excel export** — Download a styled `.xlsx` file matching the score sheet format, with the original image embedded on a second sheet
- **History** — Save scans to a local database and re-view or re-download them later
- **HEIC support** — Handles iPhone photos (HEIC) by auto-converting to JPEG

## Score Sheet Format

Supports the **Sunnyvale Rod & Gun Club Bullseye Pistol League** score sheet with:

- .22 Match (Rimfire) and C.F. Match (Centerfire) sections
- 20 Shots Slow Fire, Timed Fire, and Rapid Fire (two 10-shot strings each)
- X = bullseye (scored as 10), M = miss (scored as 0)
- Totals in "score-Xcount" format (e.g., "187-3" = 187 points with 3 Xs)

## Setup

### Prerequisites

- Python 3.9+
- An [Anthropic API key](https://console.anthropic.com/)
- macOS (for HEIC conversion via `sips`; JPEG/PNG work on any OS)

### Installation

```bash
# Clone the repo
git clone <repo-url>
cd bullseye

# Create virtual environment and install dependencies
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Running

```bash
source venv/bin/activate
python app.py
```

Open **http://localhost:5050** in your browser.

### First-time setup

1. Click **Settings** in the top-right corner
2. Paste your Anthropic API key
3. Click **Save** — the key is stored in your browser's localStorage only

## Usage

1. **Upload** — Drag and drop or click to upload a score sheet photo
2. **Review** — Check the extracted scores against the original image (click the preview to open the floating viewer)
3. **Edit** — Click any cell to correct misread values; totals update automatically
4. **Save** — Click "Save to History" to persist the scan
5. **Download** — Click "Download Excel" to get the `.xlsx` file

## Project Structure

```
bullseye/
├── app.py              # Flask backend (API extraction, Excel export, history)
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html      # Single-page frontend
└── data/
    └── IMG_6016.HEIC   # Sample score sheet
```
