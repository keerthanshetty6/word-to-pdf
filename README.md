# Word to PDF Utility

Batch convert `.doc` and `.docx` files to PDF using Microsoft Word automation.

---

## Requirements

- Windows
- Microsoft Word installed
- Python 3.9+
- pywin32

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Usage

### Basic

```bash
python word_to_pdf.py "C:\path\to\word" "C:\path\to\pdf"
```

### Recursive (include subfolders)

```bash
python word_to_pdf.py -r "C:\path\to\word" "C:\path\to\pdf"
```

### Overwrite existing PDFs

```bash
python word_to_pdf.py --no-skip "C:\path\to\word" "C:\path\to\pdf"
```

---

## How It Works

- Opens Microsoft Word in the background
- Converts each `.doc` / `.docx` file
- Saves it as `.pdf`
- Skips existing PDFs by default
- Handles common Word automation errors automatically

---

## Common Issues

### "Call was rejected by callee"

This usually means Microsoft Word is busy or blocked by a dialog.

**Fix:**

- Close all Word windows
- End `WINWORD.EXE` in Task Manager
- Make sure no file opens in Protected View
- Avoid converting directly from OneDrive/SharePoint folders (copy locally first if needed)

---

## License

MIT License
