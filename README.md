# Office → PDF Converter

> Convert `.doc`, `.docx`, `.xlsx`, `.xls`, `.pptx`, `.ppt` to PDF while keeping **fonts, layout, images, and tables exactly as they are** in the original file.

Share this repo with friends so everyone can convert files the same way.

---

## Why does my PDF look broken?

The most common causes of bad conversions:

| Problem | Root cause |
|---|---|
| Font substituted / wrong font | The font in your file isn't installed on the converter machine |
| Layout shifts, columns move | Different rendering engine (Google Docs, online tool) interprets styles differently |
| Images missing or blurry | Embedded EMF/WMF graphics not supported by the converter |
| SmartArt becomes a white box | SmartArt is rendered by Office only, not third-party tools |

**The solution:** use the actual Microsoft Office renderer (via COM automation) instead of a third-party engine. This repo does that automatically on Windows if you have Office installed.

---

## How it works — Engine priority

```
Is Windows + MS Office installed?
  YES → use Office COM engine  ✅ (pixel-perfect, same as Print → Save as PDF)
  NO  → use LibreOffice headless  ✅ (good fidelity, free, cross-platform)
```

You can force an engine with `--engine office` or `--engine libreoffice`.

---

## Quick start

### 1. Clone the repo

```bash
git clone https://github.com/YOUR_USERNAME/office-to-pdf.git
cd office-to-pdf
```

### 2. Install

**Windows** (with Microsoft Office installed — recommended for best quality):
```bat
install_windows.bat
```

**Linux / macOS** (uses LibreOffice):
```bash
chmod +x install_linux_mac.sh
./install_linux_mac.sh
```

### 3. Convert

#### Easiest: right-click (Windows)

`install_windows.bat` also adds **Convert to PDF** to the right-click menu. Then:

1. Right-click any `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx` file (or a whole folder)
2. Click **Convert to PDF**
3. The PDF appears next to the original file

You can select several files at once; they are converted one after another.
On **Windows 11** the item is under **Show more options** (or hold **Shift** while right-clicking).

| Want to... | Run |
|---|---|
| Add only the right-click menu | `install_context_menu.bat` |
| Remove the right-click menu | `uninstall_context_menu.bat` |
| Moved the `office-to-pdf` folder? | Run `install_context_menu.bat` again |

#### From the command line

```bash
# Single file
python convert.py report.docx
python convert.py slides.pptx --output slides_final.pdf
python convert.py budget.xlsx

# Whole folder (batch convert everything inside)
python convert.py C:\Users\Pele\Documents\submissions\

# Force a specific engine
python convert.py file.docx --engine libreoffice
python convert.py file.docx --engine office
```

---

## Tips for the best output

### ✅ Embed fonts in your Word / PowerPoint file before converting

**Word:** File → Options → Save → check **"Embed fonts in the file"**

**PowerPoint:** File → Options → Save → check **"Embed fonts in the file"**

This ensures the fonts travel with the document and won't be substituted.

### ✅ Use the Office COM engine on Windows
If you have Microsoft Office installed, the `--engine office` path uses Office itself to render the PDF, which is identical to doing *File → Export → PDF* manually inside Word/Excel/PowerPoint.

### ✅ For Excel: set print area first
Open the workbook → select cells → Page Layout → Print Area → Set Print Area.
This prevents extra blank pages or cut-off columns.

### ✅ Avoid complex SmartArt if using LibreOffice
SmartArt is a proprietary Office format. If your file has SmartArt and you're on Linux/macOS, consider flattening it to an image inside Office first (right-click SmartArt → Save as Picture).

---

## Requirements

| Engine | Requirement |
|---|---|
| Office COM (Windows) | Windows + Microsoft Office 2016 or newer + `pip install pywin32` |
| LibreOffice | [LibreOffice](https://www.libreoffice.org/download/) installed (free) |
| Python | Python 3.8 or newer (tick "Add python.exe to PATH" when installing on Windows) |

---

## Research: Tools & Services Compared

We compared the conversion tools out there so you know which approach to use when:

### 🟢 Highest fidelity (use these)

| Tool | Type | Quality | Cost | Notes |
|---|---|---|---|---|
| **MS Office COM** | Local (Windows) | ⭐⭐⭐⭐⭐ | Free (need Office) | Identical to manual Export PDF |
| **Microsoft Graph API** | Cloud API | ⭐⭐⭐⭐⭐ | Free tier / Azure cost | Uses Office 365 renderer; needs Microsoft account |
| **Adobe Acrobat Pro** | Desktop/API | ⭐⭐⭐⭐⭐ | Paid (~$25/mo) | Gold standard but expensive |
| **Aspose** | Cloud/library API | ⭐⭐⭐⭐ | Paid | Very accurate, handles SmartArt |

### 🟡 Good fidelity (free options)

| Tool | Type | Quality | Cost | Notes |
|---|---|---|---|---|
| **LibreOffice headless** | Local CLI | ⭐⭐⭐⭐ | Free | Best free option; some SmartArt / font issues |
| **Gotenberg** | Self-hosted Docker API | ⭐⭐⭐⭐ | Free | Wraps LibreOffice; great for server deployment |
| **IronPDF** | Library | ⭐⭐⭐⭐ | Paid (free trial) | Good .NET integration |

### 🔴 Avoid for complex documents

| Tool | Why it breaks |
|---|---|
| Google Docs import/export | Completely re-renders layout; loses custom styles, fonts, positioning |
| Smallpdf / ILovePDF / Convertio | Online tools; upload your data to their servers; inconsistent engine |
| Word "Save as PDF" on a different machine | Different fonts installed → font substitution |
| `python-docx` + reportlab | Code-level rebuild, not a render; loses 90% of formatting |

---

## File structure

```
office-to-pdf/
├── convert.py                  ← main script
├── install_windows.bat         ← Windows setup (pywin32 + right-click menu)
├── install_context_menu.bat    ← adds "Convert to PDF" to the right-click menu
├── uninstall_context_menu.bat  ← removes it
├── pdf.ico                     ← icon shown in the right-click menu
├── install_linux_mac.sh        ← Linux/macOS setup helper
├── requirements.txt            ← Python dependencies
└── README.md                   ← this file
```

---

## License

MIT — free to use, modify, and share.

---

*Built at อ.เต้ Research Lab · CMU*
