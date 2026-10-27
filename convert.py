"""
Office to PDF Converter
======================
Converts .doc, .docx, .xls, .xlsx, .ppt, .pptx -> PDF with maximum fidelity.

Priority engine order:
  1. Microsoft Office COM (Windows only, requires MS Office installed)
     -> Uses the actual Office renderer = same result as File > Export > PDF
  2. LibreOffice headless (cross-platform, free)

Usage:
  python convert.py file.docx
  python convert.py file.pptx --output my_output.pdf
  python convert.py folder/                # convert every supported file in folder
  python convert.py file.docx --engine libreoffice

Right-click menu (Windows):
  Run install_context_menu.bat once. After that, right-click any Word,
  Excel or PowerPoint file (or a folder) and choose "Convert to PDF".
"""

from __future__ import annotations

import argparse
import contextlib
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

SUPPORTED = {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"}


# ──────────────────────────────────────────────
# Engine: Microsoft Office COM (Windows only)
# ──────────────────────────────────────────────
def _office_com_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def convert_with_office(src: Path, dst: Path) -> None:
    """
    Use MS Office COM automation to export a PDF.

    Word and Excel are started as a *separate, hidden* instance (DispatchEx)
    so we never touch documents the user already has open.
    PowerPoint only allows one instance, so we attach to it and only quit
    it if it was not already running with presentations open.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    ext = src.suffix.lower()
    src_abs = str(src.resolve())
    dst_abs = str(dst.resolve())

    try:
        if ext in {".doc", ".docx"}:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            try:
                doc = app.Documents.Open(src_abs, ReadOnly=True, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(
                    OutputFileName=dst_abs,
                    ExportFormat=17,          # wdExportFormatPDF
                    OptimizeFor=0,            # wdExportOptimizeForPrint
                    CreateBookmarks=1,        # headings -> PDF bookmarks
                    DocStructureTags=True,
                    BitmapMissingFonts=True,  # never substitute a missing font
                    UseISO19005_1=False,
                )
                doc.Close(False)
            finally:
                app.Quit()

        elif ext in {".xls", ".xlsx"}:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                wb = app.Workbooks.Open(src_abs, ReadOnly=True)
                wb.ExportAsFixedFormat(0, dst_abs)  # xlTypePDF
                wb.Close(False)
            finally:
                app.Quit()

        elif ext in {".ppt", ".pptx"}:
            app = win32com.client.Dispatch("PowerPoint.Application")
            was_busy = app.Presentations.Count > 0
            try:
                prs = app.Presentations.Open(
                    src_abs, ReadOnly=True, Untitled=False, WithWindow=False
                )
                prs.SaveAs(dst_abs, 32)  # ppSaveAsPDF
                prs.Close()
            finally:
                if not was_busy:
                    app.Quit()
        else:
            raise ValueError(f"Unsupported extension for Office COM: {ext}")
    finally:
        pythoncom.CoUninitialize()


# ──────────────────────────────────────────────
# Engine: LibreOffice headless
# ──────────────────────────────────────────────
def _find_libreoffice() -> str | None:
    candidates = [
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        found = shutil.which(c)
        if found:
            return found
        if os.path.isfile(c):
            return c
    return None


def convert_with_libreoffice(src: Path, dst: Path) -> None:
    lo = _find_libreoffice()
    if not lo:
        raise RuntimeError(
            "No conversion engine found.\n"
            "  - On Windows: install Microsoft Office and run install_windows.bat, or\n"
            "  - Install LibreOffice (free): https://www.libreoffice.org/download/"
        )

    # Convert into a private temp folder, then move into place. This avoids
    # clobbering an unrelated PDF and lets several conversions run at once.
    with tempfile.TemporaryDirectory(prefix="office-to-pdf-") as tmp:
        cmd = [
            lo,
            "--headless",
            "--norestore",
            f"-env:UserInstallation={(Path(tmp) / 'profile').as_uri()}",
            "--convert-to", "pdf",
            "--outdir", tmp,
            str(src.resolve()),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        produced = Path(tmp) / (src.stem + ".pdf")
        if result.returncode != 0 or not produced.exists():
            raise RuntimeError(
                "LibreOffice conversion failed:\n"
                + (result.stderr or result.stdout or "(no output)")
            )
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(dst))


# ──────────────────────────────────────────────
# One-at-a-time lock (used by the right-click menu)
# ──────────────────────────────────────────────
@contextlib.contextmanager
def _single_instance_lock():
    """
    When a user selects 10 files and clicks "Convert to PDF", Windows starts
    10 copies of this script at the same moment. Office does not like that,
    so the copies wait in line and convert one after another.
    """
    lock_path = Path(tempfile.gettempdir()) / "office-to-pdf.lock"
    fh = open(lock_path, "a+b")
    try:
        if sys.platform == "win32":
            import msvcrt
            announced = False
            while True:
                try:
                    fh.seek(0)
                    msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                    break
                except OSError:
                    if not announced:
                        print("  Waiting for other conversions to finish...")
                        announced = True
                    time.sleep(0.5)
        else:
            import fcntl
            fcntl.flock(fh.fileno(), fcntl.LOCK_EX)
        yield
    finally:
        try:
            if sys.platform == "win32":
                import msvcrt
                fh.seek(0)
                msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(fh.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass
        fh.close()


# ──────────────────────────────────────────────
# Main conversion dispatcher
# ──────────────────────────────────────────────
def convert(src: Path, dst: Path | None = None, engine: str = "auto") -> Path:
    src = Path(src).resolve()
    if not src.exists():
        raise FileNotFoundError(f"Input file not found: {src}")
    if src.name.startswith("~$"):
        raise ValueError("This is an Office temporary/lock file, not a real document.")
    if src.suffix.lower() not in SUPPORTED:
        raise ValueError(
            f"Unsupported file type '{src.suffix}'. "
            f"Supported: {', '.join(sorted(SUPPORTED))}"
        )

    dst = Path(dst).resolve() if dst else src.with_suffix(".pdf")

    print(f"  Converting : {src.name}")
    print(f"  Output     : {dst}")

    use_office = engine == "office" or (engine == "auto" and _office_com_available())
    if engine == "office" and not _office_com_available():
        raise RuntimeError("Office engine needs Windows + Microsoft Office + pywin32.")

    if use_office:
        print("  Engine     : Microsoft Office")
        convert_with_office(src, dst)
    else:
        print("  Engine     : LibreOffice")
        convert_with_libreoffice(src, dst)

    if not dst.exists():
        raise RuntimeError("Conversion finished but the PDF was not created.")

    print(f"  Done       : {max(1, dst.stat().st_size // 1024)} KB")
    return dst


def _collect(path: Path) -> list[Path]:
    if path.is_dir():
        return sorted(
            f for f in path.iterdir()
            if f.is_file() and f.suffix.lower() in SUPPORTED and not f.name.startswith("~$")
        )
    return [path]


# ──────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────
def main() -> int:
    parser = argparse.ArgumentParser(
        prog="office-to-pdf",
        description="Convert .doc/.docx/.xls/.xlsx/.ppt/.pptx to PDF with maximum fidelity.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python convert.py report.docx
  python convert.py slides.pptx --output slides_final.pdf
  python convert.py ./my_folder/
  python convert.py file.docx --engine libreoffice
        """,
    )
    parser.add_argument("inputs", nargs="+", help="File(s) or folder(s) to convert")
    parser.add_argument("--output", "-o", default=None,
                        help="Output PDF path (only when converting a single file)")
    parser.add_argument("--engine", choices=["auto", "office", "libreoffice"],
                        default="auto", help="Conversion engine (default: auto)")
    parser.add_argument("--from-menu", action="store_true",
                        help=argparse.SUPPRESS)  # used by the right-click menu
    args = parser.parse_args()

    if args.from_menu and sys.platform == "win32":
        os.system("title Office to PDF")

    files: list[Path] = []
    for raw in args.inputs:
        files.extend(_collect(Path(raw)))

    if not files:
        print("No supported files found.")
        return _finish(args.from_menu, ok=False)

    if args.output and len(files) > 1:
        print("--output can only be used with a single file.")
        return _finish(args.from_menu, ok=False)

    print(f"\nOffice to PDF: {len(files)} file(s)\n")
    errors: list[tuple[Path, Exception]] = []

    lock = _single_instance_lock() if args.from_menu else contextlib.nullcontext()
    with lock:
        for f in files:
            try:
                convert(f, dst=args.output, engine=args.engine)
            except Exception as e:  # keep going with the other files
                print(f"  ERROR      : {e}")
                errors.append((f, e))
            print()

    if errors:
        print(f"{len(errors)} of {len(files)} file(s) failed:")
        for f, e in errors:
            print(f"  - {f.name}: {e}")
        return _finish(args.from_menu, ok=False)

    print(f"All {len(files)} file(s) converted. PDFs are next to the originals.")
    return _finish(args.from_menu, ok=True)


def _finish(from_menu: bool, ok: bool) -> int:
    """From the right-click menu: close by itself on success, wait on error."""
    if from_menu:
        if ok:
            print("\nThis window closes in 3 seconds...")
            time.sleep(3)
        else:
            input("\nPress Enter to close this window...")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
