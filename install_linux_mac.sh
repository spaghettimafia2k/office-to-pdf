#!/usr/bin/env bash
# Office to PDF Converter — Linux / macOS Setup
# ======================================
# Installs LibreOffice (the free conversion engine).

set -e

echo "==================================================="
echo " Office to PDF Converter - Linux/macOS Setup"
echo "==================================================="
echo

# Detect OS
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    echo "Detected: Linux"
    if command -v apt-get &>/dev/null; then
        sudo apt-get update -qq
        sudo apt-get install -y libreoffice
    elif command -v dnf &>/dev/null; then
        sudo dnf install -y libreoffice
    elif command -v pacman &>/dev/null; then
        sudo pacman -S --noconfirm libreoffice-still
    else
        echo "[WARNING] Unknown package manager. Install LibreOffice manually:"
        echo "  https://www.libreoffice.org/download/"
    fi

elif [[ "$OSTYPE" == "darwin"* ]]; then
    echo "Detected: macOS"
    if command -v brew &>/dev/null; then
        brew install --cask libreoffice
    else
        echo "[WARNING] Homebrew not found. Install LibreOffice manually:"
        echo "  https://www.libreoffice.org/download/"
    fi
fi

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] Python 3 is not installed."
    echo "Install it from https://www.python.org/downloads/"
    exit 1
fi

echo
echo "[OK] Python 3 found: $(python3 --version)"
echo "[OK] LibreOffice: $(libreoffice --version 2>/dev/null || echo 'see above')"
echo
echo "==================================================="
echo " Setup complete!"
echo
echo " Usage:"
echo "   python3 convert.py yourfile.docx"
echo "   python3 convert.py yourfile.pptx --output final.pdf"
echo "   python3 convert.py /path/to/folder/"
echo "==================================================="
