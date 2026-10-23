#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
if [ "$(uname -s)" != "Darwin" ]; then
    echo "Сборку нужно запускать в macOS."
    exit 1
fi
python3 -m venv .venv-mac
.venv-mac/bin/python -m pip install -r requirements.txt 'pyinstaller==6.22.0'
.venv-mac/bin/python -m PyInstaller --noconfirm --clean --windowed --onedir \
    --target-architecture arm64 --name YMMatcher YM_matcher_v2.py
ditto -c -k --sequesterRsrc --keepParent dist/YMMatcher.app dist/YMMatcher-M1.zip
echo "Готово: dist/YMMatcher-M1.zip. Распакуйте и положите Excel-файлы рядом с YMMatcher.app."
