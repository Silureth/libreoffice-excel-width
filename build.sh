#!/usr/bin/env bash
# Builds ExcelColumnWidth.oxt. Keep this file name: the command URLs refer to it.
set -euo pipefail
cd "$(dirname "$0")"
python3 make_icons.py src/icons
rm -f ExcelColumnWidth.oxt
find src -name __pycache__ -prune -exec rm -rf {} +
(cd src && zip -qrX ../ExcelColumnWidth.oxt META-INF description.xml description Addons.xcu Accelerators.xcu WindowState.xcu python icons)
echo "built $(pwd)/ExcelColumnWidth.oxt"
