#!/usr/bin/env bash
# build_mac.sh — Local macOS build via PyInstaller
# Usage: bash build_mac.sh
set -e

echo "📦 Installing dependencies..."
pip install --upgrade pip
pip install openpyxl pyinstaller

echo "🔨 Building macOS .app..."
pyinstaller \
  --name "Excel to DDL Automator" \
  --windowed \
  --onefile \
  --clean \
  --add-data "assets:assets" \
  main.py

echo ""
echo "✅ Build complete!"
echo "   App:  dist/Excel to DDL Automator.app"
echo "   Binary: dist/Excel to DDL Automator"
echo ""
echo "To create a DMG manually:"
echo "  hdiutil create -volname 'Excel to DDL Automator' \\"
echo "    -srcfolder 'dist/Excel to DDL Automator.app' \\"
echo "    -ov -format UDZO dist/ExcelToDDL_mac.dmg"
