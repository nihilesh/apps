#!/usr/bin/env bash
set -e

echo "================================================="
echo " Archaeological Literature Extractor"
echo "================================================="

# Locate Python 3 binary on macOS or Linux
if command -v python3 &>/dev/null; then
    PYTHON_BIN="python3"
elif command -v python &>/dev/null; then
    PYTHON_BIN="python"
else
    echo "Error: Python 3 is required."
    exit 1
fi

echo "[1/2] Installing dependencies..."
$PYTHON_BIN -m pip install --quiet -r requirements.txt

echo "[2/2] Running research extraction script..."
$PYTHON_BIN scan_research_data.py

echo "================================================="
echo " Complete! Results saved to Search_Results_Populated.xlsx"
echo "================================================="
