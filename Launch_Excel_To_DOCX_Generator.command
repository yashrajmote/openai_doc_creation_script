#!/bin/bash
# Excel to DOCX Generator - Launcher Script
# This script launches the drag and drop app

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Change to the script directory
cd "$SCRIPT_DIR"

# Activate virtual environment and run the app
if [ -d "venv" ]; then
    source venv/bin/activate
    python3 drag_drop_app.py
else
    echo "Error: Virtual environment not found!"
    echo "Please run: python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt"
    read -p "Press Enter to exit..."
fi
