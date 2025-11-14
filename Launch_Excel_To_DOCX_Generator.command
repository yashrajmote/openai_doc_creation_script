#!/bin/bash
# Excel to DOCX Generator - Launcher Script
# This script launches the drag and drop app

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Change to the script directory
cd "$SCRIPT_DIR"

echo "=================================="
echo "Excel to DOCX Generator"
echo "=================================="
echo "Script directory: $SCRIPT_DIR"
echo ""

# Activate virtual environment and run the app
if [ -d "venv" ]; then
    echo "✓ Virtual environment found"
    echo "Activating virtual environment..."
    source venv/bin/activate
    
    echo "Starting application..."
    python3 drag_drop_app.py
    
    EXIT_CODE=$?
    echo ""
    echo "Application exited with code: $EXIT_CODE"
    
    if [ $EXIT_CODE -ne 0 ]; then
        echo "❌ Application encountered an error"
    fi
else
    echo "❌ Error: Virtual environment not found!"
    echo "Please run: python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt"
fi

echo ""
read -p "Press Enter to close this window..."
