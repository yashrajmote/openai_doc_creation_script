# Excel to DOCX Generator

A drag-and-drop app that converts Excel job applications into individual Word documents.

## Quick Start

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the app:
   ```bash
   python3 drag_drop_app.py
   ```

3. Drag your Excel file onto the app

## Excel Format

Your Excel file needs two columns:
- **Company**: Company names
- **Position**: Job titles

## What It Does

- Filters out empty rows, dates, and job board names
- Creates one DOCX file per valid company
- Files named: `{index}_{company}_{position}.docx`

## Command Line

```bash
python3 excel_to_docx_generator.py your_file.xlsx
```
