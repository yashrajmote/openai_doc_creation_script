# Excel to DOCX Generator

A drag-and-drop app that converts Excel job applications into individual Word documents using AI and custom templates.

## Quick Start

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Set up templates:
   - Place your Java resume template as `templates/java_resume_template.dotx`
   - Place your C# resume template as `templates/csharp_resume_template.dotx`

3. Run the app:
   ```bash
   python3 drag_drop_app.py
   ```

4. Drag your Excel file onto the app

## Excel Format

Your Excel file needs these columns:
- **Company**: Company names
- **Position**: Job titles
- **Description** (or similar): Job descriptions for AI tailoring

## What It Does

- **Smart Template Selection**: Automatically chooses Java or C# template based on job description
- **Individual Resume Files**: Creates one copy of the appropriate template per job (no AI modification)
- **Master Document**: Generates a combined document with all AI responses for reference
- **Template Copying**: Simply copies your template files with job-specific naming

## Template Setup

1. Create a `templates/` folder in your project directory
2. Add your resume templates:
   - `java_resume_template.dotx` - For Java/Spring positions
   - `csharp_resume_template.dotx` - For C#/.NET positions
3. Templates should contain your base resume content and formatting

## Output Files

- **Individual Resume Files**: `CompanyName_Position.docx` (exact copies of your templates)
- **Master Document**: `AI_Generated_Resumes.docx` (contains all AI responses for reference)

## Command Line

```bash
python3 excel_to_docx_generator.py your_file.xlsx
```
