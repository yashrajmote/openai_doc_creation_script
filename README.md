# Excel to DOCX Generator

A Python script that processes Excel files containing company and job position data, filters out invalid entries, and generates individual DOCX files for each valid company.

## Features

- **Smart Filtering**: Automatically filters out:
  - Empty cells (missing company or position data)
  - Date entries in the company column (e.g., "10-Aug", "11-Aug")
  - Job board names (LinkedIn, Indeed, Handshake, Glassdoor, JobRight, GitHub)
- **DOCX Generation**: Creates professional-looking Word documents for each valid company
- **Batch Processing**: Processes entire Excel files at once
- **Error Handling**: Robust error handling with detailed progress reporting
- **Customizable Output**: Choose your own output directory

## Requirements

- Python 3.7+
- Required packages (see requirements.txt):
  - pandas
  - python-docx
  - openpyxl

## Installation

1. Clone or download this repository
2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage
```bash
python excel_to_docx_generator.py your_file.xlsx
```

### Specify Output Directory
```bash
python excel_to_docx_generator.py your_file.xlsx -o /path/to/output/directory
```

### Example
```bash
python excel_to_docx_generator.py job_applications.xlsx -o ./my_docx_files
```

## Excel File Format

Your Excel file must contain two columns:
- **Company**: Company names
- **Position**: Job positions/titles

The script will automatically detect and process these columns.

## Output

The script generates:
- Individual DOCX files for each valid company
- Files are named: `{index}_{company_name}_{position}.docx`
- Each DOCX contains:
  - Company information
  - Position details
  - Application date
  - Notes section for follow-up

## Filtering Rules

The script will skip entries where:
1. Company or Position cells are empty
2. Company field contains dates (e.g., "10-Aug", "11-Aug")
3. Company field contains job board names:
   - LinkedIn
   - Indeed
   - Handshake
   - Glassdoor
   - JobRight
   - GitHub

## Error Handling

The script provides detailed feedback:
- ✓ Successfully created files
- ✗ Skipped entries with reasons
- Processing summary with statistics
- Error messages for troubleshooting

## Example Output

```
Reading Excel file: job_applications.xlsx
Found 77 rows in Excel file
✓ Created: 001_Old_Mission_Full_Stack_Software_Engineer.docx
✓ Created: 002_Blue_Mountain_Quality_Resources_LLC_Full_Stack_Software_Engineer.docx
✗ Skipped row 4: date in company field
✗ Skipped row 5: job board name
...

==================================================
PROCESSING SUMMARY
==================================================
Total rows processed: 77
Valid entries: 45
Skipped entries: 32
DOCX files created: 45
Output directory: ./generated_docx_files
```

## Troubleshooting

1. **"Excel file not found"**: Check the file path and ensure the file exists
2. **"Excel file must contain 'Company' and 'Position' columns"**: Verify your Excel file has the correct column headers
3. **Permission errors**: Ensure you have write permissions to the output directory
4. **Import errors**: Run `pip install -r requirements.txt` to install dependencies

## License

This project is open source and available under the MIT License.
