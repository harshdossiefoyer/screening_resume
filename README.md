# Resume Information Extractor

A Python script that automatically extracts key information from PDF resumes including names, email addresses, phone numbers, and graduation/passout years.

## Features

- Extracts the following information from PDF resumes:
  - Full Name
  - Email Address (with common typo corrections)
  - Phone Number
  - Graduation/Passout Year
- Processes multiple PDF files in a directory
- Exports results to an Excel file with formatted columns
- Handles various formats of:
  - Phone numbers (including international formats)
  - Email addresses
  - Name patterns
  - Education year mentions

## Requirements

```bash
pip install -r requirements.txt
```

Key dependencies:
- pdfplumber==0.11.6
- pandas==2.2.3
- openpyxl==3.1.5

## Usage

1. Place your PDF resumes in a directory
2. Update the `folder_path` in "def(main)" function  to point to your resume directory:
```python
folder_path = r"path/to/your/resume/folder"
```

3. Run the script:
```bash
python auto.py
```

4. The script will generate `resume_info.xlsx` with the extracted information

## Output Format

The script generates an Excel file with the following columns:
- Filename
- Name
- Email
- Phone
- Passout Year

