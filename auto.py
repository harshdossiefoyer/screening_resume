import os
import re
import pandas as pd
from pathlib import Path
import pdfplumber

# Define regex patterns
NAME_PATTERN = r"^[A-Za-z\s\.\-']{2,}(?:\s[A-Za-z\s\.\-']{2,})+$"  # Handles names with initials, hyphens, etc.
EMAIL_PATTERN = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
PHONE_PATTERN = r"\b(\(\d{3}\)\s?\d{3}-\d{4}|\d{3}-\d{3}-\d{4}|\d{3}\.\d{3}\.\d{4}|\d{3}\s\d{3}\s\d{4}|\d{10}|\+\d{1,2}\s?\d{10})\b"
PASSOUT_PATTERN = r"(?:Bachelor|B\.Tech|B\.Sc|M\.Tech|M\.Sc|Master|Degree|Graduated|Diploma|University|College|Institute|[A-Z][a-z]+\s+University)\b[^0-9]*?(\d{4}|\d{4}\s*[-–—]\s*\d{4})\b"

# Common email domain corrections
EMAIL_CORRECTIONS = {
    "gamil.com": "gmail.com",
    "gmial.com": "gmail.com",
    "yaho.com": "yahoo.com",
    "hotmal.com": "hotmail.com"
}

def is_valid_name(line):
    """Filter out lines that are unlikely to be names."""
    invalid_keywords = ["mobile", "contact", "portfolio", "link", "email", "phone", "resume", "address"]
    return not any(keyword in line.lower() for keyword in invalid_keywords)

def correct_email(email):
    """Correct common email domain typos."""
    if not email:
        return email
    domain = email.split("@")[-1].lower()
    corrected_domain = EMAIL_CORRECTIONS.get(domain, domain)
    return email.replace(f"@{domain}", f"@{corrected_domain}")

def extract_passout_year(text):
    """Extract the passout year, preferring the most recent year in a range."""
    matches = re.finditer(PASSOUT_PATTERN, text, re.IGNORECASE)
    for match in matches:
        year_text = match.group(1)
        if "-" in year_text or "–" in year_text or "—" in year_text:
            # Take the end year from a range (e.g., 2018-2022 -> 2022)
            year = year_text.split()[-1]
        else:
            year = year_text
        if year.isdigit() and 1980 <= int(year) <= 2025:  # Valid year range
            return year
    return None

def extract_info(text):
    """Extract name, email, phone, and passout year from text."""
    name = None
    email = None
    phone = None
    passout_year = None

    # Name extraction
    lines = text.splitlines()
    for line in lines[:10]:  # Check first 10 lines
        line = line.strip()
        if line and re.match(NAME_PATTERN, line) and is_valid_name(line):
            name = line
            break

    # Email extraction
    email_match = re.search(EMAIL_PATTERN, text)
    if email_match:
        email = correct_email(email_match.group())

    # Phone extraction
    phone_match = re.search(PHONE_PATTERN, text)
    if phone_match:
        phone = phone_match.group()

    # Passout year extraction
    passout_year = extract_passout_year(text)

    return name, email, phone, passout_year

def process_resumes(folder_path):
    """Process all PDF resume files in the folder and extract info."""
    results = []
    folder = Path(folder_path)

    if not folder.exists():
        print(f"Folder {folder_path} does not exist.")
        return results

    for file_path in folder.glob("*.pdf"):
        try:
            with pdfplumber.open(file_path) as pdf:
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    text += page_text + "\n"
                name, email, phone, passout_year = extract_info(text)
                if name or email or phone or passout_year:  # Only add if at least one field is found
                    results.append({
                        "Filename": file_path.name,
                        "Name": name or "Not found",
                        "Email": email or "Not found",
                        "Phone": phone or "Not found",
                        "Passout Year": passout_year or "Not found"
                    })
        except Exception as e:
            print(f"Error processing {file_path.name}: {e}")

    return results

def save_to_excel(results, output_file="resume_info.xlsx"):
    """Save extracted info to an Excel file."""
    if not results:
        print("No data to save to Excel.")
        return
    
    # Convert results to DataFrame
    df = pd.DataFrame(results, columns=["Filename", "Name", "Email", "Phone", "Passout Year"])
    
    # Save to Excel
    df.to_excel(output_file, index=False, engine="openpyxl")
    
    # Adjust column widths
    from openpyxl import load_workbook
    wb = load_workbook(output_file)
    ws = wb.active
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)  # Cap width at 50
        ws.column_dimensions[column].width = adjusted_width
    wb.save(output_file)

def main():
    folder_path = r"D:\resume\UI UX"  # Specified folder path
    results = process_resumes(folder_path)
    
    if results:
        save_to_excel(results)
        print(f"Processed {len(results)} resumes. Results saved to resume_info.xlsx")
        for result in results:
            print(f"File: {result['Filename']}, Name: {result['Name']}, Email: {result['Email']}, Phone: {result['Phone']}, Passout Year: {result['Passout Year']}")
    else:
        print("No resumes processed or no valid data found.")

if __name__ == "__main__":
    main()