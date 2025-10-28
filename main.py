from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os
from pypdf import PdfReader
import requests
import re
import openpyxl
from openpyxl.cell.cell import MergedCell
from datetime import datetime

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:4200"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
HUGGING_FACE_TOKEN = "YOUR_HUGGING_FACE_TOKEN_HERE"
HUGGING_FACE_API_URL = "https://api-inference.huggingface.co/models/Helios9/BioMed_NER"

class PersonName(BaseModel):
    last_name: str
    first_name: str
    initial: str

class Address(BaseModel):
    street: str
    city: str
    state: str
    zip_code: str

class ContactInfo(BaseModel):
    telephone: str

class PatientInfo(BaseModel):
    birthdate: str

def log(title, obj=None):
    """Professional logging for medical billing agent"""
    print(f"[MEDICAL-BILLING-AGENT] {title}")
    if obj is not None:
        try:
            preview = str(obj)
            if len(preview) > 400:
                print(preview[:400] + " ... [TRUNCATED]")
            else:
                print(preview)
        except Exception as e:
            print(f"(unable to print object: {e})")

def read_pdf_text(file_path):
    """Extract all text from encounter PDF"""
    log(f"Step 1: Checking file exists: {file_path}")
    if not os.path.exists(file_path):
        log("ERROR: PDF file not found!")
        return ""
    log("Step 2: Reading PDF file text...")
    reader = PdfReader(file_path)
    text = ""
    for i, page in enumerate(reader.pages):
        page_text = page.extract_text() or ""
        log(f"*** Page {i+1} text preview:", page_text[:500])
        text += page_text + "\n"
    log("Step 3: Finished reading PDF. Total length:", len(text))
    return text

def extract_name_by_label(pdf_text, label_keywords):
    """Generic function to extract a name following specific labels"""
    log(f"Searching for name with labels: {label_keywords}")
    
    for keyword in label_keywords:
        patterns = [
            rf"{keyword}\s*:\s*([A-Z\-]+)\s*,\s*([A-Z\-]+)(?:\s+([A-Z]))?",
            rf"{keyword}\s+Name\s*:\s*([A-Z\-]+)\s*,\s*([A-Z\-]+)(?:\s+([A-Z]))?",
            rf"\d+\.\s*{keyword}.*?Name.*?:\s*([A-Z\-]+)\s*,\s*([A-Z\-]+)(?:\s+([A-Z]))?",
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, pdf_text, re.MULTILINE | re.IGNORECASE)
            for match in matches:
                last = match.group(1).strip().title()
                first = match.group(2).strip().title()
                initial = match.group(3).strip() if match.group(3) else ""
                
                if len(last) >= 2 and len(first) >= 2:
                    log(f"✓ Found {keyword} name: Last='{last}', First='{first}', Initial='{initial}'")
                    return PersonName(
                        last_name=last,
                        first_name=first,
                        initial=initial
                    )
    return None

def extract_birthdate(pdf_text):
    """Extract patient birthdate from encounter PDF"""
    log("Step 8: Extracting PATIENT BIRTHDATE from PDF")
    
    # Date patterns - common medical document formats
    patterns = [
        # Pattern 1: "DOB: MM/DD/YYYY" or "Date of Birth: MM/DD/YYYY"
        r"(?:DOB|Date of Birth|Birth Date|Birthdate)\s*:?\s*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})",
        
        # Pattern 2: "MM/DD/YYYY" after "DOB" or "Birth"
        r"(?:DOB|Birth)\s*[:\s]*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})",
        
        # Pattern 3: "YYYY-MM-DD" format
        r"(?:DOB|Date of Birth|Birth Date)\s*:?\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})",
    ]
    
    for i, pattern in enumerate(patterns):
        match = re.search(pattern, pdf_text, re.MULTILINE | re.IGNORECASE)
        if match:
            if i == 2:  # Pattern 3 (YYYY-MM-DD)
                year = match.group(1)
                month = match.group(2).zfill(2)
                day = match.group(3).zfill(2)
                formatted = f"{month}/{day}/{year}"
            else:  # Patterns 1 & 2 (MM/DD/YYYY)
                month = match.group(1).zfill(2)
                day = match.group(2).zfill(2)
                year = match.group(3)
                formatted = f"{month}/{day}/{year}"
            
            log(f"✓ Pattern {i+1} matched birthdate!")
            log(f"  Birthdate: '{formatted}'")
            return PatientInfo(birthdate=formatted)
    
    log("WARNING: Could not extract birthdate. Using placeholder.")
    return PatientInfo(birthdate="")

def extract_telephone(pdf_text):
    """Extract telephone number from encounter PDF"""
    log("Step 7: Extracting TELEPHONE from PDF")
    
    # Telephone patterns - common US formats
    patterns = [
        # Pattern 1: "(123) 456-7890"
        r"(?:Phone|Telephone|Tel|Contact|Cell|Mobile)\s*:?\s*\(?(\d{3})\)?[\s\-\.]?(\d{3})[\s\-\.]?(\d{4})",
        
        # Pattern 2: "123-456-7890" or "123.456.7890"
        r"\b(\d{3})[\-\.](\d{3})[\-\.](\d{4})\b",
        
        # Pattern 3: "1234567890" (10 digits)
        r"\b(\d{10})\b",
    ]
    
    for i, pattern in enumerate(patterns):
        match = re.search(pattern, pdf_text, re.MULTILINE | re.IGNORECASE)
        if match:
            if i == 2:  # Pattern 3 (10 digits continuous)
                phone = match.group(1)
                formatted = f"({phone[:3]}) {phone[3:6]}-{phone[6:]}"
            else:
                formatted = f"({match.group(1)}) {match.group(2)}-{match.group(3)}"
            
            log(f"✓ Pattern {i+1} matched telephone!")
            log(f"  Telephone: '{formatted}'")
            return ContactInfo(telephone=formatted)
    
    log("WARNING: Could not extract telephone. Using placeholder.")
    return ContactInfo(telephone="")

def extract_address(pdf_text):
    """Extract patient address from encounter PDF"""
    log("Step 6: Extracting PATIENT ADDRESS from PDF")
    
    # Address patterns
    patterns = [
        # Pattern 1: "Address: 123 Main St, City, ST 12345"
        r"Address\s*:\s*([^,\n]+),\s*([A-Za-z\s]+),?\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)",
        
        # Pattern 2: Multi-line with labels
        r"(?:Street|Address)\s*:\s*([^\n]+).*?City\s*:\s*([A-Za-z\s]+).*?State\s*:\s*([A-Z]{2}).*?(?:Zip|ZIP|Zip Code)\s*:\s*(\d{5}(?:-\d{4})?)",
        
        # Pattern 3: Standard format
        r"([0-9]+\s+[A-Za-z\s]+(?:Street|St|Avenue|Ave|Road|Rd|Lane|Ln|Drive|Dr))\s*,?\s*([A-Za-z\s]+),?\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)",
    ]
    
    for i, pattern in enumerate(patterns):
        match = re.search(pattern, pdf_text, re.MULTILINE | re.IGNORECASE | re.DOTALL)
        if match:
            street = match.group(1).strip()
            city = match.group(2).strip()
            state = match.group(3).strip()
            zip_code = match.group(4).strip()
            
            log(f"✓ Pattern {i+1} matched address!")
            log(f"  Street: '{street}'")
            log(f"  City: '{city}'")
            log(f"  State: '{state}'")
            log(f"  ZIP: '{zip_code}'")
            
            return Address(
                street=street,
                city=city,
                state=state,
                zip_code=zip_code
            )
    
    log("WARNING: Could not extract address. Using placeholder.")
    return Address(street="", city="", state="", zip_code="")

def extract_patient_name(pdf_text):
    """Extract patient name from encounter PDF"""
    log("Step 4: Extracting PATIENT name from PDF")
    result = extract_name_by_label(pdf_text, ["Patient", "Pt"])
    if result:
        return result
    log("WARNING: Could not extract patient name. Using placeholder.")
    return PersonName(last_name="UNKNOWN", first_name="PATIENT", initial="")

def extract_policyholder_name(pdf_text):
    """Extract policyholder name from encounter PDF"""
    log("Step 5: Extracting POLICYHOLDER name from PDF")
    result = extract_name_by_label(pdf_text, ["Policyholder", "Policy Holder", "Subscriber", "Guarantor", "Insured"])
    if result:
        return result
    log("INFO: No separate policyholder found. Assuming same as patient.")
    return None

def find_name_section(sheet, section_keywords, section_number=None):
    """Find a specific name section in Excel"""
    log(f"Scanning for section with keywords: {section_keywords}, number: {section_number}")
    
    section_header_location = None
    
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            if cell.value:
                cell_text = str(cell.value).upper().strip()
                number_match = True
                if section_number:
                    number_match = section_number in cell_text or section_number.replace(".", "") in cell_text
                keywords_match = all(keyword.upper() in cell_text for keyword in section_keywords)
                
                if number_match and keywords_match:
                    section_header_location = (cell.row, cell.column)
                    log(f"✓ Found section header at row {cell.row}, col {cell.column}: '{cell.value}'")
                    break
        if section_header_location:
            break
    
    if not section_header_location:
        log(f"ERROR: Could not find section with keywords {section_keywords}")
        return {}
    
    base_row, base_col = section_header_location
    subheader_coords = {}
    
    search_row_start = base_row
    search_row_end = base_row + 3
    search_col_start = max(1, base_col - 1)
    search_col_end = min(sheet.max_column + 1, base_col + 10)
    
    for search_row in range(search_row_start, search_row_end):
        for search_col in range(search_col_start, search_col_end):
            cell = sheet.cell(row=search_row, column=search_col)
            if cell.value:
                cell_text = str(cell.value).strip().lower()
                
                if cell_text == "last" and "Last" not in subheader_coords:
                    subheader_coords["Last"] = (search_row, search_col)
                    log(f"  ✓ Found 'Last' at row {search_row}, col {search_col}")
                elif cell_text == "first" and "First" not in subheader_coords:
                    subheader_coords["First"] = (search_row, search_col)
                    log(f"  ✓ Found 'First' at row {search_row}, col {search_col}")
                elif cell_text == "initial" and "Initial" not in subheader_coords:
                    subheader_coords["Initial"] = (search_row, search_col)
                    log(f"  ✓ Found 'Initial' at row {search_row}, col {search_col}")
                
                if len(subheader_coords) == 3:
                    return subheader_coords
    
    return subheader_coords

def find_field_write_location(sheet, section_number, section_keywords):
    """
    Find the EXACT cell where data should be written for a labeled field.
    For APTP forms, the data cell is typically BELOW the label cell.
    
    Returns: (row, col) tuple or None
    """
    log(f"Searching for field #{section_number} with keywords: {section_keywords}")
    
    # Search for the label cell
    for row_idx in range(1, sheet.max_row + 1):
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            if cell.value:
                cell_text = str(cell.value).upper().strip()
                
                # Check if this is the field we're looking for
                has_number = section_number in cell_text or section_number.replace(".", "") in cell_text
                has_keywords = all(kw.upper() in cell_text for kw in section_keywords)
                
                if has_number and has_keywords:
                    log(f"✓ Found field label at row {row_idx}, col {col_idx}: '{cell.value}'")
                    
                    # Data entry is typically 1 row BELOW the label
                    data_row = row_idx + 1
                    data_col = col_idx
                    
                    # Make sure it's not another label or merged header
                    test_cell = sheet.cell(row=data_row, column=data_col)
                    if not isinstance(test_cell, MergedCell):
                        log(f"  → Data entry cell: row {data_row}, col {data_col}")
                        return (data_row, data_col)
                    
                    # If merged, try 2 rows below
                    data_row = row_idx + 2
                    test_cell = sheet.cell(row=data_row, column=data_col)
                    if not isinstance(test_cell, MergedCell):
                        log(f"  → Data entry cell (2 rows below): row {data_row}, col {data_col}")
                        return (data_row, data_col)
    
    log(f"ERROR: Could not find field #{section_number}")
    return None

def find_data_row(sheet, subheader_coords):
    """Find the first writable row below subheaders"""
    if not subheader_coords:
        return None
    
    sample_sub = list(subheader_coords.values())[0]
    subheader_row = sample_sub[0]
    
    for offset in range(1, 5):
        candidate_row = subheader_row + offset
        first_col = list(subheader_coords.values())[0][1]
        test_cell = sheet.cell(row=candidate_row, column=first_col)
        
        if not isinstance(test_cell, MergedCell):
            log(f"✓ Data row identified: row {candidate_row}")
            return candidate_row
    
    return subheader_row + 1

def write_to_cell_safe(sheet, row, col, value):
    """Safely write to a cell, handling merged cells"""
    cell = sheet.cell(row=row, column=col)
    
    if isinstance(cell, MergedCell):
        for merged_range in sheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                top_left = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                top_left.value = value
                log(f"  → Wrote '{value}' to merged cell top-left: row {merged_range.min_row}, col {merged_range.min_col}")
                return
    else:
        cell.value = value
        log(f"  → Wrote '{value}' to cell: row {row}, col {col}")

def fill_name_section(sheet, subheader_coords, person_name, section_label):
    """Fill a name section in the Excel form"""
    if not subheader_coords or len(subheader_coords) < 3:
        log(f"ERROR: Could not locate all {section_label} name fields!")
        return False
    
    data_row = find_data_row(sheet, subheader_coords)
    if not data_row:
        log(f"ERROR: Could not identify data row for {section_label}!")
        return False
    
    log(f"Writing {section_label} name data...")
    
    if "Last" in subheader_coords:
        _, col = subheader_coords["Last"]
        write_to_cell_safe(sheet, data_row, col, person_name.last_name)
    
    if "First" in subheader_coords:
        _, col = subheader_coords["First"]
        write_to_cell_safe(sheet, data_row, col, person_name.first_name)
    
    if "Initial" in subheader_coords:
        _, col = subheader_coords["Initial"]
        write_to_cell_safe(sheet, data_row, col, person_name.initial)
    
    return True

def fill_address_field(sheet, address):
    """Fill street address field (Field 2)"""
    log("=" * 60)
    log("FILLING FIELD 2: PATIENT'S ADDRESS (Street)")
    log("=" * 60)
    
    coords = find_field_write_location(sheet, "2", ["PATIENT", "ADDRESS"])
    
    if coords:
        row, col = coords
        write_to_cell_safe(sheet, row, col, address.street)
        log(f"✓ Successfully filled Street Address: '{address.street}'")
        return True
    else:
        log("ERROR: Could not find Field 2 (Patient's Address)")
        return False

def fill_city_field(sheet, address):
    """Fill city field (Field 3)"""
    log("=" * 60)
    log("FILLING FIELD 3: CITY")
    log("=" * 60)
    
    coords = find_field_write_location(sheet, "3", ["CITY"])
    
    if coords:
        row, col = coords
        write_to_cell_safe(sheet, row, col, address.city)
        log(f"✓ Successfully filled City: '{address.city}'")
        return True
    else:
        log("ERROR: Could not find Field 3 (City)")
        return False

def fill_zipcode_field(sheet, address):
    """Fill ZIP code field (Field 5)"""
    log("=" * 60)
    log("FILLING FIELD 5: ZIP CODE")
    log("=" * 60)
    
    coords = find_field_write_location(sheet, "5", ["ZIP"])
    
    if coords:
        row, col = coords
        write_to_cell_safe(sheet, row, col, address.zip_code)
        log(f"✓ Successfully filled ZIP Code: '{address.zip_code}'")
        return True
    else:
        log("ERROR: Could not find Field 5 (ZIP CODE)")
        return False

def fill_telephone_field(sheet, contact_info):
    """Fill telephone field (Field 6)"""
    log("=" * 60)
    log("FILLING FIELD 6: TELEPHONE")
    log("=" * 60)
    
    coords = find_field_write_location(sheet, "6", ["TELEPHONE"])
    
    if coords:
        row, col = coords
        write_to_cell_safe(sheet, row, col, contact_info.telephone)
        log(f"✓ Successfully filled Telephone: '{contact_info.telephone}'")
        return True
    else:
        log("ERROR: Could not find Field 6 (TELEPHONE)")
        return False

def fill_birthdate_field(sheet, patient_info):
    """Fill patient birthdate field (Field 7)"""
    log("=" * 60)
    log("FILLING FIELD 7: PATIENT BIRTHDATE")
    log("=" * 60)
    
    coords = find_field_write_location(sheet, "7", ["PATIENT", "BIRTHDATE"])
    
    if coords:
        row, col = coords
        write_to_cell_safe(sheet, row, col, patient_info.birthdate)
        log(f"✓ Successfully filled Birthdate: '{patient_info.birthdate}'")
        return True
    else:
        log("ERROR: Could not find Field 7 (PATIENT BIRTHDATE)")
        return False

def write_to_excel(form_path, patient_name, policyholder_name, address, contact_info, patient_info):
    """Fill all fields: patient name, address, city, ZIP, telephone, birthdate, policyholder name"""
    log("Step 9: Confirm Excel file exists", form_path)
    if not os.path.exists(form_path):
        log("ERROR: Excel APTP form not found!")
        return ""
    
    log(f"Step 10: Loading Excel form: {form_path}")
    wb = openpyxl.load_workbook(form_path)
    sheet = wb.active
    
    # Fill Patient's Name (Section 1)
    log("=" * 60)
    log("FILLING SECTION 1: PATIENT'S NAME")
    log("=" * 60)
    patient_coords = find_name_section(sheet, ["PATIENT", "NAME"], section_number="1")
    fill_name_section(sheet, patient_coords, patient_name, "PATIENT")
    
    # Fill Address Field 2
    fill_address_field(sheet, address)
    
    # Fill City Field 3
    fill_city_field(sheet, address)
    
    # Fill ZIP Code Field 5
    fill_zipcode_field(sheet, address)
    
    # Fill Telephone Field 6
    fill_telephone_field(sheet, contact_info)
    
    # Fill Birthdate Field 7
    fill_birthdate_field(sheet, patient_info)
    
    # Fill Policyholder's Name (Section 14)
    log("=" * 60)
    log("FILLING SECTION 14: POLICYHOLDER'S NAME")
    log("=" * 60)
    policyholder_coords = find_name_section(sheet, ["POLICYHOLDER", "NAME"], section_number="14")
    
    if policyholder_name is None:
        log("Using patient name for policyholder (same person)")
        policyholder_name = patient_name
    
    fill_name_section(sheet, policyholder_coords, policyholder_name, "POLICYHOLDER")
    
    # Save
    output_file = os.path.join(OUTPUT_DIR, "APTP_Completed_Form.xlsx")
    try:
        wb.save(output_file)
        log("=" * 60)
        log("✓ Successfully saved completed Excel form!")
        log("=" * 60)
    except Exception as e:
        log(f"ERROR: Saving Excel file failed: {str(e)}")
        return ""
    
    return output_file

@app.get("/")
def read_root():
    log("Root endpoint called")
    return {"message": "Medical Form Filler Backend is running! 🏥"}

@app.post("/upload/")
async def upload_files(encounter_pdf: UploadFile = File(...), aptp_form: UploadFile = File(...)):
    log("=" * 80)
    log("MEDICAL FORM AUTOMATION - STARTED")
    log("=" * 80)
    
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    pdf_path = os.path.join(UPLOAD_DIR, encounter_pdf.filename)
    form_path = os.path.join(UPLOAD_DIR, aptp_form.filename)
    
    log(f"Received Encounter PDF: {encounter_pdf.filename}")
    log(f"Received APTP Form: {aptp_form.filename}")

    with open(pdf_path, "wb") as pdf_f:
        pdf_f.write(await encounter_pdf.read())
    with open(form_path, "wb") as form_f:
        form_f.write(await aptp_form.read())
    
    log("✓ Files saved to disk")
    
    try:
        pdf_text = read_pdf_text(pdf_path)
        
        # Extract all data
        patient_name = extract_patient_name(pdf_text)
        policyholder_name = extract_policyholder_name(pdf_text)
        address = extract_address(pdf_text)
        contact_info = extract_telephone(pdf_text)
        patient_info = extract_birthdate(pdf_text)
        
        # Fill Excel form
        output_file_path = write_to_excel(form_path, patient_name, policyholder_name, address, contact_info, patient_info)
        
        if not output_file_path:
            return {"error": "Failed to process form"}
        
        filename = os.path.basename(output_file_path)
        log("=" * 80)
        log(f"✓ SUCCESS! Download ready: {filename}")
        log("=" * 80)
        
        return {
            "message": "Files processed successfully! 🎉",
            "download_url": f"/download/{filename}",
            "extracted_data": {
                "patient": {
                    "last_name": patient_name.last_name,
                    "first_name": patient_name.first_name,
                    "initial": patient_name.initial,
                    "birthdate": patient_info.birthdate
                },
                "address": {
                    "street": address.street,
                    "city": address.city,
                    "state": address.state,
                    "zip_code": address.zip_code
                },
                "contact": {
                    "telephone": contact_info.telephone
                },
                "policyholder": {
                    "last_name": policyholder_name.last_name if policyholder_name else patient_name.last_name,
                    "first_name": policyholder_name.first_name if policyholder_name else patient_name.first_name,
                    "initial": policyholder_name.initial if policyholder_name else patient_name.initial
                }
            }
        }
    
    except Exception as e:
        log(f"ERROR: Processing failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return {"error": f"Processing failed: {str(e)}"}

@app.get("/download/{filename}")
def download_file(filename: str):
    log(f"Download requested: {filename}")
    file_path = os.path.join(OUTPUT_DIR, filename)
    return FileResponse(
        path=file_path, 
        filename=filename, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
