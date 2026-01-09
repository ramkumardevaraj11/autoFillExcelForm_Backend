from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os
from pypdf import PdfReader
import re
import openpyxl
from openpyxl.cell.cell import MergedCell
from datetime import datetime
from fastapi import HTTPException
import pdfplumber


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
    sex: str

def log(title, obj=None):
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

def extract_insurance_company(pdf_path):
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        # Grab all potential text lines from page 1
        lines = [line.strip() for line in (page.extract_text() or "").splitlines() if line.strip()]
        candidates = []
        for line in lines:
            ll = line.lower()
            # Must have 'insurance' and be more than one word
            if ("insurance" in ll
                and len(line.split()) > 1
                and not ll.startswith("new york")
                and not ll.startswith("verification")
                and not ll.startswith("law")
                and not ll.startswith("page")
                and not ll.isupper()  # skip ALL CAPS narratives
            ):
                candidates.append(line)  # collect all plausible company lines
        
        # Extra robustness: pick the one where all words are alpha and not just 'LAW'
        companies = [
            cand for cand in candidates
            if all(word.isalpha() or "." in word for word in cand.split())
            and cand.count(" ") < 5  # not too long, helps isolate "CURE AUTO INSURANCE"
        ]
        if companies:
            company = companies[0].strip()
            log(f"✓ Insurance company selected: '{company}'")
            return company
        if candidates:
            company = candidates[0].strip()
            log(f"✓ Insurance company fallback: '{company}'")
            return company

    log("WARNING: Insurance company not found on page 1; returning blank.")
    return ""

def extract_name_by_label(pdf_text, label_keywords):
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

def extract_sex(pdf_text):
    log("Step 9: Extracting SEX from PDF")
    patterns = [
        r"(?:Sex|Gender)\s*:?[\s>]*(M|F|Male|Female|Other|Unspecified|X)\b",
        r"\b(Male|Female|Other|M|F|X)\b.*?(?:Sex|Gender)",
    ]
    for i, pattern in enumerate(patterns):
        match = re.search(pattern, pdf_text, re.IGNORECASE)
        if match:
            sex_raw = match.group(1).strip().lower()
            if sex_raw in ['male', 'm']:
                norm = "M"
            elif sex_raw in ['female', 'f']:
                norm = "F"
            elif sex_raw in ['other', 'x']:
                norm = "Other"
            else:
                norm = sex_raw.capitalize()
            log(f"✓ Pattern {i+1} matched sex/gender: '{sex_raw}' (Normalized: '{norm}')")
            return norm
    log("WARNING: Could not extract sex. Using placeholder.")
    return ""

def extract_birthdate(pdf_text):
    log("Step 8: Extracting PATIENT BIRTHDATE from PDF")
    patterns = [
        r"(?:DOB|Date of Birth|Birth Date|Birthdate)\s*:?\s*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})",
        r"(?:DOB|Birth)\s*[:\s]*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})",
        r"(?:DOB|Date of Birth|Birth Date)\s*:?\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})",
    ]
    for i, pattern in enumerate(patterns):
        match = re.search(pattern, pdf_text, re.MULTILINE | re.IGNORECASE)
        if match:
            if i == 2:
                year = match.group(1)
                month = match.group(2).zfill(2)
                day = match.group(3).zfill(2)
                formatted = f"{month}/{day}/{year}"
            else:
                month = match.group(1).zfill(2)
                day = match.group(2).zfill(2)
                year = match.group(3)
                formatted = f"{month}/{day}/{year}"
            log(f"✓ Pattern {i+1} matched birthdate!")
            log(f"  Birthdate: '{formatted}'")
            return formatted
    log("WARNING: Could not extract birthdate. Using placeholder.")
    return ""

def extract_telephone(pdf_text):
    log("Step 7: Extracting TELEPHONE from PDF")
    patterns = [
        r"(?:Phone|Telephone|Tel|Contact|Cell|Mobile)\s*:?\s*\(?(\d{3})\)?[\s\-\.]?(\d{3})[\s\-\.]?(\d{4})",
        r"\b(\d{3})[\-\.](\d{3})[\-\.](\d{4})\b",
        r"\b(\d{10})\b",
    ]
    for i, pattern in enumerate(patterns):
        match = re.search(pattern, pdf_text, re.MULTILINE | re.IGNORECASE)
        if match:
            if i == 2:
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
    log("Step 6: Extracting PATIENT ADDRESS from PDF")
    patterns = [
        r"Address\s*:\s*([^,\n]+),\s*([A-Za-z\s]+),?\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)",
        r"(?:Street|Address)\s*:\s*([^\n]+).*?City\s*:\s*([A-Za-z\s]+).*?State\s*:\s*([A-Z]{2}).*?(?:Zip|ZIP|Zip Code)\s*:\s*(\d{5}(?:-\d{4})?)",
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
    log("Step 4: Extracting PATIENT name from PDF")
    result = extract_name_by_label(pdf_text, ["Patient", "Pt"])
    if result:
        return result
    log("WARNING: Could not extract patient name. Using placeholder.")
    return PersonName(last_name="UNKNOWN", first_name="PATIENT", initial="")

def extract_policyholder_name(pdf_text):
    log("Step 5: Extracting POLICYHOLDER name from PDF")
    result = extract_name_by_label(pdf_text, ["Policyholder", "Policy Holder", "Subscriber", "Guarantor", "Insured"])
    if result:
        return result
    log("INFO: No separate policyholder found. Assuming same as patient.")
    return None

def find_name_section(sheet, section_keywords, section_number=None):
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
    log(f"Searching for field #{section_number} with keywords: {section_keywords}")
    for row_idx in range(1, sheet.max_row + 1):
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            if cell.value:
                cell_text = str(cell.value).upper().strip()
                has_number = section_number in cell_text or section_number.replace(".", "") in cell_text
                has_keywords = all(kw.upper() in cell_text for kw in section_keywords)
                if has_number and has_keywords:
                    log(f"✓ Found field label at row {row_idx}, col {col_idx}: '{cell.value}'")
                    data_row = row_idx + 1
                    data_col = col_idx
                    test_cell = sheet.cell(row=data_row, column=data_col)
                    if not isinstance(test_cell, MergedCell):
                        log(f"  → Data entry cell: row {data_row}, col {data_col}")
                        return (data_row, data_col)
                    data_row = row_idx + 2
                    test_cell = sheet.cell(row=data_row, column=data_col)
                    if not isinstance(test_cell, MergedCell):
                        log(f"  → Data entry cell (2 rows below): row {data_row}, col {data_col}")
                        return (data_row, data_col)
    log(f"ERROR: Could not find field #{section_number}")
    return None

def write_to_cell_safe(sheet, row, col, value):
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

def fill_sex_field(sheet, sex_value):
    """
    Fill the '8. SEX' fields with a tick (✓) in the correct cell for M/F.
    Assumes the Excel row has: '8. SEX' label, and two adjacent cells for M/F (e.g., next two columns).
    Leaves the other cell empty.
    """
    log("=" * 60)
    log("FILLING FIELD 8: SEX (with tick)")
    log("=" * 60)
    coords = find_field_write_location(sheet, "8", ["SEX"])
    if coords:
        row, col = coords  # cell under "8. SEX" label
        tick = "✓"
        m_col, f_col = col + 1, col + 2
        m_cell = sheet.cell(row=row, column=m_col)
        f_cell = sheet.cell(row=row, column=f_col)
        # Clear both cells first
        m_cell.value = ""
        f_cell.value = ""
        if sex_value.upper() == "M":
            m_cell.value = tick
        elif sex_value.upper() == "F":
            f_cell.value = tick
        log(f"✓ Successfully filled Sex: '{sex_value}' with tick at {'M' if sex_value.upper() == 'M' else 'F' if sex_value.upper() == 'F' else 'NONE'}")
        return True
    else:
        log("ERROR: Could not find Field 8 (SEX)")
        return False

def fill_field(sheet, value, section_number, section_keywords, log_name):
    log("=" * 60)
    log(f"FILLING FIELD {section_number}: {log_name}")
    log("=" * 60)
    coords = find_field_write_location(sheet, section_number, section_keywords)
    if coords:
        row, col = coords
        write_to_cell_safe(sheet, row, col, value)
        log(f"✓ Successfully filled {log_name}: '{value}'")
        return True
    else:
        log(f"ERROR: Could not find Field {section_number} ({log_name})")
        return False

def fill_name_section(sheet, subheader_coords, person_name, section_label):
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

def find_data_row(sheet, subheader_coords):
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

def write_to_excel(form_path, patient_name, policyholder_name, address, contact_info, patient_info, sex_value, insurance_company):
    log("Step 10: Loading Excel form: {form_path}")
    wb = openpyxl.load_workbook(form_path)
    sheet = wb.active

    patient_coords = find_name_section(sheet, ["PATIENT", "NAME"], section_number="1")
    fill_name_section(sheet, patient_coords, patient_name, "PATIENT")
    fill_field(sheet, address.street, "2", ["PATIENT", "ADDRESS"], "PATIENT'S ADDRESS (Street)")
    fill_field(sheet, address.city, "3", ["CITY"], "CITY")
    fill_field(sheet, address.zip_code, "5", ["ZIP"], "ZIP CODE")
    fill_field(sheet, contact_info.telephone, "6", ["TELEPHONE"], "TELEPHONE")
    fill_field(sheet, patient_info.birthdate, "7", ["PATIENT", "BIRTHDATE"], "PATIENT BIRTHDATE")
    fill_sex_field(sheet, sex_value)
    fill_field(sheet, insurance_company, "9", ["INSURANCE", "COMPANY"], "INSURANCE COMPANY")

    policyholder_coords = find_name_section(sheet, ["POLICYHOLDER", "NAME"], section_number="14")
    if policyholder_name is None:
        policyholder_name = patient_name
    fill_name_section(sheet, policyholder_coords, policyholder_name, "POLICYHOLDER")

    output_file = os.path.join(OUTPUT_DIR, "APTP_Completed_Form.xlsx")
    try:
        wb.save(output_file)
        log("✓ Successfully saved completed Excel form!")
    except Exception as e:
        log(f"ERROR: Saving Excel file failed: {str(e)}")
        return ""
    return output_file

def extract_psa_pdf_values(pdf_path, pdf_text):
    today_date = datetime.today().strftime('%m/%d/%Y')
    insurance_company = extract_insurance_company(pdf_path)
    date_of_loss = extract_date_of_accident(pdf_text)
    return today_date, insurance_company, date_of_loss

def extract_claim_number(pdf_text):
    log("Step: Extracting CLAIM NUMBER (enhanced multi-line lookahead)")
    lines = [line.strip() for line in pdf_text.splitlines() if line.strip()]
    header_keywords = ["CLAIM NUMBER", "CLAIM #", "POLICY NUMBER"]
    for i, line in enumerate(lines):
        for key in header_keywords:
            if key.replace(" ", "") in line.upper().replace(" ", ""):
                # Inline (on same line)
                match_inline = re.search(rf"{key}[:\s\-]*([A-Z0-9\-]{{6,}})", line.upper())
                if match_inline:
                    claim = match_inline.group(1)
                    log(f"✓ Claim number found inline: '{claim}'")
                    return claim
                # Look for best candidate next 10 lines (not zip, not address, 6+ chars)
                for j in range(1, 11):
                    if i+j < len(lines):
                        candidate = lines[i+j].strip()
                        if (re.fullmatch(r"[A-Z0-9\-]{6,}", candidate) and 
                            not re.fullmatch(r"\d{5}", candidate) and 
                            not any(w in candidate.lower() for w in ("street", "avenue", "drive", "suite")) and
                            not re.search(r"[a-z]", candidate)):
                            log(f"✓ Claim candidates found: '{candidate}'")
                            return candidate
    log("WARNING: Could not extract claim number; using placeholder.")
    return ""

def extract_date_of_accident(pdf_text):
    log("Step: Extracting DATE OF ACCIDENT from PDF")
    lines = [line.strip() for line in pdf_text.splitlines() if line.strip()]
    
    # Pattern 1: Inline with label "DATE OF ACCIDENT: 16 May 2024"
    pattern_inline = r"DATE\s+OF\s+ACCIDENT[:\s]*(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})"
    match = re.search(pattern_inline, pdf_text, re.IGNORECASE)
    if match:
        day = match.group(1).zfill(2)
        month_name = match.group(2)
        year = match.group(3)
        month_map = {
            "jan": "01", "feb": "02", "mar": "03", "apr": "04",
            "may": "05", "jun": "06", "jul": "07", "aug": "08",
            "sep": "09", "oct": "10", "nov": "11", "dec": "12"
        }
        month = month_map.get(month_name.lower()[:3], "01")
        formatted = f"{month}/{day}/{year}"
        log(f"✓ DATE OF ACCIDENT found (text month): '{formatted}'")
        return formatted
    
    # Pattern 2: Numeric formats "05/16/2024" or "2024-05-16"
    pattern_numeric = r"DATE\s+OF\s+ACCIDENT[:\s]*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})"
    match = re.search(pattern_numeric, pdf_text, re.IGNORECASE)
    if match:
        month = match.group(1).zfill(2)
        day = match.group(2).zfill(2)
        year = match.group(3)
        formatted = f"{month}/{day}/{year}"
        log(f"✓ DATE OF ACCIDENT found (numeric): '{formatted}'")
        return formatted
    
    # Pattern 3: Lookahead in next lines after "DATE OF ACCIDENT" label
    for i, line in enumerate(lines):
        if "DATE OF ACCIDENT" in line.upper():
            for j in range(1, 5):
                if i+j < len(lines):
                    candidate = lines[i+j].strip()
                    # Try text month format
                    match_text = re.match(r"(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})", candidate, re.IGNORECASE)
                    if match_text:
                        day = match_text.group(1).zfill(2)
                        month_name = match_text.group(2)
                        year = match_text.group(3)
                        month_map = {
                            "jan": "01", "feb": "02", "mar": "03", "apr": "04",
                            "may": "05", "jun": "06", "jul": "07", "aug": "08",
                            "sep": "09", "oct": "10", "nov": "11", "dec": "12"
                        }
                        month = month_map.get(month_name.lower()[:3], "01")
                        formatted = f"{month}/{day}/{year}"
                        log(f"✓ DATE OF ACCIDENT found (lookahead text): '{formatted}'")
                        return formatted
                    # Try numeric format
                    match_num = re.match(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})", candidate)
                    if match_num:
                        month = match_num.group(1).zfill(2)
                        day = match_num.group(2).zfill(2)
                        year = match_num.group(3)
                        formatted = f"{month}/{day}/{year}"
                        log(f"✓ DATE OF ACCIDENT found (lookahead numeric): '{formatted}'")
                        return formatted
    
    log("WARNING: Could not extract DATE OF ACCIDENT; using placeholder.")
    return ""

def extract_psa_patient_name(pdf_text):
    """
    Extracts PSA patient name from Bill PDF.
    Handles formats:
    - "WALKING SAINT JEAN" (3 words: FIRST MIDDLE LAST)
    - "PEREZ, EDUAR" (LAST, FIRST)
    - "Patient Name: LASTNAME, FIRSTNAME M"
    Returns: (last, first, middle_initial)
    """
    log("Step: Extracting PSA Patient Name from PDF")
    
    # Pattern 1: Look for the structured patient ID line "WALKING  SAINT JEAN-10084-KC43773"
    # This is a strong anchor because it appears before the claim reference
    pattern_id = r"([A-Z]+)\s+([A-Z]+)\s+([A-Z]+)-\d+-[A-Z0-9]+"
    match = re.search(pattern_id, pdf_text)
    if match:
        first = match.group(1).title()
        middle = match.group(2).title()
        last = match.group(3).title()
        log(f"✓ PSA Patient Name (ID format): First='{first}', Middle='{middle}', Last='{last}'")
        return (last, first, middle[0] if middle else "")
    
    # Pattern 2: Standard comma-separated "LASTNAME, FIRSTNAME M" after "PATIENT" or "Patient Name"
    pattern_comma = r"(?:Patient(?:'s)?\s+Name(?:\s+and\s+Address)?|PATIENT)[:\s]*([A-Z][A-Z]+)\s*,\s*([A-Z][A-Z]+)(?:\s+([A-Z]))?"
    match = re.search(pattern_comma, pdf_text, re.IGNORECASE)
    if match:
        last = match.group(1).title()
        first = match.group(2).title()
        middle = match.group(3)[0] if match.group(3) else ""
        if len(last) >= 2 and len(first) >= 2:
            log(f"✓ PSA Patient Name (comma format): Last='{last}', First='{first}', Middle='{middle}'")
            return (last, first, middle)
    
    # Pattern 3: Three consecutive ALL-CAPS words (avoiding form template words)
    # Exclude common form words
    excluded = {"PATIENT", "FIRST", "CONSULT", "CONDITION", "ACCIDENT", "YES", "NO", "DATE", "NAME", "ADDRESS"}
    words = re.findall(r"\b[A-Z]{2,}\b", pdf_text)
    # Find first sequence of 3 words not in excluded set
    for i in range(len(words) - 2):
        if words[i] not in excluded and words[i+1] not in excluded and words[i+2] not in excluded:
            first = words[i].title()
            middle = words[i+1].title()
            last = words[i+2].title()
            log(f"✓ PSA Patient Name (3-word sequence): First='{first}', Middle='{middle}', Last='{last}'")
            return (last, first, middle[0])
    
    log("WARNING: PSA Patient Name not found. Using placeholder.")
    return ("UNKNOWN", "PATIENT", "")

def extract_psa_birthdate(pdf_text):
    log("Step: Extracting PSA PATIENT DATE OF BIRTH from PDF")
    # Pattern 1: MM/DD/YYYY or MM-DD-YYYY
    patterns = [
        r"(?:DOB|Date of Birth|Birth Date|Birthdate)[:\s]*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})",
        r"(?:DOB|Date of Birth|Birth Date|Birthdate)[:\s]*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})",
        # Pattern for "24 Mar 1989"
        r"(?:DOB|Date of Birth|Birth Date|Birthdate)[:\s]*(\d{1,2})\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[ ,\-]+(\d{4})",
        # Loose catch-all: dd Mon yyyy, not attached to label
        r"\b(\d{1,2})\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[ ,\-]+(\d{4})\b"
    ]
    month_map = {
        "jan": "01", "feb": "02", "mar": "03", "apr": "04",
        "may": "05", "jun": "06", "jul": "07", "aug": "08",
        "sep": "09", "oct": "10", "nov": "11", "dec": "12"
    }
    for idx, pat in enumerate(patterns):
        match = re.search(pat, pdf_text, re.IGNORECASE)
        if match:
            if idx == 0:  # MM/DD/YYYY
                month = match.group(1).zfill(2)
                day = match.group(2).zfill(2)
                year = match.group(3)
            elif idx == 1:  # YYYY-MM-DD
                year = match.group(1)
                month = match.group(2).zfill(2)
                day = match.group(3).zfill(2)
            else:  # Text month
                day = match.group(1).zfill(2)
                month_str = match.group(2).lower()[:3]
                month = month_map.get(month_str, "01")
                year = match.group(3)
            formatted = f"{month}/{day}/{year}"
            log(f"✓ PSA DOB extracted: {formatted}")
            return formatted
    log("WARNING: PSA DOB not found.")
    return ""

def fill_psa_excel(psa_xlsx_path, date_appeal_submitted, insurance_company, claim_number, date_of_loss, psa_last, psa_first, psa_middle, psa_dob):
    wb = openpyxl.load_workbook(psa_xlsx_path)
    sheet = wb.active

    def find_and_fill(header, value):
        for row in sheet.iter_rows():
            for cell in row:
                if (
                    cell.value and
                    isinstance(cell.value, str) and
                    header in cell.value.upper()
                ):
                    below_row = cell.row + 1
                    below_col = cell.column
                    right_row = cell.row
                    right_col = cell.column + 1
                    
                    # Get the actual cell objects
                    below = sheet.cell(row=below_row, column=below_col)
                    right = sheet.cell(row=right_row, column=right_col)
                    
                    # Check if BELOW is truly empty (not merged, no value)
                    if not isinstance(below, MergedCell) and below.value in (None, ""):
                        write_to_cell_safe(sheet, below_row, below_col, value)
                        log(f"Filled '{header}' BELOW at Row {below_row}, Col {below_col}: {value}")
                        return True
                    
                    # Check if RIGHT is truly empty (not merged, no value)
                    if not isinstance(right, MergedCell) and right.value in (None, ""):
                        write_to_cell_safe(sheet, right_row, right_col, value)
                        log(f"Filled '{header}' RIGHT at Row {right_row}, Col {right_col}: {value}")
                        return True
                    
                    # If both below and right are occupied, try 2 rows below header
                    below2_row = cell.row + 2
                    below2 = sheet.cell(row=below2_row, column=below_col)
                    if not isinstance(below2, MergedCell) and below2.value in (None, ""):
                        write_to_cell_safe(sheet, below2_row, below_col, value)
                        log(f"Filled '{header}' 2 ROWS BELOW at Row {below2_row}, Col {below_col}: {value}")
                        return True
                    
                    log(f"WARNING: All target cells for '{header}' are occupied. Skipping fill.")
                    return False
        
        log(f"Could not find field to fill: {header}")
        return False


    find_and_fill("DATE APPEAL SUBMITTED", date_appeal_submitted)
    find_and_fill("INSURANCE COMPANY", insurance_company)
    find_and_fill("CLAIM NUMBER", claim_number)
    find_and_fill("CLAIM #", claim_number)
    find_and_fill("DATE OF LOSS", date_of_loss)

    # -- PSA Patient Name fields fill-in --
    find_and_fill("LAST NAME", psa_last)
    find_and_fill("FIRST NAME", psa_first)
    find_and_fill("MIDDLE INITIAL", psa_middle)
    find_and_fill("DATE OF BIRTH", psa_dob)  # NEW

    wb.save(psa_xlsx_path)

@app.post("/upload/")
async def upload_files(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
):
    try:
        if not (pdf_file and excel_file):
            raise HTTPException(
                status_code=400,
                detail="Both pdf_file and excel_file are required."
            )

        os.makedirs(UPLOAD_DIR, exist_ok=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        pdf_filename = pdf_file.filename or ""
        excel_filename = excel_file.filename or ""
        pdf_path = os.path.join(UPLOAD_DIR, pdf_filename)
        excel_path = os.path.join(UPLOAD_DIR, excel_filename)
        with open(pdf_path, "wb") as f:
            f.write(await pdf_file.read())
        with open(excel_path, "wb") as f:
            f.write(await excel_file.read())
        pdf_name = pdf_filename.lower()
        excel_name = excel_filename.lower()

        # EXTENSION CHECK
        if not (pdf_name.endswith(".pdf") and (excel_name.endswith(".xls") or excel_name.endswith(".xlsx"))):
            raise HTTPException(
                status_code=400,
                detail="Files must be .pdf and .xls or .xlsx extensions."
            )
        # ----------- APTP/ENCOUNTER logic ----------
        if "encounter" in pdf_name and "aptp" in excel_name:
            try:
                pdf_text = read_pdf_text(pdf_path)
                patient_name = extract_patient_name(pdf_text)
                policyholder_name = extract_policyholder_name(pdf_text)
                address = extract_address(pdf_text)
                contact_info = extract_telephone(pdf_text)
                birthdate_str = extract_birthdate(pdf_text)
                sex_value = extract_sex(pdf_text)
                insurance_company = extract_insurance_company(pdf_path)
                if not isinstance(insurance_company, str):
                    insurance_company = ""
                patient_info = PatientInfo(birthdate=birthdate_str, sex=sex_value)
                output_file_path = write_to_excel(
                    excel_path, patient_name, policyholder_name, address, contact_info, patient_info, sex_value, insurance_company
                )
                if not output_file_path:
                    raise HTTPException(status_code=500, detail="Failed to process form")
                # Save as original Excel filename in output for auto-download
                output_final_path = os.path.join(OUTPUT_DIR, excel_filename)
                os.replace(output_file_path, output_final_path)
                return {
                    "message": "APTP Form processed successfully.",
                    "download_url": f"/download/{excel_filename}"
                }
            except Exception as e:
                log(f"APTP processing failed: {str(e)}")
                import traceback
                traceback.print_exc()
                raise HTTPException(status_code=500, detail=f"APTP Processing failed: {str(e)}")

        # ----------- PSA/BILL logic ----------------
        if pdf_name.startswith("bill") and "psa" in excel_name:
            log("Received Bill PDF and PSA Excel files. Extracting/auto-filling fields...")
            pdf_text = read_pdf_text(pdf_path)
            date_appeal, ins_company, date_of_loss = extract_psa_pdf_values(pdf_path, pdf_text)
            claim_number = extract_claim_number(pdf_text)
            psa_last, psa_first, psa_middle = extract_psa_patient_name(pdf_text)
            psa_dob = extract_psa_birthdate(pdf_text)  # NEW - reuse existing APTP function
            
            log(f"Extracted PSA PATIENT: LAST={psa_last}, FIRST={psa_first}, MIDDLE={psa_middle}")
            log(f"Extracted PSA DATE OF BIRTH: {psa_dob}")

            fill_psa_excel(
                excel_path, date_appeal, ins_company, claim_number, date_of_loss,
                psa_last, psa_first, psa_middle, psa_dob  # NEW parameter
            )
            output_psa_path = os.path.join(OUTPUT_DIR, excel_filename)
            os.replace(excel_path, output_psa_path)
            return {
                "message": "Bill PDF & PSA Excel processed successfully.",
                "download_url": f"/download/{excel_filename}"
            }

        raise HTTPException(
            status_code=400,
            detail="Invalid file combination. Only (Encounter PDF + APTP Excel) and (Bill PDF + PSA Excel) allowed."
        )
    except Exception as ex:
        log(f"SERVER ERROR: {str(ex)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Server error: {str(ex)}")

@app.get("/download/{filename}")
def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.isfile(file_path):
        raise HTTPException(status_code=404, detail="File not found.")
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )