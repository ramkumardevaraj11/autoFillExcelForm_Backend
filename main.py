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

def extract_insurance_company(pdf_text):
    log("Step 10: Extracting INSURANCE COMPANY from PDF")
    lines = pdf_text.splitlines()
    stop_words = ("address", "city", "zip", "state", "sex", "dob", "date of birth", "patient", "subscriber", "group",
                  "plan", "policyholder", "phone", "tel", "member", "mrn", "claim")
    for i, line in enumerate(lines):
        if re.search(r"Primary\s*Ins\s*:", line, re.IGNORECASE):
            after_colon = line.split(":", 1)[-1].strip()
            company_lines = []
            if after_colon:
                company_lines.append(after_colon)
            for j in range(i+1, len(lines)):
                nextline = lines[j].strip()
                if (not nextline or
                    any(sw in nextline.lower() for sw in stop_words) or
                    re.match(r"^\d{5,}$", nextline)):
                    break
                company_lines.append(nextline)
            company = " ".join(company_lines)
            company = re.sub(r"\s{2,}", " ", company).strip(" ,.-")
            log(f"✓ Found Insurance Company: '{company}'")
            return company
    log("WARNING: Could not extract insurance company. Using placeholder.")
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

@app.post("/upload/")
async def upload_files(encounter_pdf: UploadFile = File(...), aptp_form: UploadFile = File(...)):
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    pdf_path = os.path.join(UPLOAD_DIR, encounter_pdf.filename)
    form_path = os.path.join(UPLOAD_DIR, aptp_form.filename)
    with open(pdf_path, "wb") as pdf_f:
        pdf_f.write(await encounter_pdf.read())
    with open(form_path, "wb") as form_f:
        form_f.write(await aptp_form.read())
    try:
        pdf_text = read_pdf_text(pdf_path)
        patient_name = extract_patient_name(pdf_text)
        policyholder_name = extract_policyholder_name(pdf_text)
        address = extract_address(pdf_text)
        contact_info = extract_telephone(pdf_text)
        birthdate_str = extract_birthdate(pdf_text)
        sex_value = extract_sex(pdf_text)
        insurance_company = extract_insurance_company(pdf_text)
        patient_info = PatientInfo(birthdate=birthdate_str, sex=sex_value)
        output_file_path = write_to_excel(
            form_path, patient_name, policyholder_name, address, contact_info, patient_info, sex_value, insurance_company
        )
        if not output_file_path:
            return {"error": "Failed to process form"}
        filename = os.path.basename(output_file_path)
        return {
            "message": "Files processed successfully! 🎉",
            "download_url": f"/download/{filename}",
            "extracted_data": {
                "patient": {
                    "last_name": patient_name.last_name,
                    "first_name": patient_name.first_name,
                    "initial": patient_name.initial,
                    "birthdate": patient_info.birthdate,
                    "sex": patient_info.sex
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
                "insurance_company": insurance_company,
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
    file_path = os.path.join(OUTPUT_DIR, filename)
    return FileResponse(
        path=file_path, 
        filename=filename, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
