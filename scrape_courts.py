#!/usr/bin/env python3
"""
Sindh High Court Cause List Parser
Version 3.0 (No EXE Required)

Requirements:
    pip install openpyxl PyMuPDF
"""

import re
from pathlib import Path
from datetime import datetime
import openpyxl
import fitz  # PyMuPDF library (Replaces pdftotext.exe)

# -------------------------------------------------------
# CONFIG
# -------------------------------------------------------
OUTPUT_NAME = "Sindh_Cause_List.xlsx"

# -------------------------------------------------------
# REGEX
# -------------------------------------------------------
CASE_START = re.compile(
    r'^(\d+)\.\s+(.+)$'
)
BENCH = re.compile(
    r'^(MR|MRS)\.\s+JUSTICE',
    re.I
)
SECTION = re.compile(
    r'^FOR\s+',
    re.I
)
DATE = re.compile(
    r'([A-Za-z]+)\s+(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})'
)
HEADERS = [
    "HIGH COURT OF SINDH",
    "Designed & Developed",
    "Printed at",
    "Page #",
    "Daily List"
]

# -------------------------------------------------------
# PDF -> TEXT (NO EXE REQUIRED)
# -------------------------------------------------------
def extract_text(pdf_file: Path) -> str:
    """
    Convert PDF to plain text using PyMuPDF (fitz).
    This replaces Poppler's pdftotext.exe entirely.
    """
    full_text = ""
    try:
        with fitz.open(pdf_file) as doc:
            for page in doc:
                # get_text("text") maintains a good layout structure
                full_text += page.get_text("text") + "\f"
    except Exception as e:
        print(f"Error reading {pdf_file.name}: {e}")
    
    return full_text


# -------------------------------------------------------
# DATE
# -------------------------------------------------------
def extract_date(text: str):
    m = DATE.search(text)
    if not m:
        return "", "", ""
    weekday = m.group(1)
    day = m.group(2)
    month = m.group(3)
    year = m.group(4)
    return day, month, year


# -------------------------------------------------------
# CLEANING
# -------------------------------------------------------
def should_skip(line: str):
    s = line.strip()
    if not s:
        return True
    # Remove common headers
    for h in HEADERS:
        if h.lower() in s.lower():
            return True
    # Remove page numbers
    if re.search(r"Page\s*#?\s*\d+", s, re.I):
        return True
    # Remove weekday/date line
    if DATE.search(s):
        return True
    return False


def clean_page(page_text: str):
    cleaned = []
    for line in page_text.splitlines():
        if should_skip(line):
            continue
        # normalize whitespace
        line = re.sub(r"\s+$", "", line)
        cleaned.append(line)
    return cleaned


# -------------------------------------------------------
# SPLIT PDF INTO PAGES
# -------------------------------------------------------
def get_pages(text: str):
    pages = []
    for page in text.split("\f"):
        page = clean_page(page)
        if page:
            pages.append(page)
    return pages


# -------------------------------------------------------
# BENCH DETECTION
# -------------------------------------------------------
def split_into_benches(pages):
    benches = []
    current = None
    for page in pages:
        for line in page:
            text = line.strip()
            if not text:
                continue
            # ---------------------------------------
            # New Bench
            # ---------------------------------------
            if BENCH.match(text):
                if current is not None:
                    benches.append(current)
                current = {
                    "judge": text,
                    "lines": []
                }
                continue
            # Ignore anything before first judge
            if current is None:
                continue
            current["lines"].append(text)
    if current is not None:
        benches.append(current)
    return benches


# -------------------------------------------------------
# SPLIT BENCH INTO SECTIONS
# -------------------------------------------------------
def split_sections(bench):
    sections = []
    current = {
        "section": "",
        "lines": []
    }
    for line in bench["lines"]:
        if SECTION.match(line):
            if current["lines"]:
                sections.append(current)
            current = {
                "section": line,
                "lines": []
            }
            continue
        current["lines"].append(line)
    if current["lines"]:
        sections.append(current)
    return sections


# -------------------------------------------------------
# SPLIT SECTION INTO CASE BLOCKS
# -------------------------------------------------------
def split_case_blocks(section):
    cases = []
    current = None
    for line in section["lines"]:
        line = line.rstrip()
        # ---------------------------------------------
        # Start of new case
        # ---------------------------------------------
        m = CASE_START.match(line)
        if m:
            if current:
                cases.append(current)
            current = {
                "serial": m.group(1),
                "case_no": m.group(2),
                "section": section["section"],
                "lines": []
            }
            continue
        if current is None:
            continue
        current["lines"].append(line)
    if current:
        cases.append(current)
    return cases


# -------------------------------------------------------
# SPLIT ALL BENCHES INTO CASES
# -------------------------------------------------------
def collect_case_blocks(benches):
    all_cases = []
    for bench in benches:
        sections = split_sections(bench)
        for section in sections:
            blocks = split_case_blocks(section)
            for block in blocks:
                block["judge"] = bench["judge"]
                all_cases.append(block)
    return all_cases


# -------------------------------------------------------
# PARSE ONE CASE BLOCK
# -------------------------------------------------------
def parse_case(block, global_sr, day, month, year):
    lines = [x.strip() for x in block["lines"] if x.strip()]
    category = ""
    petitioner = []
    respondent = []
    pet_adv = []
    res_adv = []
    state = "category"
    i = 0
    while i < len(lines):
        line = lines[i]
        # -------------------------
        # CATEGORY
        # -------------------------
        if state == "category":
            if line.startswith("("):
                category = line.strip("()")
                i += 1
                state = "petitioner"
                continue
            state = "petitioner"
        # -------------------------
        # PETITIONER
        # -------------------------
        if state == "petitioner":
            if line.upper() == "VS":
                state = "respondent"
                i += 1
                continue
            petitioner.append(line)
            i += 1
            continue
        # -------------------------
        # RESPONDENT
        # -------------------------
        if state == "respondent":
            if line == "--":
                state = "respondent_adv"
                i += 1
                continue
            respondent.append(line)
            i += 1
            continue
        # -------------------------
        # RESPONDENT ADVOCATE
        # -------------------------
        if state == "respondent_adv":
            res_adv.append(line)
            i += 1
            continue
            
    if len(respondent):
        actual_respondent = respondent[0]
        petitioner_advocate = respondent[1:]
    else:
        actual_respondent = ""
        petitioner_advocate = []

    record = {
        "Sr_No": block["serial"],
        "Global_Sr": global_sr,
        "Bench": block["judge"],
        "Section": block["section"],
        "Case_No": block["case_no"],
        "Case_Category": category,
        "Petitioner": " ".join(petitioner),
        "Respondent": actual_respondent,
        "Petitioner_Advocate": " ".join(petitioner_advocate),
        "Respondent_Advocate": " ".join(res_adv),
        "Day": day,
        "Month": month,
        "Year": year,
    }
    return record


# -------------------------------------------------------
# PARSE ALL CASES
# -------------------------------------------------------
def parse_all_cases(case_blocks, day, month, year):
    records = []
    global_sr = 1
    for block in case_blocks:
        record = parse_case(
            block,
            global_sr,
            day,
            month,
            year
        )
        records.append(record)
        print(
            f"\rParsed {global_sr}",
            end="",
            flush=True
        )
        global_sr += 1
    print()
    return records


# -------------------------------------------------------
# BUILD RECORDS
# -------------------------------------------------------
def build_records(text):
    day, month, year = extract_date(text)
    pages = get_pages(text)
    print(f"Pages : {len(pages)}")

    benches = split_into_benches(pages)
    print(f"Benches : {len(benches)}")

    case_blocks = collect_case_blocks(benches)
    print(f"Case Blocks : {len(case_blocks)}")

    records = parse_all_cases(
        case_blocks,
        day,
        month,
        year
    )
    return records


# -------------------------------------------------------
# SAVE TO EXCEL
# -------------------------------------------------------
def save_excel(records, output_dir: Path):
    if not records:
        print("No records to save.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sindh Cause List"

    ws.append([
        "Sr_No",
        "Global_Sr",
        "Bench",
        "Section",
        "Case_No",
        "Case_Category",
        "Petitioner",
        "Respondent",
        "Petitioner_Advocate",
        "Respondent_Advocate",
        "Day",
        "Month",
        "Year"
    ])

    for r in records:
        ws.append([
            r["Sr_No"],
            r["Global_Sr"],
            r["Bench"],
            r["Section"],
            r["Case_No"],
            r["Case_Category"],
            r["Petitioner"],
            r["Respondent"],
            r["Petitioner_Advocate"],
            r["Respondent_Advocate"],
            r["Day"],
            r["Month"],
            r["Year"]
        ])

    # Auto-size columns
    for column_cells in ws.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

    # Save Excel in the same folder where PDFs are located
    filename = output_dir / OUTPUT_NAME
    try:
        wb.save(filename)
    except PermissionError:
        filename = output_dir / (
            "Sindh_Cause_List_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
            + ".xlsx"
        )
        wb.save(filename)

    print()
    print("=" * 60)
    print("Excel Saved")
    print(filename)
    print("=" * 60)


# -------------------------------------------------------
# MAIN
# -------------------------------------------------------
def main():
    # 1. Ask user for the folder path
    folder_path_input = input("Enter the path of the folder containing PDFs: ").strip()
    
    # Remove quotes if the user copy-pasted a path with quotes
    folder_path_input = folder_path_input.strip('"').strip("'")
    
    input_folder = Path(folder_path_input)

    if not input_folder.exists() or not input_folder.is_dir():
        print(f"Error: The folder '{input_folder}' does not exist or is not a valid directory.")
        return

    # 2. Get all PDFs in that folder
    pdfs = sorted(input_folder.glob("*.pdf"))
    if not pdfs:
        print(f"No PDF files found in {input_folder}")
        return

    all_records = []
    for pdf in pdfs:
        print()
        print("=" * 70)
        print("Processing:", pdf.name)
        print("=" * 70)

        text = extract_text(pdf)
        if text.strip():
            records = build_records(text)
            print()
            print(f"Extracted {len(records)} records")
            all_records.extend(records)
        else:
            print(f"Warning: Could not extract text from {pdf.name}. It might be empty or a scanned image.")

    save_excel(all_records, input_folder)

    print()
    print("Finished.")
    print("Total Records:", len(all_records))


if __name__ == "__main__":
    main()