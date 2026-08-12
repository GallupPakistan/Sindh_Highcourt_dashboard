#!/usr/bin/env python3
"""
Sindh High Court Cause List - Multi-Bench Downloader, Parser & Combined Master Builder
Version 1.0

Requirements:
    pip install openpyxl PyMuPDF requests --break-system-packages

Description:
    1. Downloads Daily List cause list PDFs for all 5 Sindh High Court benches
       (Karachi Principal Seat, Hyderabad, Sukkur, Larkana, Mirpurkhas)
       for every date in a given range.
    2. Saves every downloaded PDF into a single shared folder.
    3. Parses each PDF and extracts case data (same parsing logic as the
       single-bench script), tagging each record with its City/Bench source.
    4. Combines all parsed records from all benches and all dates into
       one single combined master Excel file.

Usage:
    python sindh_causelist_multibench.py [start_date] [end_date] [output_folder]

    Dates in DD-MM-YYYY format. Defaults:
        start_date     -> 13-07-2026
        end_date       -> today
        output_folder  -> "sindh_causelist_master" in the current directory

    Example:
        python sindh_causelist_multibench.py 13-07-2026 12-08-2026 sindh_causelist_master
"""

import re
import sys
import requests
from pathlib import Path
from datetime import datetime, timedelta
import openpyxl
import fitz  # PyMuPDF

# -------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------
BASE_URL = "https://sindhhighcourt.gov.pk/causelist/causelist_files/"

# City name -> PDF filename prefix
BENCHES = {
    "Karachi":    "1AD",
    "Hyderabad":  "2AD",
    "Sukkur":     "3AD",
    "Larkana":    "4AD",
    "Mirpurkhas": "13AD",
}

DEFAULT_START = "13-07-2026"
MASTER_FILENAME = "Sindh_Cause_List_Master_Combined.xlsx"

# -------------------------------------------------------
# REGEX PATTERNS (same parsing logic as the original single-bench script)
# -------------------------------------------------------
CASE_START = re.compile(r'^(\d+)\.\s+(.+)$')
BENCH      = re.compile(r'^(MR|MRS)\.\s+JUSTICE', re.I)
SECTION    = re.compile(r'^FOR\s+', re.I)
DATE       = re.compile(r'([A-Za-z]+)\s+(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})')
HEADERS    = [
    "HIGH COURT OF SINDH",
    "Designed & Developed",
    "Printed at",
    "Page #",
    "Daily List"
]

HEADERS_REQ = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://sindhhighcourt.gov.pk/causelist.php",
}


# -------------------------------------------------------
# BUILD PDF URL FOR A GIVEN CITY + DATE
# -------------------------------------------------------
def build_pdf_url(prefix: str, date: datetime) -> tuple[str, str]:
    """
    Constructs the PDF filename and download URL for a given bench prefix and date.
    Example: prefix="2AD", 28 July 2026 -> 2AD28JUL26.pdf
    """
    day      = date.strftime("%d")
    month    = date.strftime("%b").upper()
    year     = date.strftime("%y")
    filename = f"{prefix}{day}{month}{year}.pdf"
    url      = BASE_URL + filename
    return url, filename


# -------------------------------------------------------
# DOWNLOAD PDF
# -------------------------------------------------------
def download_pdf(url: str, save_path: Path) -> bool:
    """
    Downloads the PDF from the given URL to save_path.
    Skips the download if the file already exists.
    Returns True on success or if the file already exists, False if unavailable.
    """
    if save_path.exists():
        print(f"    [SKIP] Already downloaded: {save_path.name}")
        return True

    try:
        response = requests.get(url, headers=HEADERS_REQ, timeout=60)

        if response.status_code == 200:
            content_type = response.headers.get("Content-Type", "")
            if "pdf" not in content_type.lower() and len(response.content) < 5000:
                print(f"    [SKIP] Not a PDF (likely not published): {save_path.name}")
                return False

            save_path.write_bytes(response.content)
            size_kb = len(response.content) / 1024
            print(f"    [OK] Downloaded: {save_path.name} ({size_kb:.1f} KB)")
            return True

        elif response.status_code == 404:
            print(f"    [SKIP] Not found (holiday / not published): {save_path.name}")
            return False

        else:
            print(f"    [ERROR] HTTP {response.status_code}: {save_path.name}")
            return False

    except requests.exceptions.ConnectionError:
        print(f"    [ERROR] Connection failed: {save_path.name}")
        return False
    except requests.exceptions.Timeout:
        print(f"    [ERROR] Timed out: {save_path.name}")
        return False
    except Exception as e:
        print(f"    [ERROR] Unexpected error for {save_path.name}: {e}")
        return False


# -------------------------------------------------------
# PDF TEXT EXTRACTION
# -------------------------------------------------------
def extract_text(pdf_file: Path) -> str:
    """Extracts plain text from a PDF file using PyMuPDF."""
    full_text = ""
    try:
        with fitz.open(pdf_file) as doc:
            for page in doc:
                full_text += page.get_text("text") + "\f"
    except Exception as e:
        print(f"    [ERROR] Failed to read PDF {pdf_file.name}: {e}")
    return full_text


# -------------------------------------------------------
# DATE EXTRACTION FROM PDF TEXT
# -------------------------------------------------------
def extract_date(text: str):
    """Extracts day, month, and year from the PDF text."""
    m = DATE.search(text)
    if not m:
        return "", "", ""
    return m.group(2), m.group(3), m.group(4)


# -------------------------------------------------------
# LINE CLEANING
# -------------------------------------------------------
def should_skip(line: str) -> bool:
    """Returns True if a line is a header, footer, or page marker to be ignored."""
    s = line.strip()
    if not s:
        return True
    for h in HEADERS:
        if h.lower() in s.lower():
            return True
    if re.search(r"Page\s*#?\s*\d+", s, re.I):
        return True
    if DATE.search(s):
        return True
    return False


def clean_page(page_text: str) -> list:
    """Removes header/footer lines and normalizes whitespace."""
    cleaned = []
    for line in page_text.splitlines():
        if should_skip(line):
            continue
        line = re.sub(r"\s+$", "", line)
        cleaned.append(line)
    return cleaned


# -------------------------------------------------------
# PAGE SPLITTING
# -------------------------------------------------------
def get_pages(text: str) -> list:
    """Splits full PDF text into individual cleaned pages."""
    pages = []
    for page in text.split("\f"):
        page = clean_page(page)
        if page:
            pages.append(page)
    return pages


# -------------------------------------------------------
# BENCH (JUDGE) DETECTION
# -------------------------------------------------------
def split_into_benches(pages: list) -> list:
    """Groups lines under their respective judge/bench headings."""
    benches = []
    current = None
    for page in pages:
        for line in page:
            text = line.strip()
            if not text:
                continue
            if BENCH.match(text):
                if current is not None:
                    benches.append(current)
                current = {"judge": text, "lines": []}
                continue
            if current is None:
                continue
            current["lines"].append(text)
    if current is not None:
        benches.append(current)
    return benches


# -------------------------------------------------------
# SECTION SPLITTING
# -------------------------------------------------------
def split_sections(bench: dict) -> list:
    """Splits a bench's lines into logical sections (e.g. FOR HEARING, FOR ORDERS)."""
    sections = []
    current  = {"section": "", "lines": []}
    for line in bench["lines"]:
        if SECTION.match(line):
            if current["lines"]:
                sections.append(current)
            current = {"section": line, "lines": []}
            continue
        current["lines"].append(line)
    if current["lines"]:
        sections.append(current)
    return sections


# -------------------------------------------------------
# CASE BLOCK SPLITTING
# -------------------------------------------------------
def split_case_blocks(section: dict) -> list:
    """Identifies individual case blocks within a section."""
    cases   = []
    current = None
    for line in section["lines"]:
        line = line.rstrip()
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
# COLLECT ALL CASE BLOCKS FROM ALL BENCHES (JUDGES)
# -------------------------------------------------------
def collect_case_blocks(benches: list) -> list:
    """Flattens all judge-benches/sections into a single list of case blocks."""
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
# PARSE A SINGLE CASE BLOCK
# -------------------------------------------------------
def parse_case(block: dict, global_sr: int, day: str, month: str, year: str,
                city: str) -> dict:
    """
    Parses a single case block into a structured record.
    Extracts: category, petitioner, respondent, and advocate names.
    """
    lines      = [x.strip() for x in block["lines"] if x.strip()]
    category   = ""
    petitioner = []
    respondent = []
    res_adv    = []
    state      = "category"
    i          = 0

    while i < len(lines):
        line = lines[i]

        if state == "category":
            if line.startswith("("):
                category = line.strip("()")
                i += 1
                state = "petitioner"
                continue
            state = "petitioner"

        if state == "petitioner":
            if line.upper() == "VS":
                state = "respondent"
                i += 1
                continue
            petitioner.append(line)
            i += 1
            continue

        if state == "respondent":
            if line == "--":
                state = "respondent_adv"
                i += 1
                continue
            respondent.append(line)
            i += 1
            continue

        if state == "respondent_adv":
            res_adv.append(line)
            i += 1
            continue

    actual_respondent   = respondent[0] if respondent else ""
    petitioner_advocate = respondent[1:] if len(respondent) > 1 else []

    return {
        "City":                city,
        "Sr_No":               block["serial"],
        "Global_Sr":           global_sr,
        "Bench":               block["judge"],
        "Section":             block["section"],
        "Case_No":             block["case_no"],
        "Case_Category":       category,
        "Petitioner":          " ".join(petitioner),
        "Respondent":          actual_respondent,
        "Petitioner_Advocate": " ".join(petitioner_advocate),
        "Respondent_Advocate": " ".join(res_adv),
        "Day":                 day,
        "Month":               month,
        "Year":                year,
    }


# -------------------------------------------------------
# PARSE ALL CASE BLOCKS FOR ONE PDF
# -------------------------------------------------------
def parse_all_cases(case_blocks: list, day: str, month: str, year: str,
                     city: str) -> list:
    """Iterates over all case blocks and returns a list of parsed records."""
    records   = []
    global_sr = 1
    for block in case_blocks:
        record = parse_case(block, global_sr, day, month, year, city)
        records.append(record)
        global_sr += 1
    return records


# -------------------------------------------------------
# BUILD ALL RECORDS FROM ONE PDF'S EXTRACTED TEXT
# -------------------------------------------------------
def build_records(text: str, city: str) -> list:
    """Orchestrates the full parsing pipeline from raw text to structured records."""
    day, month, year = extract_date(text)
    pages            = get_pages(text)
    benches          = split_into_benches(pages)
    case_blocks      = collect_case_blocks(benches)
    records          = parse_all_cases(case_blocks, day, month, year, city)
    return records


# -------------------------------------------------------
# SAVE COMBINED RECORDS TO ONE MASTER EXCEL FILE
# -------------------------------------------------------
def save_master_excel(records: list, output_dir: Path) -> None:
    """Writes all combined records (all cities, all dates) to one master Excel file."""
    if not records:
        print("\n[WARN] No records found across any bench/date. Master file not created.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sindh Cause List Master"

    column_headers = [
        "City", "Sr_No", "Global_Sr", "Bench", "Section",
        "Case_No", "Case_Category", "Petitioner", "Respondent",
        "Petitioner_Advocate", "Respondent_Advocate",
        "Day", "Month", "Year"
    ]
    ws.append(column_headers)

    for r in records:
        ws.append([
            r["City"], r["Sr_No"], r["Global_Sr"], r["Bench"], r["Section"],
            r["Case_No"], r["Case_Category"], r["Petitioner"], r["Respondent"],
            r["Petitioner_Advocate"], r["Respondent_Advocate"],
            r["Day"], r["Month"], r["Year"]
        ])

    for column_cells in ws.columns:
        max_length    = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

    filename = output_dir / MASTER_FILENAME
    try:
        wb.save(filename)
    except PermissionError:
        filename = output_dir / (
            "Sindh_Cause_List_Master_Combined_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
            + ".xlsx"
        )
        wb.save(filename)

    print()
    print("=" * 60)
    print("[OK] Master combined Excel file saved successfully.")
    print(f"     {filename}")
    print(f"     Total records: {len(records)}")
    print("=" * 60)


# -------------------------------------------------------
# DATE RANGE HELPER
# -------------------------------------------------------
def daterange(start: datetime, end: datetime):
    """Yields each date from start to end (inclusive)."""
    days = (end - start).days
    for n in range(days + 1):
        yield start + timedelta(days=n)


# -------------------------------------------------------
# MAIN ENTRY POINT
# -------------------------------------------------------
def main():
    print("=" * 60)
    print("  Sindh High Court — Multi-Bench Cause List Downloader")
    print("  (Karachi, Hyderabad, Sukkur, Larkana, Mirpurkhas)")
    print("=" * 60)

    start_str  = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_START
    end_str    = sys.argv[2] if len(sys.argv) > 2 else datetime.today().strftime("%d-%m-%Y")
    folder_arg = sys.argv[3] if len(sys.argv) > 3 else "sindh_causelist_master"

    start_date = datetime.strptime(start_str, "%d-%m-%Y")
    end_date   = datetime.strptime(end_str, "%d-%m-%Y")

    pdf_folder    = Path(folder_arg) / "pdfs"
    output_folder = Path(folder_arg)
    pdf_folder.mkdir(parents=True, exist_ok=True)
    output_folder.mkdir(parents=True, exist_ok=True)

    print(f"\n[INFO] Date range   : {start_date.strftime('%d %b %Y')} -> {end_date.strftime('%d %b %Y')}")
    print(f"[INFO] Benches      : {', '.join(BENCHES.keys())}")
    print(f"[INFO] PDF folder   : {pdf_folder}")
    print(f"[INFO] Output folder: {output_folder}\n")

    all_records   = []
    total_pdfs_ok = 0
    total_pdfs_ct = 0

    for current_date in daterange(start_date, end_date):
        print("-" * 60)
        print(f"DATE: {current_date.strftime('%d %B %Y')}")
        print("-" * 60)

        for city, prefix in BENCHES.items():
            total_pdfs_ct += 1
            url, filename = build_pdf_url(prefix, current_date)
            pdf_path = pdf_folder / filename

            print(f"  [{city}] {filename}")
            success = download_pdf(url, pdf_path)

            if not success:
                continue

            total_pdfs_ok += 1

            text = extract_text(pdf_path)
            if not text.strip():
                print(f"    [WARN] No extractable text in {filename}")
                continue

            records = build_records(text, city)
            print(f"    [INFO] Parsed {len(records)} case records")
            all_records.extend(records)

    print("\n" + "=" * 60)
    print(f"[SUMMARY] PDFs downloaded/found: {total_pdfs_ok} / {total_pdfs_ct} attempted")
    print(f"[SUMMARY] Total case records parsed: {len(all_records)}")
    print("=" * 60)

    save_master_excel(all_records, output_folder)

    print("\n[OK] All tasks completed successfully.")


if __name__ == "__main__":
    main()
