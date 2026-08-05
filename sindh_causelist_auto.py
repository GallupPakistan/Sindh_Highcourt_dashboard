#!/usr/bin/env python3
"""
Sindh High Court Cause List - Automated Downloader & Parser
Version 4.1 (Non-interactive, GitHub Actions compatible)

Requirements:
    pip install openpyxl PyMuPDF requests

Description:
    1. Builds today's PDF URL based on the current date
    2. Downloads the PDF if not already present
    3. Parses the PDF and extracts case data
    4. Saves everything to an Excel file in the specified folder

Usage:
    python sindh_causelist_auto.py [output_folder]
    If output_folder is not given, defaults to "cause_lists" in the current directory.
    (This removes the interactive input() prompt so it can run unattended
    in GitHub Actions / Task Scheduler without anyone typing a path.)
"""

import re
import sys
import requests
from pathlib import Path
from datetime import datetime
import openpyxl
import fitz  # PyMuPDF

# -------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------
BASE_URL    = "https://sindhhighcourt.gov.pk/causelist/causelist_files/"
PDF_PREFIX  = "1AD"          # Section prefix — update if a different section is needed
OUTPUT_NAME = f"Sindh_Cause_List_{datetime.today().strftime('%d %B %Y')}.xlsx"
# -------------------------------------------------------
# REGEX PATTERNS
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

# -------------------------------------------------------
# STEP 1: BUILD PDF URL FROM TODAY'S DATE
# -------------------------------------------------------
def build_pdf_url(date: datetime) -> tuple[str, str]:
    """
    Constructs the PDF filename and download URL based on the given date.
    Example: 13 July 2026 -> 1AD13JUL26.pdf
    """
    day      = date.strftime("%d")  # e.g. 13
    month    = date.strftime("%b").upper()       # e.g. JUL
    year     = date.strftime("%y")               # e.g. 26
    filename = f"{PDF_PREFIX}{day}{month}{year}.pdf"
    url      = BASE_URL + filename
    return url, filename


# -------------------------------------------------------
# STEP 2: DOWNLOAD PDF
# -------------------------------------------------------
def download_pdf(url: str, save_path: Path) -> bool:
    """
    Downloads the PDF from the given URL to save_path.
    Skips the download if the file already exists.
    Returns True on success or if the file already exists.
    """
    if save_path.exists():
        print(f"[INFO] PDF already exists: {save_path.name}")
        print("       Skipping download.")
        return True

    print(f"[INFO] Downloading PDF ...")
    print(f"       URL: {url}")

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Referer": "https://sindhhighcourt.gov.pk/causelist.php",
    }

    try:
        response = requests.get(url, headers=headers, timeout=60)

        if response.status_code == 200:
            # Verify the response is actually a PDF and not an error page
            content_type = response.headers.get("Content-Type", "")
            if "pdf" not in content_type.lower() and len(response.content) < 5000:
                print("[ERROR] Server did not return a PDF file.")
                print(f"        Content-Type: {content_type}")
                print("        The cause list may not have been uploaded yet.")
                return False

            save_path.write_bytes(response.content)
            size_kb = len(response.content) / 1024
            print(f"[OK]   PDF downloaded successfully. ({size_kb:.1f} KB)")
            return True

        elif response.status_code == 404:
            print("[ERROR] PDF not found on server (404).")
            print("        The cause list may not have been uploaded yet.")
            print(f"        URL: {url}")
            return False

        else:
            print(f"[ERROR] Download failed. HTTP Status: {response.status_code}")
            return False

    except requests.exceptions.ConnectionError:
        print("[ERROR] Network connection failed. Please check your internet connection.")
        return False
    except requests.exceptions.Timeout:
        print("[ERROR] Request timed out. The server may be slow — please try again later.")
        return False
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
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
        print(f"[ERROR] Failed to read PDF: {e}")
    return full_text


# -------------------------------------------------------
# DATE EXTRACTION
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
# BENCH DETECTION
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
# COLLECT ALL CASE BLOCKS FROM ALL BENCHES
# -------------------------------------------------------
def collect_case_blocks(benches: list) -> list:
    """Flattens all benches/sections into a single list of case blocks."""
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
def parse_case(block: dict, global_sr: int, day: str, month: str, year: str) -> dict:
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
# PARSE ALL CASE BLOCKS
# -------------------------------------------------------
def parse_all_cases(case_blocks: list, day: str, month: str, year: str) -> list:
    """Iterates over all case blocks and returns a list of parsed records."""
    records   = []
    global_sr = 1
    for block in case_blocks:
        record = parse_case(block, global_sr, day, month, year)
        records.append(record)
        print(f"\r[INFO] Parsed {global_sr} records", end="", flush=True)
        global_sr += 1
    print()
    return records


# -------------------------------------------------------
# BUILD ALL RECORDS FROM EXTRACTED TEXT
# -------------------------------------------------------
def build_records(text: str) -> list:
    """Orchestrates the full parsing pipeline from raw text to structured records."""
    day, month, year = extract_date(text)
    pages            = get_pages(text)
    print(f"[INFO] Pages   : {len(pages)}")

    benches     = split_into_benches(pages)
    print(f"[INFO] Benches : {len(benches)}")

    case_blocks = collect_case_blocks(benches)
    print(f"[INFO] Cases   : {len(case_blocks)}")

    records = parse_all_cases(case_blocks, day, month, year)
    return records


# -------------------------------------------------------
# SAVE RECORDS TO EXCEL
# -------------------------------------------------------
def save_excel(records: list, output_dir: Path) -> None:
    """Writes all parsed records to an Excel (.xlsx) file with auto-sized columns."""
    if not records:
        print("[WARN] No records found. Excel file will not be created.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sindh Cause List"

    column_headers = [
        "Sr_No", "Global_Sr", "Bench", "Section",
        "Case_No", "Case_Category", "Petitioner", "Respondent",
        "Petitioner_Advocate", "Respondent_Advocate",
        "Day", "Month", "Year"
    ]
    ws.append(column_headers)

    for r in records:
        ws.append([
            r["Sr_No"], r["Global_Sr"], r["Bench"], r["Section"],
            r["Case_No"], r["Case_Category"], r["Petitioner"], r["Respondent"],
            r["Petitioner_Advocate"], r["Respondent_Advocate"],
            r["Day"], r["Month"], r["Year"]
        ])

    # Auto-size each column based on content length
    for column_cells in ws.columns:
        max_length    = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

    filename = output_dir / OUTPUT_NAME
    try:
        wb.save(filename)
    except PermissionError:
        # If the file is open elsewhere, save with a timestamped name
        filename = output_dir / (
            "Sindh_Cause_List_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
            + ".xlsx"
        )
        wb.save(filename)

    print()
    print("=" * 60)
    print("[OK]   Excel file saved successfully.")
    print(f"       {filename}")
    print("=" * 60)


# -------------------------------------------------------
# MAIN ENTRY POINT
# -------------------------------------------------------
def main():
    print("=" * 60)
    print("  Sindh High Court — Cause List Automated Downloader")
    print("=" * 60)

    # 1. Get output folder from command-line argument, or default to "cause_lists"
    #    (No input() prompt here — this lets the script run unattended
    #    in GitHub Actions or Task Scheduler.)
    folder_arg    = sys.argv[1] if len(sys.argv) > 1 else "cause_lists"
    output_folder = Path(folder_arg)

    if not output_folder.exists():
        print(f"\n[INFO] Folder does not exist. Creating it ...")
        output_folder.mkdir(parents=True, exist_ok=True)
        print(f"[OK]   Folder created: {output_folder}")

    # 2. Build today's PDF URL
    today             = datetime.today()
    url, pdf_filename = build_pdf_url(today)
    pdf_path          = output_folder / pdf_filename

    print(f"\n[INFO] Date     : {today.strftime('%d %B %Y')}")
    print(f"[INFO] Filename : {pdf_filename}")
    print(f"[INFO] URL      : {url}")
    print()

    # 3. Download PDF
    print("-" * 60)
    print("STEP 1: PDF Download")
    print("-" * 60)
    success = download_pdf(url, pdf_path)

    if not success:
        print()
        print("[INFO] Process stopped — PDF not available yet (e.g. holiday or not uploaded).")
        # Exit 0 (not an error) so the GitHub Actions workflow doesn't show a false failure
        # on days when the court hasn't published a list yet.
        sys.exit(0)

    # 4. Extract and parse PDF text
    print()
    print("-" * 60)
    print("STEP 2: PDF Parsing")
    print("-" * 60)
    text = extract_text(pdf_path)

    if not text.strip():
        print("[ERROR] No text could be extracted from the PDF.")
        print("        The file may be a scanned image and not machine-readable.")
        sys.exit(1)

    records = build_records(text)
    print(f"[OK]   Total records extracted: {len(records)}")

    # 5. Save to Excel
    print()
    print("-" * 60)
    print("STEP 3: Saving to Excel")
    print("-" * 60)
    save_excel(records, output_folder)

    print()
    print("[OK]   All tasks completed successfully.")


if __name__ == "__main__":
    main()
