"""
split_email_10d_forms.py
========================
Splits a signed multi-page 10-D PDF into individual per-pool PDFs.

Each page's pool name is read directly from the page text
(looks for "Pool Name: XXXX YYYY-ZZZZ") so the filenames always
match what's printed on the form.

Each page's distribution date is also read from the page text
and used to sort files into date subfolders.

Pages with no pool name are skipped (no PDF written).

HOW TO USE:
  1. Source Excel (for reference / printing): Template and Instructions\\10D Forms.xlsx
  2. Save the signed combined PDF under Sent\\MM.YYYY\\, e.g.:
       Sent\\04.2026\\10D Forms_April 2026_signed.pdf
  3. Run: python split_email_10d_forms.py
     Optional: python split_email_10d_forms.py --email-after
     (after a successful split, runs Step 3 create_10d_emails.py for that month)

  DESTINATION_FOLDER_BASE must match SENT_BASE in create_10d_emails.py.

REQUIREMENTS:
  - pip install PyMuPDF
"""

import argparse
import fitz  # PyMuPDF
import os
import re
import subprocess
import sys
from datetime import datetime


# ============================================================
#  CONFIGURATION — edit these paths as needed
# ============================================================

# False = faster saves, slightly larger PDFs (usually fine for email).
# True = zlib-compress streams (slower on many pages).
PDF_SAVE_DEFLATE = False

# Above this many pages, use one-line progress (faster in CMD). Override with env SPLIT_10D_VERBOSE=1
COMPACT_PROGRESS_AFTER_PAGES = 20

# Master workbook (not read by this script; documented for your workflow)
EXCEL_WORKBOOK = (
    r"S:\Lenders\Trimont\Reporting\Z - 10-D Letters\Template and Instructions\10D Forms.xlsx"
)

# Signed combined PDFs live here, one per reporting month:
#   <Sent>\04.2026\10D Forms_April 2026_signed.pdf
#   <Sent>\05.2026\10D Forms_May 2026_signed.pdf
SOURCE_FOLDER = r"S:\Lenders\Trimont\Reporting\Z - 10-D Letters\Sent"

# Base folder where split PDFs are saved (MM.YYYY subfolder + M.DD.YYYY inside).
# Keep identical to SENT_BASE in ..\Step 3. Send 10 - D Emails\create_10d_emails.py
DESTINATION_FOLDER_BASE = r"S:\Lenders\Trimont\Reporting\Z - 10-D Letters\Sent"


# --- Pre-compiled regex (avoids re-parsing patterns every page) ---
_RE_SUBMISSION_DATE = re.compile(
    r"(?is)date\s*of\s*submission\s*:\s*"
    r"([A-Za-z]+\s+\d{1,2},\s*\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})"
)
_RE_LABELED_DATE = re.compile(
    r"(?im)^\s*date\s*:\s*"
    r"([A-Za-z]+\s+\d{1,2},\s*\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*$"
)
_RE_NUMERIC_DATES = re.compile(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b")
_RE_POOL_NAME = re.compile(r"(?i)pool\s*name\s*:\s*(.+)")
_RE_MM_YYYY = re.compile(r"^(\d{2})\.(\d{4})$")
_RE_PARENT_MM_YYYY = re.compile(r"^\d{2}\.\d{4}$")
_RE_FILENAME_MONTH = re.compile(r"(\w+)\s*(\d{4})")

_DATE_FORMATS = ("%m/%d/%Y", "%m-%d-%Y", "%B %d, %Y", "%b %d, %Y")
_NUMERIC_DATE_FORMATS = ("%m/%d/%Y", "%m-%d-%Y")


# ============================================================
#  HELPER FUNCTIONS
# ============================================================

def _normalize_date_folder_name(value):
    """Convert a parsed datetime into M.DD.YYYY folder format."""
    return f"{value.month}.{value.day}.{value.year}"


def _parse_date_string_to_folder(raw, formats):
    raw = raw.strip()
    for fmt in formats:
        try:
            return _normalize_date_folder_name(datetime.strptime(raw, fmt))
        except ValueError:
            continue
    return None


def extract_distribution_date_folder_from_page(page_text):
    """
    Extract distribution date from page text and return folder name (M.DD.YYYY).
    Returns None if no date can be extracted.
    """
    submission_match = _RE_SUBMISSION_DATE.search(page_text)
    if submission_match:
        parsed = _parse_date_string_to_folder(
            submission_match.group(1), _DATE_FORMATS
        )
        if parsed:
            return parsed

    labeled_match = _RE_LABELED_DATE.search(page_text)
    if labeled_match:
        parsed = _parse_date_string_to_folder(labeled_match.group(1), _DATE_FORMATS)
        if parsed:
            return parsed

    for match in _RE_NUMERIC_DATES.findall(page_text):
        parsed = _parse_date_string_to_folder(match, _NUMERIC_DATE_FORMATS)
        if parsed:
            return parsed

    return None


def extract_pool_name_from_page(page_text):
    """
    Extract the pool name directly from the PDF page text.

    Looks for the line "Pool Name: XXXX YYYY-ZZZZ" which appears
    near the top of every 10-D form page. Returns just the pool
    name portion (e.g., "BANK 2019-BNK23").

    Returns None if no pool name is found.
    """
    match = _RE_POOL_NAME.search(page_text)
    if match:
        pool_name = match.group(1).strip().split("\n")[0].strip()
        return pool_name
    return None


def _month_folder_rank_for_pdf(pdf_path):
    """
    If the PDF sits in Sent\\MM.YYYY\\, return a sortable int (newer month = larger).
    PDFs loose in Sent\\ (not under MM.YYYY) get 0.
    """
    parent = os.path.basename(os.path.dirname(pdf_path))
    m = _RE_MM_YYYY.match(parent)
    if m:
        month, year = int(m.group(1)), int(m.group(2))
        return year * 100 + month
    return 0


def get_most_recent_signed_pdf(folder):
    """
    Find the signed PDF to split under ``folder`` and direct MM.YYYY subfolders.

    - Only files with 'signed' in the name (case-insensitive) and extension .pdf.
    - Prefers the newest MM.YYYY month folder (e.g. 05.2026 over 04.2026).
    - Within that folder, picks the most recently modified file.
    """
    search_folders = [folder]
    try:
        for item in os.listdir(folder):
            subfolder = os.path.join(folder, item)
            if os.path.isdir(subfolder):
                search_folders.append(subfolder)
    except FileNotFoundError:
        raise FileNotFoundError(f"Source folder does not exist: {folder}") from None

    all_signed_pdfs = []
    for search_dir in search_folders:
        try:
            for f in os.listdir(search_dir):
                full_path = os.path.join(search_dir, f)
                if (f.lower().endswith(".pdf")
                        and "signed" in f.lower()
                        and os.path.isfile(full_path)):
                    all_signed_pdfs.append(full_path)
        except (OSError, PermissionError):
            continue

    if not all_signed_pdfs:
        raise FileNotFoundError(
            f"No signed PDF files found in {folder} or its MM.YYYY subfolders."
        )

    all_signed_pdfs.sort(
        key=lambda p: (_month_folder_rank_for_pdf(p), os.path.getmtime(p)),
        reverse=True,
    )
    return all_signed_pdfs[0]


def detect_month_folder_from_pdf(pdf_path):
    """
    Figure out which MM.YYYY month folder this PDF belongs to.
    First checks if the PDF is already inside a month folder.
    Falls back to extracting the month from the filename.
    """
    parent = os.path.basename(os.path.dirname(pdf_path))
    if _RE_PARENT_MM_YYYY.match(parent):
        return parent, os.path.dirname(pdf_path)

    filename = os.path.basename(pdf_path)
    match = _RE_FILENAME_MONTH.search(filename)
    if match:
        month_name = match.group(1)
        year = match.group(2)
        try:
            month_num = datetime.strptime(month_name, "%B").month
            folder_name = f"{month_num:02d}.{year}"
            folder_path = os.path.join(DESTINATION_FOLDER_BASE, folder_name)
            return folder_name, folder_path
        except ValueError:
            pass

    # Last resort: use current date
    now = datetime.now()
    folder_name = f"{now.month:02d}.{now.year}"
    folder_path = os.path.join(DESTINATION_FOLDER_BASE, folder_name)
    return folder_name, folder_path


def _compact_progress_enabled(total_pages):
    if os.environ.get("SPLIT_10D_VERBOSE", "").strip() == "1":
        return False
    return total_pages > COMPACT_PROGRESS_AFTER_PAGES


def _save_single_page_pdf(src_doc, page_index, output_path):
    """Write one page to a new PDF. Uses garbage=0; deflate optional (see PDF_SAVE_DEFLATE)."""
    with fitz.open() as new_doc:
        new_doc.insert_pdf(src_doc, from_page=page_index, to_page=page_index)
        try:
            new_doc.save(
                output_path,
                garbage=0,
                deflate=PDF_SAVE_DEFLATE,
            )
        except TypeError:
            new_doc.save(output_path, garbage=0)


def _create_10d_emails_script_path():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(root, "Step 3. Send 10 - D Emails", "create_10d_emails.py")


# ============================================================
#  MAIN
# ============================================================

def main(email_after=False):
    print("=" * 60)
    print("  10-D Forms PDF Splitter")
    print("=" * 60)
    print(f"  Excel workbook: {EXCEL_WORKBOOK}")
    print()

    # --- Find the signed PDF ---
    try:
        input_pdf_path = get_most_recent_signed_pdf(SOURCE_FOLDER)
        print(f"Found signed PDF: {input_pdf_path}")
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        return 1

    # --- Determine output month folder ---
    month_name, output_folder = detect_month_folder_from_pdf(input_pdf_path)
    os.makedirs(output_folder, exist_ok=True)
    print(f"Output folder: {output_folder}")
    print()

    # --- Open the PDF ---
    try:
        doc = fitz.open(input_pdf_path)
        print(f"PDF has {len(doc)} page(s).")
    except Exception as e:
        print(f"ERROR: Failed to open PDF: {e}")
        return 1

    total_pages = len(doc)
    last_distribution_folder = None
    compact = _compact_progress_enabled(total_pages)

    success_count = 0
    skipped_count = 0
    error_count = 0
    saved_filenames = set()

    if compact:
        print(
            f"Compact progress ({total_pages} pages). "
            f"Set SPLIT_10D_VERBOSE=1 for full per-page log.\n"
        )

    for i in range(total_pages):
        page = doc[i]
        page_text = page.get_text("text")

        date_from_page = extract_distribution_date_folder_from_page(page_text)
        if date_from_page:
            last_distribution_folder = date_from_page

        pool_name = extract_pool_name_from_page(page_text)
        if not pool_name:
            skipped_count += 1
            if compact:
                print()
            print(
                f"Page {i + 1}/{total_pages} — SKIP: No pool name on page; PDF not created."
            )
            continue

        if date_from_page:
            distribution_folder = date_from_page
        elif last_distribution_folder:
            distribution_folder = last_distribution_folder
            if compact:
                print()
            print(
                f"  WARNING: No date on this page; using previous page's date: "
                f"{distribution_folder}"
            )
        else:
            distribution_folder = "UnknownDate"
            if compact:
                print()
            print("  WARNING: No date found; using fallback folder: UnknownDate")

        if not compact:
            print(f"\nPage {i + 1}/{total_pages}")
            print(f"  Pool Name: {pool_name}")
            print(f"  Distribution folder: {distribution_folder}")

        filename_safe = "".join(
            c if c.isalnum() or c in " ._-" else "_"
            for c in pool_name
        ).strip()
        pdf_filename = f"{filename_safe}-10D.pdf"

        if pdf_filename in saved_filenames:
            if compact:
                print()
            print(f"  WARNING: Duplicate pool '{pool_name}' (page {i+1}) — overwriting previous PDF")
        saved_filenames.add(pdf_filename)

        dated_output_folder = os.path.join(output_folder, distribution_folder)
        os.makedirs(dated_output_folder, exist_ok=True)
        output_path = os.path.join(dated_output_folder, pdf_filename)

        try:
            _save_single_page_pdf(doc, i, output_path)
            success_count += 1
            if compact:
                print(f"\r  [{i + 1}/{total_pages}] {pdf_filename}", end="", flush=True)
            else:
                print(f"  Saved: {output_path}")
        except Exception as e:
            error_count += 1
            if compact:
                print()
            print(f"  ERROR saving: {e}")

    if compact:
        print()

    doc.close()

    # --- Summary ---
    print()
    print("=" * 60)
    print("  DONE!")
    print(f"  PDFs written: {success_count}")
    if skipped_count:
        print(f"  Pages skipped (no pool name): {skipped_count}")
    if error_count:
        print(f"  Errors: {error_count}")
    print(f"  Output: {output_folder}")
    print("=" * 60)

    if error_count:
        return 1

    if email_after and success_count == 0:
        print("\nSkipping email step: no PDFs were created.")
        return 0

    if email_after:
        email_script = _create_10d_emails_script_path()
        if not os.path.isfile(email_script):
            print(f"\nERROR: Email script not found:\n  {email_script}")
            return 1
        print()
        print("=" * 60)
        print("  Starting Outlook draft creation (same month folder)...")
        print("=" * 60)
        proc = subprocess.run(
            [
                sys.executable,
                email_script,
                "--month",
                month_name,
                "-y",
                "--no-pause",
            ],
            cwd=os.path.dirname(email_script),
        )
        if proc.returncode != 0:
            return proc.returncode

    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Split signed 10-D PDF into per-pool files."
    )
    parser.add_argument(
        "--email-after",
        action="store_true",
        help="After splitting, run create_10d_emails.py for this PDF's month folder.",
    )
    args = parser.parse_args()
    code = main(email_after=args.email_after)
    if sys.stdin.isatty():
        input("\nPress Enter to exit...")
    raise SystemExit(code if code is not None else 0)
