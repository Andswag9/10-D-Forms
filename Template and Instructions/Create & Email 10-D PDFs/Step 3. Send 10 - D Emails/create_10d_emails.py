"""
create_10d_emails.py
====================
Creates Outlook DRAFT emails for each split 10-D PDF in the current month's
Sent folder. Each draft has:
  - Subject:  MM_YY 10D for [Pool Name]
  - Body:     Standard boilerplate with pool name inserted
  - To/CC:    Configured recipients below
  - Attachment: The individual pool's signed PDF

HOW TO USE:
  1. Make sure the split PDFs are already in their date folders
     (e.g., .../Sent/04.2026/4.12.2026/BANK 2019-BNK23-10D.pdf)
  2. Double-click the .bat file, OR run:
         python create_10d_emails.py
  3. It will ask you which month to process (defaults to next month).
  4. Check your Outlook Drafts folder — one draft per pool PDF.

  Non-interactive (e.g. after split_email_10d_forms.py --email-after):
         python create_10d_emails.py --month 04.2026 -y --no-pause

REQUIREMENTS:
  - Windows with Outlook desktop installed
  - pip install pywin32
"""

import argparse
import glob
import os
import re
import sys
from datetime import date

# ============================================================
#  CONFIGURATION — edit these values as needed
# ============================================================

# Base path where Sent folders live
SENT_BASE = r"S:\Lenders\Trimont\Reporting\Z - 10-D Letters\Sent"

# Email recipients (same for every pool)
TO_RECIPIENTS = "c3po@trimont.com"
CC_RECIPIENTS = "llau@gantryinc.com; ctowner@gantryinc.com"

# Email body template.  {pool_name} gets replaced per email.
EMAIL_BODY = (
    "Hello,\r\n"
    "\r\n"
    "Attached is the monthly 10-D for {pool_name}.  "
    "Please let me know if you have any questions.\r\n"
    "\r\n"
    "Regards,"
)

# Your Outlook signature can be appended when you open/send if configured in Outlook.
# To force a specific sender account, set this. Leave as None for default.
SENDER_ACCOUNT = None  # e.g., "anakinnd9@outlook.com"


# ============================================================
#  HELPER FUNCTIONS
# ============================================================

def get_outlook():
    """Connect to the running Outlook application via COM."""
    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 is not installed.")
        print("Run this command and try again:")
        print("    pip install pywin32")
        sys.exit(1)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        print(f"ERROR: Could not connect to Outlook.\n  {e}")
        print("Make sure Outlook is open and try again.")
        sys.exit(1)


def pool_name_from_filename(pdf_filename):
    """
    Extract the pool name from a split PDF filename.

    The split script names files like:
        BANK 2019-BNK23-10D.pdf
        CGCMT 2016-C1-10D.pdf
        GSMS 2017-GS7-10D.pdf

    We strip the '-10D' suffix and the '.pdf' extension
    to get the clean pool name for the email subject/body.
    """
    name = os.path.splitext(pdf_filename)[0]  # remove .pdf
    # Remove trailing "-10D" (case-insensitive)
    name = re.sub(r'-10D$', '', name, flags=re.IGNORECASE)
    return name.strip()


def find_month_folder(reporting_month, reporting_year):
    """
    Find the month folder under SENT_BASE.
    Format: MM.YYYY  (e.g., 04.2026)
    """
    folder_name = f"{reporting_month:02d}.{reporting_year}"
    folder_path = os.path.join(SENT_BASE, folder_name)
    if not os.path.isdir(folder_path):
        return None
    return folder_path


def find_distribution_date_folders(month_folder):
    """
    Find subfolders that look like distribution-date folders.
    These are named like: M.DD.YYYY or MM.DD.YYYY
    (e.g., 4.12.2026  or  04.17.2026)
    Also matches folders that already have "(sent ...)" appended.
    """
    date_folders = []
    for item in os.listdir(month_folder):
        full_path = os.path.join(month_folder, item)
        if not os.path.isdir(full_path):
            continue
        # Match date-like folder names: M.DD.YYYY or variations
        if re.match(r'^\d{1,2}\.\d{1,2}\.\d{4}', item):
            date_folders.append(full_path)
    date_folders.sort()
    return date_folders


def find_pool_pdfs(folder):
    """Find all PDF files in a folder (non-recursive)."""
    pdfs = glob.glob(os.path.join(folder, "*.pdf"))
    pdfs.sort()
    return pdfs


def create_draft_email(outlook, subject, body, to, cc, attachment_path,
                       save_msg_folder=None):
    """
    Create a single Outlook draft email with an attachment.
    The email is saved to Drafts (not sent).

    If save_msg_folder is provided, also saves a copy of the email
    as an .msg file in that folder (next to the PDF).
    """
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.Subject = subject
    mail.Body = body
    mail.To = to
    mail.CC = cc

    if attachment_path and os.path.isfile(attachment_path):
        mail.Attachments.Add(attachment_path)

    # If a specific sender account is configured, set it
    if SENDER_ACCOUNT:
        for account in outlook.Session.Accounts:
            if account.SmtpAddress.lower() == SENDER_ACCOUNT.lower():
                mail._oleobj_.Invoke(
                    *(64209, 0, 8, 0, account)  # MailItem.SendUsingAccount
                )
                break

    mail.Save()  # Saves to Drafts

    # Save a copy as .msg in the distribution date folder
    if save_msg_folder:
        # Sanitize subject for use as filename (replace illegal chars)
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', subject)
        msg_path = os.path.join(save_msg_folder, f"{safe_name}.msg")
        try:
            # 3 = olMSG format (Unicode .msg)
            mail.SaveAs(msg_path, 3)
        except Exception as e:
            print(f"    -> WARNING: Draft created, but could not save .msg file: {e}")

    return True


# ============================================================
#  MAIN
# ============================================================

def _default_next_reporting_month():
    today = date.today()
    if today.month == 12:
        return 1, today.year + 1
    return today.month + 1, today.year


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    parser = argparse.ArgumentParser(
        description="Create Outlook draft emails for split 10-D PDFs."
    )
    parser.add_argument(
        "--month",
        metavar="MM.YYYY",
        help="Reporting month folder (e.g. 04.2026). If omitted, you are prompted.",
    )
    parser.add_argument(
        "-y",
        "--yes",
        action="store_true",
        help="Skip the confirmation prompt before creating drafts.",
    )
    parser.add_argument(
        "--no-pause",
        action="store_true",
        help="Do not wait for Enter at exit (for scripted runs).",
    )
    args = parser.parse_args(argv)

    print("=" * 60)
    print("  10-D Email Draft Creator")
    print("=" * 60)
    print()

    if args.month:
        try:
            parts = args.month.strip().split(".")
            proc_month = int(parts[0])
            proc_year = int(parts[1])
        except (ValueError, IndexError):
            print("ERROR: Use --month MM.YYYY (e.g., 04.2026)")
            return 1, args.no_pause
    else:
        default_month, default_year = _default_next_reporting_month()
        user_input = input(
            f"Which month folder to process? (MM.YYYY) "
            f"[default: {default_month:02d}.{default_year}]: "
        ).strip()
        if user_input:
            try:
                parts = user_input.split(".")
                proc_month = int(parts[0])
                proc_year = int(parts[1])
            except (ValueError, IndexError):
                print("ERROR: Please enter month as MM.YYYY (e.g., 04.2026)")
                return 1, args.no_pause
        else:
            proc_month = default_month
            proc_year = default_year

    month_code = f"{proc_month:02d}"
    year_code = f"{proc_year % 100:02d}"

    month_folder = find_month_folder(proc_month, proc_year)
    if not month_folder:
        print("ERROR: Month folder not found at:")
        print(f"  {os.path.join(SENT_BASE, f'{proc_month:02d}.{proc_year}')}")
        return 1, args.no_pause

    print(f"Month folder: {month_folder}")
    print()

    date_folders = find_distribution_date_folders(month_folder)
    if not date_folders:
        print("ERROR: No distribution-date subfolders found.")
        print("  Expected folders like 4.12.2026 or 4.17.2026")
        return 1, args.no_pause

    print(f"Found {len(date_folders)} distribution date folder(s):")
    for df in date_folders:
        print(f"  {os.path.basename(df)}")
    print()

    all_pdfs = []
    for df in date_folders:
        for pdf_path in find_pool_pdfs(df):
            all_pdfs.append(pdf_path)

    if not all_pdfs:
        print("ERROR: No PDF files found in the distribution date folders.")
        return 1, args.no_pause

    # Deduplicate by pool name — same pool in multiple date folders = only email once
    seen_pools = {}
    unique_pdfs = []
    for pdf_path in all_pdfs:
        pool = pool_name_from_filename(os.path.basename(pdf_path))
        if pool in seen_pools:
            print(f"  WARNING: Duplicate pool '{pool}' — skipping {pdf_path}")
            print(f"           (already found in {seen_pools[pool]})")
        else:
            seen_pools[pool] = pdf_path
            unique_pdfs.append(pdf_path)
    all_pdfs = unique_pdfs

    print(f"Found {len(all_pdfs)} pool PDF(s) to email:")
    for p in all_pdfs:
        print(f"  {os.path.basename(p)}")
    print()

    if not args.yes:
        confirm = input(
            f"Create {len(all_pdfs)} draft email(s) in Outlook? (y/n): "
        ).strip().lower()
        if confirm != "y":
            print("Cancelled.")
            return 0, args.no_pause

    print()
    print("Connecting to Outlook...")
    outlook = get_outlook()
    print("Connected!")
    print()

    success_count = 0
    skipped_count = 0
    error_count = 0

    for pdf_path in all_pdfs:
        pdf_filename = os.path.basename(pdf_path)
        pool_name = pool_name_from_filename(pdf_filename)
        dist_folder = os.path.basename(os.path.dirname(pdf_path))

        subject = f"{month_code}_{year_code} 10D for {pool_name}"
        body = EMAIL_BODY.format(pool_name=pool_name)
        dist_folder_path = os.path.dirname(pdf_path)

        # Skip if .msg already exists (draft was previously created)
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', subject)
        msg_path = os.path.join(dist_folder_path, f"{safe_name}.msg")
        if os.path.isfile(msg_path):
            print(f"  SKIP: {subject}")
            print(f"    .msg already exists — draft was previously created")
            print()
            skipped_count += 1
            continue

        print(f"  Creating draft: {subject}")
        print(f"    Attachment: {pdf_filename}")
        print(f"    Dist. date folder: {dist_folder}")

        try:
            create_draft_email(
                outlook, subject, body,
                TO_RECIPIENTS, CC_RECIPIENTS,
                pdf_path,
                save_msg_folder=dist_folder_path
            )
            success_count += 1
            print("    -> Saved to Drafts")
            print(f"    -> .msg saved to: {dist_folder_path}")
        except Exception as e:
            error_count += 1
            print(f"    -> ERROR: {e}")

        print()

    print("=" * 60)
    print("  DONE!")
    print(f"  Drafts created: {success_count}")
    if skipped_count:
        print(f"  Skipped (already sent): {skipped_count}")
    if error_count:
        print(f"  Errors: {error_count}")
    print("  Check your Outlook Drafts folder.")
    print("=" * 60)

    return (1 if error_count else 0), args.no_pause


if __name__ == "__main__":
    code, no_pause = main()
    if sys.stdin.isatty() and not no_pause:
        input("\nPress Enter to exit...")
    raise SystemExit(code)
