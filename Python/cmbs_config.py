# =============================================================================
# CMBS REPORTING TOOL - CONFIGURATION
# Edit this file to change paths, test mode, and servicer mappings.
# =============================================================================

# --- Mode ---
TEST_MODE = True  # Set False to write to real S: drive folders
VALIDATION_MODE = True        # Highlight all script-touched cells yellow (set False for production)
USE_CURRENT_AS_PREV = True    # Use current month folder as prior month (set False for production)

# --- Paths ---
PROD_ROOT  = r"S:\Lenders"
TEST_ROOT  = r"S:\Lenders\Z. CMBS Test"
LOG_FOLDER = r"S:\Lenders\Z. CMBS Test\Excel Log"   # Always writes log here in test mode

# --- Servicer code → S:\Lenders subfolder name ---
SERVICER_FOLDER_MAP = {
    "K":  "KeyBank",
    "M":  "Midland",
    "TM": "Trimont",
    "WF": "Trimont",
}

# --- CREFC subfolder name variants to try when scanning deal folders ---
CREFC_FOLDER_VARIANTS = ["CREFC", "CREFCs", "CREFC Reports", "CMSAs", "CFEFC"]

# --- IRP column indices (1-based, matching PIRPXLPU layout) ---
IRP_COL_TRANS_ID   = 1    # A  - Transaction ID
IRP_COL_LOAN_ID    = 3    # C  - Loan ID  (primary key)
IRP_COL_LOAN_NUM   = 154  # EX - Loan Number (fallback key for Midland)
IRP_COL_END_BAL    = 7    # G  - Ending Scheduled Balance
IRP_COL_DIST_DATE  = 5    # E  - Distribution Date
IRP_COL_BEG_BAL    = 6    # F  - Beginning Scheduled Balance
IRP_COL_PAID_THRU  = 8    # H  - Paid Through Date
IRP_COL_SCHED_INT  = 23   # W  - Scheduled Interest Amount
IRP_COL_SCHED_PRIN = 24   # X  - Scheduled Principal Amount
IRP_COL_TOTAL_RES  = 104  # CZ - Total Reserve Balance
IRP_COL_RPT_BEGIN  = 135  # EE - Reporting Period Begin Date
IRP_COL_RPT_END    = 136  # EF - Reporting Period End Date
IRP_COL_SERVICER   = 133  # EC - Master Servicer
IRP_COL_DET_DATE   = 153  # EW - Determination Date

# Columns copied in-place for Periodic (position-matched)
PERIODIC_UPDATE_COLS = [
    IRP_COL_DIST_DATE, IRP_COL_BEG_BAL, IRP_COL_PAID_THRU,
    IRP_COL_SCHED_INT, IRP_COL_SCHED_PRIN, IRP_COL_TOTAL_RES,
    IRP_COL_RPT_BEGIN, IRP_COL_RPT_END,
    26,   # Z  - Negative Amortization/Deferred Interest
    27,   # AA - Unscheduled Principal Collections
    28,   # AB - Other Principal Adjustments
    30,   # AD - Prepayment Premium/YM Received
    31,   # AE - Prepayment Interest Excess
    32,   # AF - Liquidation/Prepayment Code
    37,   # AK - Total P&I Advance Outstanding
    38,   # AL - Total T&I Advance Outstanding
    39,   # AM - Other Expense Advance Outstanding
    40,   # AN - Payment Status of Loan
]

# Periodic file layout constants
PERIODIC_HEADER_ROW = 6
PERIODIC_FIRST_DATA = 7
PERIODIC_LOAN_COL   = 3   # Col C in the CREFC file
PERIODIC_TRANS_COL  = 1   # Col A in the CREFC file

# Property file layout
PROP_TRANS_COL    = 1     # A
PROP_LOAN_COL     = 2     # B
PROP_DIST_DATE_COL= 5     # E
PROP_ALLOC_PCT_COL= 20    # T - allocation %
PROP_ALLOC_BAL_COL= 21    # U - allocated ending balance
PROP_HEADER_ROW   = 6
PROP_FIRST_DATA   = 7

# Supplemental tabs that get only an A1 date update
SUPP_DATE_ONLY_TABS = ["Watchlist", "Delq Loan Status", "REO Status",
                        "Hist Mod & Corr", "Hist Mod & Corr"]

# Supplemental Total Loan tab
SUPP_TOTAL_LOAN_TAB  = "Total Loan"
SUPP_TOTAL_LOAN_COL  = 13   # M - Total Scheduled P&I Due (from IRP col 25)
SUPP_IRP_PI_COL      = 25   # IRP col 25 = Total Scheduled P&I Due

# Supplemental Comp Finan Status tab
SUPP_COMP_FINAN_TAB  = "Comp Finan Status"
SUPP_CF_BAL_COL      = 9    # I - Ending Scheduled Balance
SUPP_CF_DATE_COL     = 10   # J - Paid Through Date

# Supplemental Res LOC tab
SUPP_RES_LOC_TAB     = "Res LOC Report"
PIRPXLLR_SHEET       = "LL_Res_LOC"
PIRPXLLR_DATA_START  = 2    # First data row in LL_Res_LOC
PIRPXLLR_TRANS_COL   = 1    # Col A = Transaction ID in PIRPXLLR
PIRPXLLR_LOAN_COL    = 3    # Col C = Loan ID
PIRPXLLR_PROSP_COL   = 4    # Col D = Prospectus Loan ID
PIRPXLLR_PTD_COL     = 6    # Col F = Paid Through Date (filter)
PIRPXLLR_PASTE_COLS  = 18   # Number of columns to paste (A:R)

# --- Deal Tracker sheet name inside the xlsm (still used for lookup) ---
DEAL_TRACKER_SHEET = "Deal Tracker"
DT_LENDER_COL      = 2     # B
DT_TRANS_COL       = 3     # C
DT_DESC_COL        = 4     # D
DT_FIRST_DATA_ROW  = 5

# =============================================================================
# DEAL FOLDER OVERRIDES
# =============================================================================
# Use this table for any deal where the automatic S: drive scan fails.
# Common reasons: non-standard folder naming, nested subfolders, typos
# in the CREFC subfolder name, or series IDs that don't match the regex.
#
# Key   = Transaction ID exactly as it appears in the IRP (case-sensitive)
# Value = Full path to the DEAL folder on production (NOT the CREFC subfolder).
#         The tool will still scan inside for CREFC/CREFCs/etc. subfolders.
#         Do NOT include a trailing backslash.
#
# To add a new deal: copy one of the lines below and fill in the details.
# =============================================================================
DEAL_FOLDER_OVERRIDES = {
    # BANK5: series "20255YR15" / "20255YR18" have 5 leading digits so regex misses them
    "BANK5 20255YR15":  r"S:\Lenders\Trimont\Reporting\2025-5YR15 BANK5 (Del Prado)",
    "BANK5 20255YR18":  r"S:\Lenders\Trimont\Reporting\2025-5YR18 BANK5 (Naugatuck)",

    # KeyBank deals nested inside 1--ACTIVE POOLS subfolder
    "MSC 2021-L5":      r"S:\Lenders\KeyBank\Reporting\1--ACTIVE POOLS\5 -- MSC 2021-L5 (K2H Airpark and Little Boston)",
    "BBCMS 2021-C12":   r"S:\Lenders\KeyBank\Reporting\1--ACTIVE POOLS\6 -- BBCMS 2021-C12 (MCP Ind)",
    "BMO 2023-5C2":     r"S:\Lenders\KeyBank\Reporting\1--ACTIVE POOLS\7 -- BMO 2023-5C2 (Hawaii St. Office)",

    # BBCMS 2025-C32: CREFC subfolder has a typo ("CFEFC" not "CREFC")
    # Override points to the deal folder; "CFEFC" added to CREFC_FOLDER_VARIANTS below
    "BBCMS 2025-C32":   r"S:\Lenders\Midland\Reporting\2025-C32 BBCMS",
}

# Filename lookup override when IRP Transaction ID doesn't match filenames on disk.
# Key = Transaction ID (as in IRP), Value = string used in CREFC filenames.
# Example: IRP has "BANK5 20255YR15" but files are named "BANK5 2025-5YR15".
DEAL_FILENAME_OVERRIDES = {
    "BANK5 20255YR15": "BANK5 2025-5YR15",
    "BANK5 20255YR18": "BANK5 2025-5YR18",
}
