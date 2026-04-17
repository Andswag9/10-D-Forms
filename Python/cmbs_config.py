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

# --- Active Conduit Pool Tracking List ---
# Fallback source when the IRP's Master Servicer column is empty (e.g. new deals).
# Expected location: same directory as this config file.
# File is matched by glob so the monthly date in the filename doesn't matter.
TRACKING_LIST_GLOB = "Active Conduit Pool Tracking List*.xlsx"
TRACKING_LIST_SHEET = "Active Pools"
TRACKING_LIST_HEADER_ROW = 3
TRACKING_LIST_SERVICER_COL = 1   # A - Master Servicer
TRACKING_LIST_POOL_COL     = 2   # B - Pool

# Master copy of the tracking list on S: drive.
# Local tracking list is refreshed from this at the start of each run.
TRACKING_LIST_SOURCE = r"S:\Reporting\Pooling List\Active Conduit Pool & Lender Abstract Tracking List.xlsx"

# Column mapping (source_col, dest_col), 1-based. Source is the master on S:,
# dest is the local copy. Source col E is skipped.
TRACKING_LIST_COL_MAP = [
    (1, 1),   # Master Servicer
    (2, 2),   # Pool
    (3, 3),   # Cashiered or Non-cashiered
    (4, 4),   # GID
    (6, 5),   # Loan Number (source F -> dest E)
    (7, 6),   # Borrower Name (source G -> dest F)
]

# Map Master Servicer names (as they appear in the tracking list) → servicer codes
# used in SERVICER_FOLDER_MAP above.
TRACKING_SERVICER_TO_CODE = {
    "Key Bank":   "K",
    "KeyBank":    "K",
    "Midland":    "M",
    "Trimont":    "TM",
    "Wells Fargo": "WF",
}

# --- CREFC subfolder name variants to try when scanning deal folders ---
CREFC_FOLDER_VARIANTS = ["CREFC", "CREFCs", "CREFC Reports", "CMSAs", "CFEFC"]

# --- Sample CREFC templates for new deals (no prior month files) ---
# Located one directory above the Python scripts directory.
SAMPLE_TEMPLATE_GLOB = "!_SAMPLE_CREFC*.xls"

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
                        "Hist Mod & Corr"]

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
# DEAL FOLDER & FILENAME OVERRIDES
# =============================================================================
# Loaded from deal_overrides.json in the same directory as this file.
# Edit that JSON file (not this .py) to add or change deal mappings.
#
# folder_overrides:   Transaction ID -> full deal folder path (not CREFC subfolder)
# filename_overrides: Transaction ID -> string used in CREFC filenames on disk
# =============================================================================
import json as _json
import os as _os

_OVERRIDES_PATH = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "deal_overrides.json")
try:
    with open(_OVERRIDES_PATH, "r", encoding="utf-8") as _f:
        _overrides = _json.load(_f)
except (FileNotFoundError, _json.JSONDecodeError):
    _overrides = {}

DEAL_FOLDER_OVERRIDES = {k: v for k, v in _overrides.get("folder_overrides", {}).items() if k != "_comment"}
DEAL_FILENAME_OVERRIDES = {k: v for k, v in _overrides.get("filename_overrides", {}).items() if k != "_comment"}
DEAL_SERVICER_OVERRIDES = {k: v for k, v in _overrides.get("servicer_overrides", {}).items() if k != "_comment"}
