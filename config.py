# --- Data source ---
# Reading directly from your real, live folder - this script only ever
# READS files, never writes to them, so it's safe to point straight at
# the primary source with no replication/copy step needed.
DATA_SOURCE = "local"

ONEDRIVE_FOLDER = r"C:\Users\Srija.Gouda\OneDrive - Salcomp Manufacturing\SCRAP\DMT LOGS1"

FILES = {
    "LINE-1": "DMT LOG LINE-1.xlsx",
    "LINE-2": "DMT LOG LINE-2.xlsx",
    "LINE-3": "DMT LOG LINE-3.xlsx",
    "LINE-4": "DMT LOG LINE-4.xlsx",
    "LINE 6": "DMT LOG LINE 6.xlsx",
    "LINE-7": "DMT LOG LINE-7.xlsx",
    "LINE-8": "DMT LOG LINE-8.xlsx",
    "LINE-BAT 1& BAT 2": "DMT LOG LINE-BAT 1& BAT 2.xlsx",
}

OUTPUT_PATH = r"C:\Users\Srija.Gouda\OneDrive - Salcomp Manufacturing\Defect Dashboard\defect_log_dashboard_v2.html"

# --- Production reports (FPY/FY per tester) ---
# DISABLED (2026-08-28, per Srija's request) to save pipeline run time and
# storage - each run was scanning/unzipping these HPR/DPR/WPR folders even
# though the yield cards weren't needed. get_yield_data() already handles
# an empty dict gracefully (just skips straight to "not available yet" for
# each tester group), so no other code changes were needed.
# To turn this back on later, just restore the three folder paths below:
#   "HPR": r"C:\Users\Srija.Gouda\OneDrive - Salcomp Manufacturing\Dashbaord\ENABLE HPR",
#   "DPR": r"C:\Users\Srija.Gouda\OneDrive - Salcomp Manufacturing\Dashbaord\ENABLE DPR",
#   "WPR": r"C:\Users\Srija.Gouda\OneDrive - Salcomp Manufacturing\Dashbaord\ENABLE WPR",
PRODUCTION_REPORTS_FOLDERS = {}
TESTER_GROUPS = ("FT", "HT", "GT")

# --- Stage-based model costs (primary) ---
# Reads real, already-rolled-up prices per build stage (SMT, MI, PCBA,
# M-Assembly, FG, etc.) - a unit costs less if it fails early than if it
# fails late, since fewer components/labor have gone into it by then.
STAGE_COST_FILE = r"C:\Users\Srija.Gouda\OneDrive - Salcomp Manufacturing\Dashbaord\Cost\Final_Cost_for_Dashboard.xlsx"

# --- Flat model costs (fallback) ---
# Used only when a scrap event's model/station can't be matched to a
# specific stage price above - just two columns, MODEL and COST. Keep this
# file in your Dashboard folder and update it there anytime; no code
# changes needed, ever.
MODEL_COSTS_FILE = r"C:\Users\Srija.Gouda\OneDrive - Salcomp Manufacturing\Dashbaord\Cost\COST.xlsx"

# --- Component scrap costing ---
# Read directly from the master MFG_TX1 inventory layer cost report -
# matched by Salcomp Part Number, MFG_TX1 org only, quantity-weighted
# average across all lots (same methodology as the BOM costing work).
# If a part's exact number isn't found, we retry with the "PART NUMBER
# WITHOUT REVISION" value from the Component Scrap sheet (strips _01/_02/
# etc. suffixes) before giving up on that part.
INVENTORY_LAYER_COST_REPORT = r"C:\Users\Srija.Gouda\OneDrive - Salcomp Manufacturing\Dashbaord\Cost\Inventory Layer Cost Report_Jul'26.xlsx"

# --- Component part number corrections ---
# Explicit fixes for known DMT log typos where the part number itself is
# wrong (an extra/missing letter, or a digit typo) - NOT just a revision
# number difference, which the base-code matching in build_data.py already
# handles automatically. Add more pairs here anytime a new typo turns up
# in the "no cost match found" note - no code changes needed elsewhere.
# Format: "as typed in the DMT log": "correct Salcomp Part Number"
COMPONENT_PART_ALIASES = {
    "CC03809TEP_01": "CC03809EP_01",
    "CE01484TEP_01": "CE01484EP_01",
    "CE01536TEP_01": "CE01536EP_01",
    "CNO2374": "CN02374EP",
    "FJ02008MEP_02": "FJ02008EP_02",
    "SB02413EP_01": "SB02481EP_01",
    "SB02646WEP_06": "SB02646EP_06",
    "ZZ03085EP": "ZZ03084EP",
}

# --- Scrap cost trend chart start point ---
# Only show the weekly scrap cost trend from this week onward (format:
# "YYYY-Www", e.g. "2026-W27") - keeps early sparse/near-zero weeks off
# the chart. Set to None to show the full history instead.
COST_TREND_START_WEEK = "2026-W27"
