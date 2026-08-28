"""
Main entry point. Run this on a schedule, or let watch_and_update.py call
it automatically.

Two modes, controlled by config.USE_GRAPH_API:
- True (recommended if files are co-authored live in the browser): reads
  files directly from Microsoft's cloud via Graph API, reflecting live
  edits from anyone's open session within seconds.
- False: reads from a local OneDrive-synced folder instead. Simpler (no
  login needed) but can lag behind live co-authored edits, since it
  depends on OneDrive's own local sync timing rather than the live cloud
  state.

Steps:
1. Get the 8 files (from the cloud via Graph, or from a local folder)
2. Parse DEFECT LOG sheets and rebuild the combined dataset + insights
3. Inject the fresh data into the dashboard HTML template
4. Save the finished HTML to config.OUTPUT_PATH

Usage:
    python run_dashboard_update.py
"""
import json
import logging
import os
import sys
from datetime import datetime

import config
import build_data

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("dashboard_update.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "dashboard_template.html")


def get_model_costs():
    """Loads model costs fresh from the Excel file every run (see
    config.MODEL_COSTS_FILE). Returns {} if the file doesn't exist yet."""
    import build_data
    costs_file = getattr(config, "MODEL_COSTS_FILE", None)
    try:
        costs = build_data.load_model_costs_from_excel(costs_file)
        if costs:
            log.info(f"Loaded {len(costs)} model costs from {costs_file}")
        else:
            log.info("No model costs file found yet - scrap cost view will be empty")
        return costs
    except Exception:
        log.exception(f"Failed to read model costs file: {costs_file}")
        return {}


def get_stage_cost_data():
    """Loads stage-based costs fresh every run (see config.STAGE_COST_FILE).
    Returns None if the file doesn't exist yet, so scrap costing falls back
    to the flat MODEL_COSTS_FILE."""
    import stage_cost
    stage_file = getattr(config, "STAGE_COST_FILE", None)
    try:
        data = stage_cost.load_stage_cost_data(stage_file)
        if data:
            log.info(f"Loaded stage cost data from {stage_file} "
                      f"({len(data['product_code_to_stage_prices'])} product codes, "
                      f"{len(data['model_to_product_code'])} models mapped)")
        else:
            log.info("No stage cost file found yet - falling back to flat model costs")
        return data
    except Exception:
        log.exception(f"Failed to read stage cost file: {stage_file}")
        return None


def get_component_costs():
    """Loads MFG_TX1 weighted-average part costs fresh every run (see
    config.INVENTORY_LAYER_COST_REPORT). Returns {} if the file doesn't
    exist yet or can't be opened, so the Component Scrap tab shows
    quantities with $0 cost rather than failing the whole run."""
    inv_file = getattr(config, "INVENTORY_LAYER_COST_REPORT", None)
    try:
        costs = build_data.load_component_costs(inv_file)
        if costs:
            log.info(f"Loaded {len(costs)} MFG_TX1 part costs from {inv_file}")
        else:
            log.info(f"No component cost data loaded from {inv_file} - "
                      f"Component Scrap tab will show quantities only")
        return costs
    except Exception:
        log.exception(f"Failed to read inventory layer cost report: {inv_file}")
        return {}


def resolve_local_paths():
    """Build {line_name: full_local_path} and confirm each file exists."""
    local_paths = {}
    missing = []
    for line_name, filename in config.FILES.items():
        full_path = os.path.join(config.ONEDRIVE_FOLDER, filename)
        if not os.path.exists(full_path):
            missing.append(full_path)
            continue
        local_paths[line_name] = full_path
    if missing:
        raise FileNotFoundError(
            "Could not find these files - check config.ONEDRIVE_FOLDER and "
            "config.FILES match what's actually in File Explorer:\n" + "\n".join(missing)
        )
    return local_paths


def get_dashboard_data():
    """
    Returns (dashboard_data, insights_data), sourced according to
    config.DATA_SOURCE:
      "csv"   - reads CSVs a Power Automate flow exported from the live
                Excel table (bypasses local-sync lag entirely)
      "graph" - reads files live from Microsoft Graph
      "local" - reads a locally-synced folder (can lag for live co-authored files)
    """
    import build_data

    source = getattr(config, "DATA_SOURCE", "local")

    if source == "csv":
        import csv_ingest
        log.info(f"Reading latest CSV exports from {config.CSV_EXPORT_FOLDER} ...")
        all_rows = []
        missing = []
        for line_name in config.FILES.keys():
            latest = csv_ingest.find_latest_csv(config.CSV_EXPORT_FOLDER, line_name)
            if latest is None:
                missing.append(line_name)
                continue
            all_rows.extend(csv_ingest.extract_defect_rows_from_csv(latest, line_name))
        if missing:
            log.warning(f"No CSV export found yet for: {', '.join(missing)} - "
                        f"check the Power Automate flow(s) for those lines have run at least once")
        return build_data.build_from_rows(all_rows, model_costs=get_model_costs(), cost_trend_start_week=getattr(config, 'COST_TREND_START_WEEK', None), stage_cost_data=get_stage_cost_data())

    elif source == "graph":
        import graph_read
        log.info("Reading live data from Microsoft cloud (Graph API, no Table required)...")
        all_rows = graph_read.get_all_rows(config.GRAPH_FILE_PATHS)
        return build_data.build_from_rows(all_rows, model_costs=get_model_costs(), cost_trend_start_week=getattr(config, 'COST_TREND_START_WEEK', None), stage_cost_data=get_stage_cost_data())

    else:
        log.info(f"Reading files from local folder {config.ONEDRIVE_FOLDER} ...")
        sources = resolve_local_paths()
        return build_data.build_all(sources, model_costs=get_model_costs(), cost_trend_start_week=getattr(config, 'COST_TREND_START_WEEK', None), stage_cost_data=get_stage_cost_data(), component_costs=get_component_costs(), component_aliases=getattr(config, 'COMPONENT_PART_ALIASES', None))


def get_yield_data():
    """
    Reads the latest HPR/DPR/WPR files (if configured) and computes FPY/FY
    per tester for each. Returns None for a report type if no matching file
    has been found yet, so the dashboard can show "not available yet"
    instead of erroring.
    """
    import parse_production_report as ppr

    folders = getattr(config, "PRODUCTION_REPORTS_FOLDERS", {})
    testers = getattr(config, "TESTER_GROUPS", ("FT", "HT", "GT"))

    result = {}
    for key, report_type in [("hourly", "HPR"), ("daily", "DPR"), ("weekly", "WPR")]:
        folder = folders.get(report_type)
        if not folder or not os.path.isdir(folder):
            log.info(f"No {report_type} folder configured/found - skipping FPY/FY for {key}")
            result[key] = None
            continue
        latest = ppr.find_latest_report(folder, report_type)
        if latest is None:
            log.warning(f"No {report_type} file found yet in {folder}")
            result[key] = None
            continue
        try:
            events = ppr.parse_report(latest)
            result[key] = {
                "testers": ppr.compute_fpy_fy(events, testers=testers),
                "period": ppr.extract_report_period(latest),
                "filename": os.path.basename(latest),
            }
            log.info(f"{report_type} yield computed from {os.path.basename(latest)}")
        except Exception:
            log.exception(f"Failed to parse {report_type} file: {latest}")
            result[key] = None
    return result


def main():
    start = datetime.now()
    log.info("Starting dashboard update run")

    try:
        dashboard_data, insights_data = get_dashboard_data()
        log.info(f"Combined {dashboard_data['daily']['ALL']['total']} records for today "
                  f"({dashboard_data['max_date']})")

        dashboard_data["yield_data"] = get_yield_data()

        log.info("Rendering dashboard HTML...")
        with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
            template = f.read()

        # IMPORTANT: escape "</" before embedding JSON into a <script> block.
        # Without this, if ANY string field in the data (a problem
        # description, part number, reason, note, etc.) happens to contain
        # something like "</script>", the browser's HTML parser closes the
        # script tag right there - even though it's just inside a JSON
        # string, not real markup - and every line of JavaScript after that
        # point in the file silently never runs. That's a real crash mode:
        # the whole dashboard would render as static "-" placeholders with
        # a frozen clock, since nothing after the break executes at all.
        def safe_json(obj):
            return json.dumps(obj).replace("</", "<\\/")

        html = template.replace("__DATA_JSON__", safe_json(dashboard_data))
        html = html.replace("__INSIGHTS_JSON__", safe_json(insights_data))

        os.makedirs(os.path.dirname(config.OUTPUT_PATH), exist_ok=True)
        with open(config.OUTPUT_PATH, "w", encoding="utf-8") as f:
            f.write(html)

        elapsed = (datetime.now() - start).total_seconds()
        log.info(f"Dashboard updated successfully in {elapsed:.1f}s -> {config.OUTPUT_PATH}")

    except Exception:
        log.exception("Dashboard update failed")
        raise


if __name__ == "__main__":
    main()
