"""
Parses HPR (Hourly Production Report) and DPR (Daily Production Report)
CSV files from Outlook attachments.

These files have several yield-summary tables at the top (by TAN, by
Station ID) followed by the real per-test-event detail table, which starts
at the row beginning "Serial Number,Unit State,Part Number,...". This
parser skips straight to that detail table and ignores the summary tables
above it (they're redundant - we compute the same stats ourselves from the
detail rows, filtered/grouped however the dashboard needs).
"""
import csv
import collections
import glob
import os
import re
import zipfile
import io


def find_detail_header_index(lines):
    for i, line in enumerate(lines):
        if line.startswith("Serial Number,Unit State,Part Number"):
            return i
    raise ValueError("Could not find detail table header in report")


def read_report_lines(path):
    """
    Returns the raw lines of the CSV report, transparently handling the
    fact that Outlook sends these zipped up (e.g. '..._HPR_....csv.zip').
    If path ends in .zip, extracts the first .csv found inside it.
    Otherwise reads it directly as a plain .csv.
    """
    if path.lower().endswith(".zip"):
        with zipfile.ZipFile(path) as zf:
            csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
            if not csv_names:
                raise ValueError(f"No .csv file found inside zip: {path}")
            with zf.open(csv_names[0]) as f:
                text = io.TextIOWrapper(f, encoding='utf-8', errors='replace')
                return text.readlines()
    else:
        with open(path, newline='', encoding='utf-8', errors='replace') as f:
            return f.readlines()


def parse_report(path):
    """Returns a list of dicts, one per test event."""
    lines = read_report_lines(path)
    header_idx = find_detail_header_index(lines)
    detail_lines = lines[header_idx:]
    reader = csv.DictReader(detail_lines)
    return list(reader)


def summarize(events):
    """
    Aggregates raw test events into the stats useful for the dashboard:
    total tests, unique units, pass/fail counts, and per-station breakdown.
    """
    unique_units = set(e['Serial Number'] for e in events)
    passes = sum(1 for e in events if e.get('Test Result') == 'pass')
    fails = sum(1 for e in events if e.get('Test Result') == 'fail')
    station_counts = collections.Counter(e.get('Station ID', 'Unknown') for e in events)
    model_counts = collections.Counter(e.get('Model Number', 'Unknown') for e in events if e.get('Model Number'))

    return {
        "total_tests": len(events),
        "unique_units": len(unique_units),
        "pass": passes,
        "fail": fails,
        "yield_pct": round(100 * passes / len(events), 2) if events else 0,
        "by_station": station_counts.most_common(20),
        "by_model": model_counts.most_common(20),
    }


if __name__ == "__main__":
    import sys
    events = parse_report(sys.argv[1])
    print(summarize(events))


def extract_report_period(filepath):
    """
    Pulls the actual covered date/time range out of the filename itself,
    e.g. '..._20260731_100000-0500_to_20260731_110000-0500.csv' ->
    a human-readable string like '2026-07-31 10:00 to 2026-07-31 11:00'.
    Returns None if the filename doesn't match the expected pattern.
    """
    fname = os.path.basename(filepath)
    m = re.search(r'(\d{8})_(\d{6})-\d{4}_to_(\d{8})_(\d{6})-\d{4}', fname)
    if not m:
        return None
    start_date, start_time, end_date, end_time = m.groups()

    def fmt_date(d):
        return f"{d[0:4]}-{d[4:6]}-{d[6:8]}"

    def fmt_time(t):
        return f"{t[0:2]}:{t[2:4]}"

    start_str = f"{fmt_date(start_date)} {fmt_time(start_time)}"
    end_str = f"{fmt_date(end_date)} {fmt_time(end_time)}"
    if start_date == end_date:
        return f"{fmt_date(start_date)}, {fmt_time(start_time)}-{fmt_time(end_time)}"
    return f"{start_str} to {end_str}"


def find_latest_report(folder, report_type):
    """
    report_type: 'HPR', 'DPR', or 'WPR'
    Finds the most recently modified file matching that report type in the
    given folder (Power Automate saves these with timestamped names, e.g.
    '20260731090000_HPR_20260731_100000-0500_to_20260731_110000-0500.csv.zip').
    """
    pattern = os.path.join(folder, f"*_{report_type}_*")
    matches = glob.glob(pattern)
    if not matches:
        return None
    return max(matches, key=os.path.getmtime)


def tester_group(station_id, testers):
    """Matches a Station ID to one of the requested tester prefixes
    (e.g. 'HT91' -> 'HT', 'GTC083' -> 'GT', 'FTDS12' -> 'FT')."""
    for t in testers:
        if station_id.startswith(t):
            return t
    return None


def compute_fpy_fy(events, testers=("FT", "HT", "GT")):
    """
    First Pass Yield (FPY): % of units whose FIRST attempt at that tester
    passed. Final Yield (FY): % of units whose LAST (most recent) attempt
    passed - always >= FPY, since retries can convert a fail into a pass.

    Returns {tester: {"units": int, "fpy": float, "fy": float}}
    """
    unit_events = collections.defaultdict(list)
    for e in events:
        station_id = e.get("Station ID") or ""
        grp = tester_group(station_id, testers)
        if grp is None:
            continue
        sn = e.get("Serial Number")
        t = e.get("Test Start Time (local time)") or e.get("Test Start Time (Pacific time)") or ""
        result = (e.get("Test Result") or "").strip().lower()
        unit_events[(sn, grp)].append((t, result))

    per_group = {t: {"units": 0, "first_pass": 0, "final_pass": 0} for t in testers}
    for (sn, grp), evs in unit_events.items():
        evs.sort(key=lambda x: x[0])
        per_group[grp]["units"] += 1
        if evs[0][1] == "pass":
            per_group[grp]["first_pass"] += 1
        if evs[-1][1] == "pass":
            per_group[grp]["final_pass"] += 1

    result = {}
    for t in testers:
        d = per_group[t]
        units = d["units"]
        result[t] = {
            "units": units,
            "fpy": round(100 * d["first_pass"] / units, 2) if units else 0,
            "fy": round(100 * d["final_pass"] / units, 2) if units else 0,
        }
    return result
