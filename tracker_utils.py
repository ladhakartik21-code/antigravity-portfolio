"""
Shared output path router & Validation & Lessons auto-summarizer for AntiGravity Wealth.
Ensures every generated test Excel uses YYYY-MM-DD date-first naming inside structured
subfolders under '2. Output Generated\\Test Analysis\\' and logs a concise benchmark row
into 'AnitGravity Wealth.xlsx' -> 'Validation & Lessons'.
"""

import os
import shutil
import datetime
from typing import Optional, Dict, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

APP_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_ROOT = os.path.dirname(os.path.dirname(APP_DIR))

CATEGORY_FOLDERS = {
    "intraday": "Intraday_SR_Tests",
    "breakout": "MultiYear_Breakout_Tests",
}


def get_ist_now() -> datetime.datetime:
    utc = datetime.datetime.now(datetime.timezone.utc)
    return utc + datetime.timedelta(hours=5, minutes=30)


def _find_output_base() -> Optional[str]:
    candidates = [
        os.path.join(WORKSPACE_ROOT, "2. Output Generated"),
        r"C:\Users\KARTIK\Desktop\AntiGravity\2. Output Generated",
    ]
    for p in candidates:
        if os.path.isdir(p):
            return p
    return None


def resolve_test_output_path(
    report_slug: str,
    category: str = "intraday",
    dt: Optional[datetime.datetime] = None,
) -> str:
    """
    Returns an absolute date-first file path inside:
    2. Output Generated/Test Analysis/<Category_Subfolder>/YYYY-MM-DD_HHMMSS_<slug>.xlsx
    Falls back cleanly to APP_DIR on Render/Linux cloud containers.
    """
    ts = dt or get_ist_now()
    safe_slug = report_slug.strip().replace(" ", "_")
    filename = f"{ts.strftime('%Y-%m-%d_%H%M%S')}_{safe_slug}.xlsx"

    out_base = _find_output_base()
    if out_base:
        test_root = None
        for root_name in ("Test Analysis", "Test result", "Test Results"):
            cand = os.path.join(out_base, root_name)
            if os.path.isdir(cand):
                test_root = cand
                break
        if not test_root:
            test_root = os.path.join(out_base, "Test Analysis")
            os.makedirs(test_root, exist_ok=True)

        subfolder = CATEGORY_FOLDERS.get(category.lower(), category)
        dest_dir = os.path.join(test_root, subfolder)
        os.makedirs(dest_dir, exist_ok=True)
        return os.path.join(dest_dir, filename)

    return os.path.join(APP_DIR, filename)


def _estimate_row_height(values, col_widths, min_h: int = 44) -> int:
    max_lines = 1
    for val, w in zip(values, col_widths):
        if val is None:
            continue
        eff_w = max(8, w - 3)
        lines = 0
        for part in str(val).split("\n"):
            lines += max(1, (len(part) + int(eff_w) - 1) // int(eff_w))
        if lines > max_lines:
            max_lines = lines
    return max(min_h, max_lines * 15 + 8)


def _sync_fallback_copies(master_path: str):
    web_dir = os.path.dirname(APP_DIR)
    for dst in (
        os.path.join(APP_DIR, "AnitGravity Wealth.xlsx"),
        os.path.join(web_dir, "cloud", "AnitGravity Wealth.xlsx"),
    ):
        if os.path.isdir(os.path.dirname(dst)):
            try:
                shutil.copy2(master_path, dst)
            except Exception:
                pass


def record_test_to_validation_sheet(
    experiment: str,
    date_regime: str,
    sample_rules: str,
    measured_outcome: str,
    finding: str,
    rule_upgrade: str,
    test_filepath: Optional[str] = None,
    dedup_key: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Appends or updates a benchmark row in 'AnitGravity Wealth.xlsx' -> 'Validation & Lessons'.
    Safely skips on cloud containers or if user currently holds Excel open.
    """
    out_base = _find_output_base()
    if not out_base:
        return {"updated": False, "reason": "Not running inside local Windows workspace"}

    tracker_path = os.path.join(out_base, "AnitGravity Wealth.xlsx")
    if not os.path.exists(tracker_path):
        return {"updated": False, "reason": f"Tracker not found at {tracker_path}"}

    # Verify non-locked writeability before touching anything
    try:
        with open(tracker_path, "a+b"):
            pass
    except PermissionError:
        return {"updated": False, "reason": "Tracker Excel is currently open in Microsoft Excel"}

    # Timestamped safety backup before automated write
    backup_dir = os.path.join(os.path.dirname(APP_DIR), "backups")
    os.makedirs(backup_dir, exist_ok=True)
    bk_name = f"AnitGravity_Wealth_auto_{get_ist_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    shutil.copy2(tracker_path, os.path.join(backup_dir, bk_name))

    wb = openpyxl.load_workbook(tracker_path)
    if "Validation & Lessons" not in wb.sheetnames:
        return {"updated": False, "reason": "'Validation & Lessons' sheet missing"}

    ws = wb["Validation & Lessons"]

    # Attach short file reference into measured_outcome if not already mentioned
    if test_filepath:
        fname = os.path.basename(test_filepath)
        if fname not in measured_outcome:
            measured_outcome = f"{measured_outcome}\nReport: {fname}"

    # Check if dedup_key matches an existing row to update rather than duplicate same-day exports
    target_row = None
    next_num = 1
    for r in range(5, ws.max_row + 1):
        val_id = str(ws.cell(r, 1).value or "").strip()
        exp_val = str(ws.cell(r, 2).value or "").strip()
        if not val_id and not exp_val:
            if target_row is None:
                target_row = r
            continue
        if val_id.upper().startswith("VAL-"):
            try:
                num = int(val_id.split("-")[1])
                next_num = max(next_num, num + 1)
            except Exception:
                pass
        if dedup_key and dedup_key.lower() in f"{exp_val} {ws.cell(r, 3).value or ''}".lower():
            target_row = r
            audit_id = val_id or f"VAL-{next_num:02d}"
            break
    else:
        if target_row is None:
            target_row = ws.max_row + 1
        audit_id = f"VAL-{next_num:02d}"

    # Styling aligned cleanly with existing sheet style
    thin_side = Side(style="thin", color="CBD5E1")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    fill_row = PatternFill(
        start_color="FFFFFF" if target_row % 2 == 1 else "F8FAFC",
        end_color="FFFFFF" if target_row % 2 == 1 else "F8FAFC",
        fill_type="solid",
    )
    font_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_reg = Font(name="Calibri", size=10, color="1E293B")

    row_values = [audit_id, experiment, date_regime, sample_rules, measured_outcome, finding, rule_upgrade]
    col_widths = [12, 30, 18, 28, 34, 42, 44]

    for c_idx, val in enumerate(row_values, 1):
        cell = ws.cell(target_row, c_idx, val)
        cell.border = border_cell
        cell.fill = fill_row
        cell.font = font_bold if c_idx in (1, 2) else font_reg
        cell.alignment = Alignment(
            horizontal="center" if c_idx in (1, 3) else "left",
            vertical="center",
            wrap_text=True,
        )

    if test_filepath and os.path.exists(test_filepath):
        ws.cell(target_row, 1).hyperlink = os.path.abspath(test_filepath)

    ws.row_dimensions[target_row].height = _estimate_row_height(row_values, col_widths, min_h=45)
    wb.save(tracker_path)

    # Preserve Excel formula cache across Executive Summary & Journey via COM
    try:
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb_com = excel.Workbooks.Open(os.path.abspath(tracker_path))
            excel.CalculateFull()
            wb_com.Save()
            wb_com.Close()
        finally:
            excel.Quit()
    except Exception:
        pass

    _sync_fallback_copies(tracker_path)
    return {"updated": True, "audit_id": audit_id, "row": target_row, "tracker_path": tracker_path}
