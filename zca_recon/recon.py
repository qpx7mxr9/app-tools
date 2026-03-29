"""
zca_recon/recon.py

ZCA Reconciliation Engine
-------------------------
Reads a source export CSV, reconciles it against the tracking sheet,
writes status/date/package back to the sheet, and exports filtered CSVs.

Entry points (called from Excel via xlwings RunPython):
    run_reconciliation()  - Main flow: import CSV or skip, then offer exports
    export_update()       - Export rows flagged as Discrepancy or Partial
    export_add()          - Export rows flagged as Not Found in CSV
"""

import xlwings as xw
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime

# ── Sheet/column identifiers ──────────────────────────────────────────────────
SHEET_NAME  = "Common Area"
STATUS_HDR  = "Common Area Status"
EXT_HDR     = "Extension Number"
DATE_HDR    = "Common Area (Last Update)"
PKG_HDR     = "Common Area Package"
DATASRC_HDR = "Data Source"
DATAST_HDR  = "Data Status"
DASH_LABEL  = "ZP CA Last Update:"
AP          = "'"  # apostrophe used in desk phone column headers

# ── Export column map ─────────────────────────────────────────────────────────
# Each tuple: (output CSV header, source sheet column)
# Indices 9 & 10 (Phone Number, Outbound Caller ID) are filled dynamically
# based on the user's Zoom Temp vs Actual selection at export time.
EXPORT_COLS = [
    ("Display Name",                          "Display Name"),
    ("Package",                               PKG_HDR),
    ("Site Name",                             "Site Name"),
    ("Site Code",                             "Site Code"),
    ("Common Area Template",                  "Common Area Template"),
    ("Language",                              "Language"),
    ("Department",                            "Department"),
    ("Cost Center",                           "Cost Center"),
    ("Extension Number",                      "Extension Number"),
    ("Phone Number",                          None),   # resolved at export time
    ("Outbound Caller ID",                    None),   # resolved at export time
    ("Select Outbound Caller ID",             "Select Outbound Caller ID"),
    (f"Desk Phone 1{AP}s Brand",              f"Desk Phone 1{AP}s Brand"),
    (f"Desk Phone 1{AP}s Model",              f"Desk Phone 1{AP}s Model"),
    (f"Desk Phone 1{AP}s MAC Address",        f"Desk Phone 1{AP}s MAC Address"),
    (f"Desk Phone 1{AP}s Provision Template", f"Desk Phone 1{AP}s Provision Template"),
    (f"Desk Phone 2{AP}s Brand",              f"Desk Phone 2{AP}s Brand"),
    (f"Desk Phone 2{AP}s Model",              f"Desk Phone 2{AP}s Model"),
    (f"Desk Phone 2{AP}s MAC Address",        f"Desk Phone 2{AP}s MAC Address"),
    (f"Desk Phone 2{AP}s Provision Template", f"Desk Phone 2{AP}s Provision Template"),
    (f"Desk Phone 3{AP}s Brand",              f"Desk Phone 3{AP}s Brand"),
    (f"Desk Phone 3{AP}s Model",              f"Desk Phone 3{AP}s Model"),
    (f"Desk Phone 3{AP}s MAC Address",        f"Desk Phone 3{AP}s MAC Address"),
    (f"Desk Phone 3{AP}s Provision Template", f"Desk Phone 3{AP}s Provision Template"),
]


# ── Dialog helpers ────────────────────────────────────────────────────────────

def _root():
    """Create a hidden tkinter root window for dialog parenting."""
    r = tk.Tk()
    r.withdraw()
    r.lift()
    try:
        r.attributes("-topmost", True)
    except Exception:
        pass
    return r

def info(title, msg):
    r = _root(); messagebox.showinfo(title, msg, parent=r); r.destroy()

def ask_ok_cancel(title, msg):
    r = _root(); result = messagebox.askokcancel(title, msg, parent=r); r.destroy(); return result

def ask_yes_no(title, msg):
    r = _root(); result = messagebox.askyesno(title, msg, parent=r); r.destroy(); return result

def ask_yes_no_cancel(title, msg):
    """Returns True=Yes, False=No, None=Cancel."""
    r = _root(); result = messagebox.askyesnocancel(title, msg, parent=r); r.destroy(); return result

def pick_csv():
    """Open a file picker and return the selected CSV path, or empty string."""
    r = _root()
    path = filedialog.askopenfilename(
        parent=r, title="Select Source Export CSV",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
    r.destroy()
    return path or ""

def get_save_path(suggested, title):
    """Open a save-as dialog and return the chosen path, or empty string."""
    r = _root()
    path = filedialog.asksaveasfilename(
        parent=r, title=title, initialfile=suggested,
        defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    r.destroy()
    return path or ""


# ── Utilities ─────────────────────────────────────────────────────────────────

def strip_unwanted_packages(val):
    """
    Remove 'Zoom Meetings' entries from a comma-separated package string.
    e.g. 'Zoom Phone, Zoom Meetings' -> 'Zoom Phone'
    """
    if not val:
        return ""
    parts = [p.strip() for p in str(val).split(",")]
    return ", ".join(p for p in parts if "zoom meetings" not in p.lower())

def _get_sheet(wb):
    """Return the tracking worksheet, or None with an error dialog."""
    try:
        return wb.sheets[SHEET_NAME]
    except Exception:
        info("Error", f"Could not find '{SHEET_NAME}' sheet.")
        return None

def _read_df(ws):
    """
    Read the entire used range of ws into a DataFrame.
    DataFrame index = Excel row numbers (1-based) so we can write back directly.
    """
    data = ws.used_range.value
    if not data or len(data) < 2:
        return pd.DataFrame()
    headers = [str(h).strip() if h else f"_col{i}" for i, h in enumerate(data[0])]
    df = pd.DataFrame(data[1:], columns=headers)
    df.index = range(2, 2 + len(df))
    return df

def _write(ws, excel_row, headers, col_name, value):
    """Write value to (excel_row, col_name) if col_name exists in headers."""
    if col_name in headers:
        ws.range(excel_row, headers.index(col_name) + 1).value = value

def _color_status(ws, df):
    """
    Apply background color to the status column based on cell value:
      Complete -> green
      Setup    -> red/pink
      other    -> clear
    """
    if STATUS_HDR not in df.columns:
        return
    col_idx = list(df.columns).index(STATUS_HDR) + 1
    for row, val in zip(df.index, df[STATUS_HDR]):
        cell = ws.range(row, col_idx)
        val = str(val).strip() if val else ""
        if val == "Complete":
            cell.color = (198, 239, 206)
        elif val == "Setup":
            cell.color = (255, 199, 206)
        else:
            cell.color = None

def _stamp_dashboard(wb):
    """
    Write the current timestamp next to the DASH_LABEL cell on the Dashboard sheet.
    Falls back to J19 if the label cell is not found.
    """
    try:
        dash = wb.sheets["Dashboard"]
    except Exception:
        return
    now_str = datetime.now().strftime("%m-%d-%Y %H:%M")
    data = dash.used_range.value or []
    for r_idx, row in enumerate(data):
        if not row:
            continue
        for c_idx, cell in enumerate(row):
            if cell and str(cell).strip() == DASH_LABEL:
                dash.range(r_idx + 1, c_idx + 2).value = now_str
                return
    try:
        dash.range("J19").value = now_str
    except Exception:
        pass


# ── Public entry points (called from Excel) ───────────────────────────────────

def run_reconciliation():
    """
    Main entry point.
    1. Shows intro prompt.
    2. Asks whether to import a source CSV.
    3. Routes to _run_with_csv or _run_without_csv.
    """
    wb = xw.Book.caller()
    if not ask_ok_cancel(
        "CA Reconciliation",
        "COMMON AREA RECONCILIATION\n\n"
        "What you will need:\n"
        "  •  Source export CSV\n"
        "     (Admin Portal > Phone > Common Area Phones > Export)\n\n"
        "After reconciling you will be asked if you want to:\n"
        "  •  Export an UPDATE file (discrepancy/in-progress)\n"
        "  •  Export an ADD file (not yet provisioned)\n\n"
        "Click OK to continue, or Cancel to exit."
    ):
        return

    if ask_yes_no("CA Reconciliation",
                  "Do you want to import a source export CSV to reconcile against?"):
        _run_with_csv(wb)
    else:
        _run_without_csv(wb)


def export_update():
    """
    Standalone export: rows where status=Setup and data status is
    Discrepancy or Partial (i.e. found in source but fields don't match).
    """
    _export(xw.Book.caller(), "update")


def export_add():
    """
    Standalone export: rows where status=Setup and data status is
    Not Found in CSV (i.e. not present in source system at all).
    """
    _export(xw.Book.caller(), "add")


# ── Reconciliation logic ──────────────────────────────────────────────────────

def _run_with_csv(wb):
    """
    Import source CSV, compare each extension against the tracking sheet,
    write status/package/date back, color-code, stamp dashboard.

    Status logic:
      All key fields match              -> Complete  / Verified
      Display name or site name match   -> Setup     / Discrepancy
      No fields match                   -> Setup     / Partial
      Extension not in CSV              -> Setup     / Not Found in CSV
    """
    csv_path = pick_csv()
    if not csv_path:
        return

    try:
        df_csv = pd.read_csv(csv_path, dtype=str).fillna("")
        df_csv.columns = [c.strip() for c in df_csv.columns]
    except Exception as e:
        info("Error", f"Could not load CSV:\n{e}"); return

    if EXT_HDR not in df_csv.columns:
        info("Error", f"'{EXT_HDR}' not found in CSV.\n\nFound: {', '.join(df_csv.columns)}"); return

    # Build lookup keyed by lowercase extension for case-insensitive matching
    df_csv["_key"] = df_csv[EXT_HDR].str.strip().str.lower()
    lookup = df_csv.drop_duplicates("_key").set_index("_key")

    ws = _get_sheet(wb)
    if ws is None:
        return

    df = _read_df(ws)
    if df.empty:
        info("Error", "No data on tracking sheet."); return

    headers = list(df.columns)
    if EXT_HDR not in headers:
        info("Error", f"'{EXT_HDR}' not found on sheet."); return
    if STATUS_HDR not in headers:
        info("Error", f"'{STATUS_HDR}' not found on sheet."); return

    today = datetime.now().strftime("%m-%d-%Y %H:%M")
    cnt = dict(complete=0, disc=0, progress=0, incomplete=0)

    def sv(row, col):
        """Sheet value — normalized for comparison."""
        v = row.get(col, "") if col in row.index else ""
        return str(v).strip().lower()

    def cv(cr, col):
        """CSV value — normalized for comparison."""
        v = cr.get(col, "") if col in cr.index else ""
        return str(v).strip().lower()

    for excel_row, row in df.iterrows():
        ext_val = str(row.get(EXT_HDR, "")).strip()
        if not ext_val:
            continue

        key = ext_val.lower()

        if key in lookup.index:
            cr = lookup.loc[key]

            # Sync package from source CSV -> tracking sheet
            if PKG_HDR in headers and "Package" in cr.index:
                _write(ws, excel_row, headers, PKG_HDR,
                       strip_unwanted_packages(cr.get("Package", "")))

            _write(ws, excel_row, headers, DATASRC_HDR, "Source CSV")

            # Field comparisons
            disp  = sv(row, "Display Name")  == cv(cr, "Display Name")
            site  = sv(row, "Site Name")     == cv(cr, "Site Name")
            phone = sv(row, "Phone Number (Zoom Temp)") == cv(cr, "Phone Number")
            ocid  = sv(row, "Outbound Caller ID (Zoom Temp)") == cv(cr, "Outbound Caller ID")
            dp    = sv(row, f"Desk Phone 1{AP}s Brand") == cv(cr, f"Desk Phone 1{AP}s Brand")

            if disp and site and phone and ocid and dp:
                status = "Complete"
                _write(ws, excel_row, headers, DATAST_HDR, "Verified")
                cnt["complete"] += 1
            elif disp or site:
                status = "Setup"
                _write(ws, excel_row, headers, DATAST_HDR, "Discrepancy")
                cnt["disc"] += 1
            else:
                status = "Setup"
                _write(ws, excel_row, headers, DATAST_HDR, "Partial")
                cnt["progress"] += 1
        else:
            status = "Setup"
            _write(ws, excel_row, headers, DATASRC_HDR, "Sheet Only")
            _write(ws, excel_row, headers, DATAST_HDR,  "Not Found in CSV")
            cnt["incomplete"] += 1

        _write(ws, excel_row, headers, STATUS_HDR, status)
        _write(ws, excel_row, headers, DATE_HDR,   today)

    _color_status(ws, _read_df(ws))
    _stamp_dashboard(wb)

    info("CA Reconciliation",
         f"Reconciliation complete.\n\n"
         f"  Complete:     {cnt['complete']}\n"
         f"  Discrepancy:  {cnt['disc']}\n"
         f"  In Progress:  {cnt['progress']}\n"
         f"  Incomplete:   {cnt['incomplete']}")

    if ask_yes_no("CA Reconciliation", "Export UPDATE file (discrepancy/in-progress)?"):
        _export(wb, "update")
    if ask_yes_no("CA Reconciliation", "Export ADD file (not yet provisioned)?"):
        _export(wb, "add")


def _run_without_csv(wb):
    """
    No CSV import path.
    If no statuses exist yet: offer to mark all rows as Setup Incomplete
    and export an Add file.
    If statuses already exist: skip to export prompts.
    """
    ws = _get_sheet(wb)
    if ws is None:
        return

    df = _read_df(ws)
    headers = list(df.columns)

    if STATUS_HDR not in headers:
        info("Error", f"'{STATUS_HDR}' not found on sheet."); return

    statuses = df[STATUS_HDR].dropna().astype(str).str.strip()
    count = (statuses != "").sum()

    if count == 0:
        info("CA Reconciliation", "No statuses found. No rows have been reconciled yet.")
        ans = ask_yes_no_cancel(
            "CA Reconciliation",
            "Mark all rows as Setup Incomplete and export an Add file?")
        if ans is None:
            return
        if ans:
            today = datetime.now().strftime("%m-%d-%Y %H:%M")
            for excel_row, row in df.iterrows():
                if str(row.iloc[0]).strip():
                    _write(ws, excel_row, headers, STATUS_HDR,  "Setup")
                    _write(ws, excel_row, headers, DATE_HDR,    today)
                    _write(ws, excel_row, headers, DATASRC_HDR, "Manual")
                    _write(ws, excel_row, headers, DATAST_HDR,  "Not Found in CSV")
            _color_status(ws, _read_df(ws))
            _stamp_dashboard(wb)
            _export(wb, "add")
    else:
        info("CA Reconciliation",
             f"{count} row(s) already have statuses. Skipping import.")
        if ask_yes_no("CA Reconciliation", "Export UPDATE file?"):
            _export(wb, "update")
        if ask_yes_no("CA Reconciliation", "Export ADD file?"):
            _export(wb, "add")


def _export(wb, export_type):
    """
    Build and save a filtered CSV.

    export_type='update': rows where status=Setup AND data status in
                          [Discrepancy, Partial] — exists in source but wrong
    export_type='add':    rows where status=Setup AND data status =
                          Not Found in CSV — not in source system yet

    If no Data Status column exists, all Setup rows are included in both exports.
    User is asked whether to use Zoom Temp phone numbers or actual numbers.
    """
    ws = _get_sheet(wb)
    if ws is None:
        return

    df = _read_df(ws)
    headers = list(df.columns)

    if STATUS_HDR not in headers:
        info("Error", f"'{STATUS_HDR}' not found on sheet."); return

    use_temp = ask_yes_no(
        "Phone Number Source",
        "Which phone numbers should the export use?\n\n"
        "  Yes  = Zoom Temp Numbers\n"
        "  No   = Actual Numbers"
    )

    date_str = datetime.now().strftime("%Y%m%d")
    if export_type == "update":
        suggested, title = f"CA_Update_{date_str}.csv", "Save Update CSV"
    else:
        suggested, title = f"CA_Add_{date_str}.csv", "Save Add CSV"

    save_path = get_save_path(suggested, title)
    if not save_path:
        return

    # Resolve phone/OCID source columns based on user selection
    cols = list(EXPORT_COLS)
    cols[9]  = ("Phone Number",       "Phone Number (Zoom Temp)" if use_temp else "Phone Number")
    cols[10] = ("Outbound Caller ID", "Outbound Caller ID (Zoom Temp)" if use_temp else "Outbound Caller ID")

    # Row filter
    s = df[STATUS_HDR].astype(str).str.strip()
    d = df.get(DATAST_HDR, pd.Series([""] * len(df), index=df.index)).astype(str).str.strip()

    if DATAST_HDR not in headers:
        mask = s == "Setup"
    elif export_type == "update":
        mask = (s == "Setup") & d.isin(["Discrepancy", "Partial"])
    else:
        mask = (s == "Setup") & (d == "Not Found in CSV")

    filtered = df[mask]

    if filtered.empty:
        info("Nothing to Export",
             "No matching rows found for this export type."); return

    # Build output rows
    out_rows = []
    for _, row in filtered.iterrows():
        out_row = {}
        for export_hdr, src_hdr in cols:
            val = ""
            if src_hdr and src_hdr in headers:
                val = str(row.get(src_hdr, "")).strip()
            if export_hdr == "Package" and not val and "Package" in headers:
                val = str(row.get("Package", "")).strip()
            if export_hdr == "Package":
                val = strip_unwanted_packages(val)
            out_row[export_hdr] = val
        out_rows.append(out_row)

    out_df = pd.DataFrame(out_rows, columns=[h for h, _ in cols])
    out_df.to_csv(save_path, index=False)
    info("Export Complete", f"{len(out_df)} row(s) exported.\n{save_path}")
