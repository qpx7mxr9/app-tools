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
from datetime import datetime
from . import dialogs as dlg

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


# ── Workbook resolver ─────────────────────────────────────────────────────────

def _get_wb():
    """
    Get the calling workbook. Works via RunPython or Application.Run.
    Falls back to finding the first open non-addin workbook.
    """
    try:
        return xw.Book.caller()
    except Exception:
        pass
    try:
        for app in xw.apps:
            for book in app.books:
                if not book.name.endswith(('.xlam', '.xla')):
                    return book
    except Exception:
        pass
    return None


# ── Utilities ─────────────────────────────────────────────────────────────────

def strip_unwanted_packages(val):
    if not val:
        return ""
    parts = [p.strip() for p in str(val).split(",")]
    return ", ".join(p for p in parts if "zoom meetings" not in p.lower())


def _get_sheet(wb):
    try:
        return wb.sheets[SHEET_NAME]
    except Exception:
        dlg.info("Error", f"Could not find '{SHEET_NAME}' sheet.")
        return None


def _read_df(ws):
    data = ws.used_range.value
    if not data or len(data) < 2:
        return pd.DataFrame()
    headers = [str(h).strip() if h else f"_col{i}" for i, h in enumerate(data[0])]
    df = pd.DataFrame(data[1:], columns=headers)
    df.index = range(2, 2 + len(df))
    return df


def _write(ws, excel_row, headers, col_name, value):
    """Write value to cell — uses tuple syntax for Mac compatibility."""
    if col_name in headers:
        col_idx = headers.index(col_name) + 1
        try:
            ws.range((excel_row, col_idx)).value = value
        except Exception as e:
            _log(f"Write failed row={excel_row} col={col_name}({col_idx}): {e}")


LOG_PATH = "/tmp/zca_recon.log"

def _log(msg):
    try:
        with open(LOG_PATH, "a") as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
    except Exception:
        pass


def _color_status(ws, df):
    """Color-code the status column: Complete=green, Setup=red."""
    if STATUS_HDR not in df.columns:
        return
    col_idx = list(df.columns).index(STATUS_HDR) + 1
    for row, val in zip(df.index, df[STATUS_HDR]):
        cell = ws.range((row, col_idx))
        val = str(val).strip() if val else ""
        if val == "Complete":
            cell.color = (198, 239, 206)
        elif val == "Setup":
            cell.color = (255, 199, 206)
        else:
            cell.color = None


def _stamp_dashboard(wb):
    """Write last-run timestamp to the Dashboard sheet."""
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
                dash.range((r_idx + 1, c_idx + 2)).value = now_str
                return
    try:
        dash.range("J19").value = now_str
    except Exception:
        pass


# ── Public entry points ───────────────────────────────────────────────────────

def run_reconciliation():
    import os
    open(LOG_PATH, "w").close()
    wb = _get_wb()
    _log(f"wb={wb.name if wb else 'None'}")
    if not wb:
        dlg.info("Error", "Could not find open workbook."); return
    action = dlg.show_intro()
    _log(f"action={action}")
    if action is None:
        return
    if action == "import":
        _run_with_csv(wb)
    else:
        _run_without_csv(wb)


def export_update():
    wb = _get_wb()
    if not wb:
        dlg.info("Error", "Could not find open workbook."); return
    _export(wb, "update")


def export_add():
    wb = _get_wb()
    if not wb:
        dlg.info("Error", "Could not find open workbook."); return
    _export(wb, "add")


# ── Reconciliation logic ──────────────────────────────────────────────────────

def _run_with_csv(wb):
    _log("_run_with_csv start")
    csv_path = dlg.pick_csv()
    _log(f"csv_path={csv_path}")
    if not csv_path:
        return

    try:
        df_csv = pd.read_csv(csv_path, dtype=str).fillna("")
        df_csv.columns = [c.strip() for c in df_csv.columns]
    except Exception as e:
        dlg.info("Error", f"Could not load CSV:\n{e}"); return

    if EXT_HDR not in df_csv.columns:
        dlg.info("Error", f"'{EXT_HDR}' not found in CSV.\n\nFound: {', '.join(df_csv.columns)}"); return

    def _norm_ext(val):
        """Normalize extension: strip whitespace, drop .0 from Excel floats."""
        s = str(val).strip()
        try:
            s = str(int(float(s)))   # "1001.0" -> "1001"
        except (ValueError, TypeError):
            pass
        return s.lower()

    df_csv["_key"] = df_csv[EXT_HDR].apply(_norm_ext)
    lookup = df_csv.drop_duplicates("_key").set_index("_key")

    ws = _get_sheet(wb)
    if ws is None:
        return

    df = _read_df(ws)
    if df.empty:
        dlg.info("Error", "No data on tracking sheet."); return

    headers = list(df.columns)
    if EXT_HDR not in headers:
        dlg.info("Error", f"'{EXT_HDR}' not found on sheet."); return
    if STATUS_HDR not in headers:
        dlg.info("Error", f"'{STATUS_HDR}' not found on sheet."); return

    today = datetime.now().strftime("%m-%d-%Y %H:%M")
    cnt = dict(complete=0, disc=0, progress=0, incomplete=0)
    _log(f"Sheet rows={len(df)}  CSV rows={len(df_csv)}  Headers={headers[:5]}...")

    def sv(row, col):
        v = row.get(col, "") if col in row.index else ""
        return str(v).strip().lower()

    def cv(cr, col):
        v = cr.get(col, "") if col in cr.index else ""
        return str(v).strip().lower()

    # Log a sample so we can verify key format matches between sheet and CSV
    sample_keys = list(lookup.index[:5])
    _log(f"CSV sample keys: {sample_keys}")

    for excel_row, row in df.iterrows():
        raw_ext = row.get(EXT_HDR, "")
        ext_val = _norm_ext(raw_ext)
        if not ext_val:
            continue

        key = ext_val

        if key in lookup.index:
            cr = lookup.loc[key]

            if PKG_HDR in headers and "Package" in cr.index:
                _write(ws, excel_row, headers, PKG_HDR,
                       strip_unwanted_packages(cr.get("Package", "")))

            _write(ws, excel_row, headers, DATASRC_HDR, "Source CSV")

            # Compare identifying fields (sheet vs CSV)
            disp = sv(row, "Display Name") == cv(cr, "Display Name")
            site = sv(row, "Site Name")    == cv(cr, "Site Name")

            # Phone / OCID: compare actual columns; treat both-empty as matching
            s_ph, c_ph   = sv(row, "Phone Number"),       cv(cr, "Phone Number")
            s_oc, c_oc   = sv(row, "Outbound Caller ID"), cv(cr, "Outbound Caller ID")
            phone = (not s_ph and not c_ph) or (s_ph == c_ph)
            ocid  = (not s_oc and not c_oc) or (s_oc == c_oc)

            dp = sv(row, f"Desk Phone 1{AP}s Brand") == cv(cr, f"Desk Phone 1{AP}s Brand")

            if disp and site and phone and ocid and dp:
                # All key fields match
                status = "Complete"
                _write(ws, excel_row, headers, DATAST_HDR, "Verified")
                cnt["complete"] += 1
            elif disp and site:
                # Name + site match; phone / desk phone details differ → Partial
                status = "Setup"
                _write(ws, excel_row, headers, DATAST_HDR, "Partial")
                cnt["progress"] += 1
            else:
                # Display name or site doesn't match → real discrepancy
                status = "Setup"
                _write(ws, excel_row, headers, DATAST_HDR, "Discrepancy")
                cnt["disc"] += 1
        else:
            status = "Setup"
            _write(ws, excel_row, headers, DATASRC_HDR, "Sheet Only")
            _write(ws, excel_row, headers, DATAST_HDR,  "Not Found in CSV")
            cnt["incomplete"] += 1

        _write(ws, excel_row, headers, STATUS_HDR, status)
        _write(ws, excel_row, headers, DATE_HDR,   today)

    _log(f"Counts: {cnt}")
    _color_status(ws, _read_df(ws))
    _stamp_dashboard(wb)

    exports = dlg.show_results(cnt)
    if "update" in exports:
        _export(wb, "update")
    if "add" in exports:
        _export(wb, "add")


def _run_without_csv(wb):
    _log("_run_without_csv start")
    ws = _get_sheet(wb)
    if ws is None:
        _log("sheet not found"); return

    df = _read_df(ws)
    headers = list(df.columns)
    _log(f"rows={len(df)}  STATUS_HDR in headers={STATUS_HDR in headers}")

    if STATUS_HDR not in headers:
        dlg.info("Error", f"'{STATUS_HDR}' not found on sheet."); return

    statuses = df[STATUS_HDR].dropna().astype(str).str.strip()
    count = (statuses != "").sum()
    _log(f"existing status count={count}")

    if count == 0:
        exports = dlg.show_results({"complete": 0, "disc": 0, "progress": 0, "incomplete": len(df)})
        if "add" in exports:
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
        try:
            df2 = _read_df(ws)
            _log(f"df2 rows={len(df2)}")
            empty_s = pd.Series([""] * len(df2), index=df2.index)
            s = (df2[STATUS_HDR]  if STATUS_HDR  in df2.columns else empty_s).astype(str).str.strip()
            d = (df2[DATAST_HDR]  if DATAST_HDR  in df2.columns else empty_s).astype(str).str.strip()
            cnt = dict(
                complete   = int((s == "Complete").sum()),
                disc       = int((d == "Discrepancy").sum()),
                progress   = int((d == "Partial").sum()),
                incomplete = int((d == "Not Found in CSV").sum()),
            )
            _log(f"cnt={cnt}")
            exports = dlg.show_results(cnt)
            _log(f"exports={exports}")
            if "update" in exports:
                _export(wb, "update")
            if "add" in exports:
                _export(wb, "add")
        except Exception as e:
            import traceback
            _log(f"ERROR in skip path: {e}\n{traceback.format_exc()}")
            dlg.info("Error", str(e))


def _export(wb, export_type):
    _log(f"_export start type={export_type}")
    ws = _get_sheet(wb)
    if ws is None:
        _log("sheet not found in export"); return

    df = _read_df(ws)
    headers = list(df.columns)

    if STATUS_HDR not in headers:
        dlg.info("Error", f"'{STATUS_HDR}' not found on sheet."); return

    phone_choice = dlg.ask_phone_source()
    _log(f"phone_choice={phone_choice}")
    if phone_choice is None:
        return
    use_temp = phone_choice == "temp"

    date_str = datetime.now().strftime("%Y%m%d")
    if export_type == "update":
        suggested, title = f"CA_Update_{date_str}.csv", "Save Update CSV"
    else:
        suggested, title = f"CA_Add_{date_str}.csv",    "Save Add CSV"

    save_path = dlg.get_save_path(suggested, title)
    _log(f"save_path={save_path}")
    if not save_path:
        return

    cols = list(EXPORT_COLS)
    cols[9]  = ("Phone Number",       "Phone Number (Zoom Temp)" if use_temp else "Phone Number")
    cols[10] = ("Outbound Caller ID", "Outbound Caller ID (Zoom Temp)" if use_temp else "Outbound Caller ID")

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
        dlg.info("Nothing to Export", "No matching rows found for this export type."); return

    def _cell(row, col_name):
        """Return cell value as a clean string; blank for None/NaN."""
        if col_name not in headers:
            return ""
        raw = row.get(col_name, "")
        try:
            if pd.isna(raw):
                return ""
        except TypeError:
            pass
        return str(raw).strip()

    out_rows = []
    for _, row in filtered.iterrows():
        out_row = {}
        for export_hdr, src_hdr in cols:
            val = _cell(row, src_hdr) if src_hdr else ""
            # Package: prefer Common Area Package, fall back to column-B "Package"
            if export_hdr == "Package" and not val:
                val = _cell(row, "Package")
            if export_hdr == "Package":
                val = strip_unwanted_packages(val)
            out_row[export_hdr] = val
        out_rows.append(out_row)

    out_df = pd.DataFrame(out_rows, columns=[h for h, _ in cols])
    _log(f"filtered={len(filtered)} out_rows={len(out_df)}")
    try:
        out_df.to_csv(save_path, index=False)
        _log("CSV saved ok")
        dlg.info("Export Complete", f"{len(out_df)} row(s) exported.\n{save_path}")
    except Exception as e:
        _log(f"CSV save error: {e}")
        dlg.info("Export Error", str(e))
