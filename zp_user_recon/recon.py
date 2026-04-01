"""
zp_user_recon/recon.py

Zoom Phone User Reconciliation
--------------------------------
Compares the Users sheet (DGW) against a Zoom Phone Users CSV export and
writes a status to each row's "ZP User Status" column.

Status values:
    Setup Complete      - All fields match
    Setup in Progress   - Email found in CSV but some fields differ
    Setup Discrepancy   - Device/MAC mismatch that needs investigation
    Setup Incomplete    - Email not found in Zoom Phone CSV

DGW columns used:
    Email | First Name | Last Name | Package | Site Code | Site Name |
    Extension Number | Phone Number (Zoom Temp) | Outbound Caller ID (Zoom Temp) |
    ZP User Status | Desk Phone 1's Brand/Model/MAC Address/Provision Template

Zoom Phone CSV columns used:
    Email | Package | Site Code | Site Name | Extension Number |
    Phone Number | Outbound Caller ID |
    Desk Phone 1's Brand/Model/MAC Address/Provision Template

Entry point (called from Excel via xlwings RunPython):
    run_zp_reconciliation()
"""

import xlwings as xw
# pandas imported lazily inside functions — keeps startup fast
from datetime import datetime
from zca_recon import dialogs as dlg

# ── Constants ─────────────────────────────────────────────────────────────────
USERS_SHEET     = "Users"
H_EMAIL         = "Email"
H_FIRST         = "First Name"
H_LAST          = "Last Name"
H_PACKAGE       = "Package"
H_SITE_CODE     = "Site Code"
H_SITE_NAME     = "Site Name"
H_EXT           = "Extension Number"
H_PHONE_TEMP    = "Phone Number (Zoom Temp)"
H_OUTBOUND_TEMP = "Outbound Caller ID (Zoom Temp)"
H_STATUS        = "ZP User Status"
H_DP1_BRAND     = "Desk Phone 1's Brand"
H_DP1_MODEL     = "Desk Phone 1's Model"
H_DP1_MAC       = "Desk Phone 1's MAC Address"
H_DP1_PROV      = "Desk Phone 1's Provision Template"

C_EMAIL   = "Email"
C_PACKAGE = "Package"
C_SITE_CODE = "Site Code"
C_SITE_NAME = "Site Name"
C_EXT     = "Extension Number"
C_PHONE   = "Phone Number"
C_OUTBOUND = "Outbound Caller ID"
C_DP1_BRAND = "Desk Phone 1's Brand"
C_DP1_MODEL = "Desk Phone 1's Model"
C_DP1_MAC   = "Desk Phone 1's MAC Address"
C_DP1_PROV  = "Desk Phone 1's Provision Template"

STATUS_COMPLETE   = "Setup Complete"
STATUS_PROGRESS   = "Setup in Progress"
STATUS_DISCREP    = "Setup Discrepancy"
STATUS_INCOMPLETE = "Setup Incomplete"

H_ZP_PACKAGE  = "ZP User Package"   # written on Setup Complete with CSV package value
H_CUSTOMER_DATA_STATUS = "Customer Data Status"   # Initial, Removed, Addition, Change
H_TCS_DATA_STATUS      = "TCS Data Status"         # Approved, Removed — if Removed, skip exports

BRAND_WORKPLACE_APP = "workplace app"

# Actual (non-temp) phone columns on the sheet
H_PHONE    = "Phone Number"
H_OUTBOUND = "Outbound Caller ID"

CHANGES_HDR    = "ZP Changes"
DASH_LABEL     = "ZP Recon Last Update:"
MISMATCH_COLOR = (255, 175, 100)   # orange — cell value differs from CSV

_STATUS_COLORS = {
    STATUS_COMPLETE:   (198, 239, 206),   # green
    STATUS_PROGRESS:   (255, 235, 156),   # yellow
    STATUS_DISCREP:    (255, 199, 206),   # red/pink
    STATUS_INCOMPLETE: (252, 228, 214),   # light orange
}

import tempfile as _tempfile, os as _os
LOG_PATH = _os.path.join(_tempfile.gettempdir(), "zp_user_recon.log")


# ── Workbook / sheet helpers ──────────────────────────────────────────────────

def _get_wb():
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


def _get_sheet(wb, name):
    try:
        return wb.sheets[name]
    except Exception:
        dlg.info("Error", f"Could not find '{name}' sheet.")
        return None


def _read_df(ws):
    import pandas as pd
    data = ws.used_range.value
    if not data or len(data) < 2:
        return pd.DataFrame()
    headers = [str(h).strip() if h else f"_col{i}" for i, h in enumerate(data[0])]
    df = pd.DataFrame(data[1:], columns=headers)
    df.index = range(2, 2 + len(df))
    return df


def _find_col(headers, name):
    """Return 1-based column index matching name (case-insensitive), or 0."""
    for i, h in enumerate(headers):
        if str(h).strip().lower() == name.lower():
            return i + 1
    return 0


# ── Normalization helpers ─────────────────────────────────────────────────────

def _norm_email(v):
    s = str(v or "").strip().lower()
    return s.replace("mailto:", "").replace("\xa0", "").strip()


def _is_blank_or_na(v):
    s = str(v or "").strip().lower()
    return s in ("", "na", "n/a", "no temp number specified")


def _digits_only(s):
    return "".join(c for c in str(s or "") if c.isdigit())


def _norm_ext(v):
    """Normalize extension number: strip whitespace, drop trailing .0."""
    s = str(v or "").strip()
    try:
        s = str(int(float(s)))
    except (ValueError, TypeError):
        pass
    return s


def _text_equal(a, b):
    return str(a or "").strip().lower() == str(b or "").strip().lower()


def _ext_equal(a, b):
    return _norm_ext(a) == _norm_ext(b)


def _phone_equal_na_ok(dgw_val, csv_val):
    d_blank = _is_blank_or_na(dgw_val)
    c_blank = _is_blank_or_na(csv_val)
    if d_blank and c_blank:
        return True
    if d_blank != c_blank:
        return False
    d = _digits_only(str(dgw_val))
    c = _digits_only(str(csv_val))
    if len(d) > 10:
        d = d[-10:]
    if len(c) > 10:
        c = c[-10:]
    return d == c


def _norm_mac(v):
    return str(v or "").upper().replace(":", "").replace("-", "").replace(".", "").strip()


# ── Device comparison ─────────────────────────────────────────────────────────

def _device_compare(d_brand, d_model, d_mac, d_prov,
                    z_brand, z_model, z_mac, z_prov):
    """
    Compare Desk Phone 1 fields between DGW and Zoom CSV.

    Returns (is_match: bool, is_discrepancy: bool).

    Rules:
    - If both sides are blank/N-A: match, no discrepancy
    - If either side is "workplace app": match, no discrepancy
    - MAC mismatch (when both non-blank): discrepancy
    - Brand/model mismatch: not a match, but not a discrepancy
    """
    db = str(d_brand or "").strip().lower()
    zb = str(z_brand or "").strip().lower()
    dm = _norm_mac(d_mac)
    zm = _norm_mac(z_mac)

    db_blank = not db or db in ("", "na", "n/a")
    zb_blank = not zb or zb in ("", "na", "n/a")
    dm_blank = not dm or dm in ("", "NA", "N/A")
    zm_blank = not zm or zm in ("", "NA", "N/A")

    # Both sides entirely blank → match
    if db_blank and zb_blank and dm_blank and zm_blank:
        return True, False

    # Workplace App on either side → match
    if BRAND_WORKPLACE_APP in db or BRAND_WORKPLACE_APP in zb:
        return True, False

    # MAC present on both sides and different → discrepancy
    if not dm_blank and not zm_blank and dm != zm:
        return False, True

    # Compare brand + model
    brand_match = (db == zb) or db_blank or zb_blank
    model_match = _text_equal(d_model, z_model)

    if brand_match and model_match:
        return True, False

    return False, False


def _is_softphone_match(d_brand, d_model, z_brand, z_model, z_mac):
    """
    Returns True if the DGW row is a Softphone user (Brand=Zoom, Model=Softphone)
    and the CSV shows no device assigned (blank brand, model, and MAC).
    These users are complete — no physical phone needed.
    """
    db = str(d_brand or "").strip().lower()
    dm = str(d_model or "").strip().lower()
    if db != "zoom" or dm != "softphone":
        return False
    zb = str(z_brand or "").strip()
    zm = str(z_model or "").strip()
    zmc = _norm_mac(z_mac)
    csv_no_device = (not zb and not zm and
                     (not zmc or zmc.lower() in ("", "na", "n/a")))
    return csv_no_device


# ── Logging ───────────────────────────────────────────────────────────────────

def _log(msg):
    try:
        with open(LOG_PATH, "a") as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
    except Exception:
        pass


# ── Sheet write helper ────────────────────────────────────────────────────────

def _write(ws, excel_row, headers, col_name, value):
    """Write value to a named column — skips silently if column not on sheet."""
    if col_name in headers:
        col_idx = headers.index(col_name) + 1
        try:
            ws.range((excel_row, col_idx)).value = value
        except Exception as e:
            _log(f"Write failed row={excel_row} col={col_name}: {e}")


# ── Color helpers ─────────────────────────────────────────────────────────────

def _apply_colors(ws, df, headers):
    """Color-code the ZP User Status column."""
    if H_STATUS not in headers:
        return
    col_idx = headers.index(H_STATUS) + 1
    for excel_row, row in df.iterrows():
        val = str(row.get(H_STATUS, "") or "").strip()
        cell = ws.range((excel_row, col_idx))
        cell.color = _STATUS_COLORS.get(val, None)


def _highlight_mismatches(ws, excel_row, headers, mismatch_cols, compare_cols):
    """Highlight mismatched cells orange; only clear our orange on now-matching cells."""
    for col_name in compare_cols:
        if col_name not in headers:
            continue
        col_idx = headers.index(col_name) + 1
        cell = ws.range((excel_row, col_idx))
        if col_name in mismatch_cols:
            cell.color = MISMATCH_COLOR
        elif cell.color == MISMATCH_COLOR:
            cell.color = None   # only remove our color, not user highlights


def _clear_mismatch_highlights(ws, df, headers, compare_cols):
    """Remove only our orange mismatch highlights — leave other cell colors untouched."""
    for col_name in compare_cols:
        if col_name not in headers:
            continue
        col_idx = headers.index(col_name) + 1
        for excel_row in df.index:
            cell = ws.range((excel_row, col_idx))
            if cell.color == MISMATCH_COLOR:
                cell.color = None


# ── Dashboard stamp ───────────────────────────────────────────────────────────

def _stamp_dashboard(wb):
    """Write last-run timestamp next to DASH_LABEL on CA Tools or Dashboard sheet."""
    now_str = datetime.now().strftime("%m/%d/%Y %I:%M %p")
    for sheet_name in ("CA Tools", "Dashboard"):
        try:
            dash = wb.sheets[sheet_name]
        except Exception:
            _log(f"stamp_dashboard: sheet '{sheet_name}' not found")
            continue
        data = dash.used_range.value or []
        for r_idx, row in enumerate(data):
            if not row:
                continue
            for c_idx, cell in enumerate(row):
                if cell and str(cell).strip() == DASH_LABEL:
                    dash.range((r_idx + 1, c_idx + 3)).value = now_str
                    _log(f"stamp_dashboard: wrote '{now_str}' to {sheet_name} row={r_idx+1} col={c_idx+3}")
                    return
        _log(f"stamp_dashboard: label '{DASH_LABEL}' not found on '{sheet_name}'")


# ── Export helpers ───────────────────────────────────────────────────────────

# Zoom template column order (matches zoomus_user_template exactly)
_ZOOM_TEMPLATE_COLS = [
    "Email", "First Name", "Last Name", "Package",
    "Site Code", "Site Name", "User Template",
    "Extension Number", "Phone Number", "Outbound Caller ID",
    "Select Outbound Caller ID", "SMS", "User Status",
    "Desk Phone 1's Brand", "Desk Phone 1's Model",
    "Desk Phone 1's MAC Address", "Desk Phone 1's Provision Template",
    "Desk Phone 2's Brand", "Desk Phone 2's Model",
    "Desk Phone 2's MAC Address", "Desk Phone 2's Provision Template",
    "Desk Phone 3's Brand", "Desk Phone 3's Model",
    "Desk Phone 3's MAC Address", "Desk Phone 3's Provision Template",
]

_UPDATE_STATUSES = {STATUS_PROGRESS, STATUS_DISCREP}
_ADD_STATUSES    = {STATUS_INCOMPLETE}


def _export(wb, ws, df, headers, mode, phone_source="temp"):
    import pandas as pd
    from datetime import date
    import csv

    use_temp = phone_source == "temp"
    phone_sh_col   = H_PHONE_TEMP  if use_temp else H_PHONE
    outbound_sh_col = H_OUTBOUND_TEMP if use_temp else H_OUTBOUND

    if mode == "update":
        statuses  = _UPDATE_STATUSES
        suggested = f"ZPU_Update_{date.today().strftime('%Y%m%d')}.csv"
        title     = "Save ZP Update CSV"
        # Add Changes column at end for update exports
        export_cols = _ZOOM_TEMPLATE_COLS + [CHANGES_HDR]
    else:
        statuses  = _ADD_STATUSES
        suggested = f"ZPU_Add_{date.today().strftime('%Y%m%d')}.csv"
        title     = "Save ZP Add CSV"
        export_cols = _ZOOM_TEMPLATE_COLS

    if H_STATUS not in headers:
        dlg.info("Export Error", f"'{H_STATUS}' column not found.")
        return

    # Column mapping: Zoom template col → DGW sheet col
    # Phone Number / Outbound Caller ID use whichever source was chosen
    col_map = {
        "Phone Number":        phone_sh_col,
        "Outbound Caller ID":  outbound_sh_col,
    }

    all_statuses = [str(row.get(H_STATUS, "") or "").strip() for _, row in df.iterrows()]
    _log(f"export mode={mode} statuses_looking_for={statuses}")
    _log(f"export unique statuses on sheet: {sorted(set(all_statuses))}")
    _log(f"export total rows in df: {len(df)}")

    rows = []
    for _, row in df.iterrows():
        status = str(row.get(H_STATUS, "") or "").strip()
        if status not in statuses:
            continue
        # Skip rows where TCS Data Status = "Removed"
        tcs_status = str(row.get(H_TCS_DATA_STATUS, "") or "").strip().lower()
        if tcs_status == "removed":
            continue
        out = {}
        for col in export_cols:
            sheet_col = col_map.get(col, col)   # remap phone cols, rest are direct
            if sheet_col in headers:
                out[col] = row.get(sheet_col, "")
            else:
                out[col] = ""
        rows.append(out)

    if not rows:
        dlg.info("Export", f"No rows to export for {mode.upper()}.")
        return

    save_path = dlg.get_save_path(suggested, title)
    if not save_path:
        return

    try:
        with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=export_cols)
            writer.writeheader()
            writer.writerows(rows)
        src_label = "Zoom Temp" if use_temp else "Actual"
        dlg.info("Export Complete",
                 f"{mode.upper()} export saved ({src_label} numbers):\n"
                 f"{save_path}\n\n{len(rows)} rows.")
    except Exception as e:
        dlg.info("Export Error", str(e))


# ── Main entry point ──────────────────────────────────────────────────────────

def run_zp_reconciliation():
    import pandas as pd
    open(LOG_PATH, "w").close()
    _log("run_zp_reconciliation start")

    wb = _get_wb()
    if not wb:
        dlg.info("Error", "Could not find open workbook.")
        return

    ws = _get_sheet(wb, USERS_SHEET)
    if ws is None:
        return

    df = _read_df(ws)
    if df.empty:
        dlg.info("Error", f"No data found in the '{USERS_SHEET}' sheet.")
        return

    headers = list(df.columns)

    # ── Find required DGW columns ─────────────────────────────────────────────
    required = {
        H_EMAIL:         _find_col(headers, H_EMAIL),
        H_FIRST:         _find_col(headers, H_FIRST),
        H_LAST:          _find_col(headers, H_LAST),
        H_PACKAGE:       _find_col(headers, H_PACKAGE),
        H_SITE_CODE:     _find_col(headers, H_SITE_CODE),
        H_SITE_NAME:     _find_col(headers, H_SITE_NAME),
        H_EXT:           _find_col(headers, H_EXT),
        H_PHONE_TEMP:    _find_col(headers, H_PHONE_TEMP),
        H_OUTBOUND_TEMP: _find_col(headers, H_OUTBOUND_TEMP),
        H_STATUS:        _find_col(headers, H_STATUS),
        H_DP1_BRAND:     _find_col(headers, H_DP1_BRAND),
        H_DP1_MODEL:     _find_col(headers, H_DP1_MODEL),
        H_DP1_MAC:       _find_col(headers, H_DP1_MAC),
        H_DP1_PROV:      _find_col(headers, H_DP1_PROV),
    }

    missing = [name for name, col in required.items() if col == 0]
    if missing:
        dlg.info("Error",
                 f"Users sheet is missing required columns:\n\n"
                 + "\n".join(f"  - {m}" for m in missing))
        return

    d = required  # alias for brevity

    # ── Intro dialog ──────────────────────────────────────────────────────────
    action = dlg.show_zp_intro()
    if action != "import":
        return

    # ── Pick Zoom Phone CSV ───────────────────────────────────────────────────
    csv_path = dlg.pick_csv("Select Zoom Phone Users CSV")
    _log(f"csv_path={csv_path}")
    if not csv_path:
        dlg.info("Cancelled", "No CSV selected. No changes were made.")
        return

    try:
        df_csv = pd.read_csv(csv_path, dtype=str).fillna("")
        df_csv.columns = [str(c).strip() for c in df_csv.columns]
        _log(f"CSV: {len(df_csv)} rows, cols={list(df_csv.columns[:6])}")
    except Exception as e:
        dlg.info("Error", f"Could not load CSV:\n{e}")
        return

    # ── Verify CSV columns ────────────────────────────────────────────────────
    csv_required = [C_EMAIL, C_PACKAGE, C_EXT, C_PHONE, C_OUTBOUND,
                    C_DP1_BRAND, C_DP1_MODEL, C_DP1_MAC, C_DP1_PROV]
    csv_cols_lower = {c.lower(): c for c in df_csv.columns}
    csv_col_map = {}
    missing_csv = []

    for req in csv_required:
        found = csv_cols_lower.get(req.lower())
        if found:
            csv_col_map[req] = found
        else:
            missing_csv.append(req)

    if missing_csv:
        dlg.info("Error",
                 f"Zoom Phone CSV is missing required columns:\n\n"
                 + "\n".join(f"  - {m}" for m in missing_csv))
        return

    # Optional CSV columns
    csv_col_map[C_SITE_CODE] = csv_cols_lower.get(C_SITE_CODE.lower(), "")
    csv_col_map[C_SITE_NAME] = csv_cols_lower.get(C_SITE_NAME.lower(), "")

    # ── Ask which phone columns to compare against ────────────────────────────
    phone_source = dlg.ask_phone_source()
    if phone_source is None:
        dlg.info("Cancelled", "No changes were made.")
        return
    use_temp      = phone_source == "temp"
    phone_sh_col  = H_PHONE_TEMP if use_temp else H_PHONE
    outbound_sh_col = H_OUTBOUND_TEMP if use_temp else H_OUTBOUND
    _log(f"phone_source={phone_source}  phone_col={phone_sh_col}")

    compare_cols = [
        H_PACKAGE,
        H_EXT,
        phone_sh_col,
        outbound_sh_col,
        H_DP1_BRAND,
        H_DP1_MODEL,
        H_DP1_MAC,
    ]

    # ── Build email → CSV row lookup ─────────────────────────────────────────
    def _csv_col(name):
        return csv_col_map.get(name, "")

    lookup = {}
    for _, row in df_csv.iterrows():
        email = _norm_email(row.get(_csv_col(C_EMAIL), ""))
        if email and email not in lookup:
            lookup[email] = row

    _log(f"CSV lookup size: {len(lookup)}")

    # Clear previous mismatch highlights before processing
    _clear_mismatch_highlights(ws, df, headers, compare_cols)

    # ── Process each DGW row ──────────────────────────────────────────────────
    cnt = dict(complete=0, progress=0, discrep=0, incomplete=0)
    total_rows = len(df)

    prog = dlg.ProgressWindow(f"Reconciling 0 of {total_rows} rows...", wb=wb, title="ZP User Recon")

    try:
        for i, (excel_row, row) in enumerate(df.iterrows()):
            if i % 5 == 0:
                prog.update(f"Reconciling {i + 1} of {total_rows} rows...")

            em = _norm_email(str(row.iloc[d[H_EMAIL] - 1]))
            if not em:
                continue

            if em not in lookup:
                ws.range((excel_row, d[H_STATUS])).value = STATUS_INCOMPLETE
                _write(ws, excel_row, headers, CHANGES_HDR, "")
                _highlight_mismatches(ws, excel_row, headers, set(), compare_cols)
                cnt["incomplete"] += 1
                continue

            cr = lookup[em]

            def sheet_val(hdr):
                col = d.get(hdr, 0)
                return row.iloc[col - 1] if col else ""

            def csv_val(col_name):
                col = _csv_col(col_name)
                return cr.get(col, "") if col else ""

            pkg_match   = _text_equal(sheet_val(H_PACKAGE), csv_val(C_PACKAGE))
            ext_match   = _ext_equal(sheet_val(H_EXT), csv_val(C_EXT))
            phone_match = _phone_equal_na_ok(sheet_val(phone_sh_col), csv_val(C_PHONE))
            out_match   = _phone_equal_na_ok(sheet_val(outbound_sh_col), csv_val(C_OUTBOUND))

            # Softphone check — DGW Brand=Zoom/Model=Softphone + CSV has no device
            softphone = _is_softphone_match(
                sheet_val(H_DP1_BRAND), sheet_val(H_DP1_MODEL),
                csv_val(C_DP1_BRAND),   csv_val(C_DP1_MODEL),
                csv_val(C_DP1_MAC),
            )

            if softphone:
                dev_match, dev_discrep = True, False
            else:
                dev_match, dev_discrep = _device_compare(
                    sheet_val(H_DP1_BRAND), sheet_val(H_DP1_MODEL),
                    sheet_val(H_DP1_MAC),   sheet_val(H_DP1_PROV),
                    csv_val(C_DP1_BRAND),   csv_val(C_DP1_MODEL),
                    csv_val(C_DP1_MAC),     csv_val(C_DP1_PROV),
                )

            all_match = pkg_match and ext_match and phone_match and out_match and dev_match

            mismatches = set()
            if not pkg_match:   mismatches.add(H_PACKAGE)
            if not ext_match:   mismatches.add(H_EXT)
            if not phone_match: mismatches.add(phone_sh_col)
            if not out_match:   mismatches.add(outbound_sh_col)
            if not dev_match:   mismatches.add(H_DP1_BRAND)
            if dev_discrep:     mismatches.add(H_DP1_MAC)

            if all_match:
                ws.range((excel_row, d[H_STATUS])).value = STATUS_COMPLETE
                changes_note = "Softphone" if softphone else ""
                _write(ws, excel_row, headers, CHANGES_HDR, changes_note)
                _write(ws, excel_row, headers, H_ZP_PACKAGE, csv_val(C_PACKAGE))
                _highlight_mismatches(ws, excel_row, headers, set(), compare_cols)
                cnt["complete"] += 1
            elif dev_discrep:
                ws.range((excel_row, d[H_STATUS])).value = STATUS_DISCREP
                _write(ws, excel_row, headers, CHANGES_HDR, ", ".join(sorted(mismatches)))
                _highlight_mismatches(ws, excel_row, headers, mismatches, compare_cols)
                cnt["discrep"] += 1
            else:
                ws.range((excel_row, d[H_STATUS])).value = STATUS_PROGRESS
                _write(ws, excel_row, headers, CHANGES_HDR, ", ".join(sorted(mismatches)))
                _highlight_mismatches(ws, excel_row, headers, mismatches, compare_cols)
                cnt["progress"] += 1

    except Exception as _loop_err:
        import traceback as _tb
        _log(f"Loop error: {_loop_err}\n{_tb.format_exc()}")
        prog.close()
        dlg.info("Reconciliation Error", str(_loop_err))
        return

    _log(f"Counts: {cnt}")

    prog.update("Applying status colors...")
    _apply_colors(ws, _read_df(ws), headers)
    _stamp_dashboard(wb)
    prog.close()

    # ── Re-read sheet so exports see the freshly written statuses ────────────
    df_fresh  = _read_df(ws)
    headers_f = list(df_fresh.columns)
    _log(f"df_fresh rows={len(df_fresh)}  H_STATUS in headers={H_STATUS in headers_f}")

    # ── Results dialog + optional exports ────────────────────────────────────
    _log("Showing results dialog...")
    exports = dlg.show_zp_results({
        "complete":   cnt["complete"],
        "discrep":    cnt["discrep"],
        "progress":   cnt["progress"],
        "incomplete": cnt["incomplete"],
    })
    _log(f"Exports selected: {exports}")
    if "update" in exports:
        _log("Calling UPDATE export...")
        _export(wb, ws, df_fresh, headers_f, "update", phone_source)
    if "add" in exports:
        _log("Calling ADD export...")
        _export(wb, ws, df_fresh, headers_f, "add", phone_source)
