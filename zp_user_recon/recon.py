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
import pandas as pd
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

BRAND_WORKPLACE_APP = "workplace app"

COLOR_GOOD   = (198, 239, 206)   # green
COLOR_CHANGE = (255, 235, 156)   # yellow

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


# ── Logging ───────────────────────────────────────────────────────────────────

def _log(msg):
    try:
        with open(LOG_PATH, "a") as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
    except Exception:
        pass


# ── Main entry point ──────────────────────────────────────────────────────────

def run_zp_reconciliation():
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

    # ── Build email → CSV row lookup ─────────────────────────────────────────
    def _csv_col(name):
        return csv_col_map.get(name, "")

    lookup = {}
    for _, row in df_csv.iterrows():
        email = _norm_email(row.get(_csv_col(C_EMAIL), ""))
        if email and email not in lookup:
            lookup[email] = row

    _log(f"CSV lookup size: {len(lookup)}")

    # ── Process each DGW row ──────────────────────────────────────────────────
    cnt = dict(complete=0, progress=0, discrep=0, incomplete=0)

    for excel_row, row in df.iterrows():
        em = _norm_email(str(row.iloc[d[H_EMAIL] - 1]))
        if not em:
            continue

        if em not in lookup:
            ws.range((excel_row, d[H_STATUS])).value = STATUS_INCOMPLETE
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
        phone_match = _phone_equal_na_ok(sheet_val(H_PHONE_TEMP), csv_val(C_PHONE))
        out_match   = _phone_equal_na_ok(sheet_val(H_OUTBOUND_TEMP), csv_val(C_OUTBOUND))

        dev_match, dev_discrep = _device_compare(
            sheet_val(H_DP1_BRAND), sheet_val(H_DP1_MODEL),
            sheet_val(H_DP1_MAC),   sheet_val(H_DP1_PROV),
            csv_val(C_DP1_BRAND),   csv_val(C_DP1_MODEL),
            csv_val(C_DP1_MAC),     csv_val(C_DP1_PROV),
        )

        all_match = pkg_match and ext_match and phone_match and out_match and dev_match

        if all_match:
            ws.range((excel_row, d[H_STATUS])).value = STATUS_COMPLETE
            cnt["complete"] += 1
        elif dev_discrep:
            ws.range((excel_row, d[H_STATUS])).value = STATUS_DISCREP
            cnt["discrep"] += 1
        else:
            ws.range((excel_row, d[H_STATUS])).value = STATUS_PROGRESS
            cnt["progress"] += 1

    _log(f"Counts: {cnt}")

    # ── Summary ───────────────────────────────────────────────────────────────
    dlg.info("Zoom Phone Reconciliation",
             f"Reconciliation complete.\n\n"
             f"Setup Complete:      {cnt['complete']}\n"
             f"Setup in Progress:   {cnt['progress']}\n"
             f"Setup Discrepancy:   {cnt['discrep']}\n"
             f"Setup Incomplete:    {cnt['incomplete']}")
