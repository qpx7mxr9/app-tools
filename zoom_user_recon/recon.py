"""
zoom_user_recon/recon.py

Zoom User Status Audit
-----------------------
Populates three columns in the Users sheet:
  Zoom User Status       - Active - In Account / Inactive - In Account /
                           Not In Account / Pending Activation / Not Found
  Zoom License Status    - Zoom license string (users IN the account only)
  Zoom User External Info - Account type + number (users NOT in the account)

Import files prompted in sequence:
  1. Zoom Users Export  (Zoom Admin > User Management > Users > Export)
     Columns used: Email | Licenses | User Status
  2. Domain Data  (optional) - Email | Account Type | Zoom Acct Number
  3. Pending Users (optional) - Email

Entry points (called from Excel via xlwings RunPython):
    run_zoom_user_audit()   - Main audit flow
    clear_zoom_results()    - Clears the three output columns
"""

import xlwings as xw
import pandas as pd
from datetime import datetime
from zca_recon import dialogs as dlg

# ── Constants ─────────────────────────────────────────────────────────────────
WS_NAME      = "Users"
COL_EMAIL    = "Email"
COL_STATUS   = "Zoom User Status"
COL_LICENSE  = "Zoom License Status"
COL_EXTERNAL = "Zoom User External Info"
DASH_LABEL   = "Zoom User Last Update:"

import tempfile as _tempfile, os as _os
LOG_PATH = _os.path.join(_tempfile.gettempdir(), "zoom_user_recon.log")

# Status → (background RGB, font RGB)
_STATUS_COLORS = {
    "active - in account":   ((198, 239, 206), (0, 97, 0)),
    "inactive - in account": ((255, 235, 156), (156, 101, 0)),
    "not in account":        ((255, 199, 206), (156, 0, 6)),
    "not found":             ((242, 242, 242), (128, 128, 128)),
    "pending activation":    ((221, 235, 247), (31, 73, 125)),
}


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
    """Return 1-based column index, or 0 if not found."""
    for i, h in enumerate(headers):
        if str(h).strip().lower() == name.lower():
            return i + 1
    return 0


def _find_df_col(df, *names):
    """Return first matching column name in df, or None."""
    for name in names:
        for c in df.columns:
            if str(c).strip().lower() == name.lower():
                return c
    return None


# ── Normalization ─────────────────────────────────────────────────────────────

def _norm_email(v):
    s = str(v or "").strip().lower()
    s = s.replace("mailto:", "").replace("\xa0", "").strip()
    return s


# ── File loading ──────────────────────────────────────────────────────────────

def _load_file(path):
    """Load CSV or Excel file into a DataFrame with stripped column names."""
    ext = path.rsplit(".", 1)[-1].lower()
    df = pd.read_csv(path, dtype=str) if ext == "csv" else pd.read_excel(path, dtype=str)
    df = df.fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ── Coloring ──────────────────────────────────────────────────────────────────

def _apply_colors(ws, df, status_col_idx):
    """Apply background (and attempt font) colors to the status column."""
    for excel_row, row in df.iterrows():
        raw = row.iloc[status_col_idx - 1] if status_col_idx - 1 < len(row) else ""
        val = str(raw).strip().lower()
        cell = ws.range((excel_row, status_col_idx))
        if val in _STATUS_COLORS:
            bg, _fg = _STATUS_COLORS[val]
            cell.color = bg
        else:
            cell.color = None


# ── Dashboard stamp ───────────────────────────────────────────────────────────

def _stamp_dashboard(wb):
    try:
        dash = wb.sheets["Dashboard"]
    except Exception:
        return
    now_str = datetime.now().strftime("%m-%d-%Y")
    data = dash.used_range.value or []
    for r_idx, row in enumerate(data):
        if not row:
            continue
        for c_idx, cell in enumerate(row):
            if cell and DASH_LABEL in str(cell).strip():
                dash.range((r_idx + 1, c_idx + 2)).value = now_str
                return
    try:
        dash.range("J18").value = now_str
    except Exception:
        pass


# ── Logging ───────────────────────────────────────────────────────────────────

def _log(msg):
    try:
        with open(LOG_PATH, "a") as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
    except Exception:
        pass


# ── Main entry point ──────────────────────────────────────────────────────────

def run_zoom_user_audit():
    open(LOG_PATH, "w").close()
    _log("run_zoom_user_audit start")

    wb = _get_wb()
    if not wb:
        dlg.info("Error", "Could not find open workbook.")
        return

    ws = _get_sheet(wb, WS_NAME)
    if ws is None:
        return

    df = _read_df(ws)
    if df.empty:
        dlg.info("Error", "No user data found in the Users sheet.")
        return

    headers = list(df.columns)

    email_col    = _find_col(headers, COL_EMAIL) or 1
    status_col   = _find_col(headers, COL_STATUS)
    license_col  = _find_col(headers, COL_LICENSE)
    external_col = _find_col(headers, COL_EXTERNAL)

    missing_hdrs = [n for n, c in [
        (COL_STATUS, status_col), (COL_LICENSE, license_col), (COL_EXTERNAL, external_col)
    ] if not c]
    if missing_hdrs:
        dlg.info("Column Not Found",
                 f"Could not find required column headers in the Users sheet:\n\n"
                 f"{', '.join(missing_hdrs)}")
        return

    # ── Intro dialog — choose optional files ──────────────────────────────────
    intro = dlg.show_zu_intro()
    if intro["action"] != "start":
        return

    # ── Step 1: Zoom Users Export (required) ──────────────────────────────────
    zoom_path = dlg.pick_file_any("Step 1 \u2013 Select Zoom Users Export")
    if not zoom_path:
        dlg.info("Cancelled", "Import cancelled. No changes were made.")
        return

    try:
        df_zoom = _load_file(zoom_path)
        _log(f"Zoom Users: {len(df_zoom)} rows, cols={list(df_zoom.columns)}")
    except Exception as e:
        dlg.info("Error", f"Could not load Zoom Users Export:\n{e}")
        return

    z_email_col   = _find_df_col(df_zoom, "Email") or df_zoom.columns[0]
    z_license_col = _find_df_col(df_zoom, "Licenses", "License") or df_zoom.columns[1]
    z_status_col  = _find_df_col(df_zoom, "User Status", "Status") or df_zoom.columns[2]

    # ── Step 2: Domain Data (optional) ────────────────────────────────────────
    df_domain = None
    if intro["domain"]:
        domain_path = dlg.pick_file_any("Step 2 \u2013 Select Domain Data")
        if domain_path:
            try:
                df_domain = _load_file(domain_path)
                _log(f"Domain Data: {len(df_domain)} rows")
            except Exception as e:
                dlg.info("Warning", f"Could not load Domain Data:\n{e}\n\nContinuing without it.")

    d_email_col = _find_df_col(df_domain, "Email") if df_domain is not None else None
    d_acct_type = _find_df_col(df_domain, "Account Type") if df_domain is not None else None
    d_acct_num  = _find_df_col(df_domain, "Zoom Acct Number", "Acct Number") if df_domain is not None else None

    # ── Step 3: Pending Users (optional) ──────────────────────────────────────
    df_pending = None
    if intro["pending"]:
        pending_path = dlg.pick_file_any("Step 3 \u2013 Select Pending Users")
        if pending_path:
            try:
                df_pending = _load_file(pending_path)
                _log(f"Pending Users: {len(df_pending)} rows")
            except Exception as e:
                dlg.info("Warning", f"Could not load Pending Users:\n{e}\n\nContinuing without it.")

    p_email_col = _find_df_col(df_pending, "Email") if df_pending is not None else None

    # ── Build lookup dictionaries ──────────────────────────────────────────────
    zoom_dict = {}
    for _, row in df_zoom.iterrows():
        email = _norm_email(row.get(z_email_col, ""))
        if email and email not in zoom_dict:
            zoom_dict[email] = (
                str(row.get(z_status_col, "")).strip(),
                str(row.get(z_license_col, "")).strip(),
            )

    domain_dict = {}
    if df_domain is not None and d_email_col:
        for _, row in df_domain.iterrows():
            email = _norm_email(row.get(d_email_col, ""))
            if email and email not in domain_dict:
                acct_type = str(row.get(d_acct_type, "") if d_acct_type else "").strip()
                acct_num  = str(row.get(d_acct_num, "")  if d_acct_num  else "").strip()
                try:
                    acct_num = str(int(float(acct_num)))  # strip trailing .0
                except (ValueError, TypeError):
                    pass
                domain_dict[email] = (acct_type, acct_num)

    pending_set = set()
    if df_pending is not None and p_email_col:
        for _, row in df_pending.iterrows():
            email = _norm_email(row.get(p_email_col, ""))
            if email:
                pending_set.add(email)

    _log(f"zoom_dict={len(zoom_dict)}  domain_dict={len(domain_dict)}  pending_set={len(pending_set)}")

    # ── Process each user row ─────────────────────────────────────────────────
    cnt = dict(active=0, inactive=0, domain=0, pending=0, missing=0)

    for excel_row, row in df.iterrows():
        ue = _norm_email(str(row.iloc[email_col - 1]))
        if not ue:
            continue

        ws.range((excel_row, status_col)).value   = ""
        ws.range((excel_row, license_col)).value  = ""
        ws.range((excel_row, external_col)).value = ""
        ws.range((excel_row, status_col)).color   = None
        ws.range((excel_row, license_col)).color  = None
        ws.range((excel_row, external_col)).color = None

        if ue in zoom_dict:
            z_stat, z_lic = zoom_dict[ue]
            if z_stat.lower() == "active":
                ws.range((excel_row, status_col)).value = "Active - In Account"
                cnt["active"] += 1
            else:
                ws.range((excel_row, status_col)).value = "Inactive - In Account"
                cnt["inactive"] += 1
            ws.range((excel_row, license_col)).value  = z_lic
            ws.range((excel_row, external_col)).value = ""

        elif ue in pending_set:
            ws.range((excel_row, status_col)).value   = "Pending Activation"
            ws.range((excel_row, license_col)).value  = ""
            ws.range((excel_row, external_col)).value = ""
            cnt["pending"] += 1

        elif ue in domain_dict:
            d_type, d_num = domain_dict[ue]
            ws.range((excel_row, status_col)).value  = "Not In Account"
            ws.range((excel_row, license_col)).value = ""
            t = d_type.lower()
            if "business" in t:
                ext_info = "Business Account"
                if d_num:
                    ext_info += f" | Acct #: {d_num}"
            elif "pro" in t:
                ext_info = "Pro Account"
                if d_num:
                    ext_info += f" | Acct #: {d_num}"
            elif "free with credit" in t:
                ext_info = "Free Account (Credit Card)"
            elif "free" in t:
                ext_info = "Free Account"
            else:
                ext_info = d_type
                if d_num:
                    ext_info += f" | Acct #: {d_num}"
            ws.range((excel_row, external_col)).value = ext_info
            cnt["domain"] += 1

        else:
            ws.range((excel_row, status_col)).value   = "Not Found"
            ws.range((excel_row, license_col)).value  = ""
            ws.range((excel_row, external_col)).value = ""
            cnt["missing"] += 1

    _log(f"Counts: {cnt}")

    # Color-code status column
    df_fresh = _read_df(ws)
    _apply_colors(ws, df_fresh, status_col)

    # Stamp dashboard
    _stamp_dashboard(wb)

    # ── Summary ───────────────────────────────────────────────────────────────
    total = sum(cnt.values())
    lines = [
        "AUDIT COMPLETE",
        "",
        f"{total} users processed:",
        "",
        f"  Active \u2013 In Account:    {cnt['active']}",
        f"  Inactive \u2013 In Account:  {cnt['inactive']}",
        f"  Not In Account:         {cnt['domain']}",
    ]
    if df_pending is not None:
        lines.append(f"  Pending Activation:     {cnt['pending']}")
    lines += [
        f"  Not Found:              {cnt['missing']}",
        "",
        "COLOR KEY:",
        "  Green  \u2013 Active in Zoom account",
        "  Yellow \u2013 Inactive in Zoom account",
        "  Red    \u2013 Not in Zoom account (see External Info)",
    ]
    if df_pending is not None:
        lines.append("  Blue   \u2013 Pending account activation")
    lines.append("  Gray   \u2013 Not found in any source")
    if cnt["domain"] > 0:
        lines += [
            "",
            f"Note: {cnt['domain']} user(s) marked Not In Account need a",
            "Zoom license before Zoom Phone can be assigned.",
        ]

    dlg.show_zu_results(cnt, has_pending=df_pending is not None)


# ── Clear results entry point ─────────────────────────────────────────────────

def clear_zoom_results():
    wb = _get_wb()
    if not wb:
        dlg.info("Error", "Could not find open workbook.")
        return

    ws = _get_sheet(wb, WS_NAME)
    if ws is None:
        return

    df = _read_df(ws)
    headers = list(df.columns)

    status_col   = _find_col(headers, COL_STATUS)
    license_col  = _find_col(headers, COL_LICENSE)
    external_col = _find_col(headers, COL_EXTERNAL)

    if not all([status_col, license_col, external_col]):
        dlg.info("Error", "Could not find one or more output columns.")
        return

    if not dlg.ask_yes_no("Clear Zoom Status Results",
                          f"Clear all values in:\n"
                          f"  - {COL_STATUS}\n"
                          f"  - {COL_LICENSE}\n"
                          f"  - {COL_EXTERNAL}\n\nContinue?"):
        return

    last_row = len(df) + 1
    for col in [status_col, license_col, external_col]:
        rng = ws.range((2, col), (last_row, col))
        rng.value = ""
        rng.color = None

    dlg.info("Done", "Results cleared.")
