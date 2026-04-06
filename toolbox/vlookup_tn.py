"""
toolbox/vlookup_tn.py

Zoom Temp TN Lookup
-------------------
Fills temp phone numbers across all named lookup ranges using the
'lookup_tn_sheet' named range as the source table.

Named ranges processed (header row skipped in each):
    UsersLookup, CommonAreaLookup, CallQueuesLookup, SLGLookup, ARLookup

For each row: looks up the key in column 0 against lookup_tn_sheet,
writes the result (or "No Temp Number Specified") to the column immediately
to the right — equivalent to VBA's cell.Offset(0, 1).
"""

import xlwings as xw
from datetime import datetime

LOG_PATH = "/tmp/toolbox.log"

LOOKUP_NAMES = [
    "UsersLookup",
    "CommonAreaLookup",
    "CallQueuesLookup",
    "SLGLookup",
    "ARLookup",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _log(msg):
    try:
        with open(LOG_PATH, "a") as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
    except Exception:
        pass


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


# ── Main entry point ──────────────────────────────────────────────────────────

def run_vlookup_zoom_temp_tn():
    from zca_recon import dialogs as dlg

    open(LOG_PATH, "w").close()
    _log("run_vlookup_zoom_temp_tn start")

    wb = _get_wb()
    if not wb:
        dlg.info("Error", "Could not find open workbook.")
        return

    # ── Build lookup dict from named table range ───────────────────────────
    try:
        tbl = wb.names["lookup_tn_sheet"].refers_to_range
        tbl_data = tbl.value
        if not tbl_data:
            dlg.info("Error", "Named range 'lookup_tn_sheet' is empty.")
            return
        # Normalise to list-of-rows regardless of single vs multi row
        if not isinstance(tbl_data[0], list):
            tbl_data = [tbl_data]
        lookup = {}
        for row in tbl_data:
            if row and row[0] is not None:
                k = str(row[0]).strip()
                v = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
                if k:
                    lookup[k] = v
        _log(f"Lookup table: {len(lookup)} entries")
    except Exception as e:
        import traceback as _tb
        _log(f"Lookup table error: {e}\n{_tb.format_exc()}")
        dlg.info("Error", f"Could not read 'lookup_tn_sheet':\n{e}")
        return

    success_count = 0
    missing_count = 0
    processed    = 0

    prog = dlg.ProgressWindow("Starting Zoom Temp TN lookup...")

    try:
        for nm_name in LOOKUP_NAMES:
            try:
                rng = wb.names[nm_name].refers_to_range
            except Exception:
                _log(f"Named range not found: {nm_name} — skipping")
                continue

            ws        = rng.sheet
            nrows     = rng.shape[0]
            col_key   = rng.column          # 1-based column of the key
            row_start = rng.row             # 1-based first row of range

            _log(f"{nm_name}: {nrows} rows  col={col_key}  start_row={row_start}")

            for i in range(1, nrows):       # i=0 is header — skip
                key_cell = ws.range((row_start + i, col_key))
                raw = key_cell.value
                if raw is None:
                    continue
                key = str(raw).strip()
                if not key or key.lower() == "none":
                    continue

                result = lookup.get(key)
                out_cell = ws.range((row_start + i, col_key + 1))

                if result:
                    out_cell.value = result
                    success_count += 1
                else:
                    out_cell.value = "No Temp Number Specified"
                    missing_count += 1

                processed += 1
                if processed % 5 == 0:
                    prog.update(f"Processing {nm_name}: {processed} rows...")

    except Exception as e:
        import traceback as _tb
        _log(f"Loop error: {e}\n{_tb.format_exc()}")
        prog.close()
        dlg.info("Error", str(e))
        return

    prog.close()
    _log(f"Done: processed={processed}  found={success_count}  missing={missing_count}")

    dlg.info(
        "Lookup Complete",
        f"Total processed:       {processed}\n"
        f"Found temp numbers:    {success_count}\n"
        f"Missing temp numbers:  {missing_count}"
    )
