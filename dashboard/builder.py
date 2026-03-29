"""
dashboard/builder.py

Dashboard 2 — Deployment Tracking Dashboard
--------------------------------------------
Builds and refreshes the Dashboard 2 worksheet in the workbook.
Called from Excel buttons via xlwings RunPython.

Entry points:
    build_dashboard()   - Full build: creates/resets Dashboard 2 and populates all blocks
    refresh_ca_block()  - Refresh only the Common Areas block (fast update)
"""

import sys
import platform
import xlwings as xw
from datetime import datetime

_IS_MAC = platform.system() == "Darwin"

# ── Sheet names ───────────────────────────────────────────────
DASHBOARD_SHEET = "CA Tools"
CA_SHEET        = "Common Area"

# ── Column indices (0-based) on Common Area sheet ─────────────
CA_COL_NAME     = 0   # Display Name
CA_COL_STATUS   = 35  # Common Area Status
CA_COL_DATAST   = 34  # Data Status
CA_COL_SITE     = 2   # Site Name
CA_COL_LASTRUN  = 37  # Common Area (Last Update)

# ── Colors ────────────────────────────────────────────────────
C_HEADER_BG     = (26,  45,  77)    # Dark navy
C_HEADER_FG     = (255, 255, 255)   # White
C_SECTION_BG    = (44,  62,  100)   # Medium navy
C_SECTION_FG    = (255, 255, 255)   # White
C_CARD_BG       = (242, 245, 250)   # Light blue-grey
C_LABEL_FG      = (80,  80,  80)    # Grey label
C_COMPLETE_BG   = (198, 239, 206)   # Green
C_COMPLETE_FG   = (0,   97,  0)     # Dark green
C_SETUP_BG      = (255, 199, 206)   # Red/pink
C_SETUP_FG      = (156, 0,   6)     # Dark red
C_WARN_BG       = (255, 235, 156)   # Yellow
C_WARN_FG       = (156, 101, 0)     # Dark yellow
C_NEUTRAL_BG    = (242, 242, 242)   # Light grey
C_NEUTRAL_FG    = (64,  64,  64)    # Dark grey
C_WHITE         = (255, 255, 255)
C_BORDER        = (189, 189, 189)   # Light border grey
C_PROGRESS_FILL = (0,   176, 80)    # Progress bar green
C_PROGRESS_EMPTY= (217, 217, 217)   # Progress bar empty


# ── xlwings cell helpers ──────────────────────────────────────

def _cell(ws, row, col):
    """1-based row/col cell reference."""
    return ws.range((row, col))

def _range(ws, r1, c1, r2, c2):
    return ws.range((r1, c1), (r2, c2))

def _write(ws, row, col, value, bold=False, size=11, fg=None, bg=None,
           align="left", valign="center", wrap=False, num_fmt=None, italic=False):
    cell = _cell(ws, row, col)
    cell.value = value
    if bg:
        cell.color = bg
    if num_fmt:
        cell.number_format = num_fmt
    _fmt_cell(cell, bold=bold, size=size, italic=italic, fg=fg,
              align=align, valign=valign, wrap=wrap)


def _fmt_cell(cell, bold=False, size=11, italic=False, fg=None,
              align="left", valign="center", wrap=False):
    """Apply font + alignment — cross-platform (xlwings Font object + api for alignment)."""
    try:
        cell.font.bold   = bold
        cell.font.size   = size
        cell.font.italic = italic
        if fg:
            cell.font.color = fg   # xlwings Font accepts (R,G,B) tuple
    except Exception:
        pass
    # Alignment and wrap — try COM first, fall back to AppleScript
    try:
        if _IS_MAC:
            cell.api.horizontal_alignment.set(_halign(align))
            cell.api.vertical_alignment.set(_valign(valign))
            cell.api.wrap_text.set(wrap)
        else:
            cell.api.HorizontalAlignment = _halign(align)
            cell.api.VerticalAlignment   = _valign(valign)
            cell.api.WrapText            = wrap
    except Exception:
        pass

def _merge(ws, r1, c1, r2, c2):
    _range(ws, r1, c1, r2, c2).merge()

def _fill(ws, r1, c1, r2, c2, bg):
    _range(ws, r1, c1, r2, c2).color = bg

def _font(ws, r1, c1, r2, c2, fg=None, bold=False, size=None):
    rng = _range(ws, r1, c1, r2, c2)
    try:
        rng.font.bold = bold
        if fg:
            rng.font.color = fg
        if size:
            rng.font.size = size
    except Exception:
        pass

def _border_box(ws, r1, c1, r2, c2, color=C_BORDER, weight=2):
    """Draw a thin border around a range."""
    rng = _range(ws, r1, c1, r2, c2)
    try:
        if _IS_MAC:
            rng.api.border_around(line_style=1, weight=weight,
                                   color=_rgb(color))
        else:
            rng.api.BorderAround(LineStyle=1, Weight=weight,
                                  Color=_rgb(color))
    except Exception:
        pass

def _row_height(ws, row, height):
    ws.range((row, 1)).row_height = height

def _col_width(ws, col, width):
    ws.range((1, col)).column_width = width

def _halign(a):
    return {
        "left":   -4131,
        "center": -4108,
        "right":  -4152,
    }.get(a, -4131)

def _valign(a):
    return {
        "top":    -4160,
        "center": -4108,
        "bottom": -4107,
    }.get(a, -4108)

def _rgb(t):
    """Convert (R,G,B) to Excel BGR integer."""
    r, g, b = t
    return r + g * 256 + b * 65536

def _pct(n, d):
    return round((n / d) * 100) if d > 0 else 0


# ── Data readers ──────────────────────────────────────────────

def _read_ca_data(wb):
    """
    Read Common Area sheet and return stats dict.
    Returns None if sheet not found or empty.
    """
    try:
        ws   = wb.sheets[CA_SHEET]
        data = ws.used_range.value
    except Exception:
        return None

    if not data or len(data) < 2:
        return None

    total = complete = setup = verified = discrepancy = partial = not_found = 0
    last_run   = None
    site_stats = {}

    for row in data[1:]:
        name = row[CA_COL_NAME] if len(row) > CA_COL_NAME else None
        if not name:
            continue
        total += 1

        status  = str(row[CA_COL_STATUS]).strip() if len(row) > CA_COL_STATUS and row[CA_COL_STATUS] else ""
        datast  = str(row[CA_COL_DATAST]).strip() if len(row) > CA_COL_DATAST and row[CA_COL_DATAST] else ""
        site    = str(row[CA_COL_SITE]).strip()   if len(row) > CA_COL_SITE   and row[CA_COL_SITE]   else "Unknown"
        lastrun = row[CA_COL_LASTRUN]              if len(row) > CA_COL_LASTRUN                       else None

        if status == "Complete":
            complete += 1
        else:
            setup += 1

        if datast == "Verified":           verified     += 1
        elif datast == "Discrepancy":      discrepancy  += 1
        elif datast == "Partial":          partial      += 1
        elif datast == "Not Found in CSV": not_found    += 1

        if site not in site_stats:
            site_stats[site] = {"total": 0, "complete": 0, "setup": 0}
        site_stats[site]["total"]    += 1
        site_stats[site]["complete"] += 1 if status == "Complete" else 0
        site_stats[site]["setup"]    += 1 if status != "Complete" else 0

        if lastrun and (last_run is None or str(lastrun) > str(last_run)):
            last_run = lastrun

    return {
        "total":       total,
        "complete":    complete,
        "setup":       setup,
        "verified":    verified,
        "discrepancy": discrepancy,
        "partial":     partial,
        "not_found":   not_found,
        "last_run":    last_run,
        "sites":       site_stats,
        "pct":         _pct(complete, total),
    }


# ── Dashboard layout ──────────────────────────────────────────

def _setup_columns(ws):
    """Set column widths for the dashboard."""
    widths = {
        1: 2,    # A - left margin
        2: 18,   # B - labels
        3: 12,   # C - values
        4: 12,   # D
        5: 12,   # E
        6: 12,   # F
        7: 12,   # G
        8: 12,   # H
        9: 18,   # I - right content
        10: 2,   # J - right margin
    }
    for col, width in widths.items():
        _col_width(ws, col, width)


def _draw_header(ws, row, tenant_name="Deployment Dashboard"):
    """Draw the top header bar."""
    _row_height(ws, row, 36)
    _merge(ws, row, 1, row, 10)
    _fill(ws, row, 1, row, 10, C_HEADER_BG)
    _write(ws, row, 2, f"  {tenant_name.upper()}",
           bold=True, size=16, fg=C_HEADER_FG, bg=C_HEADER_BG,
           align="left", valign="center")
    # Timestamp right side
    now = datetime.now().strftime("%m/%d/%Y %H:%M")
    _write(ws, row, 9, f"Refreshed: {now}",
           size=9, fg=(180, 200, 230), bg=C_HEADER_BG,
           align="right", valign="center", italic=True)
    return row + 1


def _draw_spacer(ws, row, height=8):
    _row_height(ws, row, height)
    _fill(ws, row, 1, row, 10, C_WHITE)
    return row + 1


def _draw_section_header(ws, row, title, subtitle=""):
    _row_height(ws, row, 26)
    _merge(ws, row, 1, row, 10)
    _fill(ws, row, 1, row, 10, C_SECTION_BG)
    label = f"  {title}"
    if subtitle:
        label += f"  ·  {subtitle}"
    _write(ws, row, 2, label,
           bold=True, size=11, fg=C_SECTION_FG, bg=C_SECTION_BG,
           align="left", valign="center")
    return row + 1


def _draw_stat_cards(ws, row, stats):
    """
    Draw 4 stat cards in one row: Total | Complete | Setup | % Complete
    """
    _row_height(ws, row,     14)
    _row_height(ws, row + 1, 32)
    _row_height(ws, row + 2, 20)
    _row_height(ws, row + 3, 8)

    cards = [
        ("TOTAL",      stats["total"],    C_NEUTRAL_BG,  C_NEUTRAL_FG,  C_LABEL_FG),
        ("COMPLETE",   stats["complete"], C_COMPLETE_BG, C_COMPLETE_FG, C_LABEL_FG),
        ("IN PROGRESS",stats["setup"],    C_SETUP_BG,    C_SETUP_FG,    C_LABEL_FG),
        ("% COMPLETE", f"{stats['pct']}%",C_NEUTRAL_BG,  C_NEUTRAL_FG,  C_LABEL_FG),
    ]

    col_positions = [2, 4, 6, 8]  # B, D, F, H

    for i, (label, value, bg, val_fg, lbl_fg) in enumerate(cards):
        c = col_positions[i]
        # Merge 2 cols per card
        _merge(ws, row,     c, row,     c + 1)
        _merge(ws, row + 1, c, row + 1, c + 1)
        _merge(ws, row + 2, c, row + 2, c + 1)

        _fill(ws, row,     c, row + 2, c + 1, bg)
        _write(ws, row,     c, label, bold=True, size=8,  fg=lbl_fg, bg=bg, align="center")
        _write(ws, row + 1, c, value, bold=True, size=20, fg=val_fg, bg=bg, align="center")
        _write(ws, row + 2, c, "",    bg=bg)
        _border_box(ws, row, c, row + 2, c + 1)

    return row + 4


def _draw_progress_bar(ws, row, pct, total_cells=16):
    """Draw a visual progress bar using cell fills."""
    _row_height(ws, row,     12)
    _row_height(ws, row + 1, 14)
    _row_height(ws, row + 2, 8)

    # Label
    _merge(ws, row + 1, 2, row + 1, 2)
    _write(ws, row + 1, 2, "PROGRESS", bold=True, size=8,
           fg=C_LABEL_FG, bg=C_WHITE, align="left", valign="center")

    # Progress cells — spread across cols 3 to 9 (7 cols, merge each into mini-bars)
    filled = round((pct / 100) * total_cells) if pct > 0 else 0
    bar_start_col = 3
    bar_end_col   = 9

    # Use individual cells as bar segments (merge col 3-9 into segments)
    # Simpler: fill a merged range partially
    _merge(ws, row + 1, bar_start_col, row + 1, bar_end_col)
    bar_cell = _range(ws, row + 1, bar_start_col, row + 1, bar_end_col)

    if pct == 100:
        bar_cell.color = C_PROGRESS_FILL
    elif pct == 0:
        bar_cell.color = C_PROGRESS_EMPTY
    else:
        bar_cell.color = C_PROGRESS_EMPTY  # fallback — real segmented bar needs API

    _write(ws, row + 1, bar_start_col,
           f"{'█' * filled}{'░' * (total_cells - filled)}  {pct}%",
           bold=False, size=9,
           fg=C_PROGRESS_FILL if pct > 0 else C_LABEL_FG,
           bg=C_NEUTRAL_BG,
           align="left", valign="center")

    _border_box(ws, row + 1, bar_start_col, row + 1, bar_end_col)

    return row + 3


def _draw_data_status(ws, row, stats):
    """Draw the Data Status breakdown row."""
    _row_height(ws, row,     16)
    _row_height(ws, row + 1, 12)
    _row_height(ws, row + 2, 26)
    _row_height(ws, row + 3, 8)

    # Section label
    _merge(ws, row, 2, row, 9)
    _write(ws, row, 2, "DATA STATUS BREAKDOWN",
           bold=True, size=8, fg=C_LABEL_FG, bg=C_WHITE, align="left")

    substats = [
        ("VERIFIED",        stats["verified"],    C_COMPLETE_BG, C_COMPLETE_FG),
        ("DISCREPANCY",     stats["discrepancy"], C_WARN_BG,     C_WARN_FG),
        ("PARTIAL",         stats["partial"],     C_WARN_BG,     C_WARN_FG),
        ("NOT IN SOURCE",   stats["not_found"],   C_SETUP_BG,    C_SETUP_FG),
    ]

    col_positions = [2, 4, 6, 8]

    for i, (label, value, bg, fg) in enumerate(substats):
        c = col_positions[i]
        _merge(ws, row + 1, c, row + 1, c + 1)
        _merge(ws, row + 2, c, row + 2, c + 1)
        _fill(ws, row + 1, c, row + 2, c + 1, bg)
        _write(ws, row + 1, c, label, bold=True, size=7,  fg=fg, bg=bg, align="center")
        _write(ws, row + 2, c, value, bold=True, size=14, fg=fg, bg=bg, align="center")
        _border_box(ws, row + 1, c, row + 2, c + 1)

    return row + 4


def _draw_site_table(ws, row, stats):
    """Draw site-by-site breakdown table."""
    sites = stats["sites"]
    if not sites:
        return row

    _row_height(ws, row, 16)
    _merge(ws, row, 2, row, 9)
    _write(ws, row, 2, "BY SITE",
           bold=True, size=8, fg=C_LABEL_FG, bg=C_WHITE, align="left")
    row += 1

    # Table header
    _row_height(ws, row, 20)
    headers = ["SITE", "TOTAL", "COMPLETE", "IN PROGRESS", "% DONE"]
    col_map  = [2,      4,       5,          6,              7]
    spans    = [(2, 3), (4, 4),  (5, 5),     (6, 6),         (7, 7)]

    for (label, (c1, c2)) in zip(headers, spans):
        _merge(ws, row, c1, row, c2)
        _fill(ws, row, c1, row, c2, C_SECTION_BG)
        _write(ws, row, c1, label, bold=True, size=9,
               fg=C_HEADER_FG, bg=C_SECTION_BG, align="center")

    _border_box(ws, row, 2, row, 7)
    row += 1

    # Table rows
    for site_name in sorted(sites.keys()):
        s = sites[site_name]
        pct = _pct(s["complete"], s["total"])
        row_bg = C_COMPLETE_BG if pct == 100 else C_SETUP_BG if pct == 0 else C_WARN_BG

        _row_height(ws, row, 18)
        data_rows = [
            (site_name, (2, 3), "left"),
            (s["total"],    (4, 4), "center"),
            (s["complete"], (5, 5), "center"),
            (s["setup"],    (6, 6), "center"),
            (f"{pct}%",     (7, 7), "center"),
        ]

        for (val, (c1, c2), align) in data_rows:
            _merge(ws, row, c1, row, c2)
            _fill(ws, row, c1, row, c2, C_WHITE)
            _write(ws, row, c1, val, size=10, fg=C_NEUTRAL_FG,
                   bg=C_WHITE, align=align, valign="center")

        # Color the % done cell
        pct_cell = _cell(ws, row, 7)
        pct_cell.color = row_bg
        _fmt_cell(pct_cell, bold=True,
                  fg=C_COMPLETE_FG if pct == 100 else C_SETUP_FG if pct == 0 else C_WARN_FG)

        _border_box(ws, row, 2, row, 7)
        row += 1

    return row + 1


def _draw_last_run(ws, row, last_run):
    """Draw last run timestamp."""
    _row_height(ws, row, 18)
    _merge(ws, row, 2, row, 5)
    last_str = str(last_run) if last_run else "Never"
    _write(ws, row, 2,
           f"Last Reconciliation Run:  {last_str}",
           size=9, italic=True, fg=C_LABEL_FG, bg=C_WHITE, align="left")
    return row + 1


def _draw_buttons(ws, row):
    """
    Add Form Control buttons wired to VBA macros.
    Falls back to a styled text row if button API is unavailable.
    """
    _row_height(ws, row, 28)
    _fill(ws, row, 1, row, 10, C_WHITE)

    # Button specs: (label, macro_name, col_start, col_end)
    buttons = [
        ("Run Reconciliation",  "ZCA_RunReconciliation", 2, 3),
        ("Export Update",       "ZCA_ExportUpdate",      5, 6),
        ("Export Add",          "ZCA_ExportAdd",         8, 9),
    ]

    added = 0
    for caption, macro, c1, c2 in buttons:
        anchor = _range(ws, row, c1, row, c2)
        try:
            left   = anchor.left   + 2
            top    = anchor.top    + 3
            width  = anchor.width  - 4
            height = anchor.height - 6

            if _IS_MAC:
                btn = ws.api.buttons.add(left, top, width, height)
                btn.caption.set(caption)
                btn.on_action.set(macro)
            else:
                btn = ws.api.Buttons().Add(left, top, width, height)
                btn.Caption  = caption
                btn.OnAction = macro
            added += 1
        except Exception:
            pass

    if added == 0:
        # Fallback: styled text row
        _merge(ws, row, 2, row, 9)
        _write(ws, row, 2,
               "  ▶  Run Reconciliation    ▶  Export Update    ▶  Export Add",
               size=9, bold=True, fg=C_SECTION_BG, bg=C_CARD_BG,
               align="left", valign="center")
        _border_box(ws, row, 2, row, 9, color=C_SECTION_BG, weight=1)

    return row + 2


# ── Public entry points ───────────────────────────────────────

def build_dashboard():
    """
    Full dashboard build — called from Excel button via RunPython.
    """
    _build(xw.Book.caller())


def build_for_workbook(wb_path):
    """
    Full dashboard build — called from a standalone setup script.
    Opens the workbook, builds the dashboard, saves and closes.
    """
    app = xw.App(visible=True)
    wb  = app.books.open(wb_path)
    try:
        _build(wb)
        wb.save()
    finally:
        wb.close()
        app.quit()


def refresh_ca_block():
    """Fast refresh — alias for build_dashboard (always full rebuild)."""
    _build(xw.Book.caller())


def _build(wb):

    # Create or clear Dashboard 2
    try:
        ws = wb.sheets[DASHBOARD_SHEET]
        ws.clear()
    except Exception:
        wb.sheets.add(DASHBOARD_SHEET, after=wb.sheets[0])
        ws = wb.sheets[DASHBOARD_SHEET]

    try:
        ws.activate()
    except Exception:
        pass

    # Hide gridlines
    try:
        if _IS_MAC:
            ws.api.display_gridlines.set(False)
        else:
            ws.api.DisplayGridlines = False
    except Exception:
        pass

    _setup_columns(ws)

    # Fill entire background white
    try:
        ws.used_range.color = C_WHITE
        ws.range("A1:J100").color = C_WHITE
    except Exception:
        pass

    # ── Draw layout ───────────────────────────────────────────
    row = 1
    _draw_spacer(ws, row, 6);  row += 1

    # Header
    row = _draw_header(ws, row, tenant_name="Common Area Tools")
    row = _draw_spacer(ws, row, 10)

    # Common Areas block
    ca = _read_ca_data(wb)
    if ca:
        row = _draw_section_header(ws, row,
                                    "COMMON AREAS",
                                    f"{ca['total']} total  ·  {ca['pct']}% complete")
        row = _draw_spacer(ws, row, 6)
        row = _draw_stat_cards(ws, row, ca)
        row = _draw_progress_bar(ws, row, ca["pct"])
        row = _draw_data_status(ws, row, ca)
        row = _draw_site_table(ws, row, ca)
        row = _draw_last_run(ws, row, ca["last_run"])
        row = _draw_spacer(ws, row, 8)
        row = _draw_buttons(ws, row)
    else:
        _write(ws, row, 2, "⚠  Common Area sheet not found.", fg=C_SETUP_FG)
        row += 2

    row = _draw_spacer(ws, row, 20)

    # Freeze top rows
    try:
        if _IS_MAC:
            ws.range("B3").select()
            wb.app.api.active_window.freeze_panes.set(True)
        else:
            ws.range("B3").api.Select()
            wb.app.api.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    ws.range("A1").select()


def refresh_ca_block():
    """
    Fast refresh — re-reads CA data and updates the dashboard in place.
    Faster than a full rebuild if only CA stats have changed.
    """
    build_dashboard()
