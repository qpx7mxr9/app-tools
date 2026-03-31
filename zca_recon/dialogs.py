"""
zca_recon/dialogs.py

Custom dialogs for CA Reconciliation.

Uses a single persistent hidden Tk root so we never call tk.Tk() more than
once per process.  Each dialog is a tk.Toplevel() child; we block with
root.wait_window() instead of root.mainloop(), which is safe to call
multiple times.  This avoids the macOS crash caused by re-initialising Tk
(TkpInit -> [NSApplication setMainMenu:] assertion) that occurs when a
second tk.Tk() is created after the first has been destroyed.
"""

import sys as _sys
import os as _os
import tkinter as tk
from tkinter import filedialog

# On Mac, permanently redirect raw fd 2 to /dev/null for this xlwings
# subprocess.  macOS Tk 8.5 writes Objective-C warnings directly to fd 2
# (bypassing Python's sys.stderr) at various points during NSApp init —
# including on the first Toplevel creation, which happens after any
# try/finally silence block.  This process is short-lived and ephemeral;
# stderr is not needed here (errors go to /tmp/zca_recon.log instead).
if _sys.platform == "darwin":
    try:
        _d = _os.open(_os.devnull, _os.O_WRONLY)
        _os.dup2(_d, 2)
        _os.close(_d)
        del _d
    except Exception:
        pass


# ── Single shared root ────────────────────────────────────────

_root = None


def _get_root():
    """Return (creating if needed) the one hidden root Tk window."""
    global _root
    if _root is not None:
        try:
            _root.winfo_id()   # raises TclError if already destroyed
            return _root
        except Exception:
            _root = None

    _root = tk.Tk()
    _root.withdraw()
    _root.update_idletasks()
    try:
        _root.attributes("-topmost", True)
    except Exception:
        pass
    _focus_python()
    return _root


def _focus_python():
    """Bring the Python process to the foreground on Mac."""
    try:
        import subprocess, platform
        if platform.system() == "Darwin":
            subprocess.Popen(
                ["osascript", "-e",
                 'tell application "System Events" to set frontmost of '
                 'first process whose name starts with "Python" to true'],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


# ── File pickers ──────────────────────────────────────────────

def pick_file_any(title="Select File"):
    """Open a file picker that accepts CSV and Excel files."""
    import sys
    if sys.platform == "darwin":
        return _macos_open_dialog(title, ["csv", "xlsx", "xls", "xlsm"])
    root = _get_root()
    path = filedialog.askopenfilename(
        parent=root, title=title,
        filetypes=[("Excel & CSV", "*.xlsx *.xls *.xlsm *.csv"),
                   ("CSV Files", "*.csv"),
                   ("Excel Files", "*.xlsx *.xls *.xlsm"),
                   ("All Files", "*.*")])
    return path or ""


def pick_csv(title="Select Source Export CSV"):
    """Open a CSV file picker. Uses osascript on macOS for reliable focus."""
    import sys
    if sys.platform == "darwin":
        return _macos_open_dialog(title, ["csv"])
    root = _get_root()
    path = filedialog.askopenfilename(
        parent=root, title=title,
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
    return path or ""


def get_save_path(suggested, title="Save CSV"):
    """Open a save-file dialog. Uses osascript on macOS for reliable focus."""
    import sys
    if sys.platform == "darwin":
        return _macos_save_dialog(title, suggested)
    root = _get_root()
    path = filedialog.asksaveasfilename(
        parent=root, title=title, initialfile=suggested,
        defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    return path or ""


def _macos_open_dialog(title, file_types=None):
    """
    Use AppleScript 'choose file' on macOS.
    Always appears on top — no Tk focus required.
    """
    import subprocess
    type_clause = ""
    if file_types:
        quoted = ", ".join(f'"{t}"' for t in file_types)
        type_clause = f" of type {{{quoted}}}"
    script = (
        f'tell application "System Events" to activate\n'
        f'set f to choose file with prompt "{title}"{type_clause}\n'
        f'return POSIX path of f'
    )
    try:
        r = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=300)
        return r.stdout.strip() or ""
    except Exception:
        return ""


def _macos_save_dialog(title, default_name):
    """
    Use AppleScript 'choose file name' on macOS.
    Always appears on top — no Tk focus required.
    """
    import subprocess
    # Default folder: Desktop
    script = (
        f'tell application "System Events" to activate\n'
        f'set f to choose file name with prompt "{title}" '
        f'default name "{default_name}"\n'
        f'return POSIX path of f'
    )
    try:
        r = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=300)
        path = r.stdout.strip()
        if path and not path.endswith(".csv"):
            path += ".csv"
        return path or ""
    except Exception:
        return ""


# ── Progress window ──────────────────────────────────────────

class ProgressWindow:
    """
    Small non-closeable status window for long-running operations.
    Call update(msg) mid-loop — uses update_idletasks() to repaint
    without processing user events (safe against Tk reentrancy).
    Works on both Mac and Windows.
    """
    def __init__(self, message="Working..."):
        root = _get_root()
        win = tk.Toplevel(root)
        win.title("CA Reconciliation")
        win.resizable(False, False)
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        _center(win, 340, 72)
        win.protocol("WM_DELETE_WINDOW", lambda: None)  # prevent close

        frame = tk.Frame(win, bg="#1F2D4E", padx=24, pady=20)
        frame.pack(fill="both", expand=True)
        self._label = tk.Label(
            frame, text=message,
            bg="#1F2D4E", fg="white",
            font=("Segoe UI", 10))
        self._label.pack()
        self._win = win
        win.update_idletasks()

    def update(self, message):
        try:
            self._label.config(text=message)
            self._win.update_idletasks()
        except Exception:
            pass

    def close(self):
        try:
            self._win.destroy()
        except Exception:
            pass


# ── Shared styled dialog helpers ─────────────────────────────

def _make_header(win, text):
    h = tk.Frame(win, bg="#1F2D4E", height=50)
    h.pack(fill="x")
    h.pack_propagate(False)
    tk.Label(h, text=text, bg="#1F2D4E", fg="white",
             font=("Segoe UI", 12, "bold")).pack(side="left", padx=18, anchor="center")

def _make_body(win):
    b = tk.Frame(win, bg="white", padx=20, pady=14)
    b.pack(fill="both", expand=True)
    return b

def _make_btn_frame(win):
    f = tk.Frame(win, bg="#F0F0F0", padx=14, pady=10)
    f.pack(fill="x")
    return f

def _add_cancel_btn(frame, cmd, text="Cancel"):
    tk.Button(frame, text=text, bg="#D8D8D8", fg="#333",
              font=("Segoe UI", 10), width=10, relief="flat",
              cursor="hand2", command=cmd).pack(side="right", padx=(4, 0))

def _add_primary_btn(frame, cmd, text="Continue ->", width=13):
    tk.Button(frame, text=text, bg="#1F2D4E", fg="white",
              font=("Segoe UI", 10, "bold"), width=width, relief="flat",
              cursor="hand2", command=cmd).pack(side="right", padx=(4, 0))

def _stat_row(parent, label, value, fg, bg, label_width=20):
    row = tk.Frame(parent, bg="white")
    row.pack(fill="x", pady=2)
    tk.Label(row, text=f"  {label}:", bg="white",
             font=("Segoe UI", 10), fg="#555",
             width=label_width, anchor="w").pack(side="left")
    tk.Label(row, text=str(value), bg=bg, fg=fg,
             font=("Segoe UI", 10, "bold"),
             width=6, relief="flat").pack(side="left")


# ── ZP User Recon dialogs ─────────────────────────────────────

def show_zp_intro():
    """ZP intro dialog. Returns: 'import' or None (cancel)."""
    result = {"action": None}
    root = _get_root(); _focus_python()
    win = tk.Toplevel(root)
    win.title("ZP User Reconciliation")
    win.resizable(False, False)
    try: win.attributes("-topmost", True)
    except Exception: pass
    _center(win, 430, 260)

    _make_header(win, "ZOOM PHONE USER RECONCILIATION")
    body = _make_body(win)

    tk.Label(body, text="What you will need:",
             bg="white", font=("Segoe UI", 10, "bold"), fg="#333").pack(anchor="w")
    tk.Label(body,
             text="  \u2022  Zoom Phone Users CSV\n"
                  "     (Admin Portal \u2192 Phone \u2192 Users \u2192 Export)",
             bg="white", font=("Segoe UI", 9), fg="#555",
             justify="left").pack(anchor="w", pady=(2, 10))
    tk.Label(body, text="After reconciling you can export:",
             bg="white", font=("Segoe UI", 10, "bold"), fg="#333").pack(anchor="w")
    tk.Label(body,
             text="  \u2022  UPDATE file \u2013 in Zoom Phone but data doesn\u2019t match\n"
                  "  \u2022  ADD file \u2013 not yet in Zoom Phone",
             bg="white", font=("Segoe UI", 9), fg="#555",
             justify="left").pack(anchor="w", pady=(2, 0))

    bf = _make_btn_frame(win)

    def on_ok():   result["action"] = "import"; win.destroy()
    def on_cancel(): win.destroy()

    _add_cancel_btn(bf, on_cancel)
    _add_primary_btn(bf, on_ok, "Import CSV ->")
    win.protocol("WM_DELETE_WINDOW", on_cancel)
    win.lift(); win.focus_force()
    root.wait_window(win)
    return result["action"]


def show_zp_results(counts):
    """
    ZP results dialog. Returns set of {"update", "add"}.
    counts = {"complete": n, "discrep": n, "progress": n, "incomplete": n}
    """
    result = {"exports": set()}
    root = _get_root(); _focus_python()
    win = tk.Toplevel(root)
    win.title("ZP Reconciliation Complete")
    win.resizable(False, False)
    try: win.attributes("-topmost", True)
    except Exception: pass
    _center(win, 480, 320)

    _make_header(win, "ZP RECONCILIATION COMPLETE")

    stats = tk.Frame(win, bg="white", padx=20, pady=14)
    stats.pack(fill="x")
    _stat_row(stats, "Setup Complete",    counts.get("complete",   0), "#00612A", "#C6EFCE", 22)
    _stat_row(stats, "Setup Discrepancy", counts.get("discrep",    0), "#9C6400", "#FFEB9C", 22)
    _stat_row(stats, "Setup in Progress", counts.get("progress",   0), "#9C6400", "#FFEB9C", 22)
    _stat_row(stats, "Setup Incomplete",  counts.get("incomplete", 0), "#9C0006", "#FFC7CE", 22)

    tk.Frame(win, bg="#E0E0E0", height=1).pack(fill="x", padx=20)

    exp = tk.Frame(win, bg="white", padx=20, pady=12)
    exp.pack(fill="x")
    tk.Label(exp, text="Select exports:", bg="white",
             font=("Segoe UI", 10, "bold"), fg="#333").pack(anchor="w", pady=(0, 6))

    var_upd = tk.BooleanVar(value=counts.get("discrep", 0) > 0 or counts.get("progress", 0) > 0)
    var_add = tk.BooleanVar(value=counts.get("incomplete", 0) > 0)
    tk.Checkbutton(exp, text="UPDATE file  (Discrepancy / In Progress)",
                   variable=var_upd, bg="white",
                   font=("Segoe UI", 10), fg="#333",
                   activebackground="white").pack(anchor="w")
    tk.Checkbutton(exp, text="ADD file  (Setup Incomplete \u2013 not in Zoom Phone)",
                   variable=var_add, bg="white",
                   font=("Segoe UI", 10), fg="#333",
                   activebackground="white").pack(anchor="w", pady=(4, 0))

    bf = _make_btn_frame(win)

    def on_done():
        if var_upd.get(): result["exports"].add("update")
        if var_add.get(): result["exports"].add("add")
        win.destroy()
    def on_skip(): win.destroy()

    _add_cancel_btn(bf, on_skip, "Skip Exports")
    _add_primary_btn(bf, on_done, "Export Selected ->", width=16)
    win.protocol("WM_DELETE_WINDOW", on_skip)
    win.lift(); win.focus_force()
    root.wait_window(win)
    return result["exports"]


# ── ZU Recon dialogs ──────────────────────────────────────────

def show_zu_intro():
    """
    ZU intro dialog with checkboxes for optional files.
    Returns: {"action": "start"|None, "domain": bool, "pending": bool}
    """
    result = {"action": None, "domain": False, "pending": False}
    root = _get_root(); _focus_python()
    win = tk.Toplevel(root)
    win.title("Zoom User Audit")
    win.resizable(False, False)
    try: win.attributes("-topmost", True)
    except Exception: pass
    _center(win, 450, 300)

    _make_header(win, "ZOOM USER STATUS AUDIT")
    body = _make_body(win)

    tk.Label(body, text="Required:",
             bg="white", font=("Segoe UI", 10, "bold"), fg="#333").pack(anchor="w")
    tk.Label(body,
             text="  \u2022  Zoom Users Export\n"
                  "     (Admin Portal \u2192 User Management \u2192 Users \u2192 Export)",
             bg="white", font=("Segoe UI", 9), fg="#555",
             justify="left").pack(anchor="w", pady=(2, 10))

    tk.Label(body, text="Optional \u2014 check to include:",
             bg="white", font=("Segoe UI", 10, "bold"), fg="#333").pack(anchor="w")

    var_domain  = tk.BooleanVar(value=False)
    var_pending = tk.BooleanVar(value=False)
    tk.Checkbutton(body,
                   text="Domain Data  (Email | Account Type | Zoom Acct Number)",
                   variable=var_domain, bg="white",
                   font=("Segoe UI", 9), fg="#555",
                   activebackground="white").pack(anchor="w", pady=(4, 0))
    tk.Checkbutton(body,
                   text="Pending Users  (Email \u2014 users awaiting activation)",
                   variable=var_pending, bg="white",
                   font=("Segoe UI", 9), fg="#555",
                   activebackground="white").pack(anchor="w", pady=(4, 0))

    bf = _make_btn_frame(win)

    def on_start():
        result["action"]  = "start"
        result["domain"]  = var_domain.get()
        result["pending"] = var_pending.get()
        win.destroy()
    def on_cancel(): win.destroy()

    _add_cancel_btn(bf, on_cancel)
    _add_primary_btn(bf, on_start, "Start Audit ->")
    win.protocol("WM_DELETE_WINDOW", on_cancel)
    win.lift(); win.focus_force()
    root.wait_window(win)
    return result


def show_zu_results(counts, has_pending=False):
    """ZU results dialog. counts = {"active", "inactive", "domain", "pending", "missing"}."""
    root = _get_root(); _focus_python()
    win = tk.Toplevel(root)
    win.title("Zoom User Audit Complete")
    win.resizable(False, False)
    try: win.attributes("-topmost", True)
    except Exception: pass
    _center(win, 460, 320 if has_pending else 290)

    _make_header(win, "ZOOM USER AUDIT COMPLETE")

    stats = tk.Frame(win, bg="white", padx=20, pady=14)
    stats.pack(fill="both", expand=True)
    _stat_row(stats, "Active \u2013 In Account",   counts.get("active",   0), "#00612A", "#C6EFCE", 24)
    _stat_row(stats, "Inactive \u2013 In Account", counts.get("inactive", 0), "#9C6400", "#FFEB9C", 24)
    _stat_row(stats, "Not In Account",             counts.get("domain",   0), "#9C0006", "#FFC7CE", 24)
    if has_pending:
        _stat_row(stats, "Pending Activation",     counts.get("pending",  0), "#1F497D", "#DCE6F1", 24)
    _stat_row(stats, "Not Found",                  counts.get("missing",  0), "#666666", "#F2F2F2", 24)

    bf = _make_btn_frame(win)

    def on_done(): win.destroy()
    tk.Button(bf, text="Done", bg="#1F2D4E", fg="white",
              font=("Segoe UI", 10, "bold"), width=10, relief="flat",
              cursor="hand2", command=on_done).pack(side="right")
    win.protocol("WM_DELETE_WINDOW", on_done)
    win.lift(); win.focus_force()
    root.wait_window(win)


# ── Intro dialog ──────────────────────────────────────────────

def show_intro():
    """
    Returns: 'import', 'skip', or None (cancel)
    """
    result = {"action": None}
    root = _get_root()
    _focus_python()

    win = tk.Toplevel(root)
    win.title("CA Reconciliation")
    win.resizable(False, False)
    try:
        win.attributes("-topmost", True)
    except Exception:
        pass
    _center(win, 420, 280)

    # Header
    header = tk.Frame(win, bg="#1F2D4E", height=50)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="COMMON AREA RECONCILIATION",
             bg="#1F2D4E", fg="white",
             font=("Segoe UI", 12, "bold")).pack(side="left", padx=18, anchor="center")

    # Body
    body = tk.Frame(win, bg="white", padx=20, pady=14)
    body.pack(fill="both", expand=True)

    tk.Label(body, text="What you will need:",
             bg="white", font=("Segoe UI", 10, "bold"),
             fg="#333").pack(anchor="w")
    tk.Label(body,
             text="  \u2022  Source export CSV\n"
                  "     (Admin Portal > Phone > Common Area Phones > Export)",
             bg="white", font=("Segoe UI", 9), fg="#555",
             justify="left").pack(anchor="w", pady=(2, 10))

    tk.Label(body, text="After reconciling you can export:",
             bg="white", font=("Segoe UI", 10, "bold"),
             fg="#333").pack(anchor="w")
    tk.Label(body,
             text="  \u2022  UPDATE file -- exists in source but data doesn't match\n"
                  "  \u2022  ADD file -- not yet in source system",
             bg="white", font=("Segoe UI", 9), fg="#555",
             justify="left").pack(anchor="w", pady=(2, 0))

    # Buttons
    btn_frame = tk.Frame(win, bg="#F0F0F0", padx=14, pady=10)
    btn_frame.pack(fill="x")

    def on_import():
        result["action"] = "import"
        win.destroy()

    def on_skip():
        result["action"] = "skip"
        win.destroy()

    def on_cancel():
        win.destroy()

    tk.Button(btn_frame, text="Cancel",
              bg="#D8D8D8", fg="#333",
              font=("Segoe UI", 10), width=10, relief="flat", cursor="hand2",
              command=on_cancel).pack(side="right", padx=(4, 0))
    tk.Button(btn_frame, text="Skip Import ->",
              bg="#607D9F", fg="white",
              font=("Segoe UI", 10), width=13, relief="flat", cursor="hand2",
              command=on_skip).pack(side="right", padx=(4, 0))
    tk.Button(btn_frame, text="Import CSV ->",
              bg="#1F2D4E", fg="white",
              font=("Segoe UI", 10, "bold"), width=13, relief="flat", cursor="hand2",
              command=on_import).pack(side="right", padx=(4, 0))

    win.protocol("WM_DELETE_WINDOW", on_cancel)
    win.lift()
    win.focus_force()
    root.wait_window(win)
    return result["action"]


# ── Results + export dialog ───────────────────────────────────

def show_results(counts):
    """
    Show reconciliation results and ask what to export.
    counts = {"complete": n, "disc": n, "progress": n, "incomplete": n}
    Returns: set of {"update", "add"} -- which exports to run
    """
    result = {"exports": set(), "confirmed": False}
    root = _get_root()
    _focus_python()

    win = tk.Toplevel(root)
    win.title("Reconciliation Complete")
    win.resizable(False, False)
    try:
        win.attributes("-topmost", True)
    except Exception:
        pass
    _center(win, 480, 340)

    # Header
    header = tk.Frame(win, bg="#1F2D4E", height=48)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="RECONCILIATION COMPLETE",
             bg="#1F2D4E", fg="white",
             font=("Segoe UI", 12, "bold")).pack(side="left", padx=18, anchor="center")

    # Stats
    stats_frame = tk.Frame(win, bg="white", padx=20, pady=14)
    stats_frame.pack(fill="x")

    stat_items = [
        ("Complete",      counts.get("complete",   0), "#00612A", "#C6EFCE"),
        ("Discrepancy",   counts.get("disc",        0), "#9C6400", "#FFEB9C"),
        ("In Progress",   counts.get("progress",    0), "#9C6400", "#FFEB9C"),
        ("Incomplete",    counts.get("incomplete",  0), "#9C0006", "#FFC7CE"),
    ]

    for label, value, fg, bg in stat_items:
        row = tk.Frame(stats_frame, bg="white")
        row.pack(fill="x", pady=2)
        tk.Label(row, text=f"  {label}:", bg="white",
                 font=("Segoe UI", 10), fg="#555", width=16,
                 anchor="w").pack(side="left")
        tk.Label(row, text=str(value), bg=bg, fg=fg,
                 font=("Segoe UI", 10, "bold"),
                 width=6, relief="flat").pack(side="left")

    # Divider
    tk.Frame(win, bg="#E0E0E0", height=1).pack(fill="x", padx=20)

    # Export options
    export_frame = tk.Frame(win, bg="white", padx=20, pady=12)
    export_frame.pack(fill="x")

    tk.Label(export_frame, text="Select exports:",
             bg="white", font=("Segoe UI", 10, "bold"),
             fg="#333").pack(anchor="w", pady=(0, 6))

    var_update = tk.BooleanVar(value=counts.get("disc", 0) > 0 or counts.get("progress", 0) > 0)
    var_add    = tk.BooleanVar(value=counts.get("incomplete", 0) > 0)

    tk.Checkbutton(export_frame,
                   text="UPDATE file  (Discrepancy / In Progress)",
                   variable=var_update,
                   bg="white", font=("Segoe UI", 10), fg="#333",
                   activebackground="white").pack(anchor="w")
    tk.Checkbutton(export_frame,
                   text="ADD file  (Not in source system yet)",
                   variable=var_add,
                   bg="white", font=("Segoe UI", 10), fg="#333",
                   activebackground="white").pack(anchor="w", pady=(4, 0))

    # Buttons
    btn_frame = tk.Frame(win, bg="#F0F0F0", padx=14, pady=10)
    btn_frame.pack(fill="x")

    def on_done():
        if var_update.get(): result["exports"].add("update")
        if var_add.get():    result["exports"].add("add")
        result["confirmed"] = True
        win.destroy()

    def on_cancel():
        win.destroy()

    tk.Button(btn_frame, text="Skip Exports",
              bg="#D8D8D8", fg="#333",
              font=("Segoe UI", 10), width=12, relief="flat", cursor="hand2",
              command=on_cancel).pack(side="right", padx=(4, 0))
    tk.Button(btn_frame, text="Export Selected ->",
              bg="#1F2D4E", fg="white",
              font=("Segoe UI", 10, "bold"), width=16, relief="flat", cursor="hand2",
              command=on_done).pack(side="right", padx=(4, 0))

    win.protocol("WM_DELETE_WINDOW", on_cancel)
    win.lift()
    win.focus_force()
    root.wait_window(win)
    return result["exports"]


# ── Phone number source dialog ────────────────────────────────

def ask_phone_source():
    """
    Ask whether to use Zoom Temp or actual phone numbers.
    Returns: 'temp', 'actual', or None (cancel)
    """
    result = {"choice": None}
    root = _get_root()
    _focus_python()

    win = tk.Toplevel(root)
    win.title("Phone Number Source")
    win.resizable(False, False)
    try:
        win.attributes("-topmost", True)
    except Exception:
        pass
    _center(win, 340, 180)

    header = tk.Frame(win, bg="#1F2D4E", height=42)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="SELECT PHONE NUMBER SOURCE",
             bg="#1F2D4E", fg="white",
             font=("Segoe UI", 10, "bold")).pack(side="left", padx=16, anchor="center")

    body = tk.Frame(win, bg="white", padx=22, pady=14)
    body.pack(fill="both", expand=True)

    var = tk.StringVar(value="temp")
    tk.Radiobutton(body, text="Zoom Temp Numbers",
                   variable=var, value="temp",
                   bg="white", font=("Segoe UI", 10)).pack(anchor="w")
    tk.Radiobutton(body, text="Actual Numbers",
                   variable=var, value="actual",
                   bg="white", font=("Segoe UI", 10)).pack(anchor="w", pady=(6, 0))

    btn_frame = tk.Frame(win, bg="#F0F0F0", padx=14, pady=10)
    btn_frame.pack(fill="x")

    def on_ok():
        result["choice"] = var.get()
        win.destroy()

    def on_cancel():
        win.destroy()

    tk.Button(btn_frame, text="Cancel",
              bg="#D8D8D8", fg="#333",
              font=("Segoe UI", 10), width=10, relief="flat", cursor="hand2",
              command=on_cancel).pack(side="right", padx=(4, 0))
    tk.Button(btn_frame, text="Continue ->",
              bg="#1F2D4E", fg="white",
              font=("Segoe UI", 10, "bold"), width=12, relief="flat", cursor="hand2",
              command=on_ok).pack(side="right", padx=(4, 0))

    win.protocol("WM_DELETE_WINDOW", on_cancel)
    win.lift()
    win.focus_force()
    root.wait_window(win)
    return result["choice"]


# ── Yes/No dialog ────────────────────────────────────────────

def ask_yes_no(title, message):
    """
    Blocking Yes/No dialog.
    Returns True if user clicked Yes, False otherwise.
    """
    import sys
    if sys.platform == "darwin":
        import subprocess
        # Embed message with newlines replaced by AppleScript return literal
        lines = str(message).split("\n")
        as_parts = " & return & ".join(f'"{l.replace(chr(92), chr(92)*2).replace(chr(34), chr(92)+chr(34))}"' for l in lines)
        ttl = title.replace("\\", "\\\\").replace('"', '\\"')
        script = (
            f'tell application "System Events" to activate\n'
            f'set r to button returned of (display dialog {as_parts} with title "{ttl}" '
            f'buttons {{"No", "Yes"}} default button "Yes")\n'
            f'return r'
        )
        try:
            r = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=120)
            return r.stdout.strip() == "Yes"
        except Exception:
            return False
    from tkinter import messagebox
    root = _get_root()
    return messagebox.askyesno(title, message, parent=root)


# ── Messages ─────────────────────────────────────────────────

def info(title, message):
    """Blocking dialog — use for errors that need acknowledgement."""
    import sys
    if sys.platform == "darwin":
        import subprocess
        msg = message.replace("\\", "\\\\").replace('"', '\\"')
        ttl = title.replace('"', '\\"')
        script = (
            f'display dialog "{msg}" with title "{ttl}" '
            f'buttons {{"OK"}} default button "OK"'
        )
        try:
            subprocess.run(["osascript", "-e", script],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=120)
        except Exception:
            pass
        return
    from tkinter import messagebox
    root = _get_root()
    messagebox.showinfo(title, message, parent=root)


def notify(title, message):
    """Non-blocking macOS notification — fires and disappears on its own."""
    import sys
    if sys.platform == "darwin":
        import subprocess
        msg = message.replace("\\", "\\\\").replace('"', '\\"')
        ttl = title.replace('"', '\\"')
        script = f'display notification "{msg}" with title "{ttl}"'
        try:
            subprocess.Popen(["osascript", "-e", script],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
        return
    # Windows fallback — just use a messagebox
    from tkinter import messagebox
    root = _get_root()
    messagebox.showinfo(title, message, parent=root)


# ── Helpers ───────────────────────────────────────────────────

def _center(win, w, h):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")
