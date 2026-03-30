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

import tkinter as tk
from tkinter import filedialog


# ── Single shared root ────────────────────────────────────────

_root = None


def _silence_fd2():
    """
    Redirect fd 2 to /dev/null and return a callable that restores it.
    Used to suppress macOS Objective-C warnings (e.g. Secure Coding) that
    write directly to the raw file descriptor, bypassing sys.stderr.
    """
    import os, sys
    if sys.platform != "darwin":
        return lambda: None
    devnull = os.open(os.devnull, os.O_WRONLY)
    saved = os.dup(2)
    os.dup2(devnull, 2)
    os.close(devnull)
    def _restore():
        os.dup2(saved, 2)
        os.close(saved)
    return _restore


def _get_root():
    """Return (creating if needed) the one hidden root Tk window."""
    global _root
    if _root is not None:
        try:
            _root.winfo_id()   # raises TclError if already destroyed
            return _root
        except Exception:
            _root = None

    # Suppress the macOS "Secure coding" Tk 8.5 warning.
    # NSApplication finishes initializing restorable state on the first event
    # loop tick, so keep fd 2 silenced through withdraw + update_idletasks.
    restore = _silence_fd2()
    try:
        _root = tk.Tk()
        _root.withdraw()
        _root.update_idletasks()
    finally:
        restore()

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
    _center(win, 400, 320)

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
