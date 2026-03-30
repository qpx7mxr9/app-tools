"""
zca_recon/dialogs.py

Custom dialogs for CA Reconciliation.
Replaces multiple yes/no popups with clean single-window forms.
"""

import tkinter as tk
from tkinter import ttk, filedialog


# ── File pickers ──────────────────────────────────────────────

def pick_csv(title="Select Source Export CSV"):
    root = _root()
    path = filedialog.askopenfilename(
        parent=root, title=title,
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
    root.destroy()
    return path or ""


def get_save_path(suggested, title="Save CSV"):
    root = _root()
    root.deiconify()
    _center(root, 1, 1)   # force root to center so dialog inherits position
    root.update()
    path = filedialog.asksaveasfilename(
        parent=root, title=title, initialfile=suggested,
        defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    root.destroy()
    return path or ""


# ── Intro dialog ──────────────────────────────────────────────

def show_intro():
    """
    Returns: 'import', 'skip', or None (cancel)
    """
    result = {"action": None}

    root = _root()
    root.deiconify()
    root.title("CA Reconciliation")
    root.resizable(False, False)
    _center(root, 420, 280)

    # Header
    header = tk.Frame(root, bg="#1F2D4E", height=50)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="COMMON AREA RECONCILIATION",
             bg="#1F2D4E", fg="white",
             font=("Segoe UI", 12, "bold")).pack(side="left", padx=18, pady=0, anchor="center")

    # Body
    body = tk.Frame(root, bg="white", padx=20, pady=14)
    body.pack(fill="both", expand=True)

    tk.Label(body, text="What you will need:",
             bg="white", font=("Segoe UI", 10, "bold"),
             fg="#333").pack(anchor="w")
    tk.Label(body,
             text="  •  Source export CSV\n"
                  "     (Admin Portal › Phone › Common Area Phones › Export)",
             bg="white", font=("Segoe UI", 9), fg="#555",
             justify="left").pack(anchor="w", pady=(2, 10))

    tk.Label(body, text="After reconciling you can export:",
             bg="white", font=("Segoe UI", 10, "bold"),
             fg="#333").pack(anchor="w")
    tk.Label(body,
             text="  •  UPDATE file — exists in source but data doesn't match\n"
                  "  •  ADD file — not yet in source system",
             bg="white", font=("Segoe UI", 9), fg="#555",
             justify="left").pack(anchor="w", pady=(2, 0))

    # Buttons
    btn_frame = tk.Frame(root, bg="#F0F0F0", padx=14, pady=10)
    btn_frame.pack(fill="x")

    def on_import():
        result["action"] = "import"
        root.destroy()

    def on_skip():
        result["action"] = "skip"
        root.destroy()

    def on_cancel():
        root.destroy()

    tk.Button(btn_frame, text="Cancel",
              bg="#D8D8D8", fg="#333",
              font=("Segoe UI", 10), width=10, relief="flat", cursor="hand2",
              command=on_cancel).pack(side="right", padx=(4, 0))
    tk.Button(btn_frame, text="Skip Import →",
              bg="#607D9F", fg="white",
              font=("Segoe UI", 10), width=13, relief="flat", cursor="hand2",
              command=on_skip).pack(side="right", padx=(4, 0))
    tk.Button(btn_frame, text="Import CSV →",
              bg="#1F2D4E", fg="white",
              font=("Segoe UI", 10, "bold"), width=13, relief="flat", cursor="hand2",
              command=on_import).pack(side="right", padx=(4, 0))

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return result["action"]


# ── Results + export dialog ───────────────────────────────────

def show_results(counts):
    """
    Show reconciliation results and ask what to export.
    counts = {"complete": n, "disc": n, "progress": n, "incomplete": n}
    Returns: set of {"update", "add"} — which exports to run
    """
    result = {"exports": set(), "confirmed": False}

    root = _root()
    root.deiconify()
    root.title("Reconciliation Complete")
    root.resizable(False, False)
    _center(root, 400, 320)

    # Header
    header = tk.Frame(root, bg="#1F2D4E", height=48)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="RECONCILIATION COMPLETE",
             bg="#1F2D4E", fg="white",
             font=("Segoe UI", 12, "bold")).pack(side="left", padx=18, anchor="center")

    # Stats
    stats_frame = tk.Frame(root, bg="white", padx=20, pady=14)
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
    tk.Frame(root, bg="#E0E0E0", height=1).pack(fill="x", padx=20)

    # Export options
    export_frame = tk.Frame(root, bg="white", padx=20, pady=12)
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
    btn_frame = tk.Frame(root, bg="#F0F0F0", padx=14, pady=10)
    btn_frame.pack(fill="x")

    def on_done():
        if var_update.get(): result["exports"].add("update")
        if var_add.get():    result["exports"].add("add")
        result["confirmed"] = True
        root.destroy()

    def on_cancel():
        root.destroy()

    tk.Button(btn_frame, text="Skip Exports",
              bg="#D8D8D8", fg="#333",
              font=("Segoe UI", 10), width=12, relief="flat", cursor="hand2",
              command=on_cancel).pack(side="right", padx=(4, 0))
    tk.Button(btn_frame, text="Export Selected →",
              bg="#1F2D4E", fg="white",
              font=("Segoe UI", 10, "bold"), width=16, relief="flat", cursor="hand2",
              command=on_done).pack(side="right", padx=(4, 0))

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return result["exports"]


# ── Phone number source dialog ────────────────────────────────

def ask_phone_source():
    """
    Ask whether to use Zoom Temp or actual phone numbers.
    Returns: 'temp', 'actual', or None (cancel)
    """
    result = {"choice": None}

    root = _root()
    root.deiconify()
    root.title("Phone Number Source")
    root.resizable(False, False)
    _center(root, 340, 180)

    header = tk.Frame(root, bg="#1F2D4E", height=42)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="SELECT PHONE NUMBER SOURCE",
             bg="#1F2D4E", fg="white",
             font=("Segoe UI", 10, "bold")).pack(side="left", padx=16, anchor="center")

    body = tk.Frame(root, bg="white", padx=22, pady=14)
    body.pack(fill="both", expand=True)

    var = tk.StringVar(value="temp")
    tk.Radiobutton(body, text="Zoom Temp Numbers",
                   variable=var, value="temp",
                   bg="white", font=("Segoe UI", 10)).pack(anchor="w")
    tk.Radiobutton(body, text="Actual Numbers",
                   variable=var, value="actual",
                   bg="white", font=("Segoe UI", 10)).pack(anchor="w", pady=(6, 0))

    btn_frame = tk.Frame(root, bg="#F0F0F0", padx=14, pady=10)
    btn_frame.pack(fill="x")

    def on_ok():
        result["choice"] = var.get()
        root.destroy()

    def on_cancel():
        root.destroy()

    tk.Button(btn_frame, text="Cancel",
              bg="#D8D8D8", fg="#333",
              font=("Segoe UI", 10), width=10, relief="flat", cursor="hand2",
              command=on_cancel).pack(side="right", padx=(4, 0))
    tk.Button(btn_frame, text="Continue →",
              bg="#1F2D4E", fg="white",
              font=("Segoe UI", 10, "bold"), width=12, relief="flat", cursor="hand2",
              command=on_ok).pack(side="right", padx=(4, 0))

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return result["choice"]


# ── Simple message ────────────────────────────────────────────

def info(title, message):
    from tkinter import messagebox
    root = _root()
    root.deiconify()
    _center(root, 1, 1)
    root.update()
    messagebox.showinfo(title, message, parent=root)
    root.destroy()


# ── Helpers ───────────────────────────────────────────────────

def _root():
    r = tk.Tk()
    r.withdraw()
    try:
        r.attributes("-topmost", True)
    except Exception:
        pass
    return r


def _center(win, w, h):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")
