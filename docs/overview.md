# App Tools — Documentation

## What This Is

A collection of Excel automation tools built as Python/xlwings add-ins.
Each tool is a self-contained Python package that connects to an Excel workbook
via xlwings. Logic runs locally on the user's machine — no data ever leaves.
Works identically on Mac and Windows.

---

## Tools

| Package | Description | Status |
|---|---|---|
| `zca_recon` | Common Area reconciliation — import source CSV, write status, export filtered CSVs | ✅ Live |

*Add new tools here as they are built.*

---

## Repo Structure

```
app-tools/
├── zca_recon/                ← Common Area reconciliation
│   ├── __init__.py
│   └── recon.py
├── scripts/
│   ├── setup_windows.bat     ← One-time setup — Windows
│   └── setup_mac.sh          ← One-time setup — Mac
├── docs/
│   └── overview.md           ← This file
├── .gitignore
├── CHANGELOG.md
└── pyproject.toml            ← Installs all tools in one shot
```

When a new tool is added, a new folder appears at the root level:
```
├── zp_users_recon/           ← example: future Users reconciliation tool
│   ├── __init__.py
│   └── recon.py
```
No changes to setup scripts needed — `pip install -e .` picks it up automatically.

---

## Installation (one time per machine)

**Windows:** run `scripts/setup_windows.bat`
**Mac:** run `bash scripts/setup_mac.sh`

This will:
1. Clone the repo to `C:\AppTools\app-tools` (Win) or `~/AppTools/app-tools` (Mac)
2. Install pip dependencies (`xlwings`, `pandas`)
3. Install all tools via `pip install -e .`
4. Install the xlwings Excel add-in
5. Schedule a silent daily auto-update at 8am

### Updating
Updates pull automatically every morning.
Manual update: `C:\AppTools\update.bat` (Win) or `~/AppTools/update.sh` (Mac).

---

## Adding a New Tool

1. Create a new folder at the repo root: `your_tool_name/`
2. Add `__init__.py` and your logic file
3. Export your entry points in `__init__.py`
4. Document it in this file under the Tools table
5. Add a VBA stub to the workbook
6. Push — users get it automatically next morning

---

## Excel VBA Stubs

Each tool gets its own small VBA stub in the workbook. That's all the VBA needed.

```vba
' ── Common Area Reconciliation ──────────────────────────────
Sub ZCA_RunReconciliation()
    RunPython "import zca_recon; zca_recon.run_reconciliation()"
End Sub
Sub ZCA_ExportUpdate()
    RunPython "import zca_recon; zca_recon.export_update()"
End Sub
Sub ZCA_ExportAdd()
    RunPython "import zca_recon; zca_recon.export_add()"
End Sub

' ── Add stubs for new tools below ───────────────────────────
```

---

## Tool Detail: zca_recon

### What it does
Reads a source system export CSV, compares each extension number against the
tracking sheet, writes status/date/package back, color-codes rows, and produces
filtered export CSVs for provisioning.

### Entry points
| Function | What It Does |
|---|---|
| `run_reconciliation()` | Full flow: import CSV → reconcile → offer exports |
| `export_update()` | Export Discrepancy/Partial rows (exist in source, wrong data) |
| `export_add()` | Export Not Found in CSV rows (not in source system yet) |

### Sheet requirements
Sheet must be named **"Common Area"** with these columns:

| Column | Purpose |
|---|---|
| Extension Number | Lookup key matched against source CSV |
| Common Area Status | Written by tool: `Complete` or `Setup` |
| Common Area (Last Update) | Timestamp of last run |
| Common Area Package | Synced from source CSV |
| Data Source | `Source CSV`, `Sheet Only`, or `Manual` |
| Data Status | `Verified`, `Discrepancy`, `Partial`, `Not Found in CSV` |

### Status logic
| Condition | Status | Data Status |
|---|---|---|
| All key fields match | Complete | Verified |
| Display name or site name matches | Setup | Discrepancy |
| Extension found, no fields match | Setup | Partial |
| Extension not in source CSV | Setup | Not Found in CSV |

### Key fields compared
- Display Name
- Site Name
- Phone Number (Zoom Temp vs CSV Phone Number)
- Outbound Caller ID (Zoom Temp vs CSV Outbound Caller ID)
- Desk Phone 1's Brand

### Export columns (24)
```
Display Name, Package, Site Name, Site Code, Common Area Template,
Language, Department, Cost Center, Extension Number,
Phone Number, Outbound Caller ID, Select Outbound Caller ID,
Desk Phone 1-3: Brand, Model, MAC Address, Provision Template
```

### Dashboard integration
Writes timestamp next to `ZP CA Last Update:` label on Dashboard sheet.
Falls back to cell `J19` if label not found.
