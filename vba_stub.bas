' ============================================================
' App Tools — VBA Module
' Paste this entire module into the workbook's VBA editor.
' This is the ONLY VBA needed. All logic lives in Python.
'
' Setup: pip install xlwings pandas && xlwings addin install
' Repo:  https://github.com/qpx7mxr9/app-tools
' ============================================================

' ── Dashboard ────────────────────────────────────────────────
Sub Dashboard_Build()
    RunPython "import sys; sys.path.insert(0,'~/AppTools/app-tools'); import dashboard; dashboard.build_dashboard()"
End Sub

Sub Dashboard_Refresh()
    RunPython "import sys; sys.path.insert(0,'~/AppTools/app-tools'); import dashboard; dashboard.refresh_ca_block()"
End Sub

' ── Common Areas ─────────────────────────────────────────────
Sub ZCA_RunReconciliation()
    RunPython "import zca_recon; zca_recon.run_reconciliation()"
End Sub

Sub ZCA_ExportUpdate()
    RunPython "import zca_recon; zca_recon.export_update()"
End Sub

Sub ZCA_ExportAdd()
    RunPython "import zca_recon; zca_recon.export_add()"
End Sub

' ── Add stubs for new tools below as they are built ──────────
