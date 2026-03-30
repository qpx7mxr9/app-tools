' ============================================================
' App Tools — VBA Module
' Paste this entire module into the workbook's VBA editor.
' This is the ONLY VBA needed. All logic lives in Python.
'
' Setup: pip install xlwings pandas && xlwings addin install
' Repo:  https://github.com/qpx7mxr9/app-tools
' ============================================================

Private Function PyPath() As String
    Dim p As String
    p = Environ("HOME") & "/AppTools/app-tools"
    If Dir(p, vbDirectory) = "" Then
        p = Environ("HOME") & "/Documents/GitHub/app-tools"
    End If
    PyPath = p
End Function

' ── CA Tools Dashboard ────────────────────────────────────
Sub Dashboard_Build()
    RunPython "import sys, os; sys.path.insert(0, os.path.expanduser('~/Documents/GitHub/app-tools')); import dashboard; dashboard.build_dashboard()"
End Sub

Sub Dashboard_Refresh()
    RunPython "import sys, os; sys.path.insert(0, os.path.expanduser('~/Documents/GitHub/app-tools')); import dashboard; dashboard.refresh_ca_block()"
End Sub

' ── Common Areas ─────────────────────────────────────────
Sub ZCA_RunReconciliation()
    RunPython "import sys, os; sys.path.insert(0, os.path.expanduser('~/Documents/GitHub/app-tools')); import zca_recon; zca_recon.run_reconciliation()"
End Sub

Sub ZCA_ExportUpdate()
    RunPython "import sys, os; sys.path.insert(0, os.path.expanduser('~/Documents/GitHub/app-tools')); import zca_recon; zca_recon.export_update()"
End Sub

Sub ZCA_ExportAdd()
    RunPython "import sys, os; sys.path.insert(0, os.path.expanduser('~/Documents/GitHub/app-tools')); import zca_recon; zca_recon.export_add()"
End Sub

' ── Zoom User Recon ──────────────────────────────────────
Sub ZUR_RunAudit()
    RunPython "import sys, os; sys.path.insert(0, os.path.expanduser('~/Documents/GitHub/app-tools')); import zoom_user_recon; zoom_user_recon.run_zoom_user_audit()"
End Sub

Sub ZUR_ClearResults()
    RunPython "import sys, os; sys.path.insert(0, os.path.expanduser('~/Documents/GitHub/app-tools')); import zoom_user_recon; zoom_user_recon.clear_zoom_results()"
End Sub

' ── ZP User Recon ─────────────────────────────────────────
Sub ZPU_RunReconciliation()
    RunPython "import sys, os; sys.path.insert(0, os.path.expanduser('~/Documents/GitHub/app-tools')); import zp_user_recon; zp_user_recon.run_zp_reconciliation()"
End Sub
