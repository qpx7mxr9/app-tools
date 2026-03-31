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

    ' ── Windows paths ─────────────────────────────────────────
    p = Environ("USERPROFILE") & "\app-tools"
    If Dir(p, vbDirectory) <> "" Then PyPath = p : Exit Function

    p = Environ("USERPROFILE") & "\Documents\GitHub\app-tools"
    If Dir(p, vbDirectory) <> "" Then PyPath = p : Exit Function

    ' ── Mac paths ─────────────────────────────────────────────
    p = Environ("HOME") & "/app-tools"
    If Dir(p, vbDirectory) <> "" Then PyPath = p : Exit Function

    p = Environ("HOME") & "/Documents/GitHub/app-tools"
    If Dir(p, vbDirectory) <> "" Then PyPath = p : Exit Function

    MsgBox "Could not find app-tools folder. " & vbCrLf & _
           "Expected: " & Environ("USERPROFILE") & "\app-tools", _
           vbExclamation, "App Tools"
    PyPath = ""
End Function

Private Sub RunTool(code As String)
    Dim p As String
    p = PyPath()
    If p = "" Then Exit Sub
    ' Normalize path separator for cross-platform
    p = Replace(p, "\", "/")
    RunPython "import sys, os; sys.path.insert(0, '" & p & "'); " & code
End Sub

' ── Dashboard ─────────────────────────────────────────────
Sub Dashboard_Build()
    RunTool "import dashboard; dashboard.build_dashboard()"
End Sub

Sub Dashboard_Refresh()
    RunTool "import dashboard; dashboard.refresh_ca_block()"
End Sub

' ── Common Areas ──────────────────────────────────────────
Sub ZCA_RunReconciliation()
    RunTool "import zca_recon; zca_recon.run_reconciliation()"
End Sub

Sub ZCA_ExportUpdate()
    RunTool "import zca_recon; zca_recon.export_update()"
End Sub

Sub ZCA_ExportAdd()
    RunTool "import zca_recon; zca_recon.export_add()"
End Sub

' ── Zoom User Recon ───────────────────────────────────────
Sub ZUR_RunAudit()
    RunTool "import zoom_user_recon; zoom_user_recon.run_zoom_user_audit()"
End Sub

Sub ZUR_ClearResults()
    RunTool "import zoom_user_recon; zoom_user_recon.clear_zoom_results()"
End Sub

' ── ZP User Recon ─────────────────────────────────────────
Sub ZPU_RunReconciliation()
    RunTool "import zp_user_recon; zp_user_recon.run_zp_reconciliation()"
End Sub
