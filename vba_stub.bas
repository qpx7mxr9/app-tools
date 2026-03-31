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

    ' ── Windows paths ─────────────────────────────────────
    p = Environ("USERPROFILE") & "\app-tools"
    If Dir(p, vbDirectory) <> "" Then PyPath = p : Exit Function

    p = Environ("USERPROFILE") & "\Documents\GitHub\app-tools"
    If Dir(p, vbDirectory) <> "" Then PyPath = p : Exit Function

    ' ── Mac paths ─────────────────────────────────────────
    p = Environ("HOME") & "/app-tools"
    If Dir(p, vbDirectory) <> "" Then PyPath = p : Exit Function

    p = Environ("HOME") & "/Documents/GitHub/app-tools"
    If Dir(p, vbDirectory) <> "" Then PyPath = p : Exit Function

    MsgBox "Could not find app-tools folder. " & vbCrLf & _
           "Expected: " & Environ("USERPROFILE") & "\app-tools", _
           vbExclamation, "App Tools"
    PyPath = ""
End Function

Private Sub XRun(code As String)
    Dim p As String : p = Replace(PyPath(), "\", "/")
    If p = "" Then Exit Sub
    Dim full As String
    full = "import sys; sys.path.insert(0, '" & p & "'); " & code
    Application.Run "xlwings.RunPython", full
End Sub

' ── Dashboard ─────────────────────────────────────────────
Sub Dashboard_Build()
    XRun "import dashboard; dashboard.build_dashboard()"
End Sub

Sub Dashboard_Refresh()
    XRun "import dashboard; dashboard.refresh_ca_block()"
End Sub

' ── Common Areas ──────────────────────────────────────────
Sub ZCA_RunReconciliation()
    XRun "import zca_recon; zca_recon.run_reconciliation()"
End Sub

Sub ZCA_ExportUpdate()
    XRun "import zca_recon; zca_recon.export_update()"
End Sub

Sub ZCA_ExportAdd()
    XRun "import zca_recon; zca_recon.export_add()"
End Sub

' ── Zoom User Recon ───────────────────────────────────────
Sub ZUR_RunAudit()
    XRun "from zoom_user_recon.recon import run_zoom_user_audit; run_zoom_user_audit()"
End Sub

Sub ZUR_ClearResults()
    XRun "from zoom_user_recon.recon import clear_zoom_results; clear_zoom_results()"
End Sub

' ── ZP User Recon ─────────────────────────────────────────
Sub ZPU_RunReconciliation()
    XRun "from zp_user_recon.recon import run_zp_reconciliation; run_zp_reconciliation()"
End Sub
