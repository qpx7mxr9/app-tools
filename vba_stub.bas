' ============================================================
' App Tools — VBA Module
' Paste this entire module into the workbook's VBA editor.
' This is the ONLY VBA needed. All logic lives in Python.
'
' Setup: pip install xlwings pandas && xlwings addin install
' Repo:  https://github.com/qpx7mxr9/app-tools
' ============================================================

Private Function FolderExists(p As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(p) And vbDirectory) = vbDirectory
    On Error GoTo 0
    If Not FolderExists Then
        FolderExists = (Dir(p & "/vba_stub.bas") <> "") Or _
                       (Dir(p & "\vba_stub.bas") <> "")
    End If
End Function

Private Function RealHome() As String
    Dim h As String
    h = Environ("HOME")
    ' Excel (App Store) runs sandboxed — HOME points to container, not real home
    If InStr(h, "/Library/Containers/") > 0 Or h = "" Then
        h = "/Users/" & Environ("USER")
    End If
    RealHome = h
End Function

Private Function PyPath() As String
    Dim p As String
    Dim isMac As Boolean
    isMac = (InStr(Application.OperatingSystem, "Mac") > 0)

    If isMac Then
        Dim home As String
        home = RealHome()
        p = home & "/app-tools"
        If FolderExists(p) Then PyPath = p: Exit Function
        p = home & "/Documents/GitHub/app-tools"
        If FolderExists(p) Then PyPath = p: Exit Function
        p = home & "/GitHub/app-tools"
        If FolderExists(p) Then PyPath = p: Exit Function
    Else
        p = Environ("USERPROFILE") & "\app-tools"
        If FolderExists(p) Then PyPath = p: Exit Function
        p = Environ("USERPROFILE") & "\Documents\GitHub\app-tools"
        If FolderExists(p) Then PyPath = p: Exit Function
        p = Environ("USERPROFILE") & "\GitHub\app-tools"
        If FolderExists(p) Then PyPath = p: Exit Function
    End If

    MsgBox "Could not find app-tools folder." & vbCrLf & "Last tried: " & p, _
           vbExclamation, "App Tools"
    PyPath = ""
End Function

Private Sub SetPythonPath(p As String)
    ' Write PYTHONPATH into xlwings.conf sheet so xlwings handles sys.path natively.
    ' Avoids injecting sys.path.insert into the code string, which breaks on Mac
    ' because xlwings uses semicolons as internal delimiters.
    Dim ws As Worksheet
    Dim found As Boolean
    Dim i As Long

    found = False
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "xlwings.conf" Then
            Set ws = ThisWorkbook.Sheets(i)
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "xlwings.conf"
        ws.Visible = xlSheetHidden
    End If

    For i = 1 To 20
        If ws.Cells(i, 1).Value = "PYTHONPATH" Then
            ws.Cells(i, 2).Value = p
            Exit Sub
        End If
    Next i
    ' Not found — add it
    For i = 1 To 20
        If ws.Cells(i, 1).Value = "" Then
            ws.Cells(i, 1).Value = "PYTHONPATH"
            ws.Cells(i, 2).Value = p
            Exit Sub
        End If
    Next i
End Sub

Private Sub XRun(code As String)
    Dim p As String
    p = PyPath()
    If p = "" Then Exit Sub
    SetPythonPath p
    Application.Run "xlwings.RunPython", code
End Sub

' -- Dashboard ---------------------------------------------
Sub Dashboard_Build()
    XRun "import dashboard; dashboard.build_dashboard()"
End Sub

Sub Dashboard_Refresh()
    XRun "import dashboard; dashboard.refresh_ca_block()"
End Sub

' -- Common Areas ------------------------------------------
Sub ZCA_RunReconciliation()
    XRun "import zca_recon; zca_recon.run_reconciliation()"
End Sub

Sub ZCA_ExportUpdate()
    XRun "import zca_recon; zca_recon.export_update()"
End Sub

Sub ZCA_ExportAdd()
    XRun "import zca_recon; zca_recon.export_add()"
End Sub

' -- Zoom User Recon ---------------------------------------
Sub ZUR_RunAudit()
    XRun "from zoom_user_recon.recon import run_zoom_user_audit; run_zoom_user_audit()"
End Sub

Sub ZUR_ClearResults()
    XRun "from zoom_user_recon.recon import clear_zoom_results; clear_zoom_results()"
End Sub

' -- ZP User Recon -----------------------------------------
Sub ZPU_RunReconciliation()
    XRun "from zp_user_recon.recon import run_zp_reconciliation; run_zp_reconciliation()"
End Sub
