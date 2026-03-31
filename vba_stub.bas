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
    ' Returns the real user home, bypassing Mac App Store sandbox container path
    Dim h As String
    h = Environ("HOME")
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

Private Sub SetupMacConf(p As String)
    ' Write a clean xlwings.conf to Excel's container HOME (Environ("HOME") on Mac).
    ' xlwings reads from this exact path on Mac via GetMacDir("$HOME", False).
    ' This fixes the "EOL while scanning string literal" error caused by a trailing
    ' newline that gets embedded inside the AppleScript shell command quotes.
    Dim confPath As String
    confPath = Environ("HOME") & "/xlwings.conf"

    ' Preserve INTERPRETER_MAC if already set
    Dim interpLine As String
    interpLine = ""
    Dim fileNum As Integer
    If Dir(confPath) <> "" Then
        Dim oneLine As String
        fileNum = FreeFile
        Open confPath For Input As #fileNum
        Do While Not EOF(fileNum)
            Line Input #fileNum, oneLine
            If InStr(LCase(oneLine), "interpreter_mac") > 0 Then
                interpLine = oneLine
            End If
        Loop
        Close #fileNum
    End If

    ' Write clean file — Print # adds exactly one newline, no trailing garbage
    fileNum = FreeFile
    Open confPath For Output As #fileNum
    If interpLine <> "" Then Print #fileNum, interpLine
    Print #fileNum, """PYTHONPATH"",""" & p & """"
    Close #fileNum
End Sub

Private Sub XRun(code As String)
    Dim p As String
    p = PyPath()
    If p = "" Then Exit Sub

    #If Mac Then
        ' On Mac: write clean xlwings.conf so xlwings injects PYTHONPATH itself.
        ' Cannot use semicolons in the code string — xlwings uses them as internal
        ' delimiters and they corrupt the AppleScript shell command quoting.
        SetupMacConf p
        Application.Run "xlwings.RunPython", code
    #Else
        ' On Windows: inject path directly via sys.path.insert (semicolons are fine).
        ' Replace backslashes with forward slashes so Python does not treat \U, \D etc.
        ' as unicode/escape sequences inside the string literal.
        Dim pFwd As String
        pFwd = Replace(p, "\", "/")
        Application.Run "xlwings.RunPython", _
            "import sys; sys.path.insert(0, '" & pFwd & "'); " & code
    #End If
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
