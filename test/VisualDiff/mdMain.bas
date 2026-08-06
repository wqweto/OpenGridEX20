Attribute VB_Name = "mdMain"
'=========================================================================
'
' Open GridEX 2000 Control
' VisualDiff runner: record goldens from the original control, self-test
' the harness (original vs recorded golden) or verify our control against
' the golden corpus. Usage: VisualDiff.exe [record|selftest|verify] [mask]
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdMain"

'=========================================================================
' API
'=========================================================================

Private Const GWL_STYLE                     As Long = -16
Private Const WS_HSCROLL                    As Long = &H100000
Private Const WS_VSCROLL                    As Long = &H200000
Private Const WS_TABSTOP                    As Long = &H10000
Private Const SB_CTL                         As Long = 2
Private Const SIF_ALL                        As Long = &H17
Private Const OBJID_CLIENT                   As Long = &HFFFFFFFC
Private Const EM_GETSEL                      As Long = &HB0
Private Const EM_GETLINECOUNT                As Long = &HBA
Private Const EM_GETFIRSTVISIBLELINE         As Long = &HCE

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpsi As SCROLLINFO) As Long
Private Declare Function GetScrollBarInfo Lib "user32" (ByVal hWnd As Long, ByVal idObject As Long, psbi As SCROLLBARINFO) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

Private Type SCROLLINFO
    cbSize                  As Long
    fMask                   As Long
    nMin                    As Long
    nMax                    As Long
    nPage                   As Long
    nPos                    As Long
    nTrackPos               As Long
End Type

Private Type SCROLLBARINFO
    cbSize                  As Long
    rcScrollBar             As RECT
    dxyLineButton           As Long
    xyThumbTop              As Long
    xyThumbBottom           As Long
    reserved                As Long
    rgstate0                As Long
    rgstate1                As Long
    rgstate2                As Long
    rgstate3                As Long
    rgstate4                As Long
    rgstate5                As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_PROGID_ORIGINAL           As String = "GridEX20.GridEX"
Private Const STR_PROGID_OURS               As String = "OpenGridEX20.GridEX"
Private Const IGNORE_EVENT_GOLDEN           As Boolean = False

'=========================================================================
' Error management
'=========================================================================

Private Sub PrintError(sFunction As String)
    PopPrintError PushError, MODULE_NAME, sFunction
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Const FUNC_NAME     As String = "Main"
    Dim aArgs()         As String
    Dim sMode           As String
    Dim sMask           As String
    Dim sProgId         As String
    Dim vFile           As Variant
    Dim sName           As String
    Dim vDoc            As Variant
    Dim oForm           As frmHost
    Dim baBits()        As Byte
    Dim baGolden()      As Byte
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lGoldenW        As Long
    Dim lGoldenH        As Long
    Dim lDiff           As Long
    Dim sReport         As String
    Dim sGolden         As String
    Dim sEvents         As String
    Dim lDpi            As Long
    Dim sDpi            As String
    Dim sError          As String

    On Error GoTo EH
    '--- goldens are per DPI: the embedded manifest makes the harness run at
    '--- the real system DPI, while __COMPAT_LAYER=DPIUNAWARE forces the
    '--- 96dpi (OS-virtualized) mode from the same executable
    lDpi = ScreenDpi()
    sDpi = "\" & lDpi
    aArgs = Split(Command$)
    If UBound(aArgs) >= 0 Then
        sMode = LCase$(aArgs(0))
    End If
    If LenB(sMode) = 0 Then
        sMode = "verify"
    End If
    If UBound(aArgs) >= 1 Then
        sMask = aArgs(1)
    Else
        sMask = "*"
    End If
    TestInit App.Path & "\VisualDiff.out.txt"
    If sMode = "verify" Or sMode = "scrollinfo-ours" Or sMode = "dump-ours" Or sMode = "windows-ours" Then
        sProgId = STR_PROGID_OURS
    Else
        sProgId = STR_PROGID_ORIGINAL
    End If
    If sMode = "scrollinfo" Or sMode = "scrollinfo-ours" Then
        For Each vFile In EnumFiles(App.Path & "\scenarios", "*.json")
            sName = Mid$(vFile, InStrRev(vFile, "\") + 1)
            sName = Left$(sName, Len(sName) - Len(".json"))
            If sName Like sMask Then
                JsonParse ReadTextFile(CStr(vFile)), vDoc
                Set oForm = New frmHost
                If oForm.RunScenario(sProgId, C2Obj(vDoc), baBits, lWidth, lHeight) Then
                    Assert "--- " & sProgId & " / " & sName & " ---", True
                    pvDumpScrollInfo oForm.hWnd, oForm.hWnd
                End If
                Unload oForm
                Set oForm = Nothing
            End If
        Next
        TestsDone
        Exit Sub
    End If
    If sMode = "windows" Or sMode = "windows-ours" Then
        For Each vFile In EnumFiles(App.Path & "\scenarios", "*.json")
            sName = Mid$(vFile, InStrRev(vFile, "\") + 1)
            sName = Left$(sName, Len(sName) - Len(".json"))
            If sName Like sMask Then
                JsonParse ReadTextFile(CStr(vFile)), vDoc
                Set oForm = New frmHost
                If oForm.RunScenario(sProgId, C2Obj(vDoc), baBits, lWidth, lHeight) Then
                    Assert "--- " & sProgId & " / " & sName & " ---", True
                    Assert "hWnd property returns &H" & Hex$(oForm.ControlHwnd()), True
                    pvDumpWindows oForm.hWnd, oForm.hWnd, 0
                End If
                Unload oForm
                Set oForm = Nothing
            End If
        Next
        TestsDone
        Exit Sub
    End If
    If sMode = "metrics" Then
        Assert "SM_CYHSCROLL=" & GetSystemMetrics(3) & " SM_CXVSCROLL=" & GetSystemMetrics(2) & sDpi & "dpi", True
        Assert FontMetrics("MS Sans Serif", 8) & sDpi & "dpi", True
        Assert FontMetrics("Tahoma", 8) & sDpi & "dpi", True
        Assert FontMetrics("Tahoma", 12) & sDpi & "dpi", True
        Assert FontMetrics("Segoe UI", 9) & sDpi & "dpi", True
        TestsDone
        Exit Sub
    End If
    If Not pvCanCreate(sProgId) Then
        TestSkip "control not available: " & sProgId
        Exit Sub
    End If
    For Each vFile In EnumFiles(App.Path & "\scenarios", "*.json")
        sName = Mid$(vFile, InStrRev(vFile, "\") + 1)
        sName = Left$(sName, Len(sName) - Len(".json"))
        If sName Like sMask Then
            JsonParse ReadTextFile(CStr(vFile)), vDoc
            Set oForm = New frmHost
            If Not oForm.RunScenario(sProgId, C2Obj(vDoc), baBits, lWidth, lHeight) Then
                Assert sMode & " " & sName & " (capture failed)", False
            Else
                sGolden = App.Path & "\golden" & sDpi & "\" & sName & ".png"
                Select Case sMode
                Case "record"
                    pvEnsureDir App.Path & "\golden" & sDpi
                    Assert "record " & sName & sDpi & "dpi", SavePng(sGolden, lWidth, lHeight, baBits)
                    '--- a scenario that drives input records what the control
                    '--- raised as well: the order of the edit events is a
                    '--- contract a picture cannot hold
                    If Not C2Obj(JsonValue(C2Obj(vDoc), "input")) Is Nothing Then
                        WriteTextFile Left$(sGolden, Len(sGolden) - Len(".png")) & ".events.txt", oForm.EventLog
                    End If
                Case "dump", "dump-ours"
                    pvEnsureDir App.Path & "\output" & sDpi
                    WriteTextFile App.Path & "\output" & sDpi & "\" & sName & ".dump.json", oForm.DumpState(True)
                    Assert "dump " & sName & sDpi & "dpi", True
                Case "selftest", "verify"
                    '--- every capture is written out, not just the ones that
                    '--- differ: output\<dpi>\ is then a full contact sheet of
                    '--- what the control renders, which is what you want when
                    '--- looking for something a golden cannot express
                    pvEnsureDir App.Path & "\output" & sDpi
                    SavePng App.Path & "\output" & sDpi & "\" & sName & ".actual.png", lWidth, lHeight, baBits
                    If Not LoadPng(sGolden, lGoldenW, lGoldenH, baGolden) Then
                        Assert sMode & " " & sName & sDpi & "dpi (no golden)", False
                    ElseIf lGoldenW <> lWidth Or lGoldenH <> lHeight Then
                        Assert sMode & " " & sName & sDpi & "dpi (golden size mismatch: golden " & lGoldenW & "x" & lGoldenH & " actual " & lWidth & "x" & lHeight & ")", False
                    Else
                        lDiff = DiffBits(baGolden, baBits, lWidth, lHeight, sReport)
                        Assert sMode & " " & sName & sDpi & "dpi" & IIf(lDiff > 0, " diff=" & lDiff & "px " & sReport, vbNullString), lDiff = 0
                    End If
                    If Not C2Obj(JsonValue(C2Obj(vDoc), "input")) Is Nothing Then
                        sEvents = Left$(sGolden, Len(sGolden) - Len(".png")) & ".events.txt"
                        WriteTextFile App.Path & "\output" & sDpi & "\" & sName & ".events.txt", oForm.EventLog
                        '--- a scenario whose gesture only lands at one of the
                        '--- two scales has no log recorded at the other, and
                        '--- pixels are all it is asked about there
                        If Not IGNORE_EVENT_GOLDEN And FileExists(sEvents) Then
                            Assert sMode & " " & sName & sDpi & "dpi events" & pvEventReport(ReadTextFile(sEvents), oForm.EventLog), _
                                ReadTextFile(sEvents) = oForm.EventLog
                        End If
                    End If
                End Select
            End If
            Unload oForm
            Set oForm = Nothing
        End If
    Next
    TestsDone
    Exit Sub
EH:
    '--- PrintError resets Err, so the assert has to read it first
    sError = "&H" & Hex$(Err.Number) & " " & Err.Description
    PrintError FUNC_NAME
    Assert "Unhandled error " & sError & " in " & sName, False
    TestsDone
End Sub

Private Function pvEventReport(sGolden As String, sActual As String) As String
    Dim vGolden         As Variant
    Dim vActual         As Variant
    Dim lIdx            As Long

    '--- the first line that differs is what a reader needs, not the two logs
    If sGolden = sActual Then
        Exit Function
    End If
    vGolden = Split(sGolden, vbCrLf)
    vActual = Split(sActual, vbCrLf)
    For lIdx = 0 To UBound(vGolden)
        If lIdx > UBound(vActual) Then
            pvEventReport = " (missing at " & lIdx + 1 & ": " & vGolden(lIdx) & ")"
            Exit Function
        End If
        If vGolden(lIdx) <> vActual(lIdx) Then
            pvEventReport = " (line " & lIdx + 1 & ": expected " & vGolden(lIdx) & " got " & vActual(lIdx) & ")"
            Exit Function
        End If
    Next
    pvEventReport = " (extra at " & UBound(vGolden) + 2 & ": " & vActual(UBound(vGolden) + 1) & ")"
End Function

Private Sub pvDumpScrollInfo(ByVal hWndRoot As Long, ByVal hWnd As Long)
    Dim hChild          As Long
    Dim sClass          As String
    Dim uSi             As SCROLLINFO
    Dim uSbi            As SCROLLBARINFO
    Dim uWin            As RECT
    Dim uRoot           As RECT

    '--- the original's horizontal bar is a child window, so its range and
    '--- its thumb rect can be read back rather than guessed from pixels
    hChild = GetWindow(hWnd, GW_CHILD)
    Do While hChild <> 0
        sClass = String$(256, 0)
        sClass = Left$(sClass, GetClassName(hChild, StrPtr(sClass), 256))
        If InStr(1, sClass, "ScrollBar", vbTextCompare) > 0 Then
            GetWindowRect hWndRoot, uRoot
            GetWindowRect hChild, uWin
            uSi.cbSize = Len(uSi)
            uSi.fMask = SIF_ALL
            GetScrollInfo hChild, SB_CTL, uSi
            uSbi.cbSize = Len(uSbi)
            GetScrollBarInfo hChild, OBJID_CLIENT, uSbi
            Assert sClass & " at (" & uWin.Left - uRoot.Left & "," & uWin.Top - uRoot.Top & ") " & _
                uWin.Right - uWin.Left & "x" & uWin.Bottom - uWin.Top & _
                " | min=" & uSi.nMin & " max=" & uSi.nMax & " page=" & uSi.nPage & " pos=" & uSi.nPos & _
                " | thumb=" & uSbi.xyThumbTop & ".." & uSbi.xyThumbBottom & " linebtn=" & uSbi.dxyLineButton, True
        End If
        pvDumpScrollInfo hWndRoot, hChild
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Loop
End Sub

Private Sub pvDumpWindows(ByVal hWndRoot As Long, ByVal hWnd As Long, ByVal lLevel As Long)
    Dim hChild          As Long
    Dim sClass          As String
    Dim uWin            As RECT
    Dim uRoot           As RECT
    Dim uClient         As RECT
    Dim lStyle          As Long
    Dim sFlags          As String
    Dim sText           As String
    Dim lSelStart       As Long
    Dim lSelEnd         As Long

    hChild = GetWindow(hWnd, GW_CHILD)
    Do While hChild <> 0
        sClass = String$(256, 0)
        Call GetClassName(hChild, StrPtr(sClass), 256)
        sClass = Left$(sClass, InStr(sClass, vbNullChar) - 1)
        Call GetWindowRect(hChild, uWin)
        Call GetWindowRect(hWndRoot, uRoot)
        Call GetClientRect(hChild, uClient)
        lStyle = GetWindowLong(hChild, GWL_STYLE)
        sFlags = vbNullString
        If (lStyle And WS_HSCROLL) <> 0 Then
            sFlags = sFlags & " WS_HSCROLL"
        End If
        If (lStyle And WS_VSCROLL) <> 0 Then
            sFlags = sFlags & " WS_VSCROLL"
        End If
        If (lStyle And WS_TABSTOP) <> 0 Then
            sFlags = sFlags & " WS_TABSTOP"
        End If
        '--- an editor's text and caret say what a picture of it cannot: which
        '--- character the click that opened it landed on
        sText = String$(1024, 0)
        sText = Left$(sText, GetWindowText(hChild, StrPtr(sText), 1024))
        If LenB(sText) <> 0 Then
            Call SendMessage(hChild, EM_GETSEL, VarPtr(lSelStart), VarPtr(lSelEnd))
            sText = " text=""" & sText & """ sel=" & lSelStart & ".." & lSelEnd & _
                " lines=" & SendMessage(hChild, EM_GETLINECOUNT, 0, 0) & " top=" & SendMessage(hChild, EM_GETFIRSTVISIBLELINE, 0, 0)
        End If
        Assert Space$(lLevel * 2) & "&H" & Hex$(hChild) & " [" & sClass & "] at (" & uWin.Left - uRoot.Left & "," & uWin.Top - uRoot.Top & ") " & _
            uWin.Right - uWin.Left & "x" & uWin.Bottom - uWin.Top & " client " & uClient.Right & "x" & uClient.Bottom & sFlags & sText, True
        pvDumpWindows hWndRoot, hChild, lLevel + 1
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Loop
End Sub

Private Sub pvEnsureDir(sPath As String)
    Dim sParent         As String

    sParent = Left$(sPath, InStrRev(sPath, "\") - 1)
    If LenB(Dir$(sParent, vbDirectory)) = 0 Then
        MkDir sParent
    End If
    If LenB(Dir$(sPath, vbDirectory)) = 0 Then
        MkDir sPath
    End If
End Sub

Private Function pvCanCreate(sProgId As String) As Boolean
    Const FUNC_NAME     As String = "pvCanCreate"
    Dim oForm           As frmHost

    '--- no Licenses.Add needed: RemoveUnusedControlInfo=0 keeps the design
    '--- license of the referenced original OCX compiled into the harness
    On Error GoTo EH
    Set oForm = New frmHost
    oForm.Controls.Add sProgId, "ctlProbe"
    pvCanCreate = True
QH:
    If Not oForm Is Nothing Then
        Unload oForm
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    GoTo QH
End Function
