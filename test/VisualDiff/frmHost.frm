VERSION 5.00
Begin VB.Form frmHost 
   Caption         =   "VisualDiff host"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   621
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   879
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' Scenario host: creates the control under test at runtime (late bound,
' no compile-time OCX reference), applies scenario props through the
' shared import engine, feeds unbound data via ObjectEvent and captures
' the client area to 24bpp DIB bits
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private WithEvents m_oExt           As VBControlExtender
Attribute m_oExt.VB_VarHelpID = -1
Private m_oUnboundRows              As Object
Private m_sClass                    As String
Private m_sEventLog                 As String
Private m_bLogEvents                As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Get EventLog() As String
    EventLog = m_sEventLog
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function RunScenario(sProgId As String, oScenario As Object, baBits() As Byte, lWidth As Long, lHeight As Long) As Boolean
    Const FUNC_NAME     As String = "RunScenario"
    Dim sApply          As String
    Dim lHwnd           As Long

    On Error GoTo EH
    lWidth = C2Lng(JsonValue(oScenario, "width"))
    If lWidth = 0 Then
        lWidth = 400
    End If
    lHeight = C2Lng(JsonValue(oScenario, "height"))
    If lHeight = 0 Then
        lHeight = 260
    End If
    Move Left, Top, Width - ScaleWidth * Screen.TwipsPerPixelX + lWidth * Screen.TwipsPerPixelX, Height - ScaleHeight * Screen.TwipsPerPixelY + lHeight * Screen.TwipsPerPixelY
    Set m_oUnboundRows = C2Obj(JsonValue(oScenario, "unbound/rows"))
    sApply = C2Str(JsonValue(oScenario, "apply"))
    If sApply = "after-show" Then
        Show
        DoEvents
        pvCreateAndApply sProgId, oScenario
    Else
        pvCreateAndApply sProgId, oScenario
        Show
    End If
    DoEvents
    '--- data rows are always fed after the display initialized: neither
    '--- control can be switched to a data mode before it has a window
    If Not m_oUnboundRows Is Nothing Then
        pvFeedData sProgId
        DoEvents
    End If
    '--- runtime props applied after the data feed (e.g. FirstItem)
    If Not C2Obj(JsonValue(oScenario, "post")) Is Nothing Then
        ImportObject m_oExt.Object, m_sClass, C2Obj(JsonValue(oScenario, "post"))
        '--- the original repaints some runtime property changes off a timer,
        '--- so two identical captures can be taken before it ever fires --
        '--- that raced the golden recording of post-bearing scenarios
        pvSettle 10
    End If
    '--- methods and indexed properties a scenario needs after the data is
    '--- in, e.g. CollapseAll or RowExpanded(2) = False -- neither shape can
    '--- be expressed as a plain property in the props/post blocks
    If Not C2Obj(JsonValue(oScenario, "calls")) Is Nothing Then
        pvRunCalls C2Obj(JsonValue(oScenario, "calls"))
        pvSettle 10
    End If
    '--- row selection: select the listed 1-based row positions
    If Not C2Obj(JsonValue(oScenario, "select")) Is Nothing Then
        pvSelectRows C2Obj(JsonValue(oScenario, "select"))
        DoEvents
    End If
    '--- the input a scenario drives -- clicks, keys and characters -- with
    '--- the event log cleared first so it holds the interaction alone and
    '--- not the flood of reads the data feed above raised
    m_sEventLog = vbNullString
    m_bLogEvents = True
    If Not C2Obj(JsonValue(oScenario, "input")) Is Nothing Then
        pvRunInput C2Obj(JsonValue(oScenario, "input"))
        pvSettle 10
    End If
    m_bLogEvents = False
    '--- always the host form client, never the control window: the two
    '--- controls do not expose the same window from hWnd -- the original's
    '--- is an inner grid window sized inside its own chrome, ours is the
    '--- whole UserControl -- while the form client holds each of them plus
    '--- their scrollbars and record navigator on identical coordinates
    RunScenario = pvCaptureStable(hWnd, lWidth, lHeight, baBits)
    Exit Function
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Function

Public Function ControlHwnd() As Long
    ControlHwnd = pvControlHwnd()
End Function

Public Function DumpState(Optional ByVal Runtime As Boolean) As String
    '--- runtime included on request: RowCount, Row and the rest only exist
    '--- once the control is populated, and that is exactly what a model
    '--- question like "does RowCount count group rows" needs
    DumpState = SnapshotToJson(m_oExt.Object, m_sClass, Runtime)
End Function

Private Function pvControlHwnd() As Long
    Const FUNC_NAME     As String = "pvControlHwnd"

    On Error GoTo EH
    pvControlHwnd = C2Lng(CallByName(m_oExt.Object, "hWnd", VbGet))
    Exit Function
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Function

Private Sub pvSettle(ByVal lRetries As Long)
    Dim lIdx            As Long

    '--- pumped a millisecond at a time so a timer tick is picked up as soon
    '--- as it fires rather than at the end of a coarse sleep
    For lIdx = 1 To lRetries
        Sleep 1
        DoEvents
    Next
End Sub

Private Function pvCaptureStable(ByVal lHwnd As Long, lWidth As Long, lHeight As Long, baBits() As Byte) As Boolean
    Dim lRetry          As Long
    Dim baPrev()        As Byte
    Dim sReport         As String

    '--- captures can race pending paints: capture repeatedly until two
    '--- consecutive shots are identical
    If Not CaptureWindowClient(lHwnd, lWidth, lHeight, baPrev) Then
        Exit Function
    End If
    For lRetry = 1 To 20
        Sleep 50
        DoEvents
        If Not CaptureWindowClient(lHwnd, lWidth, lHeight, baBits) Then
            Exit Function
        End If
        If DiffBits(baPrev, baBits, lWidth, lHeight, sReport) = 0 Then
            pvCaptureStable = True
            Exit Function
        End If
        baPrev = baBits
    Next
End Function

'=========================================================================
' Methods
'=========================================================================

Private Sub pvCreateAndApply(sProgId As String, oScenario As Object)
    Dim oProps          As Object

    m_sClass = Split(sProgId, ".")(1)
    Set m_oExt = Controls.Add(sProgId, "ctlGrid")
    m_oExt.Move 0, 0, ScaleWidth, ScaleHeight
    m_oExt.Visible = True
    pvSetDefaultFonts
    Set oProps = C2Obj(JsonValue(oScenario, "props"))
    If Not oProps Is Nothing Then
        pvApplyProps oProps
        '--- the original resets the column layout when its display first
        '--- initializes over the (empty) recordset; HoldFields preserves
        '--- the imported columns (samples persist this as MethodHoldFields)
        pvTryCall "HoldFields"
    End If
End Sub

Private Sub pvSetDefaultFonts()
    Const FUNC_NAME     As String = "pvSetDefaultFonts"

    '--- a control added at run-time inherits the host form ambient font, so
    '--- the scenario baseline is restored explicitly: the form itself is
    '--- Tahoma because ambient MS Sans Serif wrecks the original control's
    '--- layout at 144dpi, but the grid must still default to MS Sans Serif
    On Error GoTo EH
    CallByName m_oExt.Object, "Font", VbSet, pvNewFont()
    CallByName m_oExt.Object, "ColumnHeaderFont", VbSet, pvNewFont()
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Function pvNewFont() As StdFont
    Set pvNewFont = New StdFont
    pvNewFont.Name = "MS Sans Serif"
    pvNewFont.Size = 8
End Function

Private Sub pvFeedData(sProgId As String)
    Dim lItemCount      As Long

    '--- the original control cannot switch DataMode/ItemCount at runtime,
    '--- so it is fed a fabricated ADO recordset (columns preserved via
    '--- HoldFields); our control uses the unbound pipeline
    If LCase$(Split(sProgId, ".")(0)) = "opengridex20" Then
        lItemCount = C2Lng(JsonValue(m_oUnboundRows, "-1"))
        pvTrySet "DataMode", 99
        pvTrySet "ItemCount", lItemCount
        pvTryCall "Rebind"
    Else
        pvFeedAdoRows
    End If
End Sub

Private Sub pvFeedAdoRows()
    Const FUNC_NAME     As String = "pvFeedAdoRows"
    Const adVarChar     As Long = 200
    Dim oRs             As Object
    Dim oRow            As Object
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim nCol            As Integer
    Dim nColCount       As Integer

    On Error GoTo EH
    lCount = C2Lng(JsonValue(m_oUnboundRows, "-1"))
    Set oRow = C2Obj(JsonValue(m_oUnboundRows, 0))
    nColCount = C2Lng(JsonValue(oRow, "-1"))
    Set oRs = CreateObject("ADODB.Recordset")
    For nCol = 1 To nColCount
        oRs.Fields.Append "F" & nCol, adVarChar, 255
    Next
    oRs.Open
    For lIdx = 0 To lCount - 1
        Set oRow = C2Obj(JsonValue(m_oUnboundRows, lIdx))
        oRs.AddNew
        For nCol = 1 To nColCount
            oRs.Fields(nCol - 1).Value = C2Str(JsonValue(oRow, nCol - 1))
        Next
    Next
    oRs.Update
    oRs.MoveFirst
    '--- held columns display nothing unless they map to recordset fields
    Set oRow = C2Obj(CallByName(m_oExt.Object, "Columns", VbGet))
    For nCol = 1 To nColCount
        If nCol > C2Lng(CallByName(oRow, "Count", VbGet)) Then
            Exit For
        End If
        CallByName C2Obj(CallByName(oRow, "Item", VbGet, nCol)), "DataField", VbLet, "F" & nCol
    Next
    CallByName m_oExt.Object, "HoldFields", VbMethod
    CallByName m_oExt.Object, "ADORecordset", VbSet, oRs
    Exit Sub
EH:
    LogError "adofeed: &H" & Hex$(Err.Number) & " " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Sub pvApplyProps(oProps As Object)
    Const FUNC_NAME     As String = "pvApplyProps"

    On Error GoTo EH
    ImportObject m_oExt.Object, m_sClass, oProps
    Exit Sub
EH:
    '--- surface import errors in the runner log (Debug.Print is invisible
    '--- in the compiled exe)
    LogError "import: &H" & Hex$(Err.Number) & " " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Sub pvSelectRows(oList As Object)
    Const FUNC_NAME     As String = "pvSelectRows"
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim oGrid           As Object

    On Error GoTo EH
    lCount = C2Lng(JsonValue(oList, "-1"))
    If lCount = 0 Then
        Exit Sub
    End If
    Set oGrid = m_oExt.Object
    oGrid.MultiSelect = (lCount > 1)
    oGrid.SelectedItems.Clear
    For lIdx = 0 To lCount - 1
        oGrid.SelectedItems.Add C2Lng(JsonValue(oList, lIdx))
    Next
    oGrid.Refresh
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Function pvFormatEvent(Info As EventInfo) As String
    Const FUNC_NAME     As String = "pvFormatEvent"
    Dim lIdx            As Long
    Dim sParams         As String
    Dim vValue          As Variant

    '--- an event reads as name(param=value, ...), skipping the ones whose
    '--- value is an object: those are the JSRet* and JSRowData carriers,
    '--- whose identity says nothing and whose contents the test asserts
    '--- through the control instead. Shift is skipped too: a synthetic
    '--- WM_KEYDOWN carries no modifier state, so both controls read it off
    '--- the live keyboard and a modifier held while the corpus runs would
    '--- show up as a mismatch that says nothing about either of them
    On Error GoTo EH
    For lIdx = 0 To Info.EventParameters.Count - 1
        vValue = Empty
        If Info.EventParameters(lIdx).Name = "Shift" Then
            '--- environment, not behaviour
        ElseIf Not IsObject(Info.EventParameters(lIdx).Value) Then
            vValue = Info.EventParameters(lIdx).Value
            If LenB(sParams) <> 0 Then
                sParams = sParams & ", "
            End If
            sParams = sParams & Info.EventParameters(lIdx).Name & "=" & C2Str(vValue)
        End If
    Next
    pvFormatEvent = Info.Name & "(" & sParams & ")"
    Exit Function
EH:
    pvFormatEvent = Info.Name & "(?)"
End Function

Private Sub pvRunInput(oList As Object)
    Const FUNC_NAME     As String = "pvRunInput"
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim oEntry          As Object
    Dim oPoint          As Object
    Dim lX              As Long
    Dim lY              As Long
    Dim sText           As String
    Dim lJdx            As Long
    Dim hTarget         As Long
    Dim uPt             As POINTAPI

    '--- clicks and keystrokes are posted at the control the same way the
    '--- ported samples receive them, so an edit session can be scripted
    On Error GoTo EH
    lCount = C2Lng(JsonValue(oList, "-1"))
    For lIdx = 0 To lCount - 1
        Set oEntry = C2Obj(JsonValue(oList, lIdx))
        Set oPoint = C2Obj(JsonValue(oEntry, "click"))
        If Not oPoint Is Nothing Then
            '--- a scenario clicks where the capture shows it, so the point is
            '--- in the host form's client space and travels through the screen
            '--- to whichever window actually sits under it -- the original's
            '--- grid is an inner window, ours is the UserControl itself
            uPt.X = C2Lng(JsonValue(oPoint, 0))
            uPt.Y = C2Lng(JsonValue(oPoint, 1))
            Call ClientToScreen(hWnd, uPt)
            hTarget = WindowFromPoint(uPt.X, uPt.Y)
            Call ScreenToClient(hTarget, uPt)
            Call SendMessage(hTarget, WM_LBUTTONDOWN, MK_LBUTTON, ByVal pvMakeLong(uPt.X, uPt.Y))
            Call SendMessage(hTarget, WM_LBUTTONUP, 0, ByVal pvMakeLong(uPt.X, uPt.Y))
        End If
        '--- keys follow the focus, which an in-place editor takes from the grid
        If Not IsEmpty(JsonValue(oEntry, "key")) Then
            Call SendMessage(pvFocusHwnd(), WM_KEYDOWN, C2Lng(JsonValue(oEntry, "key")), ByVal 0&)
            '--- resolved again: a key that closes an in-place editor destroys
            '--- the window it went down on, and the release lands on whatever
            '--- holds the focus by then, exactly as real input would
            Call SendMessage(pvFocusHwnd(), WM_KEYUP, C2Lng(JsonValue(oEntry, "key")), ByVal 0&)
        End If
        sText = C2Str(JsonValue(oEntry, "text"))
        For lJdx = 1 To Len(sText)
            hTarget = pvFocusHwnd()
            Call SendMessage(hTarget, WM_KEYDOWN, AscW(UCase$(Mid$(sText, lJdx, 1))), ByVal 0&)
            Call SendMessage(hTarget, WM_CHAR, AscW(Mid$(sText, lJdx, 1)), ByVal 0&)
            Call SendMessage(hTarget, WM_KEYUP, AscW(UCase$(Mid$(sText, lJdx, 1))), ByVal 0&)
        Next
        pvSettle 5
    Next
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Function pvFocusHwnd() As Long
    Dim hChild          As Long
    Dim sClass          As String

    '--- an in-place editor is a child edit window the grid puts over the
    '--- cell, and synthetic keys have to go at it rather than at the grid,
    '--- which would only raise KeyPress and drop them
    hChild = pvFindEditWindow(hWnd)
    If hChild <> 0 Then
        pvFocusHwnd = hChild
        Exit Function
    End If
    pvFocusHwnd = GetFocus()
    If pvFocusHwnd = 0 Then
        pvFocusHwnd = m_oExt.hWnd
    End If
End Function

Private Function pvFindEditWindow(ByVal hParent As Long) As Long
    Dim hChild          As Long
    Dim sClass          As String
    Dim hFound          As Long

    hChild = GetWindow(hParent, GW_CHILD)
    Do While hChild <> 0
        If IsWindowVisible(hChild) <> 0 Then
            sClass = String$(256, 0)
            sClass = Left$(sClass, GetClassName(hChild, StrPtr(sClass), 256))
            '--- the original's editor is an EDIT window and ours is a VB
            '--- TextBox, whose class is ThunderRT6TextBox
            If InStr(1, sClass, "Edit", vbTextCompare) > 0 Or InStr(1, sClass, "TextBox", vbTextCompare) > 0 Then
                pvFindEditWindow = hChild
                Exit Function
            End If
            hFound = pvFindEditWindow(hChild)
            If hFound <> 0 Then
                pvFindEditWindow = hFound
                Exit Function
            End If
        End If
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Loop
End Function

Private Function pvMakeLong(ByVal lLo As Long, ByVal lHi As Long) As Long
    pvMakeLong = (lLo And &HFFFF&) Or (lHi * &H10000)
End Function

Private Sub pvRunCalls(oList As Object)
    Const FUNC_NAME     As String = "pvRunCalls"
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim oEntry          As Object
    Dim sName           As String
    Dim oArgs           As Object
    Dim lArgs           As Long
    Dim oGrid           As Object

    '--- an entry is either a bare method name or an object carrying the
    '--- name, its arguments and, for a property, the value to assign
    On Error GoTo EH
    lCount = C2Lng(JsonValue(oList, "-1"))
    Set oGrid = m_oExt.Object
    For lIdx = 0 To lCount - 1
        Set oEntry = C2Obj(JsonValue(oList, lIdx))
        If oEntry Is Nothing Then
            CallByName oGrid, C2Str(JsonValue(oList, lIdx)), VbMethod
        Else
            sName = C2Str(JsonValue(oEntry, "name"))
            Set oArgs = C2Obj(JsonValue(oEntry, "args"))
            lArgs = 0
            If Not oArgs Is Nothing Then
                lArgs = C2Lng(JsonValue(oArgs, "-1"))
            End If
            If Not IsEmpty(JsonValue(oEntry, "value")) Then
                Select Case lArgs
                Case 0
                    CallByName oGrid, sName, VbLet, JsonValue(oEntry, "value")
                Case Else
                    CallByName oGrid, sName, VbLet, JsonValue(oArgs, 0), JsonValue(oEntry, "value")
                End Select
            Else
                Select Case lArgs
                Case 0
                    CallByName oGrid, sName, VbMethod
                Case 1
                    CallByName oGrid, sName, VbMethod, JsonValue(oArgs, 0)
                Case Else
                    CallByName oGrid, sName, VbMethod, JsonValue(oArgs, 0), JsonValue(oArgs, 1)
                End Select
            End If
        End If
    Next
    oGrid.Refresh
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Sub pvTrySet(sProp As String, ByVal vValue As Variant)
    Const FUNC_NAME     As String = "pvTrySet"

    On Error GoTo EH
    CallByName m_oExt.Object, sProp, VbLet, vValue
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Sub pvTryCall(sMethod As String)
    Const FUNC_NAME     As String = "pvTryCall"

    On Error GoTo EH
    CallByName m_oExt.Object, sMethod, VbMethod
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"

    '--- the window exists but is still hidden here, so the DWM open
    '--- transition can be suppressed before the first Show
    On Error GoTo EH
    DisableWindowTransitions hWnd
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub

Private Sub m_oExt_ObjectEvent(Info As EventInfo)
    Const FUNC_NAME     As String = "m_oExt_ObjectEvent"
    Dim lRowIndex       As Long
    Dim oValues         As Object
    Dim oRow            As Object
    Dim nIdx            As Integer

    On Error GoTo EH
    If m_bLogEvents Then
        m_sEventLog = m_sEventLog & pvFormatEvent(Info) & vbCrLf
    End If
    If Info.Name = "UnboundReadData" Then
        If Not m_oUnboundRows Is Nothing Then
            lRowIndex = C2Lng(Info.EventParameters("RowIndex").Value)
            Set oValues = C2Obj(Info.EventParameters("Values").Value)
            Set oRow = C2Obj(JsonValue(m_oUnboundRows, lRowIndex - 1))
            If Not oValues Is Nothing And Not oRow Is Nothing Then
                For nIdx = 1 To C2Lng(CallByName(oValues, "ColCount", VbGet))
                    CallByName oValues, "Value", VbLet, nIdx, JsonValue(oRow, nIdx - 1)
                Next
            End If
        End If
    End If
    Exit Sub
EH:
    LogError "Critical error: " & Err.Description & " [" & FUNC_NAME & "]", Erl
End Sub
