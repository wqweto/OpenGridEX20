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
    '--- row selection: select the listed 1-based row positions
    If Not C2Obj(JsonValue(oScenario, "select")) Is Nothing Then
        pvSelectRows C2Obj(JsonValue(oScenario, "select"))
        DoEvents
    End If
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
