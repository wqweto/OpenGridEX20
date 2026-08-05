VERSION 5.00
Object = "{24F4AB9F-37F4-43D4-B383-FB6CD721B629}#1.0#0"; "OpenGridEX20.ocx"
Begin VB.Form frmMain
   Caption         =   "ModelTests"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin OpenGridEX20.GridEX GridEX1
      Height          =   2000
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3528
   End
   Begin OpenGridEX20.GridEX GridEX2
      Height          =   2000
      Left            =   3180
      TabIndex        =   1
      Top             =   60
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3528
   End
   Begin OpenGridEX20.GEXPreview GEXPreview1
      Height          =   2000
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3528
   End
   Begin OpenGridEX20.GEXPreview GEXPreview2
      Height          =   2000
      Left            =   3180
      TabIndex        =   3
      Top             =   2160
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3528
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' M2 object model tests: collection semantics + snapshot round-trip
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "frmMain"

'=========================================================================
' API
'=========================================================================

Private Const WM_VSCROLL            As Long = &H115
Private Const WM_KEYDOWN            As Long = &H100
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_LBUTTONUP          As Long = &H202
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_CHAR               As Long = &H102
Private Const MK_LBUTTON            As Long = &H1
Private Const SB_LINEUP             As Long = 0
Private Const SB_LINEDOWN           As Long = 1
Private Const SB_PAGEDOWN           As Long = 3
Private Const SB_THUMBPOSITION      As Long = 4
Private Const SB_THUMBTRACK         As Long = 5
Private Const GWL_STYLE             As Long = -16
Private Const WS_VSCROLL            As Long = &H200000

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oCanonFont                As New StdFont

'=========================================================================
' Error management
'=========================================================================

Private Sub PrintError(sFunction As String)
    PopPrintError PushError, MODULE_NAME, sFunction
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"
    Dim lIdx            As Long
    Dim sError          As String

    On Error GoTo EH
    TestInit OutputFile("ModelTests.out.txt")
    pvTestColumns
    pvTestValueList
    pvTestFormatStyles
    pvTestSortKeysGroups
    pvTestFmtConditions
    pvTestPrinterProperties
    pvTestRoundTrip
    pvTestRowData
    pvTestRowDataWeakRef
    pvTestUnbound
    pvTestScroll
    pvTestScrollProps
    pvTestItemCountInvalidate
    pvTestEditLeaveCell
    pvTestSorting
    pvTestGrouping
    pvTestGroupCaption
    pvTestGroupFooter
    pvTestAdvancedSampleParts
    pvTestKeyNav
    pvTestMouse
    pvTestAutomaticSort
    pvTestNavigator
    pvTestGeneratedSetup
    pvTestSelection
    pvTestKeyPressDrag
    pvTestSnapshotCorpus
QH:
    TestsDone
    '--- any host form a failing test left behind keeps the message loop
    '--- running long after the results are written, which reads as a hang
    For lIdx = Forms.Count - 1 To 0 Step -1
        If Not Forms(lIdx) Is Me Then
            Unload Forms(lIdx)
        End If
    Next
    Unload Me
    Exit Sub
EH:
    '--- PrintError resets Err, so the assert has to read it first
    sError = "&H" & Hex$(Err.Number) & " " & Err.Description
    PrintError FUNC_NAME
    Assert "Unhandled error " & sError, False
    GoTo QH
End Sub

'=========================================================================
' Functions
'=========================================================================

Private Sub pvTestColumns()
    Dim oCol            As JSColumn
    Dim lCount          As Long

    With GridEX1.Columns
        .Add "Alpha", , , "kA"
        Set oCol = .Add("Beta")
        oCol.Width = 777
        .Add "Gamma", jgexIcon, jgexEditNone, "kC"
        AssertEquals "Columns.Count", 3, .Count
        AssertEquals "Columns(2).Caption", "Beta", .Item(2).Caption
        '--- widths snap to whole pixels like the original
        AssertEquals "Columns(2).Width", 780, .Item(2).Width
        AssertEquals "Columns(kC).ColumnType", jgexIcon, .Item("kC").ColumnType
        AssertEquals "Columns(kC).Index", 3, .Item("kC").Index
        AssertEquals "ItemByPosition(2).Caption", "Beta", .ItemByPosition(2).Caption
        For Each oCol In GridEX1.Columns
            lCount = lCount + 1
        Next
        AssertEquals "For Each count", 3, lCount
        .Remove 1
        AssertEquals "Count after Remove", 2, .Count
        AssertEquals "Reindexed Item(1).Caption", "Beta", .Item(1).Caption
        AssertEquals "Reindexed Item(1).Index", 1, .Item(1).Index
        .Remove "kC"
        AssertEquals "Count after Remove by key", 1, .Count
        .Clear
        AssertEquals "Count after Clear", 0, .Count
    End With
End Sub

Private Sub pvTestValueList()
    Dim oCol            As JSColumn

    Set oCol = GridEX1.Columns.Add("Values")
    oCol.HasValueList = True
    With oCol.ValueList
        .Add 1, "One"
        .Add 2, "Two"
        .Add 3, "Three"
        .Item(3).Visible = False
        AssertEquals "ValueList.Count", 3, .Count
        AssertEquals "ValueList.VisibleCount", 2, .VisibleCount
        AssertEquals "ItemByValue(2).Text", "Two", .ItemByValue(2).Text
        .RemoveItemByValue 1
        AssertEquals "Count after RemoveItemByValue", 2, .Count
        AssertEquals "Reindexed Item(1).Text", "Two", .Item(1).Text
    End With
    GridEX1.Columns.Clear
End Sub

Private Sub pvTestFormatStyles()
    Dim lErr            As Long

    With GridEX1.FormatStyles
        AssertEquals "FormatStyles: 6 built-in styles", 6, .Count
        AssertEquals "FormatStyles: built-in EvenRow color", &HC1D7B0, .Item("EvenRow").BackColor
        '--- system styles are protected from Remove (error 380)
        On Error Resume Next
        Err.Clear
        .Remove "SelectedRow"
        lErr = Err.Number
        On Error GoTo 0
        AssertEquals "FormatStyles: Remove system raises 380", 380, lErr
        AssertEquals "FormatStyles: system count unchanged", 6, .Count
        '--- user styles add on top of the system ones
        .Add("Header").FontBold = True
        .Add "Totals"
        AssertEquals "FormatStyles: count with user styles", 8, .Count
        AssertEquals "Item(Header).FontBold", True, .Item("Header").FontBold
        .Remove "Header"
        AssertEquals "FormatStyles: count after user Remove", 7, .Count
        '--- Clear drops user styles but keeps the system ones
        .Clear
        AssertEquals "FormatStyles: Clear keeps system styles", 6, .Count
        AssertEquals "FormatStyles: SelectedRow survives Clear", "SelectedRow", .Item("SelectedRow").Name
    End With
End Sub

Private Sub pvTestSortKeysGroups()
    With GridEX1.SortKeys
        .Add 1, jgexSortAscending
        .Add 2, jgexSortDescending, 1
        AssertEquals "SortKeys.Count", 2, .Count
        AssertEquals "Insert at 1: Item(1).ColIndex", 2, .Item(1).ColIndex
        AssertEquals "Item(1).Index", 1, .Item(1).Index
        AssertEquals "Item(2).Index", 2, .Item(2).Index
        .Clear
    End With
    With GridEX1.Groups
        .Add 3, jgexSortAscending
        AssertEquals "Groups.Count", 1, .Count
        AssertEquals "Groups(1).ColIndex", 3, .Item(1).ColIndex
        .Clear
    End With
End Sub

Private Sub pvTestFmtConditions()
    With GridEX1.FmtConditions
        .Add 1, jgexGreaterThan, 100, , "big"
        AssertEquals "FmtConditions.Count", 1, .Count
        AssertEquals "Item(big).Operator", jgexGreaterThan, .Item("big").Operator
        AssertEquals "Item(big).Value1", 100, .Item("big").Value1
        .GroupCondition.SetCondition 2, jgexEqual, "x"
        AssertEquals "GroupCondition.ColIndex", 2, .GroupCondition.ColIndex
        '--- restore defaults: GridEX1 is the round-trip import target below
        .GroupCondition.SetCondition 0, jgexEqual, Empty
        .Clear
    End With
End Sub

Private Sub pvTestPrinterProperties()
    With GridEX1.PrinterProperties
        .HeaderString(jgexHFCenter) = "Middle"
        .FooterString(jgexHFRight) = "Page"
        AssertEquals "HeaderString(center)", "Middle", .HeaderString(jgexHFCenter)
        AssertEquals "FooterString(right)", "Page", .FooterString(jgexHFRight)
        AssertEquals "ClientWidth", .PaperWidth - .LeftMargin - .RightMargin, .ClientWidth
    End With
End Sub

Private Sub pvConfigureGrid(oGrid As GridEX)
    Dim oCol            As JSColumn

    With oGrid
        .BackColor = &HC0FFC0
        .AllowAddNew = True
        .MultiSelect = True
        .View = jgexCard
        .GroupByBoxVisible = True
        .PreviewColumn = 2
        .RecordSource = "Products"
        .DatabaseName = "test.mdb"
        Set oCol = .Columns.Add("ID", jgexText, jgexEditNone, "id")
        oCol.Width = 600
        Set oCol = .Columns.Add("Name")
        oCol.HeaderAlignment = jgexAlignCenter
        oCol.HasValueList = True
        oCol.ValueList.Add 1, "Yes"
        oCol.ValueList.Add 0, "No", -1
        .Columns.Add "Notes", jgexText, jgexEditTextBox, "notes"
        .FormatStyles.Add("Header").FontBold = True
        .FormatStyles.Add("Odd").BackColor = &HF0F0F0
        .SortKeys.Add 1, jgexSortAscending
        .Groups.Add 2, jgexSortDescending
        .FmtConditions.Add 1, jgexGreaterThan, 50, 100, "range"
        .FmtConditions.ApplyGroupCondition = True
        .PrinterProperties.HeaderString(jgexHFLeft) = "Left"
        .PrinterProperties.FooterString(jgexHFCenter) = "Footer"
        .PrinterProperties.LeftMargin = 720
        .Font.Bold = True
    End With
End Sub

Private Sub pvTestRoundTrip()
    Dim sJson1          As String
    Dim sJson2          As String
    Dim vDoc            As Variant
    Dim oProps          As Object

    '--- GridEX: configure -> export -> import into twin -> export -> compare
    pvConfigureGrid GridEX2
    Assert "checkpoint configure", True
    sJson1 = SnapshotToJson(GridEX2.Object, "GridEX", False)
    Assert "checkpoint export1", LenB(sJson1) > 0
    JsonParse sJson1, vDoc
    Set oProps = JsonValue(C2Obj(vDoc), "props")
    Assert "checkpoint parse", Not oProps Is Nothing
    ImportObject GridEX1.Object, "GridEX", oProps
    Assert "checkpoint import", True
    sJson2 = SnapshotToJson(GridEX1.Object, "GridEX", False)
    WriteTextFile OutputFile("RoundTrip1.json"), sJson1
    WriteTextFile OutputFile("RoundTrip2.json"), sJson2
    Assert "GridEX round-trip lossless", (sJson1 = sJson2)
    '--- GEXPreview pair
    With GEXPreview1
        .Zoom = jgexZoomTwoPages
        .ToolbarVisible = False
        .CloseButtonText = "Dismiss"
        .ToolbarFont.Italic = True
    End With
    sJson1 = SnapshotToJson(GEXPreview1.Object, "GEXPreview", False)
    JsonParse sJson1, vDoc
    Set oProps = JsonValue(C2Obj(vDoc), "props")
    ImportObject GEXPreview2.Object, "GEXPreview", oProps
    sJson2 = SnapshotToJson(GEXPreview2.Object, "GEXPreview", False)
    Assert "GEXPreview round-trip lossless", (sJson1 = sJson2)
End Sub

Private Sub pvTestRowData()
    Dim oRD             As JSRowData
    Dim oRD2            As JSRowData

    GridEX2.Columns.Clear
    GridEX2.Columns.Add "Alpha"
    GridEX2.Columns.Add "Beta"
    Set oRD = GridEX2.GetRowData(3)
    Set oRD2 = GridEX2.GetRowData(3)
    Assert "RowData: cached instance", oRD Is oRD2
    Assert "RowData: distinct rows distinct wrappers", Not oRD Is GridEX2.GetRowData(2)
    AssertEquals "RowData: RowIndex", 3, oRD.RowIndex
    AssertEquals "RowData: ColCount", 2, oRD.ColCount
    AssertEquals "RowData: PreviewRowVisible default", True, oRD.PreviewRowVisible
    AssertEquals "RowData: RowType default", jgexRowTypeRecord, oRD.RowType
    '--- writes through one wrapper ref must be visible through the other
    '--- because all state lives in the control internal arrays
    oRD.Value(1) = "abc"
    oRD.IconIndex(2) = 5
    oRD.DisplayValue(1) = "shown"
    oRD.CellStyle(2) = "cellstyle"
    oRD.RowStyle = "rowstyle"
    oRD.GroupCaption = "caption"
    oRD.RowHeight = 400
    AssertEquals "RowData: Value via twin ref", "abc", oRD2.Value(1)
    AssertEquals "RowData: IconIndex via twin ref", 5, oRD2.IconIndex(2)
    AssertEquals "RowData: DisplayValue via twin ref", "shown", oRD2.DisplayValue(1)
    AssertEquals "RowData: CellStyle via twin ref", "cellstyle", oRD2.CellStyle(2)
    AssertEquals "RowData: RowStyle via twin ref", "rowstyle", oRD2.RowStyle
    AssertEquals "RowData: GroupCaption via twin ref", "caption", oRD2.GroupCaption
    AssertEquals "RowData: RowHeight via twin ref", 400, oRD2.RowHeight
    '--- cell storage grows when columns are added after the fact
    GridEX2.Columns.Add "Gamma"
    AssertEquals "RowData: ColCount after Columns.Add", 3, oRD.ColCount
    oRD.Value(3) = 42
    AssertEquals "RowData: Value in new column", 42, oRD2.Value(3)
    AssertEquals "RowData: value kept in older column", "abc", GridEX2.GetRowData(3).Value(1)
End Sub

Private Sub pvTestRowDataWeakRef()
    Dim oForm           As frmWeak
    Dim oRD             As JSRowData
    Dim vValue          As Variant
    Dim lErr            As Long

    Set oForm = New frmWeak
    Load oForm
    oForm.GridEX1.Columns.Add "Alpha"
    Set oRD = oForm.GridEX1.GetRowData(1)
    oRD.Value(1) = "before unload"
    AssertEquals "WeakRef: Value before unload", "before unload", oRD.Value(1)
    Unload oForm
    Set oForm = Nothing
    '--- the wrapper holds no refcount on the control, so the unload must
    '--- have terminated the control which detached the wrapper (frTerm);
    '--- a strong reference would keep the control alive and this access
    '--- would still succeed instead of raising error 91
    On Error Resume Next
    vValue = oRD.Value(1)
    lErr = Err.Number
    On Error GoTo 0
    AssertEquals "WeakRef: orphaned access raises 91", 91, lErr
    '--- every member reaching the owner answers the same way, not just the
    '--- default one, so no orphaned access can dereference a zeroed pointer
    On Error Resume Next
    Err.Clear
    lErr = oRD.RowIndex
    lErr = Err.Number
    On Error GoTo 0
    AssertEquals "WeakRef: orphaned RowIndex raises 91", 91, lErr
End Sub

Private Sub pvTestUnbound()
    Dim oForm           As frmWeak

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "Alpha"
        .Columns.Add "Beta"
        .DataMode = jgexUnbound
        .ItemCount = 3
        AssertEquals "Unbound: RowCount = ItemCount", 3, .RowCount
        AssertEquals "Unbound: RowIndex maps 1:1", 2, .RowIndex(2)
        '--- first cell read of a row fires UnboundReadData exactly once
        oForm.EventLog = vbNullString
        AssertEquals "Unbound: fetched value", "R2C1", .GetRowData(2).Value(1)
        AssertEquals "Unbound: fetch fired once", "Read(2);", oForm.EventLog
        AssertEquals "Unbound: second col cached", "R2C2", .GetRowData(2).Value(2)
        AssertEquals "Unbound: no refire on cached row", "Read(2);", oForm.EventLog
        AssertEquals "Unbound: other row fetches on demand", "R1C2", .GetRowData(1).Value(2)
        AssertEquals "Unbound: fetch log per row", "Read(2);Read(1);", oForm.EventLog
        '--- RefreshRowIndex marks one row for lazy refetch
        oForm.EventLog = vbNullString
        .RefreshRowIndex 2
        AssertEquals "Unbound: refresh is lazy", vbNullString, oForm.EventLog
        AssertEquals "Unbound: refetched value", "R2C1", .GetRowData(2).Value(1)
        AssertEquals "Unbound: refetch fired for row 2", "Read(2);", oForm.EventLog
        AssertEquals "Unbound: row 1 still cached", "R1C1", .GetRowData(1).Value(1)
        AssertEquals "Unbound: row 1 not refetched", "Read(2);", oForm.EventLog
        '--- bookmarks round-trip, reach the event and key RefreshRowBookmark
        .RowBookmark(3) = "bk3"
        AssertEquals "Unbound: RowBookmark get", "bk3", .RowBookmark(3)
        oForm.EventLog = vbNullString
        AssertEquals "Unbound: fetch row 3", "R3C2", .GetRowData(3).Value(2)
        AssertEquals "Unbound: bookmark passed to event", "Read(3)bk3;", oForm.EventLog
        oForm.EventLog = vbNullString
        .RefreshRowBookmark "bk3"
        AssertEquals "Unbound: bookmark refetch", "R3C1", .GetRowData(3).Value(1)
        AssertEquals "Unbound: bookmark refetch log", "Read(3)bk3;", oForm.EventLog
        '--- Refetch resets data on all rows but keeps bookmarks
        .Refetch
        AssertEquals "Unbound: value refetched after Refetch", "R1C1", .GetRowData(1).Value(1)
        AssertEquals "Unbound: Refetch keeps bookmark", "bk3", .RowBookmark(3)
        '--- Rebind full-resets incl. bookmarks; sort cleared per param with
        '--- HoldSortSettings property (original default False) as default
        .SortKeys.Add 1, jgexSortAscending
        .Rebind True
        AssertEquals "Unbound: Rebind True holds sort", 1, .SortKeys.Count
        .Rebind
        AssertEquals "Unbound: Rebind clears sort by default", 0, .SortKeys.Count
        Assert "Unbound: Rebind clears bookmark", IsEmpty(.RowBookmark(3))
        '--- navigation state events with previous row/col in the args
        oForm.EventLog = vbNullString
        .Row = 2
        .Col = 2
        .Row = 2
        .FirstItem = 2
        .FirstItem = 2
        '--- Rebind selects the whole row, which is what Col = 0 means, so
        '--- both changes report coming from there
        AssertEquals "Unbound: nav event order", "RowCol(1,0);RowCol(2,0);First;", oForm.EventLog
    End With
    Unload oForm
End Sub

Private Sub pvTestScroll()
    Dim oForm           As frmWeak

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "Alpha"
        .DataMode = jgexUnbound
        .ItemCount = 50
        .Rebind
        AssertEquals "Scroll: FirstItem after Rebind", 1, .FirstItem
        oForm.EventLog = vbNullString
        '--- drive the MST-subclassed scrollbar with real WM_VSCROLL messages
        SendMessage .hWnd, WM_VSCROLL, SB_LINEDOWN, 0
        AssertEquals "Scroll: line down", 2, .FirstItem
        SendMessage .hWnd, WM_VSCROLL, SB_LINEDOWN, 0
        AssertEquals "Scroll: line down again", 3, .FirstItem
        SendMessage .hWnd, WM_VSCROLL, SB_LINEUP, 0
        AssertEquals "Scroll: line up", 2, .FirstItem
        SendMessage .hWnd, WM_VSCROLL, SB_PAGEDOWN, 0
        Assert "Scroll: page down advances", .FirstItem > 2
        AssertEquals "Scroll: FirstItemChange event count", "First;First;First;First;", oForm.EventLog
    End With
    Unload oForm
End Sub

Private Sub pvTestEditLeaveCell()
    Dim oForm           As frmWeak
    Dim hEdit           As Long

    '--- the editor only ever shows on the current cell, so moving off it
    '--- closes the editor and puts what was typed into the row's storage.
    '--- The visual corpus proves the pixels and the event order; this is the
    '--- underlying value, which a repaint of stale storage would also show
    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add("A").Width = 1500
        .Columns.Add("B").Width = 1500
        .DataMode = jgexUnbound
        .ItemCount = 5
        .Rebind
        .AllowEdit = True
        .Columns.Item(1).EditType = jgexEditTextBox
        .Columns.Item(2).EditType = jgexEditTextBox
        AssertEquals "LeaveCell: the row starts on its fetched value", "R1C1", .GetRowData(1).Value(1)
        '--- same geometry as pvTestMouse: data from y=52, rows 19px, A=0..99,
        '--- so (50, 60) is row 1 column A and (50, 80) is row 2. The click
        '--- lands past the end of the short cell text, which puts the caret
        '--- there and makes the typing an append
        AssertEquals "LeaveCell: no editor, no handle", 0, .hWndEdit
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(50, 60)
        hEdit = GetFocus()
        Assert "LeaveCell: the editor opened and took the focus", hEdit <> 0 And hEdit <> .hWnd
        AssertEquals "LeaveCell: hWndEdit is that editor", hEdit, .hWndEdit
        SendMessage hEdit, WM_CHAR, 90, 0
        '--- still open, so nothing is in the row yet
        AssertEquals "LeaveCell: typing alone does not reach the row", "R1C1", .GetRowData(1).Value(1)
        '--- moving a column over closes the editor and buffers the cell: the
        '--- row's own storage is not written until the row itself is left
        .Col = 2
        AssertEquals "LeaveCell: a column move only buffers it", "R1C1", .GetRowData(1).Value(1)
        AssertEquals "LeaveCell: and Value reads the buffer back", "R1C1Z", .Value(1)
        Assert "LeaveCell: the editor is gone", IsWindow(hEdit) = 0
        AssertEquals "LeaveCell: hWndEdit back to nothing", 0, .hWndEdit
        '--- the row move is what writes it through
        .Row = 2
        AssertEquals "LeaveCell: the row move commits it", "R1C1Z", .GetRowData(1).Value(1)
        '--- Value buffers the same way with no editor ever opened
        .Value(1) = "typed in"
        AssertEquals "LeaveCell: Value shows at once", "typed in", .Value(1)
        AssertEquals "LeaveCell: without reaching storage", "R2C1", .GetRowData(2).Value(1)
        '--- and Escape on the grid drops the whole row's buffer
        SendMessage .hWnd, WM_KEYDOWN, vbKeyEscape, 0
        AssertEquals "LeaveCell: Escape drops the buffered row", "R2C1", .Value(1)
        AssertEquals "LeaveCell: storage never saw it", "R2C1", .GetRowData(2).Value(1)
        '--- one more, this time left rather than cancelled
        .Value(1) = "kept"
        .Row = 1
        AssertEquals "LeaveCell: leaving commits what Value buffered", "kept", .GetRowData(2).Value(1)
    End With
    Unload oForm
End Sub

Private Sub pvTestItemCountInvalidate()
    Dim oForm           As frmWeak

    '--- setting ItemCount has to repaint on its own. The visual corpus cannot
    '--- see this: the harness always follows it with Rebind, whose own
    '--- invalidate covered for it, and the original control refuses an
    '--- ItemCount change at runtime at all, so no golden could exist either.
    '--- The vertical bar is a WS_VSCROLL style on the control's own window,
    '--- put there by the same paint pass, so it stands in for the repaint
    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add("Alpha").Width = 1500
        .DataMode = jgexUnbound
        .ItemCount = 2
        .Rebind
        Assert "ItemCount: two rows need no vertical bar", (GetWindowLong(.hWnd, GWL_STYLE) And WS_VSCROLL) = 0
        '--- no Rebind and no Refresh behind it -- the property alone
        .ItemCount = 500
        AssertEquals "ItemCount: the count is taken", 500, .RowCount
        Assert "ItemCount: and the bar appears without a Refresh", (GetWindowLong(.hWnd, GWL_STYLE) And WS_VSCROLL) <> 0
        '--- and back down again, which has to take the bar away
        .ItemCount = 2
        Assert "ItemCount: shrinking takes the bar away again", (GetWindowLong(.hWnd, GWL_STYLE) And WS_VSCROLL) = 0
        '--- Redraw = False still batches it, the way every other change is
        .Redraw = False
        .ItemCount = 500
        Assert "ItemCount: Redraw off holds the repaint back", (GetWindowLong(.hWnd, GWL_STYLE) And WS_VSCROLL) = 0
        .Redraw = True
        Assert "ItemCount: turning Redraw back on pays it", (GetWindowLong(.hWnd, GWL_STYLE) And WS_VSCROLL) <> 0
    End With
    Unload oForm
End Sub

Private Sub pvTestScrollProps()
    Dim oForm           As frmWeak

    '--- the M3d properties that change behaviour rather than a static
    '--- picture, so no golden can pin them down
    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add("Alpha").Width = 1500
        .Columns.Add("Beta").Width = 1500
        .Columns.Add("Gamma").Width = 1500
        .Columns.Add("Delta").Width = 1500
        .DataMode = jgexUnbound
        .ItemCount = 50
        .Rebind
        '--- two of the four columns fit, so the scroll range is known
        oForm.GridEX1.Width = 4500
        '--- ContinuousScroll: with it off the contents wait for the thumb to
        '--- be released, with it on they follow the drag
        AssertEquals "ContinuousScroll: default", False, .ContinuousScroll
        SendMessage .hWnd, WM_VSCROLL, pvMakeLong(SB_THUMBTRACK, 20), 0
        AssertEquals "ContinuousScroll off: track does not scroll", 1, .FirstItem
        SendMessage .hWnd, WM_VSCROLL, pvMakeLong(SB_THUMBPOSITION, 20), 0
        Assert "ContinuousScroll off: release scrolls", .FirstItem > 1
        .FirstItem = 1
        .ContinuousScroll = True
        SendMessage .hWnd, WM_VSCROLL, pvMakeLong(SB_THUMBTRACK, 20), 0
        Assert "ContinuousScroll on: track scrolls", .FirstItem > 1
        '--- FrozenColumns pins the leftmost columns, so LeftCol may go
        '--- further right than it could without them
        .LeftCol = 4
        AssertEquals "LeftCol clamps at the last full page", 3, .LeftCol
        .FrozenColumns = 2
        AssertEquals "FrozenColumns stored", 2, .FrozenColumns
        .LeftCol = 4
        AssertEquals "LeftCol reaches further with frozen columns", 4, .LeftCol
        .FrozenColumns = 0
        '--- Col follows the same clamping as Row
        .Col = 99
        AssertEquals "Col clamps to the last column", 4, .Col
        .Col = 0
        AssertEquals "Col clamps to the first column", 1, .Col
        '--- ColumnAutoResize spreads the columns over the client width
        AssertEquals "ColumnAutoResize: default", False, .ColumnAutoResize
        .ColumnAutoResize = True
        AssertEquals "ColumnAutoResize stored", True, .ColumnAutoResize
        .ColumnAutoResize = False
        '--- Redraw batches repaints, and the property itself round-trips
        AssertEquals "Redraw: default", True, .Redraw
        .Redraw = False
        AssertEquals "Redraw off", False, .Redraw
        .ItemCount = 60
        .Redraw = True
        AssertEquals "Redraw back on", True, .Redraw
        AssertEquals "Redraw: changes made while off are kept", 60, .RowCount
    End With
    Unload oForm
End Sub

Private Sub pvTestSorting()
    Dim oForm           As frmWeak

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "Alpha"
        .Columns.Add("Beta").SortType = jgexSortTypeNumeric
        .DataMode = jgexUnbound
        .ItemCount = 4
        .Rebind
        '--- frmWeak fills row N as "R<N>C<col>", so Alpha already ascends;
        '--- descending is the interesting direction
        .SortKeys.Add 1, jgexSortDescending
        .Refresh
        AssertEquals "Sort: SortKeys.Count", 1, .SortKeys.Count
        AssertEquals "Sort: first row is the last data row", 4, .RowIndex(1)
        AssertEquals "Sort: last row is the first data row", 1, .RowIndex(4)
        AssertEquals "Sort: cell text follows the row", "R4C1", .GetRowData(1).Value(1)
        '--- the current row rides along: Rebind left it on data row 1, which
        '--- descending sends to the bottom
        AssertEquals "Sort: current row follows its data", 4, .Row
        AssertEquals "Sort: selection follows its data", 4, .SelectedItems.Item(1).RowPosition
        AssertEquals "Sort: selection keeps its row index", 1, .SelectedItems.Item(1).RowIndex
        '--- flipping the key in place re-sorts, as the samples do from a
        '--- header click
        .SortKeys.Item(1).SortOrder = jgexSortAscending
        .Refresh
        AssertEquals "Sort: flipped key re-sorts", 1, .RowIndex(1)
        AssertEquals "Sort: current row follows the flip", 1, .Row
        '--- and clearing puts the rows back in the order they came in
        .SortKeys.Clear
        .Refresh
        AssertEquals "Sort: cleared keys restore data order", 3, .RowIndex(3)
        AssertEquals "Sort: RowCount unchanged by sorting", 4, .RowCount
    End With
    Unload oForm
End Sub

Private Sub pvTestGrouping()
    Dim oForm           As frmWeak
    Dim lIdx            As Long

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "Alpha"
        .Columns.Add "Beta"
        .DataMode = jgexUnbound
        .ItemCount = 4
        .Rebind
        '--- the unbound feed numbers every cell, so the column to group by is
        '--- overwritten with a value that repeats: two groups of two
        For lIdx = 1 To 4
            .GetRowData(lIdx).Value(1) = IIf(lIdx <= 2, "north", "south")
        Next
        .Groups.Add 1, jgexSortAscending
        .Refresh
        AssertEquals "Group: RowCount counts the group rows", 6, .RowCount
        AssertEquals "Group: position 1 is a group row", True, .IsGroupItem(1)
        AssertEquals "Group: position 2 is a record", False, .IsGroupItem(2)
        AssertEquals "Group: group row level", 1, .GroupRowLevel(1)
        AssertEquals "Group: a group row has no row index", 0, .RowIndex(1)
        AssertEquals "Group: caption comes off the grouped column", "north", .GetRowData(1).GroupCaption
        AssertEquals "Group: the group counts its records", 2, .GetRowData(1).RecordCount
        AssertEquals "Group: group row type", jgexRowTypeGroupHeader, .GetRowData(1).RowType
        AssertEquals "Group: a record keeps its own row type", jgexRowTypeRecord, .GetRowData(2).RowType
        AssertEquals "Group: current row is the first record", 2, .Row
        '--- collapsing takes the records under it off the display
        .RowExpanded(1) = False
        AssertEquals "Group: RowExpanded reports the state", False, .RowExpanded(1)
        AssertEquals "Group: a collapsed group hides its records", 4, .RowCount
        AssertEquals "Group: the next group moved up", True, .IsGroupItem(2)
        AssertEquals "Group: current row moved to the group row", 1, .Row
        .RowExpanded(1) = True
        AssertEquals "Group: expanding brings them back", 6, .RowCount
        .CollapseAll
        AssertEquals "Group: CollapseAll leaves the group rows", 2, .RowCount
        .ExpandAll
        AssertEquals "Group: ExpandAll restores every row", 6, .RowCount
        '--- DefaultGroupMode decides how the next rebuild starts out
        .DefaultGroupMode = jgexDGMCollapsed
        .RefreshGroups
        AssertEquals "Group: DefaultGroupMode collapses a rebuild", 2, .RowCount
        .DefaultGroupMode = jgexDGMExpanded
        .Groups.Clear
        .Refresh
        AssertEquals "Group: cleared groups restore the records", 4, .RowCount
    End With
    Unload oForm
End Sub

Private Sub pvTestGroupCaption()
    Dim oForm           As frmWeak
    Dim lIdx            As Long

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "Alpha"
        .Columns.Add "Beta"
        .DataMode = jgexUnbound
        .ItemCount = 4
        .Rebind
        For lIdx = 1 To 4
            .GetRowData(lIdx).Value(1) = IIf(lIdx <= 2, "north", "south")
        Next
        .Groups.Add 1, jgexSortAscending
        .Refresh
        AssertEquals "GroupCaption: plain value by default", "north", .GetRowData(1).GroupCaption
        AssertEquals "GroupCaption: default empty caption", "(none)", .Columns.Item(1).GroupEmptyStringCaption
        '--- a space joins the prefix to the value, and the caption then starts
        '--- one space earlier so the value lands where an unprefixed one does
        '--- -- 049-group-prefix pins both against the original
        .Columns.Item(1).GroupPrefix = "Region:"
        .RefreshGroups
        AssertEquals "GroupCaption: prefix leads the value", "Region: north", .GetRowData(1).GroupCaption
        '--- an empty value gives way to the empty caption, prefix and all
        .GetRowData(1).Value(1) = vbNullString
        .GetRowData(2).Value(1) = vbNullString
        .RefreshGroups
        AssertEquals "GroupCaption: empty value takes the empty caption", "Region: (none)", .GetRowData(1).GroupCaption
        .Columns.Item(1).GroupEmptyStringCaption = "<blank>"
        .RefreshGroups
        AssertEquals "GroupCaption: empty caption is settable", "Region: <blank>", .GetRowData(1).GroupCaption
        .Columns.Item(1).GroupPrefix = vbNullString
    End With
    Unload oForm
    '--- GroupFormat re-formats the value the group shows, which is the point
    '--- of the property: group by month over a column of full dates
    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "When"
        .DataMode = jgexUnbound
        .ItemCount = 2
        .Rebind
        .GetRowData(1).Value(1) = DateSerial(2026, 8, 1)
        .GetRowData(2).Value(1) = DateSerial(2026, 8, 20)
        '--- month names are localized, so the assertion sticks to digits
        .Columns.Item(1).GroupFormat = "yyyy\-mm"
        .Groups.Add 1, jgexSortAscending
        .Refresh
        AssertEquals "GroupCaption: GroupFormat applied", "2026-08", .GetRowData(1).GroupCaption
        '--- it labels the caption and nothing else: the original still breaks
        '--- groups on the raw value, so two August dates are two groups that
        '--- happen to read the same -- 052-group-format pins that
        AssertEquals "GroupCaption: GroupFormat does not group", 4, .RowCount
        AssertEquals "GroupCaption: second group reads the same", "2026-08", .GetRowData(3).GroupCaption
    End With
    Unload oForm
End Sub

Private Sub pvTestGroupFooter()
    Dim oForm           As frmWeak
    Dim lIdx            As Long

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "Alpha"
        .Columns.Add "Beta"
        .DataMode = jgexUnbound
        .ItemCount = 3
        .Rebind
        '--- north: 10 and 20, south: 5
        For lIdx = 1 To 3
            .GetRowData(lIdx).Value(1) = IIf(lIdx <= 2, "north", "south")
            .GetRowData(lIdx).Value(2) = Array(10, 20, 5)(lIdx - 1)
        Next
        .Groups.Add 1, jgexSortAscending
        .Refresh
        AssertEquals "Footer: none by default", 5, .RowCount
        '--- a footer per group, after its records
        .GroupFooterStyle = jgexCaptionGroupFooter
        .RefreshGroups
        AssertEquals "Footer: a row per group", 7, .RowCount
        AssertEquals "Footer: sits after the records", jgexRowTypeGroupFooter, .GetRowData(4).RowType
        AssertEquals "Footer: repeats the caption", "north", .GetRowData(4).GroupCaption
        AssertEquals "Footer: the header stays a header", jgexRowTypeGroupHeader, .GetRowData(1).RowType
        AssertEquals "Footer: counts the same records", 2, .GetRowData(4).RecordCount
        '--- GetSubTotal aggregates over the group either row stands for
        AssertEquals "SubTotal: sum from the header", 30, .GetRowData(1).GetSubTotal(2, jgexSum)
        AssertEquals "SubTotal: sum from the footer", 30, .GetRowData(4).GetSubTotal(2, jgexSum)
        AssertEquals "SubTotal: count", 2, .GetRowData(1).GetSubTotal(2, jgexCount)
        AssertEquals "SubTotal: value count", 2, .GetRowData(1).GetSubTotal(2, jgexValueCount)
        AssertEquals "SubTotal: avg", 15, .GetRowData(1).GetSubTotal(2, jgexAvg)
        AssertEquals "SubTotal: min", 10, .GetRowData(1).GetSubTotal(2, jgexMin)
        AssertEquals "SubTotal: max", 20, .GetRowData(1).GetSubTotal(2, jgexMax)
        AssertEquals "SubTotal: stddev", 5, .GetRowData(1).GetSubTotal(2, jgexStdDev)
        AssertEquals "SubTotal: the second group is its own", 5, .GetRowData(5).GetSubTotal(2, jgexSum)
        AssertEquals "SubTotal: none aggregates to nothing", True, IsEmpty(.GetRowData(1).GetSubTotal(2, jgexAggregateNone))
        '--- a record is not a group and has nothing to total
        AssertEquals "SubTotal: a record totals nothing", True, IsEmpty(.GetRowData(2).GetSubTotal(2, jgexSum))
        '--- collapsing hides the footer along with the records it closes
        .RowExpanded(1) = False
        AssertEquals "Footer: collapsed hides its footer too", 4, .RowCount
        AssertEquals "Footer: the next group follows straight on", True, .IsGroupItem(2)
        .RowExpanded(1) = True
        AssertEquals "Footer: expanding brings both back", 7, .RowCount
        '--- totals style keeps the same rows, only what they draw changes
        .GroupFooterStyle = jgexTotalsGroupFooter
        .RefreshGroups
        AssertEquals "Footer: totals style has the same rows", 7, .RowCount
        .GroupFooterStyle = jgexNoGroupFooter
        .RefreshGroups
        AssertEquals "Footer: turning them off drops the rows", 5, .RowCount
        AssertEquals "SubTotal: still totals with no footers", 30, .GetRowData(1).GetSubTotal(2, jgexSum)
    End With
    Unload oForm
End Sub

Private Sub pvTestAdvancedSampleParts()
    Dim oForm           As frmWeak
    Dim lIdx            As Long
    Dim oKey            As JSSortKey
    Dim oGroup          As JSGroup
    Dim oKeys           As JSSortKeys
    Dim oGroups         As JSGroups

    '--- the Advanced Sample drives sorting and grouping from three dialogs:
    '--- frmSort rebuilds SortKeys and calls RefreshSort, frmGroupBy does the
    '--- same over Groups, frmSummary reads the keys back and totals columns.
    '--- Their data comes from ADO, which is M8, so what is exercised here is
    '--- the API sequence each of them runs
    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add("Region").Tag = "Region"
        .Columns.Add("Rep").Tag = "Rep"
        .Columns.Add("Amount").Tag = "Amount"
        .DataMode = jgexUnbound
        .ItemCount = 4
        .Rebind
        For lIdx = 1 To 4
            .GetRowData(lIdx).Value(1) = IIf(lIdx <= 2, "north", "south")
            .GetRowData(lIdx).Value(2) = Array("bea", "abe", "dee", "cy")(lIdx - 1)
            .GetRowData(lIdx).Value(3) = Array(10, 20, 5, 7)(lIdx - 1)
        Next
        '--- frmSort: clear, re-add from the combos, RefreshSort
        Set oKeys = .SortKeys
        oKeys.Clear
        oKeys.Add 2, jgexSortAscending
        .RefreshSort
        AssertEquals "AdvSort: sorted by the rep column", "abe", .GetRowData(1).Value(2)
        AssertEquals "AdvSort: which is the second data row", 2, .RowIndex(1)
        AssertEquals "AdvSort: key readable back", 2, .SortKeys.Item(1).ColIndex
        '--- reading the keys back is what frmSort does to preselect its combos
        For Each oKey In .SortKeys
            AssertEquals "AdvSort: NewEnum walks the keys", 2, oKey.ColIndex
        Next
        '--- frmGroupBy: same shape over Groups, then RefreshGroups
        Set oGroups = .Groups
        oGroups.Clear
        oGroups.Add 1, jgexSortAscending
        .RefreshGroups
        AssertEquals "AdvGroup: group rows appear", 6, .RowCount
        AssertEquals "AdvGroup: the column reports itself grouped", True, .Columns.Item(1).IsGrouped
        For Each oGroup In .Groups
            AssertEquals "AdvGroup: NewEnum walks the groups", 1, oGroup.ColIndex
        Next
        '--- the sample's dialog offers "all collapsed" straight off the call
        .RefreshGroups True
        AssertEquals "AdvGroup: RefreshGroups collapses on demand", 2, .RowCount
        .RefreshGroups
        AssertEquals "AdvGroup: and expands again", 6, .RowCount
        '--- frmSummary: totals per column over each group
        .Columns.Item(3).AggregateFunction = jgexSum
        .GroupFooterStyle = jgexTotalsGroupFooter
        .RefreshGroups
        AssertEquals "AdvSummary: footers close each group", 8, .RowCount
        AssertEquals "AdvSummary: north totals its amounts", 30, .GetRowData(1).GetSubTotal(3, jgexSum)
        AssertEquals "AdvSummary: the column carries the function", jgexSum, .Columns.Item(3).AggregateFunction
        '--- and the sample's grid clears back to a plain list
        oGroups.Clear
        oKeys.Clear
        .RefreshSort
        AssertEquals "AdvSample: cleared back to the records", 4, .RowCount
        AssertEquals "AdvSample: no column left grouped", False, .Columns.Item(1).IsGrouped
    End With
    Unload oForm
End Sub

Private Sub pvTestKeyNav()
    Dim oForm           As frmWeak

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "A"
        .Columns.Add "B"
        .DataMode = jgexUnbound
        .ItemCount = 50
        .Rebind
        AssertEquals "KeyNav: initial Row", 1, .Row
        '--- a fresh bind has the whole row selected, not a cell in it
        AssertEquals "KeyNav: initial Col", 0, .Col
        '--- arrows drive the current cell through the subclassed proc
        SendMessage .hWnd, WM_KEYDOWN, vbKeyDown, 0
        AssertEquals "KeyNav: Down -> Row 2", 2, .Row
        '--- the first Right steps off the row selection onto column 1
        SendMessage .hWnd, WM_KEYDOWN, vbKeyRight, 0
        AssertEquals "KeyNav: Right off the row selection -> Col 1", 1, .Col
        SendMessage .hWnd, WM_KEYDOWN, vbKeyRight, 0
        AssertEquals "KeyNav: Right -> Col 2", 2, .Col
        SendMessage .hWnd, WM_KEYDOWN, vbKeyRight, 0
        AssertEquals "KeyNav: Right clamps at last col", 2, .Col
        SendMessage .hWnd, WM_KEYDOWN, vbKeyLeft, 0
        AssertEquals "KeyNav: Left -> Col 1", 1, .Col
        SendMessage .hWnd, WM_KEYDOWN, vbKeyUp, 0
        AssertEquals "KeyNav: Up -> Row 1", 1, .Row
        '--- End jumps to and scrolls in the last row
        SendMessage .hWnd, WM_KEYDOWN, vbKeyEnd, 0
        AssertEquals "KeyNav: End -> last Row", 50, .Row
        Assert "KeyNav: End scrolled into view", .FirstItem > 1
        SendMessage .hWnd, WM_KEYDOWN, vbKeyHome, 0
        AssertEquals "KeyNav: Home -> Row 1", 1, .Row
        AssertEquals "KeyNav: Home scrolled to top", 1, .FirstItem
        SendMessage .hWnd, WM_KEYDOWN, vbKeyPageDown, 0
        Assert "KeyNav: PageDown advances Row", .Row > 1
    End With
    Unload oForm
End Sub

Private Sub pvTestGeneratedSetup()
    Const FUNC_NAME     As String = "pvTestGeneratedSetup"
    Dim lIdx            As Long
    Dim vDoc            As Variant
    Dim oProps          As Object
    Dim oProps2         As Object
    Dim oForm           As frmWeak
    Dim oCtl            As GridEX
    Dim sName           As String
    Dim sError          As String

    '--- the same round-trip the corpus test does, but with code generated by
    '--- tools\GenSample.ps1 standing in for ImportObject: proves a ported
    '--- sample reproduces the design-time state recorded from the original
    On Error GoTo EH
    For lIdx = 1 To 3
        Select Case lIdx
        Case 1
            sName = "015-Unbound-1_Form1_GridEX1.json"
        Case 2
            sName = "016-Unbound-2_Form1_GridEX1.json"
        Case Else
            sName = "019-Unbound-Collection_frmUnboundCol_GridEX1.json"
        End Select
        JsonParse ReadTextFile(App.Path & "\..\snapshots\" & sName), vDoc
        Set oProps = JsonValue(C2Obj(vDoc), "props")
        Set oForm = New frmWeak
        Load oForm
        '--- the extender converts to the control class, while its .Object
        '--- property hands back a plain IDispatch that will not QI to GridEX
        Set oCtl = oForm.GridEX1
        Select Case lIdx
        Case 1
            SetupUnbound1 oCtl
        Case 2
            SetupUnbound2 oCtl
        Case Else
            SetupUnboundCollection oCtl
        End Select
        Set oProps2 = JsonValue(JsonParseObject(SnapshotToJson(oCtl, "GridEX", False)), "props")
        pvCanonProps oProps, oProps2
        '--- pictures cannot be written as code literals, so a ported sample
        '--- keeps them in its own .frx -- out of scope for the generator
        JsonValue(oProps, "GridImages") = Empty
        JsonValue(oProps2, "GridImages") = Empty
        If JsonDump(oProps) <> JsonDump(oProps2) Then
            WriteTextFile OutputFile(sName & ".gen-expected.txt"), JsonDump(oProps)
            WriteTextFile OutputFile(sName & ".gen-actual.txt"), JsonDump(oProps2)
        End If
        Assert "generated setup " & sName, (JsonDump(oProps) = JsonDump(oProps2))
        Unload oForm
        Set oForm = Nothing
    Next
QH:
    '--- a host form left loaded keeps the runner alive past TestsDone, so
    '--- the failure path has to unload it too
    If Not oForm Is Nothing Then
        Unload oForm
        Set oForm = Nothing
    End If
    Exit Sub
EH:
    '--- PrintError resets Err, so the assert has to read it first
    sError = "&H" & Hex$(Err.Number) & " " & Err.Description
    PrintError FUNC_NAME
    Assert "generated setup error in " & sName & ": " & sError, False
    GoTo QH
End Sub

Private Sub pvTestNavigator()
    Dim oForm           As frmWeak
    Dim lHwnd           As Long
    Dim uRect           As RECT
    Dim lY              As Long

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add("A").Width = 1500
        .Columns.Add("B").Width = 1500
        .DataMode = jgexUnbound
        .ItemCount = 20
        .Rebind
        .RecordNavigator = True
        '--- the navigator is painted on the outer control, not on the grid
        '--- surface that hWnd returns, so clicks go to the parent window.
        '--- Buttons at 96dpi: first 47..64, prev 65..82, next 162..179,
        '--- last 180..197, all in the band along the bottom edge
        lHwnd = GetParent(.hWnd)
        GetClientRect lHwnd, uRect
        lY = uRect.Bottom - 9
        .Row = 3
        SendMessage lHwnd, WM_LBUTTONDOWN, 0, pvMakeLong(70, lY)
        AssertEquals "Navigator: prev steps back", 2, .Row
        SendMessage lHwnd, WM_LBUTTONDOWN, 0, pvMakeLong(170, lY)
        AssertEquals "Navigator: next steps forward", 3, .Row
        SendMessage lHwnd, WM_LBUTTONDOWN, 0, pvMakeLong(188, lY)
        AssertEquals "Navigator: last jumps to the end", .RowCount, .Row
        SendMessage lHwnd, WM_LBUTTONDOWN, 0, pvMakeLong(55, lY)
        AssertEquals "Navigator: first jumps to the start", 1, .Row
        '--- already at the first row, so prev must not move
        SendMessage lHwnd, WM_LBUTTONDOWN, 0, pvMakeLong(70, lY)
        AssertEquals "Navigator: prev clamps at the first row", 1, .Row
        '--- a click between the buttons changes nothing
        SendMessage lHwnd, WM_LBUTTONDOWN, 0, pvMakeLong(120, lY)
        AssertEquals "Navigator: click off a button is ignored", 1, .Row
    End With
    Unload oForm
End Sub

Private Sub pvTestMouse()
    Dim oForm           As frmWeak

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add("A").Width = 1500
        .Columns.Add("B").Width = 1500
        .DataMode = jgexUnbound
        .ItemCount = 50
        .Rebind
        '--- geometry (MS Sans Serif 8.25): group-by box 0..32, header
        '--- 33..51, data from y=52; row height 19px; A=0..99, B=100..199
        oForm.EventLog = vbNullString
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(50, 80)
        SendMessage .hWnd, WM_LBUTTONUP, 0, pvMakeLong(50, 80)
        AssertEquals "Mouse: click sets Row 2", 2, .Row
        AssertEquals "Mouse: click sets Col 1", 1, .Col
        Assert "Mouse: Click event fired", InStr(oForm.EventLog, "Click;") > 0
        '--- click a cell in the second column
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(150, 80)
        AssertEquals "Mouse: click sets Col 2", 2, .Col
        '--- click the first column header
        oForm.EventLog = vbNullString
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(50, 40)
        AssertEquals "Mouse: header click fires ColumnHeaderClick", "HdrClick(A);", oForm.EventLog
        '--- AutomaticSort off by default, so the click above sorted nothing
        AssertEquals "AutoSort: off leaves the sort keys alone", 0, .SortKeys.Count
    End With
    Unload oForm
End Sub

Private Sub pvTestAutomaticSort()
    Dim oForm           As frmWeak

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add("A").Width = 1500
        .Columns.Add("B").Width = 1500
        .DataMode = jgexUnbound
        .ItemCount = 50
        .Rebind
        .AutomaticSort = True
        '--- same geometry as pvTestMouse: header band 33..51, A=0..99
        '--- an unsorted column sorts ascending and becomes the only key
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(50, 40)
        AssertEquals "AutoSort: one key after the first click", 1, .SortKeys.Count
        AssertEquals "AutoSort: on the column clicked", 1, .SortKeys.Item(1).ColIndex
        AssertEquals "AutoSort: ascending first", jgexSortAscending, .SortKeys.Item(1).SortOrder
        '--- the column reads its order back out of the keys rather than
        '--- keeping one of its own -- read-only, like the original's
        AssertEquals "AutoSort: the column reports the key's order", jgexSortAscending, .Columns.Item(1).SortOrder
        AssertEquals "AutoSort: an unsorted column reports none", jgexSortNone, .Columns.Item(2).SortOrder
        '--- clicking it again flips to descending, never back to unsorted
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(50, 40)
        AssertEquals "AutoSort: still one key", 1, .SortKeys.Count
        AssertEquals "AutoSort: descending on the second click", jgexSortDescending, .SortKeys.Item(1).SortOrder
        AssertEquals "AutoSort: the column follows it down", jgexSortDescending, .Columns.Item(1).SortOrder
        '--- and a third click comes back to ascending
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(50, 40)
        AssertEquals "AutoSort: ascending again on the third", jgexSortAscending, .SortKeys.Item(1).SortOrder
        '--- another column replaces the key rather than adding to it
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(150, 40)
        AssertEquals "AutoSort: the new column is the only key", 1, .SortKeys.Count
        AssertEquals "AutoSort: keyed on the second column", 2, .SortKeys.Item(1).ColIndex
        AssertEquals "AutoSort: ascending on a fresh column", jgexSortAscending, .SortKeys.Item(1).SortOrder
        '--- and the column the key left goes back to reporting none
        AssertEquals "AutoSort: the abandoned column reports none", jgexSortNone, .Columns.Item(1).SortOrder
        AssertEquals "AutoSort: the new one reports the key", jgexSortAscending, .Columns.Item(2).SortOrder
        '--- a grouped column flips its group instead of touching the keys
        .SortKeys.Clear
        .Groups.Add 1, jgexSortAscending
        AssertEquals "AutoSort: IsGrouped follows Groups", True, .Columns.Item(1).IsGrouped
        AssertEquals "AutoSort: an ungrouped column says so", False, .Columns.Item(2).IsGrouped
        '--- grouping counts towards SortOrder with no sort key in play at all,
        '--- which is what the original's own object model dump reports
        AssertEquals "AutoSort: a grouped column reports the group's order", jgexSortAscending, .Columns.Item(1).SortOrder
        AssertEquals "AutoSort: with no sort key behind it", 0, .SortKeys.Count
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(50, 40)
        AssertEquals "AutoSort: group flipped to descending", jgexSortDescending, .Groups.Item(1).SortOrder
        AssertEquals "AutoSort: no sort key added for it", 0, .SortKeys.Count
        AssertEquals "AutoSort: the column follows the group down", jgexSortDescending, .Columns.Item(1).SortOrder
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(50, 40)
        AssertEquals "AutoSort: group back to ascending", jgexSortAscending, .Groups.Item(1).SortOrder
        '--- the chip in the group-by box stands for the same column: it sits
        '--- at x=8, y=7, sized off the caption, so (30, 15) lands inside it
        oForm.EventLog = vbNullString
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(30, 15)
        AssertEquals "GBox: chip click fires GroupByBoxHeaderClick", "GBoxClick(1);", oForm.EventLog
        AssertEquals "GBox: chip click flips the group", jgexSortDescending, .Groups.Item(1).SortOrder
        AssertEquals "GBox: no sort key added for it", 0, .SortKeys.Count
        '--- the empty part of the box is not a chip
        oForm.EventLog = vbNullString
        SendMessage .hWnd, WM_LBUTTONDOWN, 0, pvMakeLong(300, 15)
        AssertEquals "GBox: a click off the chips does nothing", vbNullString, oForm.EventLog
        AssertEquals "GBox: and leaves the order alone", jgexSortDescending, .Groups.Item(1).SortOrder
        '--- ungrouping hands the column back to the sort keys
        .Groups.Clear
        AssertEquals "AutoSort: IsGrouped cleared with the group", False, .Columns.Item(1).IsGrouped
        AssertEquals "AutoSort: and SortOrder with it", jgexSortNone, .Columns.Item(1).SortOrder
    End With
    Unload oForm
End Sub

Private Function pvMakeLong(ByVal lLo As Long, ByVal lHi As Long) As Long
    pvMakeLong = (lLo And &HFFFF&) Or (lHi * &H10000)
End Function

Private Sub pvTestKeyPressDrag()
    Dim oForm           As frmWeak

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add("A").Width = 1500
        .DataMode = jgexUnbound
        .ItemCount = 50
        .MultiSelect = True
        .Rebind
        '--- WM_CHAR raises KeyPress with the character code
        oForm.EventLog = vbNullString
        SendMessage .hWnd, WM_CHAR, 65, 0
        AssertEquals "KeyPress: WM_CHAR raises KeyPress(65)", "Press(65);", oForm.EventLog
        '--- left-button drag over rows extends the selection range from the
        '--- mouse-down anchor (row 1 after bind) to the row under the cursor
        SendMessage .hWnd, WM_MOUSEMOVE, MK_LBUTTON, pvMakeLong(30, 80)
        Assert "Drag: row under cursor selected", .RowSelected(2)
        Assert "Drag: anchor row still selected", .RowSelected(1)
        Assert "Drag: row past cursor not selected", Not .RowSelected(4)
    End With
    Unload oForm
End Sub

Private Sub pvTestSelection()
    Dim oForm           As frmWeak

    Set oForm = New frmWeak
    Load oForm
    With oForm.GridEX1
        .Columns.Add "A"
        .DataMode = jgexUnbound
        .ItemCount = 50
        .Rebind
        '--- binding selects the first row
        Assert "Sel: row 1 selected after Rebind", .RowSelected(1)
        Assert "Sel: row 2 not selected", Not .RowSelected(2)
        '--- RowSelected Let drives selection and fires SelectionChange
        .MultiSelect = True
        oForm.EventLog = vbNullString
        .RowSelected(3) = True
        Assert "Sel: row 3 selected via property", .RowSelected(3)
        Assert "Sel: row 1 still selected (multi)", .RowSelected(1)
        AssertEquals "Sel: SelectionChange fired once", "Sel;", oForm.EventLog
        '--- setting the same value again is a no-op (no event)
        oForm.EventLog = vbNullString
        .RowSelected(3) = True
        AssertEquals "Sel: no event when unchanged", vbNullString, oForm.EventLog
        .RowSelected(3) = False
        Assert "Sel: row 3 deselected", Not .RowSelected(3)
        AssertEquals "Sel: deselect fires event", "Sel;", oForm.EventLog
    End With
    Unload oForm
End Sub

Private Sub pvTestSnapshotCorpus()
    Const FUNC_NAME     As String = "pvTestSnapshotCorpus"
    Dim vFile           As Variant
    Dim sName           As String
    Dim vDoc            As Variant
    Dim oProps          As Object
    Dim oProps2         As Object
    Dim sClass          As String
    Dim oForm           As frmWeak
    Dim oCtl            As Object
    Dim sJson1          As String
    Dim sJson2          As String
    Dim lCount          As Long
    Dim sError          As String

    On Error GoTo EH
    For Each vFile In EnumFiles(App.Path & "\..\snapshots", "*.json")
        sName = Mid$(vFile, InStrRev(vFile, "\") + 1)
        JsonParse ReadTextFile(CStr(vFile)), vDoc
        If JsonValue(C2Obj(vDoc), "mode") = "persist" Then
            sClass = JsonValue(C2Obj(vDoc), "class")
            Set oProps = JsonValue(C2Obj(vDoc), "props")
            '--- fresh control instances on a disposable host form
            Set oForm = New frmWeak
            Load oForm
            If sClass = "GEXPreview" Then
                Set oCtl = oForm.GEXPreview1.Object
            Else
                Set oCtl = oForm.GridEX1.Object
            End If
            ImportObject oCtl, sClass, oProps
            Set oProps2 = JsonValue(JsonParseObject(SnapshotToJson(oCtl, sClass, False)), "props")
            pvCanonProps oProps, oProps2
            sJson1 = JsonDump(oProps)
            sJson2 = JsonDump(oProps2)
            If sJson1 <> sJson2 Then
                WriteTextFile OutputFile(sName & ".expected.txt"), sJson1
                WriteTextFile OutputFile(sName & ".actual.txt"), sJson2
            End If
            Assert "corpus round-trip " & sName, (sJson1 = sJson2)
            Unload oForm
            Set oForm = Nothing
            lCount = lCount + 1
        End If
    Next
    Assert "corpus contains snapshots", lCount > 0
    Exit Sub
EH:
    '--- PrintError resets Err, so the assert has to read it first
    sError = "&H" & Hex$(Err.Number) & " " & Err.Description
    PrintError FUNC_NAME
    Assert "corpus error in " & sName & ": " & sError, False
End Sub

Private Sub pvCanonTwips(oExp As Object, sKey As String)
    If Not IsEmpty(JsonValue(oExp, sKey)) Then
        JsonValue(oExp, sKey) = ((C2Lng(JsonValue(oExp, sKey)) + Screen.TwipsPerPixelY \ 2) \ Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub pvStripErrors(oSide As Object, oOther As Object)
    Dim oErr            As Object
    Dim vKeys           As Variant
    Dim lIdx            As Long
    Dim sKey            As String

    Set oErr = C2Obj(JsonValue(oSide, "/$errors"))
    If Not oErr Is Nothing Then
        vKeys = JsonKeys(oErr)
        If IsArray(vKeys) Then
            For lIdx = 0 To UBound(vKeys)
                sKey = vKeys(lIdx)
                If InStr(sKey, "[") > 0 Then
                    sKey = Left$(sKey, InStr(sKey, "[") - 1)
                End If
                If LenB(sKey) <> 0 Then
                    If Not IsEmpty(JsonValue(oSide, sKey)) Then
                        JsonValue(oSide, sKey) = Empty
                    End If
                    If Not IsEmpty(JsonValue(oOther, sKey)) Then
                        JsonValue(oOther, sKey) = Empty
                    End If
                End If
            Next
        End If
        JsonValue(oSide, "/$errors") = Empty
    End If
End Sub

Private Sub pvCanonProps(oExp As Object, oAct As Object)
    Dim vKeys           As Variant
    Dim lIdx            As Long
    Dim oE              As Object
    Dim oA              As Object
    Dim lCount          As Long

    If oExp Is Nothing Or oAct Is Nothing Then
        Exit Sub
    End If
    If JsonObjectType(oExp) = "array" Then
        lCount = C2Lng(JsonValue(oExp, "-1"))
        For lIdx = 0 To lCount - 1
            Set oE = C2Obj(JsonValue(oExp, lIdx))
            Set oA = C2Obj(JsonValue(oAct, lIdx))
            pvCanonProps oE, oA
        Next
        Exit Sub
    End If
    '--- drop props either side could not read ($errors) from both sides
    '--- before comparing
    pvStripErrors oExp, oAct
    pvStripErrors oAct, oExp
    '--- StdFont quantizes bitmap font sizes (e.g. MS Sans Serif 7.8 reads
    '--- back as 8.25) so canon the expected size through a real StdFont
    If Not IsEmpty(JsonValue(oExp, "Charset")) And Not IsEmpty(JsonValue(oExp, "Size")) Then
        m_oCanonFont.Name = C2Str(JsonValue(oExp, "Name"))
        m_oCanonFont.Size = C2Dbl(JsonValue(oExp, "Size"))
        JsonValue(oExp, "Size") = CDbl(m_oCanonFont.Size)
    End If
    '--- pixel-stored twips props snap on runtime set (original too), so
    '--- quantize expected design-time twips the same way
    pvCanonTwips oExp, "ColumnHeaderHeight"
    pvCanonTwips oExp, "RowHeight"
    pvCanonTwips oExp, "Width"
    pvCanonTwips oExp, "DefaultColumnWidth"
    pvCanonTwips oExp, "CardWidth"
    pvCanonTwips oExp, "CardSpacing"
    vKeys = JsonKeys(oExp)
    If IsArray(vKeys) Then
        For lIdx = 0 To UBound(vKeys)
            Set oE = C2Obj(JsonValue(oExp, CStr(vKeys(lIdx))))
            If Not oE Is Nothing Then
                Set oA = C2Obj(JsonValue(oAct, CStr(vKeys(lIdx))))
                pvCanonProps oE, oA
            End If
        Next
    End If
End Sub

