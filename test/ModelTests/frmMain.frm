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

'=========================================================================
' API
'=========================================================================

Private Const WM_VSCROLL            As Long = &H115
Private Const WM_KEYDOWN            As Long = &H100
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_LBUTTONUP          As Long = &H202
Private Const SB_LINEUP             As Long = 0
Private Const SB_LINEDOWN           As Long = 1
Private Const SB_PAGEDOWN           As Long = 3

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oCanonFont                As New StdFont

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"

    On Error GoTo EH
    TestInit App.Path & "\ModelTests.out.txt"
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
    pvTestKeyNav
    pvTestMouse
    pvTestSelection
    pvTestSnapshotCorpus
QH:
    TestsDone
    Unload Me
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
    Assert "Unhandled error &H" & Hex$(Err.Number) & " " & Err.Description, False
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
    WriteTextFile App.Path & "\RoundTrip1.json", sJson1
    WriteTextFile App.Path & "\RoundTrip2.json", sJson2
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
    AssertEquals "WeakRef: RowIndex works after unload", 1, oRD.RowIndex
    On Error Resume Next
    vValue = oRD.Value(1)
    lErr = Err.Number
    On Error GoTo 0
    AssertEquals "WeakRef: orphaned access raises 91", 91, lErr
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
        '--- Rebind positioned the current cell on (1,1)
        AssertEquals "Unbound: nav event order", "RowCol(1,1);RowCol(2,1);First;", oForm.EventLog
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
        AssertEquals "KeyNav: initial Col", 1, .Col
        '--- arrows drive the current cell through the subclassed proc
        SendMessage .hWnd, WM_KEYDOWN, vbKeyDown, 0
        AssertEquals "KeyNav: Down -> Row 2", 2, .Row
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
    End With
    Unload oForm
End Sub

Private Function pvMakeLong(ByVal lLo As Long, ByVal lHi As Long) As Long
    pvMakeLong = (lLo And &HFFFF&) Or (lHi * &H10000)
End Function

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
                WriteTextFile App.Path & "\" & sName & ".expected.txt", sJson1
                WriteTextFile App.Path & "\" & sName & ".actual.txt", sJson2
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
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
    Assert "corpus error in " & sName & ": &H" & Hex$(Err.Number) & " " & Err.Description, False
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

