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
        AssertEquals "Columns(2).Width", 777, .Item(2).Width
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
    With GridEX1.FormatStyles
        .Add("Header").FontBold = True
        .Add "Totals"
        AssertEquals "FormatStyles.Count", 2, .Count
        AssertEquals "Item(Header).FontBold", True, .Item("Header").FontBold
        AssertEquals "Item(1).Name", "Header", .Item(1).Name
        .Remove "Header"
        AssertEquals "Count after Remove", 1, .Count
        .Clear
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
        '--- HoldSortSettings property (default True) as the default
        .SortKeys.Add 1, jgexSortAscending
        .Rebind
        AssertEquals "Unbound: Rebind holds sort by default", 1, .SortKeys.Count
        .Rebind False
        AssertEquals "Unbound: Rebind False clears sort", 0, .SortKeys.Count
        Assert "Unbound: Rebind clears bookmark", IsEmpty(.RowBookmark(3))
        '--- navigation state events with previous row/col in the args
        oForm.EventLog = vbNullString
        .Row = 2
        .Col = 2
        .Row = 2
        .FirstItem = 2
        .FirstItem = 2
        AssertEquals "Unbound: nav event order", "RowCol(0,0);RowCol(2,0);First;", oForm.EventLog
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

Private Sub pvCanonProps(oExp As Object, oAct As Object)
    Dim vKeys           As Variant
    Dim lIdx            As Long
    Dim oErr            As Object
    Dim oE              As Object
    Dim oA              As Object
    Dim sKey            As String
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
    '--- drop props the original could not read at design time ($errors)
    '--- from both sides before comparing
    Set oErr = C2Obj(JsonValue(oExp, "/$errors"))
    If Not oErr Is Nothing Then
        vKeys = JsonKeys(oErr)
        If IsArray(vKeys) Then
            For lIdx = 0 To UBound(vKeys)
                sKey = vKeys(lIdx)
                If InStr(sKey, "[") > 0 Then
                    sKey = Left$(sKey, InStr(sKey, "[") - 1)
                End If
                If LenB(sKey) <> 0 Then
                    If Not IsEmpty(JsonValue(oExp, sKey)) Then
                        JsonValue(oExp, sKey) = Empty
                    End If
                    If Not IsEmpty(JsonValue(oAct, sKey)) Then
                        JsonValue(oAct, sKey) = Empty
                    End If
                End If
            Next
        End If
        JsonValue(oExp, "/$errors") = Empty
    End If
    '--- StdFont quantizes bitmap font sizes (e.g. MS Sans Serif 7.8 reads
    '--- back as 8.25) so canon the expected size through a real StdFont
    If Not IsEmpty(JsonValue(oExp, "Charset")) And Not IsEmpty(JsonValue(oExp, "Size")) Then
        m_oCanonFont.Name = C2Str(JsonValue(oExp, "Name"))
        m_oCanonFont.Size = C2Dbl(JsonValue(oExp, "Size"))
        JsonValue(oExp, "Size") = CDbl(m_oCanonFont.Size)
    End If
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

