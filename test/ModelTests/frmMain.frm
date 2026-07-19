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
    Set oProps = JsonValue(pvToObj(vDoc), "props")
    Assert "checkpoint parse", Not oProps Is Nothing
    ImportObject GridEX1.Object, "GridEX", oProps
    Assert "checkpoint import", True
    sJson2 = SnapshotToJson(GridEX1.Object, "GridEX", False)
    pvSaveText App.Path & "\RoundTrip1.json", sJson1
    pvSaveText App.Path & "\RoundTrip2.json", sJson2
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
    Set oProps = JsonValue(pvToObj(vDoc), "props")
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

Private Function pvToObj(vValue As Variant) As Object
    Set pvToObj = vValue
End Function

Private Sub pvSaveText(sFile As String, sText As String)
    Dim lFile           As Long

    lFile = FreeFile
    Open sFile For Output As #lFile
    Print #lFile, sText;
    Close #lFile
End Sub
