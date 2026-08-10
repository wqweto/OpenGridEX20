VERSION 5.00
Object = "{24F4AB9F-37F4-43D4-B383-FB6CD721B629}#1.0#0"; "OpenGridEX20.ocx"
Begin VB.Form frmWeak
   Caption         =   "WeakRef host"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   ScaleHeight     =   3000
   ScaleWidth      =   4500
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
   Begin OpenGridEX20.GEXPreview GEXPreview1
      Height          =   800
      Left            =   3180
      TabIndex        =   1
      Top             =   60
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1411
   End
End
Attribute VB_Name = "frmWeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' Disposable host form for the JSRowData weak reference test, for fresh
' control instances in the snapshot corpus round-trip test and as event
' sink for the unbound data pipeline test (appends to EventLog)
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "frmWeak"

'=========================================================================
' Constants and member variables
'=========================================================================

Public EventLog                     As String
'--- what column 1 already reads as when RowFormat gets the buffer, which is
'--- where Format and the value list have written before the client looks
Public SeenDisplay                  As String
Public CancelResize                 As Boolean
Public CancelMove                   As Boolean
Public CancelGroup                  As Boolean
Public CancelColDrag                As Boolean
'--- what the feed should hand out, by row and column. The original refuses
'--- a write through JSRowData.Value outside UnboundReadData -- probed: it
'--- raises &H80040000 "'Value' property can not be change in this context."
'--- -- so a test wanting particular data seeds it here and lets the feed
'--- deliver it. An entry left Empty falls back to the generated "RnCn"
Private m_aSeed(1 To 64, 1 To 8)    As Variant

'=========================================================================
' Methods
'=========================================================================

'--- assigning into a Public array through an object reference writes to a
'--- copy in VB6, silently, so the seed goes in through a Sub
Public Sub SetSeed(ByVal lRow As Long, ByVal nCol As Integer, vValue As Variant)
    If lRow >= 1 And lRow <= UBound(m_aSeed, 1) And nCol >= 1 And nCol <= UBound(m_aSeed, 2) Then
        m_aSeed(lRow, nCol) = vValue
    End If
End Sub

'=========================================================================
' Error management
'=========================================================================

Private Sub PrintError(sFunction As String)
    PopPrintError PushError, MODULE_NAME, sFunction
End Sub

'=========================================================================
' Control events
'=========================================================================

'--- every event the control declares is sinked here under the name it is
'--- declared with, so a test asserts on the log the client would see and a
'--- newly raised event shows up as a diff rather than as silence. What the
'--- entry carries is the behaviour of the raise: the JSRet* and JSRowData
'--- carriers are left out -- their identity says nothing and their contents
'--- the tests read off the control -- and so are Shift and the mouse X/Y,
'--- which are environment. A position stated in twips does not survive the
'--- round trip through pixels at every DPI, so logging it would pin the
'--- machine the log was recorded on rather than the control

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeColEdit"

    On Error GoTo EH
    EventLog = EventLog & "BeforeColEdit(" & ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
    Const FUNC_NAME     As String = "GridEX1_AfterColEdit"

    On Error GoTo EH
    EventLog = EventLog & "AfterColEdit(" & ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeColUpdate"

    On Error GoTo EH
    EventLog = EventLog & "BeforeColUpdate(" & Row & "," & ColIndex & "," & OldValue & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_Error(ByVal ErrNumber As Long, ByVal DisplayMessage As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_Error"

    On Error GoTo EH
    EventLog = EventLog & "Error(" & ErrNumber & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeUpdate(ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeUpdate"

    On Error GoTo EH
    EventLog = EventLog & "BeforeUpdate;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_AfterColUpdate(ByVal ColIndex As Integer)
    Const FUNC_NAME     As String = "GridEX1_AfterColUpdate"

    On Error GoTo EH
    EventLog = EventLog & "AfterColUpdate(" & ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_AfterUpdate()
    Const FUNC_NAME     As String = "GridEX1_AfterUpdate"

    On Error GoTo EH
    EventLog = EventLog & "AfterUpdate;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_Change()
    Const FUNC_NAME     As String = "GridEX1_Change"

    On Error GoTo EH
    EventLog = EventLog & "Change;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As JSColumn)
    Const FUNC_NAME     As String = "GridEX1_ColumnHeaderClick"

    On Error GoTo EH
    EventLog = EventLog & "ColumnHeaderClick(" & Column.Caption & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_GroupByBoxHeaderClick(ByVal Group As JSGroup)
    Const FUNC_NAME     As String = "GridEX1_GroupByBoxHeaderClick"

    On Error GoTo EH
    EventLog = EventLog & "GroupByBoxHeaderClick(" & Group.ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeColMove(ByVal Column As JSColumn, ByVal NewPosition As Integer, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeColMove"

    On Error GoTo EH
    EventLog = EventLog & "BeforeColMove(" & Column.Caption & "," & NewPosition & ");"
    Cancel.Value = CancelMove
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_AfterColMove()
    Const FUNC_NAME     As String = "GridEX1_AfterColMove"

    On Error GoTo EH
    EventLog = EventLog & "AfterColMove;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeColumnDrag(ByVal Column As JSColumn, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeColumnDrag"

    On Error GoTo EH
    EventLog = EventLog & "BeforeColumnDrag(" & Column.Caption & ");"
    Cancel.Value = CancelColDrag
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeGroupDrag(ByVal Group As JSGroup, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeGroupDrag"

    On Error GoTo EH
    EventLog = EventLog & "BeforeGroupDrag(" & Group.ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_AfterGroupChange()
    Const FUNC_NAME     As String = "GridEX1_AfterGroupChange"

    On Error GoTo EH
    EventLog = EventLog & "AfterGroupChange;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeGroupChange(ByVal Group As JSGroup, ByVal ChangeOperation As jgexGroupChange, ByVal GroupPosition As Integer, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeGroupChange"

    On Error GoTo EH
    EventLog = EventLog & "BeforeGroupChange(" & Group.ColIndex & "," & ChangeOperation & "," & GroupPosition & ");"
    Cancel.Value = CancelGroup
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_AfterDelete()
    Const FUNC_NAME     As String = "GridEX1_AfterDelete"

    On Error GoTo EH
    EventLog = EventLog & "AfterDelete;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeDelete(ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeDelete"

    On Error GoTo EH
    EventLog = EventLog & "BeforeDelete;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_FetchData(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal Value As JSRetVariant)
    Const FUNC_NAME     As String = "GridEX1_FetchData"

    On Error GoTo EH
    EventLog = EventLog & "FetchData(" & RowIndex & "," & ColIndex & ")" & C2Str(RowBookmark) & ";"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As JSRetInteger)
    Const FUNC_NAME     As String = "GridEX1_FetchIcon"

    On Error GoTo EH
    EventLog = EventLog & "FetchIcon(" & RowIndex & "," & ColIndex & ")" & C2Str(RowBookmark) & ";"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    Const FUNC_NAME     As String = "GridEX1_UnboundDelete"

    On Error GoTo EH
    EventLog = EventLog & "UnboundDelete(" & RowIndex & ")" & C2Str(Bookmark) & ";"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_UnboundAddNew(ByVal NewRowBookmark As JSRetVariant, ByVal Values As JSRowData)
    Const FUNC_NAME     As String = "GridEX1_UnboundAddNew"

    On Error GoTo EH
    EventLog = EventLog & "UnboundAddNew;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As JSRowData)
    Const FUNC_NAME     As String = "GridEX1_UnboundUpdate"

    On Error GoTo EH
    EventLog = EventLog & "UnboundUpdate(" & RowIndex & ")" & C2Str(Bookmark) & ";"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_Click()
    Const FUNC_NAME     As String = "GridEX1_Click"

    On Error GoTo EH
    EventLog = EventLog & "Click;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_DblClick()
    Const FUNC_NAME     As String = "GridEX1_DblClick"

    On Error GoTo EH
    EventLog = EventLog & "DblClick;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "GridEX1_KeyDown"

    On Error GoTo EH
    EventLog = EventLog & "KeyDown(" & KeyCode & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_KeyUp(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "GridEX1_KeyUp"

    On Error GoTo EH
    EventLog = EventLog & "KeyUp(" & KeyCode & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "GridEX1_MouseDown"

    On Error GoTo EH
    EventLog = EventLog & "MouseDown(" & Button & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "GridEX1_MouseMove"

    On Error GoTo EH
    EventLog = EventLog & "MouseMove(" & Button & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "GridEX1_MouseUp"

    On Error GoTo EH
    EventLog = EventLog & "MouseUp(" & Button & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_KeyPress(KeyAscii As Integer)
    Const FUNC_NAME     As String = "GridEX1_KeyPress"

    On Error GoTo EH
    EventLog = EventLog & "KeyPress(" & KeyAscii & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_FirstItemChange()
    Const FUNC_NAME     As String = "GridEX1_FirstItemChange"

    On Error GoTo EH
    EventLog = EventLog & "FirstItemChange;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_LeftColChange()
    Const FUNC_NAME     As String = "GridEX1_LeftColChange"

    On Error GoTo EH
    EventLog = EventLog & "LeftColChange;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Const FUNC_NAME     As String = "GridEX1_RowColChange"

    On Error GoTo EH
    EventLog = EventLog & "RowColChange(" & LastRow & "," & LastCol & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_ColResize"

    On Error GoTo EH
    EventLog = EventLog & "ColResize(" & ColIndex & "," & NewColWidth & ");"
    Cancel.Value = CancelResize
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_RowResize(ByVal NewRowHeight As Long, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_RowResize"

    On Error GoTo EH
    EventLog = EventLog & "RowResize(" & NewRowHeight & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_CardResize(ByVal NewCardWidth As Long, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_CardResize"

    On Error GoTo EH
    EventLog = EventLog & "CardResize(" & NewCardWidth & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_DropList(ByVal ColIndex As Integer)
    Const FUNC_NAME     As String = "GridEX1_DropList"

    On Error GoTo EH
    EventLog = EventLog & "DropList(" & ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_ListSelected(ByVal ColIndex As Integer, ByVal ValueListIndex As Long, ByVal Value As Variant)
    Const FUNC_NAME     As String = "GridEX1_ListSelected"

    On Error GoTo EH
    EventLog = EventLog & "ListSelected(" & ColIndex & "," & ValueListIndex & "," & C2Str(Value) & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As JSRowData)
    Const FUNC_NAME     As String = "GridEX1_UnboundReadData"
    Dim nIdx            As Integer

    On Error GoTo EH
    EventLog = EventLog & "UnboundReadData(" & RowIndex & ")" & C2Str(Bookmark) & ";"
    For nIdx = 1 To Values.ColCount
        If RowIndex >= 1 And RowIndex <= UBound(m_aSeed, 1) And nIdx <= UBound(m_aSeed, 2) Then
            If Not IsEmpty(m_aSeed(RowIndex, nIdx)) Then
                Values(nIdx) = m_aSeed(RowIndex, nIdx)
            Else
                Values(nIdx) = "R" & RowIndex & "C" & nIdx
            End If
        Else
            Values(nIdx) = "R" & RowIndex & "C" & nIdx
        End If
    Next
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_RowDrag(ByVal Button As Integer, ByVal Shift As Integer)
    Const FUNC_NAME     As String = "GridEX1_RowDrag"

    On Error GoTo EH
    EventLog = EventLog & "RowDrag(" & Button & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_InitCustomEdit(ByVal ColIndex As Integer, ByVal EditBackColor As OLE_COLOR, ByVal EditForeColor As OLE_COLOR, ByVal EditFont As Font)
    Const FUNC_NAME     As String = "GridEX1_InitCustomEdit"

    On Error GoTo EH
    EventLog = EventLog & "InitCustomEdit(" & ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_ShowCustomEdit(ByVal ColIndex As Integer, ByVal EditLeft As Single, ByVal EditTop As Single, ByVal EditWidth As Single, ByVal EditHeight As Single, ByVal EditVisible As Boolean)
    Const FUNC_NAME     As String = "GridEX1_ShowCustomEdit"

    On Error GoTo EH
    EventLog = EventLog & "ShowCustomEdit(" & ColIndex & "," & EditVisible & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_EndCustomEdit(ByVal ColIndex As Integer)
    Const FUNC_NAME     As String = "GridEX1_EndCustomEdit"

    On Error GoTo EH
    EventLog = EventLog & "EndCustomEdit(" & ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_OLECompleteDrag(Effect As Long)
    Const FUNC_NAME     As String = "GridEX1_OLECompleteDrag"

    On Error GoTo EH
    EventLog = EventLog & "OLECompleteDrag(" & Effect & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_OLEDragDrop(Data As JSDataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "GridEX1_OLEDragDrop"

    On Error GoTo EH
    EventLog = EventLog & "OLEDragDrop(" & Effect & "," & Button & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_OLEDragOver(Data As JSDataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Const FUNC_NAME     As String = "GridEX1_OLEDragOver"

    On Error GoTo EH
    EventLog = EventLog & "OLEDragOver(" & Effect & "," & Button & "," & State & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Const FUNC_NAME     As String = "GridEX1_OLEGiveFeedback"

    On Error GoTo EH
    EventLog = EventLog & "OLEGiveFeedback(" & Effect & "," & DefaultCursors & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_OLESetData(Data As JSDataObject, DataFormat As Integer)
    Const FUNC_NAME     As String = "GridEX1_OLESetData"

    On Error GoTo EH
    EventLog = EventLog & "OLESetData(" & DataFormat & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_OLEStartDrag(Data As JSDataObject, AllowedEffects As Long)
    Const FUNC_NAME     As String = "GridEX1_OLEStartDrag"

    On Error GoTo EH
    EventLog = EventLog & "OLEStartDrag(" & AllowedEffects & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeDeleteEX(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeDeleteEX"

    On Error GoTo EH
    EventLog = EventLog & "BeforeDeleteEX(" & RowIndex & ")" & C2Str(Bookmark) & ";"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_ColButtonClick(ByVal ColIndex As Integer)
    Const FUNC_NAME     As String = "GridEX1_ColButtonClick"

    On Error GoTo EH
    EventLog = EventLog & "ColButtonClick(" & ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As JSRowData)
    Const FUNC_NAME     As String = "GridEX1_RowFormat"

    On Error GoTo EH
    EventLog = EventLog & "RowFormat(" & RowBuffer.RowIndex & ");"
    SeenDisplay = RowBuffer.DisplayValue(1)
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforePrinting(ByVal nPages As Long, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforePrinting"

    On Error GoTo EH
    EventLog = EventLog & "BeforePrinting(" & nPages & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    Const FUNC_NAME     As String = "GridEX1_BeforePrintPage"

    On Error GoTo EH
    EventLog = EventLog & "BeforePrintPage(" & PageNumber & "," & nPages & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_PrintingProgress(ByVal PrintProgress As Single, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_PrintingProgress"

    On Error GoTo EH
    EventLog = EventLog & "PrintingProgress(" & PrintProgress & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_NotInList(ByVal ColIndex As Integer, ByVal NewData As String, ByVal CancelUpdate As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_NotInList"

    On Error GoTo EH
    EventLog = EventLog & "NotInList(" & ColIndex & "," & NewData & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_SelectionChange()
    Const FUNC_NAME     As String = "GridEX1_SelectionChange"

    On Error GoTo EH
    EventLog = EventLog & "SelectionChange;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GEXPreview1_OnCloseClick()
    Const FUNC_NAME     As String = "GEXPreview1_OnCloseClick"

    On Error GoTo EH
    EventLog = EventLog & "OnCloseClick;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GEXPreview1_OnPrintClick(ByVal UsePrintSetupDlg As JSRetBoolean)
    Const FUNC_NAME     As String = "GEXPreview1_OnPrintClick"

    On Error GoTo EH
    EventLog = EventLog & "OnPrintClick;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GEXPreview1_BeforePaginating()
    Const FUNC_NAME     As String = "GEXPreview1_BeforePaginating"

    On Error GoTo EH
    EventLog = EventLog & "BeforePaginating;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GEXPreview1_AfterPaginating()
    Const FUNC_NAME     As String = "GEXPreview1_AfterPaginating"

    On Error GoTo EH
    EventLog = EventLog & "AfterPaginating;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GEXPreview1_ZoomChanged()
    Const FUNC_NAME     As String = "GEXPreview1_ZoomChanged"

    On Error GoTo EH
    EventLog = EventLog & "ZoomChanged;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GEXPreview1_PageChanged()
    Const FUNC_NAME     As String = "GEXPreview1_PageChanged"

    On Error GoTo EH
    EventLog = EventLog & "PageChanged;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub
