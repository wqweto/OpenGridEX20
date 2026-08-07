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

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As JSRowData)
    Const FUNC_NAME     As String = "GridEX1_UnboundReadData"
    Dim nIdx            As Integer

    On Error GoTo EH
    EventLog = EventLog & "Read(" & RowIndex & ")" & C2Str(Bookmark) & ";"
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

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Const FUNC_NAME     As String = "GridEX1_RowColChange"

    On Error GoTo EH
    EventLog = EventLog & "RowCol(" & LastRow & "," & LastCol & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_FirstItemChange()
    Const FUNC_NAME     As String = "GridEX1_FirstItemChange"

    On Error GoTo EH
    EventLog = EventLog & "First;"
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

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_ColResize"

    On Error GoTo EH
    EventLog = EventLog & "Resize(" & ColIndex & "," & NewColWidth & ");"
    Cancel.Value = CancelResize
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeColumnDrag(ByVal Column As JSColumn, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeColumnDrag"

    On Error GoTo EH
    EventLog = EventLog & "BeforeColDrag(" & Column.Caption & ");"
    Cancel.Value = CancelColDrag
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeColMove(ByVal Column As JSColumn, ByVal NewPosition As Integer, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeColMove"

    On Error GoTo EH
    EventLog = EventLog & "BeforeMove(" & Column.Caption & "," & NewPosition & ");"
    Cancel.Value = CancelMove
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_AfterColMove()
    Const FUNC_NAME     As String = "GridEX1_AfterColMove"

    On Error GoTo EH
    EventLog = EventLog & "AfterMove;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As JSColumn)
    Const FUNC_NAME     As String = "GridEX1_ColumnHeaderClick"

    On Error GoTo EH
    EventLog = EventLog & "HdrClick(" & Column.Caption & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_BeforeGroupChange(ByVal Group As JSGroup, ByVal ChangeOperation As jgexGroupChange, ByVal GroupPosition As Integer, ByVal Cancel As JSRetBoolean)
    Const FUNC_NAME     As String = "GridEX1_BeforeGroupChange"

    On Error GoTo EH
    EventLog = EventLog & "BeforeGroup(" & Group.ColIndex & "," & ChangeOperation & "," & GroupPosition & ");"
    Cancel.Value = CancelGroup
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_AfterGroupChange()
    Const FUNC_NAME     As String = "GridEX1_AfterGroupChange"

    On Error GoTo EH
    EventLog = EventLog & "AfterGroup;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_GroupByBoxHeaderClick(ByVal Group As JSGroup)
    Const FUNC_NAME     As String = "GridEX1_GroupByBoxHeaderClick"

    On Error GoTo EH
    EventLog = EventLog & "GBoxClick(" & Group.ColIndex & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_SelectionChange()
    Const FUNC_NAME     As String = "GridEX1_SelectionChange"

    On Error GoTo EH
    EventLog = EventLog & "Sel;"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub GridEX1_KeyPress(KeyAscii As Integer)
    Const FUNC_NAME     As String = "GridEX1_KeyPress"

    On Error GoTo EH
    EventLog = EventLog & "Press(" & KeyAscii & ");"
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub
