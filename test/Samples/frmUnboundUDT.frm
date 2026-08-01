VERSION 5.00
Object = "{24F4AB9F-37F4-43D4-B383-FB6CD721B629}#1.0#0"; "OpenGridEX20.ocx"
Begin VB.Form frmUnboundUDT 
   Caption         =   "Unbound sample using a user defined type array"
   ClientHeight    =   5715
   ClientLeft      =   4260
   ClientTop       =   3390
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6570
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add a row in code"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin OpenGridEX20.GridEX GridEX1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8705
   End
End
Attribute VB_Name = "frmUnboundUDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type MyRecord
    LastName As String
    FirstName As String
    Age As Integer
End Type

Dim maUDTData() As MyRecord
Dim mRowCount As Long

Private Sub cmdAdd_Click()

    'first we add the row to the dataset
    mRowCount = mRowCount + 1
    ReDim Preserve maUDTData(1 To mRowCount)
    
    'set values in the last row in the array
    maUDTData(mRowCount).LastName = "Jordan"
    maUDTData(mRowCount).FirstName = "Michael"
    maUDTData(mRowCount).Age = 27
    
    'Now adjust the ItemCount
    GridEX1.ItemCount = mRowCount
    
End Sub

Private Sub Form_Load()
Dim colTemp As JSColumn
    '--- design-time state from the sample's own .frm: these two ship
    '--- without the .frx, so the recorded snapshot is raw-only and there
    '--- is no generated setup to call
    GridEX1.DataMode = jgexUnbound
    GridEX1.AllowAddNew = True
    GridEX1.AllowDelete = True

    'Set the Column layout (This can be done at design time
    
    GridEX1.Columns.Clear
    Set colTemp = GridEX1.Columns.Add("Last Name", jgexText, jgexEditTextBox, "LastName")
    colTemp.Width = 3000
    Set colTemp = GridEX1.Columns.Add("First Name", , , "FirstName")
    colTemp.Width = 3000
    Set colTemp = GridEX1.Columns.Add("Age", , , "Age")
    colTemp.Width = 600
    colTemp.SortType = jgexSortTypeNumeric
    
    'Redim our dataset
    'first dimension is used for columns
    'second dimension is used for rows
    mRowCount = 3
    ReDim maUDTData(1 To mRowCount)
    'Filling the dataset
    maUDTData(1).LastName = "Smith"
    maUDTData(1).FirstName = "John"
    maUDTData(1).Age = 22
    
    maUDTData(2).LastName = "Wright"
    maUDTData(2).FirstName = "Steven"
    maUDTData(2).Age = 35
    
    maUDTData(3).LastName = "Connors"
    maUDTData(3).FirstName = "Henry"
    maUDTData(3).Age = 47
    
    'Tell GridEX how many rows you have in your dataset
    GridEX1.ItemCount = mRowCount
    
    'See Unbound events
    
    
End Sub


Private Sub GridEX1_ColumnHeaderClick(ByVal Column As OpenGridEX20.JSColumn)
Dim grTemp As JSGroup
Dim SortOrder As jgexSortOrderConstants
    'When clicking in a column header
    
    'If column is grouped
    If Column.IsGrouped Then
        'Find the group for this column
        For Each grTemp In GridEX1.Groups
            If grTemp.ColIndex = Column.Index Then
                'do the same as clicking in the Group by box header
                GridEX1_GroupByBoxHeaderClick grTemp
                Exit For
            End If
        Next
    Else
        SortOrder = Column.SortOrder
        'Clear SortKeys
        GridEX1.SortKeys.Clear
        'Add this new sortkey
        If SortOrder = jgexSortAscending Then
            'if the column was sorted in ascending order, sort the column in descending order
            GridEX1.SortKeys.Add Column.Index, jgexSortDescending
        Else
            'if was sorted in descending order or not sorted, sort the column in ascending order
            GridEX1.SortKeys.Add Column.Index, jgexSortAscending
        End If
    End If
    
End Sub

Private Sub GridEX1_GroupByBoxHeaderClick(ByVal Group As OpenGridEX20.JSGroup)

    'When clicking in a group by box header we change SortOrder for that group
    
    Group.SortOrder = -Group.SortOrder
    
    
End Sub


Private Sub GridEX1_UnboundAddNew(ByVal NewRowBookmark As OpenGridEX20.JSRetVariant, ByVal Values As OpenGridEX20.JSRowData)

    'Increment rowcount and redim the array
    mRowCount = mRowCount + 1
    ReDim Preserve maUDTData(1 To mRowCount)
    
    'set values in the last row in the array
    maUDTData(mRowCount).LastName = Values(1)
    maUDTData(mRowCount).FirstName = Values(2)
    maUDTData(mRowCount).Age = Values(3)
    
    
End Sub

Private Sub GridEX1_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
Dim i As Long
Dim j As Long

    'If you want to prevent deletion of a row use the BeforeDelete event
    
    'First shift the rows
    For i = RowIndex To mRowCount - 1
        maUDTData(i) = maUDTData(i + 1)
    Next
    
    'decrement rowcount and redim array
    mRowCount = mRowCount - 1
    If mRowCount > 0 Then ReDim Preserve maUDTData(1 To mRowCount)
    
End Sub


Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As OpenGridEX20.JSRowData)

    'don't worry if the grid is sorted or if the column positions are changed by the user
    Values(1) = maUDTData(RowIndex).LastName
    Values(2) = maUDTData(RowIndex).FirstName
    Values(3) = maUDTData(RowIndex).Age
    
End Sub


Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As OpenGridEX20.JSRowData)

    maUDTData(RowIndex).LastName = Values(1)
    maUDTData(RowIndex).FirstName = Values(2)
    maUDTData(RowIndex).Age = Values(3)
    
End Sub


