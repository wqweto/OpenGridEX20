VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of column objects."
'=========================================================================
'
' Open GridEX 2000 Control
' Collection of column objects
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_cItems                    As Collection

'=========================================================================
' Properties
'=========================================================================

Public Property Get Item(ByVal vntIndexKey As Variant) As JSColumn
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSColumn either by index or by key."
    Set Item = m_cItems.Item(vntIndexKey)
End Property

Public Property Get Count() As Integer
Attribute Count.VB_Description = "Returns the number of objects in a collection."
    Count = m_cItems.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Add(Optional ByVal Caption As String, Optional ByVal ColumnType As jgexColumnTypeConstants = jgexText, Optional ByVal EditType As jgexEditTypeConstants = jgexEditTextBox, Optional ByVal Key As String) As JSColumn
Attribute Add.VB_Description = "Adds a JSColumn and returns a reference to the newly created object."
    Set Add = New JSColumn
    Add.frInit m_cItems.Count + 1, Caption, ColumnType, EditType, Key
    If LenB(Key) <> 0 Then
        m_cItems.Add Add, Key
    Else
        m_cItems.Add Add
    End If
End Function

Public Function ItemByPosition(ByVal ColPosition As Integer) As JSColumn
Attribute ItemByPosition.VB_Description = "Returns a specific JSColumn based on its position."
    Dim oCol            As JSColumn

    For Each oCol In m_cItems
        If oCol.ColPosition = ColPosition Then
            Set ItemByPosition = oCol
            Exit For
        End If
    Next
End Function

Public Sub Remove(ByVal vntIndexKey As Variant)
Attribute Remove.VB_Description = "Removes a specific JSColumn from the collection."
    m_cItems.Remove vntIndexKey
    pvReindex
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
    Set m_cItems = New Collection
End Sub

Private Sub pvReindex()
    Dim lIdx            As Long
    Dim oItem           As JSColumn

    For lIdx = 1 To m_cItems.Count
        Set oItem = m_cItems.Item(lIdx)
        oItem.frSetIndex lIdx
    Next
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_cItems = New Collection
End Sub
