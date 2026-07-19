VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSValueList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains a List a ValueItem objects."
'=========================================================================
'
' Open GridEX 2000 Control
' Contains a List a ValueItem objects
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

Public Property Get Item(ByVal Index As Long) As JSValueItem
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSValueItem of the JSValueList."
    Set Item = m_cItems.Item(Index)
End Property

Public Property Get ItemByValue(ByVal Value As Variant) As JSValueItem
Attribute ItemByValue.VB_Description = "Returns a specific JSValueItem of the JSValueList."
    Dim lIdx            As Long

    lIdx = pvFindByValue(Value)
    If lIdx > 0 Then
        Set ItemByValue = m_cItems.Item(lIdx)
    End If
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number ofJS ValueItem objects in a JSValueList."
    Count = m_cItems.Count
End Property

Public Property Get VisibleCount() As Long
Attribute VisibleCount.VB_Description = "Returns the number of visible value items."
    Dim oItem           As JSValueItem

    For Each oItem In m_cItems
        If oItem.Visible Then
            VisibleCount = VisibleCount + 1
        End If
    Next
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Add(ByVal Value As Variant, Optional ByVal Text As String, Optional ByVal IconIndex As Integer) As JSValueItem
Attribute Add.VB_Description = "Adds a JSValueItem object and returns a reference to the newly created object."
    Set Add = New JSValueItem
    Add.frInit m_cItems.Count + 1, Value, Text, IconIndex
    m_cItems.Add Add
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all JSValueItem objects."
    Set m_cItems = New Collection
End Sub

Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes a specific JSValueItem from the JSValueList."
    m_cItems.Remove Index
    pvReindex
End Sub

Public Sub RemoveItemByValue(ByVal Value As Variant)
Attribute RemoveItemByValue.VB_Description = "Removes a specific JSValueItem from the JSValueList."
    Dim lIdx            As Long

    lIdx = pvFindByValue(Value)
    If lIdx > 0 Then
        m_cItems.Remove lIdx
        pvReindex
    End If
End Sub

Private Function pvFindByValue(Value As Variant) As Long
    Dim lIdx            As Long

    For lIdx = 1 To m_cItems.Count
        If m_cItems.Item(lIdx).Value = Value Then
            pvFindByValue = lIdx
            Exit For
        End If
    Next
End Function

Private Sub pvReindex()
    Dim lIdx            As Long
    Dim oItem           As JSValueItem

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
