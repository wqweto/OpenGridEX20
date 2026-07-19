VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSGridImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of JSGridImage objects."
'=========================================================================
'
' Open GridEX 2000 Control
' Collection of JSGridImage objects
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_cItems                    As Collection
Private m_hImageList                As Long

'=========================================================================
' Properties
'=========================================================================

Public Property Get Item(ByVal Index As Integer) As JSGridImage
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSGridImage of the JSGridImages Collection."
    Set Item = m_cItems.Item(Index)
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

Public Property Get hImageList() As Long
Attribute hImageList.VB_Description = "Returns/sets the handle of an external ImageList."
    hImageList = m_hImageList
End Property

Public Property Let hImageList(ByVal hValue As Long)
    m_hImageList = hValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Add(ByVal Picture As Picture) As JSGridImage
Attribute Add.VB_Description = "Adds a JSGridImage object to the collection and returns a reference to the newly created object."
    Set Add = New JSGridImage
    Add.frInit m_cItems.Count + 1, Picture
    m_cItems.Add Add
End Function

Public Sub Remove(ByVal Index As Integer)
Attribute Remove.VB_Description = "Removes a specific member from the JSGridImages collection."
    m_cItems.Remove Index
    pvReindex
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
    Set m_cItems = New Collection
End Sub

Private Sub pvReindex()
    Dim lIdx            As Long
    Dim oItem           As JSGridImage

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
