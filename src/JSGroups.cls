VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of JSGroup objects."
'=========================================================================
'
' Open GridEX 2000 Control
' Collection of JSGroup objects
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_pOwner                    As GridEX
Private m_cItems                    As Collection

'=========================================================================
' Properties
'=========================================================================

Public Property Get Item(ByVal Index As Integer) As JSGroup
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSGroup of the Collection."
    Set Item = m_cItems.Item(Index)
End Property

Public Property Get Count() As Long
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

Public Function Add(ByVal ColIndex As Integer, ByVal SortOrder As jgexSortOrderConstants, Optional ByVal Index As Integer) As JSGroup
Attribute Add.VB_Description = "Adds a JSGroup object to the collection and returns a reference to the newly created object."
    Set Add = New JSGroup
    Add.frInit m_pOwner, 0, ColIndex, SortOrder
    If Index > 0 And Index <= m_cItems.Count Then
        m_cItems.Add Add, Before:=Index
    Else
        m_cItems.Add Add
    End If
    pvReindex
    pvNotify
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
    Dim oItem           As JSGroup

    '--- detached on the way out: a level the client still holds must not
    '--- keep a pointer to a control it no longer belongs to
    For Each oItem In m_cItems
        oItem.frTerminate
    Next
    Set m_cItems = New Collection
    pvNotify
End Sub

Public Sub Remove(ByVal Index As Integer)
Attribute Remove.VB_Description = "Removes a specific member from the collection."
    Dim oItem           As JSGroup

    Set oItem = m_cItems.Item(Index)
    oItem.frTerminate
    m_cItems.Remove Index
    pvReindex
    pvNotify
End Sub

Friend Sub frInit(ByVal oCtl As GridEX)
    AssignWeakRef m_pOwner, oCtl
End Sub

Friend Sub frTerminate()
    Dim oItem           As JSGroup

    '--- items the client app still holds must not keep a pointer to a
    '--- control that is going away
    For Each oItem In m_cItems
        oItem.frTerminate
    Next
    AssignWeakRef m_pOwner, Nothing
End Sub

Private Sub pvNotify()
    If Not m_pOwner Is Nothing Then
        m_pOwner.frNotifyGroupsChanged
    End If
End Sub

Private Sub pvReindex()
    Dim lIdx            As Long
    Dim oItem           As JSGroup

    For Each oItem In m_cItems
        lIdx = lIdx + 1
        oItem.frIndex = lIdx
    Next
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_cItems = New Collection
End Sub

Private Sub Class_Terminate()
    frTerminate
End Sub
