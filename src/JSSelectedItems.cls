VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSSelectedItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a collection of selected rows in a GridEX control."
'=========================================================================
'
' Open GridEX 2000 Control
' Represents a collection of selected rows in a GridEX control
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

Public Property Get Item(ByVal Index As Long) As JSSelectedItem
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSSelectedItem of the collection."
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

'--- the three Adds differ only in which of a row's three names the caller
'--- holds: whichever it is, the other two are resolved through the owner
'--- before the item is stored, since an item keeps which record it is and a
'--- position is worthless after the next re-sort
Public Sub Add(ByVal RowPosition As Long)
Attribute Add.VB_Description = "Selects a row given its current position."
    Dim lRowIndex       As Long

    If m_pOwner Is Nothing Then
        Exit Sub
    End If
    lRowIndex = m_pOwner.RowIndex(RowPosition)
    pvAdd m_pOwner.RowBookmark(lRowIndex), lRowIndex, RowPosition
End Sub

Public Sub AddBookmark(ByVal Bookmark As Variant)
Attribute AddBookmark.VB_Description = "Selects a row its bookmark."
    Dim lRowIndex       As Long

    If m_pOwner Is Nothing Then
        Exit Sub
    End If
    lRowIndex = m_pOwner.frGetRowIndex(Bookmark)
    If lRowIndex > 0 Then
        pvAdd Bookmark, lRowIndex, m_pOwner.frGetRowPosition(lRowIndex)
    End If
End Sub

Public Sub AddRowIndex(ByVal RowIndex As Long)
Attribute AddRowIndex.VB_Description = "Selects a row given its original row index."
    If m_pOwner Is Nothing Then
        Exit Sub
    End If
    pvAdd m_pOwner.RowBookmark(RowIndex), RowIndex, m_pOwner.frGetRowPosition(RowIndex)
End Sub

Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes a specific member from the collection."
    m_cItems.Remove Index
End Sub

Public Sub RemoveBookmark(ByVal Bookmark As Variant)
Attribute RemoveBookmark.VB_Description = "Deselects a row given its bookmark."
    Dim lRowIndex       As Long

    '--- resolved to the stable space and removed by that: bookmarks compare
    '--- by key, not by =, which an ADO byte-array bookmark would raise a
    '--- type mismatch over
    If m_pOwner Is Nothing Then
        Exit Sub
    End If
    lRowIndex = m_pOwner.frGetRowIndex(Bookmark)
    If lRowIndex > 0 Then
        RemoveRowIndex lRowIndex
    End If
End Sub

Public Sub RemoveRowIndex(ByVal RowIndex As Long)
Attribute RemoveRowIndex.VB_Description = "Deselects a row given its original row index."
    Dim lIdx            As Long
    Dim oItem           As JSSelectedItem

    For Each oItem In m_cItems
        lIdx = lIdx + 1
        If oItem.RowIndex = RowIndex Then
            m_cItems.Remove lIdx
            Exit For
        End If
    Next
End Sub

Public Sub RemoveRowPosition(ByVal RowPosition As Long)
Attribute RemoveRowPosition.VB_Description = "Deselects a row given its current row position."
    Dim lIdx            As Long
    Dim oItem           As JSSelectedItem

    For Each oItem In m_cItems
        lIdx = lIdx + 1
        If oItem.RowPosition = RowPosition Then
            m_cItems.Remove lIdx
            Exit For
        End If
    Next
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all selected rows in the collection."
    Set m_cItems = New Collection
End Sub

Friend Sub frAdd(ByVal RowPosition As Long, vBookmark As Variant, ByVal RowIndex As Long)
    pvAdd vBookmark, RowIndex, RowPosition
End Sub

Friend Sub frInit(ByVal oCtl As GridEX)
    AssignWeakRef m_pOwner, oCtl
End Sub

Friend Sub frTerminate()
    AssignWeakRef m_pOwner, Nothing
End Sub

Private Sub pvAdd(vBookmark As Variant, ByVal lRowIndex As Long, ByVal lRowPosition As Long)
    Dim oItem           As JSSelectedItem

    Set oItem = New JSSelectedItem
    oItem.frInit vBookmark, lRowIndex, lRowPosition, jgexRowTypeRecord
    m_cItems.Add oItem
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
