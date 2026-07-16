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

Public Property Get Item(ByVal Index As Long) As JSSelectedItem
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSSelectedItem of the collection."
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of objects in a collection."
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

Public Sub Add(ByVal RowPosition As Long)
Attribute Add.VB_Description = "Selects a row given its current position."
End Sub

Public Sub AddBookmark(ByVal Bookmark As Variant)
Attribute AddBookmark.VB_Description = "Selects a row its bookmark."
End Sub

Public Sub AddRowIndex(ByVal RowIndex As Long)
Attribute AddRowIndex.VB_Description = "Selects a row given its original row index."
End Sub

Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes a specific member from the collection."
End Sub

Public Sub RemoveBookmark(ByVal Bookmark As Variant)
Attribute RemoveBookmark.VB_Description = "Deselects a row given its bookmark."
End Sub

Public Sub RemoveRowIndex(ByVal RowIndex As Long)
Attribute RemoveRowIndex.VB_Description = "Deselects a row given its original row index."
End Sub

Public Sub RemoveRowPosition(ByVal RowPosition As Long)
Attribute RemoveRowPosition.VB_Description = "Deselects a row given its current row position."
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all selected rows in the collection."
End Sub
