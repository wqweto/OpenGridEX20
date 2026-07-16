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

Public Property Get Item(ByVal Index As Integer) As JSGroup
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSGroup of the Collection."
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of objects in a collection."
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

Public Function Add(ByVal ColIndex As Integer, ByVal SortOrder As jgexSortOrderConstants, Optional ByVal Index As Integer) As JSGroup
Attribute Add.VB_Description = "Adds a JSGroup object to the collection and returns a reference to the newly created object."
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
End Sub

Public Sub Remove(ByVal Index As Integer)
Attribute Remove.VB_Description = "Removes a specific member from the collection."
End Sub
