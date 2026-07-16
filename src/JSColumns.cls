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

Public Property Get Item(ByVal vntIndexKey As Variant) As JSColumn
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSColumn either by index or by key."
End Property

Public Property Get Count() As Integer
Attribute Count.VB_Description = "Returns the number of objects in a collection."
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
End Sub

Public Function Add(Optional ByVal Caption As String, Optional ByVal ColumnType As jgexColumnTypeConstants = jgexText, Optional ByVal EditType As jgexEditTypeConstants = jgexEditTextBox, Optional ByVal Key As String) As JSColumn
Attribute Add.VB_Description = "Adds a JSColumn and returns a reference to the newly created object."
End Function

Public Function ItemByPosition(ByVal ColPosition As Integer) As JSColumn
Attribute ItemByPosition.VB_Description = "Returns a specific JSColumn based on its position."
End Function

Public Sub Remove(ByVal vntIndexKey As Variant)
Attribute Remove.VB_Description = "Removes a specific JSColumn from the collection."
End Sub
