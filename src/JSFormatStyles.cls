VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSFormatStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains a collection of JSFormatStyle objects."
'=========================================================================
'
' Open GridEX 2000 Control
' Contains a collection of JSFormatStyle objects
'
'=========================================================================
Option Explicit

Public Property Get Item(ByVal Index As Variant) As JSFormatStyle
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSFormatStyle of the collection."
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of objects in a collection."
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

Public Function Add(ByVal Name As String) As JSFormatStyle
Attribute Add.VB_Description = "Adds a JSFormatStyle object to the collection."
End Function

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a specific member from the collection."
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all the JSFormatStyle objects from the collection."
End Sub
