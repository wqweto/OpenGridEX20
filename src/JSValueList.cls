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

Public Property Get Item(ByVal Index As Long) As JSValueItem
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSValueItem of the JSValueList."
End Property

Public Property Get ItemByValue(ByVal Value As Variant) As JSValueItem
Attribute ItemByValue.VB_Description = "Returns a specific JSValueItem of the JSValueList."
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number ofJS ValueItem objects in a JSValueList."
End Property

Public Property Get VisibleCount() As Long
Attribute VisibleCount.VB_Description = "Returns the number of visible value items."
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

Public Function Add(ByVal Value As Variant, Optional ByVal Text As String, Optional ByVal IconIndex As Integer) As JSValueItem
Attribute Add.VB_Description = "Adds a JSValueItem object and returns a reference to the newly created object."
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all JSValueItem objects."
End Sub

Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes a specific JSValueItem from the JSValueList."
End Sub

Public Sub RemoveItemByValue(ByVal Value As Variant)
Attribute RemoveItemByValue.VB_Description = "Removes a specific JSValueItem from the JSValueList."
End Sub
