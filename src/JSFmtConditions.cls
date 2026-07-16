VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSFmtConditions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of FmtCondition objects."
'=========================================================================
'
' Open GridEX 2000 Control
' Collection of FmtCondition objects
'
'=========================================================================
Option Explicit

Public Property Get Item(ByVal Index As Variant) As JSFmtCondition
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSFmtCondition of  the collection either by index or by key."
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of objects in a collection."
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

Public Property Get GroupCondition() As JSFmtCondition
Attribute GroupCondition.VB_Description = "Returns a JSFmtCondition object that is used through the group rows."
End Property

Public Property Get ApplyGroupCondition() As Boolean
Attribute ApplyGroupCondition.VB_Description = "Controls whether the GroupCondition property is valid."
End Property

Public Property Let ApplyGroupCondition(ByVal bValue As Boolean)
End Property

Public Property Get GroupConditionCountTitle() As String
Attribute GroupConditionCountTitle.VB_Description = "Returns/sets the text displayed in a group row when one or more rows in the group meet the criteria specified in the JSGroupCondition object's properties."
End Property

Public Property Let GroupConditionCountTitle(ByVal sValue As String)
End Property

Public Property Get ShowGroupConditionCount() As Boolean
Attribute ShowGroupConditionCount.VB_Description = "Controls whether the count of rows, that meet the criteria in GroupCondition property settings, will be displayed in a group row."
End Property

Public Property Let ShowGroupConditionCount(ByVal bValue As Boolean)
End Property

Public Function Add(ByVal ColIndex As Integer, Optional ByVal Operator As jgexConditionOperatorConstants = jgexEqual, Optional ByVal Value1 As Variant, Optional ByVal Value2 As Variant, Optional ByVal Key As String) As JSFmtCondition
Attribute Add.VB_Description = "Adds a JSFmtCondition object to the collection and returns a reference to the newly created object."
End Function

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a specific member from the JSFmtConditions collection."
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
End Sub
