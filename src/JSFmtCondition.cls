VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSFmtCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains the criteria that rows must meet to appear in a different format."
'=========================================================================
'
' Open GridEX 2000 Control
' Contains the criteria that rows must meet to appear in a different format
'
'=========================================================================
Option Explicit

Public Property Get FormatStyle() As JSFormatStyle
Attribute FormatStyle.VB_Description = "Returns the JSFormatStyle object for the JSFmtCondition."
End Property

Public Property Get ColIndex() As Integer
Attribute ColIndex.VB_Description = "Returns/sets the index of the column  to be compared."
End Property

Public Property Let ColIndex(ByVal nValue As Integer)
End Property

Public Property Get Operator() As jgexConditionOperatorConstants
Attribute Operator.VB_Description = "Returns/sets the operator used for comparison."
End Property

Public Property Let Operator(ByVal eValue As jgexConditionOperatorConstants)
End Property

Public Property Get Value1() As Variant
Attribute Value1.VB_Description = "Returns/sets the value that is compared with column values."
End Property

Public Property Let Value1(ByVal vntValue As Variant)
End Property

Public Property Get Value2() As Variant
Attribute Value2.VB_Description = "Returns/sets the value that is compared with column values."
End Property

Public Property Let Value2(ByVal vntValue As Variant)
End Property

Public Property Get Index() As Integer
Attribute Index.VB_Description = "Returns a value that represents the index of an object in a collection. "
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets a string that uniquely identifies a member in a collection."
End Property

Public Sub SetCondition(ByVal ColIndex As Integer, ByVal Operator As jgexConditionOperatorConstants, ByVal Value1 As Variant, Optional ByVal Value2 As Variant)
Attribute SetCondition.VB_Description = "Sets the condition parameters."
End Sub
