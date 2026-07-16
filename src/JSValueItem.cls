VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSValueItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a text for replacement when a cell matches its value."
'=========================================================================
'
' Open GridEX 2000 Control
' Represents a text for replacement when a cell matches its value
'
'=========================================================================
Option Explicit

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns a value that represents the index of an object in a collection. "
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text that is linked to the Value."
End Property

Public Property Let Text(ByVal sValue As String)
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Returns/sets the value for the object."
End Property

Public Property Let Value(ByVal vntValue As Variant)
End Property

Public Property Get IconIndex() As Integer
Attribute IconIndex.VB_Description = "Returns/sets the index of the JSGridImage that is linked to the value."
End Property

Public Property Let IconIndex(ByVal nValue As Integer)
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Determines whether a JSValueItem is visible or hidden in a drop-down list."
End Property

Public Property Let Visible(ByVal bValue As Boolean)
End Property
