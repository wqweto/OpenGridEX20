VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a group level in a GridEX control."
'=========================================================================
'
' Open GridEX 2000 Control
' Represents a group level in a GridEX control
'
'=========================================================================
Option Explicit

Public Property Get Index() As Integer
Attribute Index.VB_Description = "Returns a value that represents the index of the JSGroup in the JSGroups collection."
End Property

Public Property Get SortOrder() As jgexSortOrderConstants
End Property

Public Property Let SortOrder(ByVal eValue As jgexSortOrderConstants)
End Property

Public Property Get ColIndex() As Integer
Attribute ColIndex.VB_Description = "Returns/sets the index of the grouped column."
End Property

Public Property Let ColIndex(ByVal nValue As Integer)
End Property
