VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSSelectedItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a selected row within a multi-select control."
'=========================================================================
'
' Open GridEX 2000 Control
' Represents a selected row within a multi-select control
'
'=========================================================================
Option Explicit

Public Property Get Bookmark() As Variant
Attribute Bookmark.VB_Description = "Returns the bookmark of a selected row."
End Property

Public Property Get RowIndex() As Long
Attribute RowIndex.VB_Description = "Returns the original index of a selected row."
End Property

Public Property Get RowPosition() As Long
Attribute RowPosition.VB_Description = "Returns the current position of a selected row."
End Property

Public Property Get RowType() As jgexRowTypeConstants
Attribute RowType.VB_Description = "Returns the type of a selected row."
End Property
