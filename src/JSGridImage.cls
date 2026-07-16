VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSGridImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an image used in a GridEX control."
'=========================================================================
'
' Open GridEX 2000 Control
' Represents an image used in a GridEX control
'
'=========================================================================
Option Explicit

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns or sets a graphic to be displayed."
End Property

Public Property Get Index() As Integer
Attribute Index.VB_Description = "Returns a value that represents the index of an object in a collection. "
End Property
