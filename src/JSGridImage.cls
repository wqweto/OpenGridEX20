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
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oPicture                  As Picture
Private m_nIndex                    As Integer

'=========================================================================
' Properties
'=========================================================================

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns or sets a graphic to be displayed."
    Set Picture = m_oPicture
End Property

Public Property Get Index() As Integer
Attribute Index.VB_Description = "Returns a value that represents the index of an object in a collection. "
    Index = m_nIndex
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(ByVal nIndex As Integer, oPicture As Picture)
    m_nIndex = nIndex
    Set m_oPicture = oPicture
End Sub

Friend Sub frSetIndex(ByVal nIndex As Integer)
    m_nIndex = nIndex
End Sub
