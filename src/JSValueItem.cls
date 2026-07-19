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
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_lIndex                    As Long
Private m_sText                     As String
Private m_vValue                    As Variant
Private m_nIconIndex                As Integer
Private m_bVisible                  As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns a value that represents the index of an object in a collection. "
    Index = m_lIndex
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text that is linked to the Value."
    Text = m_sText
End Property

Public Property Let Text(ByVal sValue As String)
    m_sText = sValue
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Returns/sets the value for the object."
    Value = m_vValue
End Property

Public Property Let Value(ByVal vntValue As Variant)
    m_vValue = vntValue
End Property

Public Property Get IconIndex() As Integer
Attribute IconIndex.VB_Description = "Returns/sets the index of the JSGridImage that is linked to the value."
    IconIndex = m_nIconIndex
End Property

Public Property Let IconIndex(ByVal nValue As Integer)
    m_nIconIndex = nValue
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Determines whether a JSValueItem is visible or hidden in a drop-down list."
    Visible = m_bVisible
End Property

Public Property Let Visible(ByVal bValue As Boolean)
    m_bVisible = bValue
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(ByVal lIndex As Long, vValue As Variant, sText As String, ByVal nIconIndex As Integer)
    m_lIndex = lIndex
    m_vValue = vValue
    m_sText = sText
    m_nIconIndex = nIconIndex
End Sub

Friend Sub frSetIndex(ByVal lIndex As Long)
    m_lIndex = lIndex
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_bVisible = True
    m_nIconIndex = -1
End Sub
