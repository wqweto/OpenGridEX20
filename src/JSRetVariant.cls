VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSRetVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' Open GridEX 2000 Control
' Variant return value holder passed to events
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_vValue                    As Variant

'=========================================================================
' Properties
'=========================================================================

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_vValue
End Property

Public Property Let Value(ByVal vntValue As Variant)
    m_vValue = vntValue
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(vValue As Variant)
    m_vValue = vValue
End Sub
