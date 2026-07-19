VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSRetInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' Open GridEX 2000 Control
' Integer return value holder passed to events
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_nValue                    As Integer

'=========================================================================
' Properties
'=========================================================================

Public Property Get Value() As Integer
Attribute Value.VB_UserMemId = 0
    Value = m_nValue
End Property

Public Property Let Value(ByVal nValue As Integer)
    m_nValue = nValue
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(ByVal nValue As Integer)
    m_nValue = nValue
End Sub
