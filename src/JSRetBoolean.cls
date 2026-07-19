VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSRetBoolean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' Open GridEX 2000 Control
' Boolean return value holder passed to events
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_bValue                    As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
    Value = m_bValue
End Property

Public Property Let Value(ByVal bValue As Boolean)
    m_bValue = bValue
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(ByVal bValue As Boolean)
    m_bValue = bValue
End Sub
