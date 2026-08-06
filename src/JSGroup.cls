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
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_pOwner                    As GridEX
Private m_lIndex                    As Long
Private m_eSortOrder                As jgexSortOrderConstants
Private m_lColIndex                 As Long
'--- where the control last laid this level's chip out, so painting and
'--- hit-testing read the same rectangle instead of each walking the row
Private m_uChipRect                 As RECT

'=========================================================================
' Properties
'=========================================================================

Public Property Get Index() As Integer
Attribute Index.VB_Description = "Returns a value that represents the index of the JSGroup in the JSGroups collection."
    Index = m_lIndex
End Property

Public Property Get SortOrder() As jgexSortOrderConstants
    SortOrder = m_eSortOrder
End Property

Public Property Let SortOrder(ByVal eValue As jgexSortOrderConstants)
    m_eSortOrder = eValue
    If Not m_pOwner Is Nothing Then
        m_pOwner.frNotifySortChanged
    End If
End Property

Public Property Get ColIndex() As Integer
Attribute ColIndex.VB_Description = "Returns/sets the index of the grouped column."
    ColIndex = m_lColIndex
End Property

Public Property Let ColIndex(ByVal nValue As Integer)
    m_lColIndex = nValue
    If Not m_pOwner Is Nothing Then
        m_pOwner.frNotifySortChanged
    End If
End Property

Friend Property Let frIndex(ByVal lIndex As Long)
    m_lIndex = lIndex
End Property

Friend Property Get frChipRect() As RECT
    frChipRect = m_uChipRect
End Property

Friend Property Let frChipRect(uValue As RECT)
    m_uChipRect = uValue
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(ByVal oCtl As GridEX, ByVal lIndex As Long, ByVal lColIndex As Long, ByVal eSortOrder As jgexSortOrderConstants)
    SetWeakRef m_pOwner, oCtl
    m_lIndex = lIndex
    m_lColIndex = lColIndex
    m_eSortOrder = eSortOrder
End Sub

Friend Sub frTerminate()
    SetWeakRef m_pOwner, Nothing
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    frTerminate
End Sub
