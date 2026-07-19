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
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_vBookmark                 As Variant
Private m_lRowIndex                 As Long
Private m_lRowPosition              As Long
Private m_eRowType                  As jgexRowTypeConstants

'=========================================================================
' Properties
'=========================================================================

Public Property Get Bookmark() As Variant
Attribute Bookmark.VB_Description = "Returns the bookmark of a selected row."
    Bookmark = m_vBookmark
End Property

Public Property Get RowIndex() As Long
Attribute RowIndex.VB_Description = "Returns the original index of a selected row."
    RowIndex = m_lRowIndex
End Property

Public Property Get RowPosition() As Long
Attribute RowPosition.VB_Description = "Returns the current position of a selected row."
    RowPosition = m_lRowPosition
End Property

Public Property Get RowType() As jgexRowTypeConstants
Attribute RowType.VB_Description = "Returns the type of a selected row."
    RowType = m_eRowType
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(vBookmark As Variant, ByVal lRowIndex As Long, ByVal lRowPosition As Long, ByVal eRowType As jgexRowTypeConstants)
    m_vBookmark = vBookmark
    m_lRowIndex = lRowIndex
    m_lRowPosition = lRowPosition
    m_eRowType = eRowType
End Sub
