VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSFormatStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains a collection of JSFormatStyle objects."
'=========================================================================
'
' Open GridEX 2000 Control
' Contains a collection of JSFormatStyle objects
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_cItems                    As Collection

'=========================================================================
' Properties
'=========================================================================

Public Property Get Item(ByVal Index As Variant) As JSFormatStyle
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a specific JSFormatStyle of the collection."
    Set Item = m_cItems.Item(Index)
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of objects in a collection."
    Count = m_cItems.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Add(ByVal Name As String) As JSFormatStyle
Attribute Add.VB_Description = "Adds a JSFormatStyle object to the collection."
    Set Add = New JSFormatStyle
    Add.frInit Name
    m_cItems.Add Add, Name
End Function

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a specific member from the collection."
    Dim oStyle          As JSFormatStyle

    '--- the built-in system styles are protected, as in the original;
    '--- the typed local keeps the Friend access early-bound
    Set oStyle = m_cItems.Item(Index)
    If oStyle.frIsSystem Then
        Err.Raise 380 '--- Invalid property value
    End If
    m_cItems.Remove Index
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all the JSFormatStyle objects from the collection."
    Dim lIdx            As Long
    Dim oStyle          As JSFormatStyle

    '--- Clear keeps the built-in system styles (matches the original)
    For lIdx = m_cItems.Count To 1 Step -1
        Set oStyle = m_cItems.Item(lIdx)
        If Not oStyle.frIsSystem Then
            m_cItems.Remove lIdx
        End If
    Next
End Sub

'=========================================================================
' Friend methods
'=========================================================================

Friend Function frItemOrNothing(sName As String) As JSFormatStyle
    Dim vItem           As Variant

    '--- non-raising lookup so painting stays Break-on-All-Errors safe
    If SearchCollection(m_cItems, sName, vItem) Then
        Set frItemOrNothing = vItem
    End If
End Function

'=========================================================================
' Functions
'=========================================================================

Private Sub pvAddSystem(sName As String, ByVal clrBack As OLE_COLOR, ByVal clrFore As OLE_COLOR)
    With Add(sName)
        .frIsSystem = True
        .BackColor = clrBack
        .ForeColor = clrFore
    End With
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_cItems = New Collection
    '--- built-in styles present on a fresh original control; protected
    '--- from Remove and preserved across Clear; creation order matches
    '--- the original (SelectedRow first)
    pvAddSystem "SelectedRow", vbHighlight, vbHighlightText
    pvAddSystem "Default", vbWindowBackground, vbWindowText
    pvAddSystem "OddRow", &HBFFFFF, vbWindowText
    pvAddSystem "EvenRow", &HC1D7B0, vbWindowText
    pvAddSystem "RowGroup", vbButtonFace, vbButtonText
    pvAddSystem "PreviewRow", vbWindowBackground, &HFF0000
End Sub
