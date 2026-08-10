VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSRowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a container for data, row and cell based formats."
'=========================================================================
'
' Open GridEX 2000 Control
' Represents a container for data, row and cell based formats
'
' The wrapper carries the row itself: an IDataModel fills it for one display
' position, RowFormat decorates it, and the grid repaints from it without
' either running again. Nothing here reads back through the control, so a
' wrapper a client keeps is a snapshot of the row it was filled with and can
' outlive the grid safely -- which is what the original does
'
' The group forwards are the exception: they ask the model, and refuse on a
' version mismatch, because the span a group row covers is only meaningful
' against the projection it was built from
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Type UcsCell
    Value                   As Variant
    IconIndex               As Long
    DisplayValue            As String
    CellStyle               As String
    CellPicture             As Picture
    Dirty                   As Boolean
End Type

Private m_pModel                    As IDataModel
Private m_lRowIndex                 As Long
Private m_lVersion                  As Long
Private m_eRowType                  As jgexRowTypeConstants
Private m_lGroupLevel               As Long
Private m_vBookmark                 As Variant
Private m_sRowStyle                 As String
Private m_sGroupCaption             As String
Private m_bGroupPrefixed            As Boolean
Private m_lGroupIconIndex           As Long
Private m_bPreviewRowVisible        As Boolean
Private m_lRowHeight                As Long
Private m_lCellCount                As Long
Private m_aCell()                   As UcsCell
Private m_bAllowUpdate              As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Get Value(ByVal ColIndex As Integer) As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_Description = "Returns/sets the value of a column."
    If ColIndex >= 1 And ColIndex <= m_lCellCount Then
        AssignVariant Value, m_aCell(ColIndex).Value
    End If
End Property

Public Property Let Value(ByVal ColIndex As Integer, ByVal vntValue As Variant)
    If Not m_bAllowUpdate Then
        Err.Raise ERR_VALUE_NUMBER, , ERR_VALUE_READONLY
    End If
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        AssignVariant m_aCell(ColIndex).Value, vntValue
        m_aCell(ColIndex).Dirty = True
    End If
End Property

Public Property Get Bookmark() As Variant
Attribute Bookmark.VB_Description = "Returns the bookmark of the row."
    AssignVariant Bookmark, m_vBookmark
End Property

Public Property Get ColCount() As Integer
Attribute ColCount.VB_Description = "Returns the number of columns."
    ColCount = m_lCellCount
End Property

Public Property Get GroupLevel() As Integer
Attribute GroupLevel.VB_Description = "Returns the group level of a row if it is a group row or 0 otherwise."
    GroupLevel = m_lGroupLevel
End Property

Public Property Get RowIndex() As Long
Attribute RowIndex.VB_Description = "Returns the row index of a row."
    If m_lRowIndex > 0 Then
        RowIndex = m_lRowIndex
    End If
End Property

Public Property Get IconIndex(ByVal ColIndex As Integer) As Integer
Attribute IconIndex.VB_Description = "Returns/sets the index of the JSGridImage displayed in a cell."
    If ColIndex >= 1 And ColIndex <= m_lCellCount Then
        IconIndex = m_aCell(ColIndex).IconIndex
    End If
End Property

Public Property Let IconIndex(ByVal ColIndex As Integer, ByVal nValue As Integer)
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        m_aCell(ColIndex).IconIndex = nValue
    End If
End Property

Public Property Get DisplayValue(ByVal ColIndex As Integer) As String
Attribute DisplayValue.VB_Description = "Returns/sets the display value of a column."
    If ColIndex >= 1 And ColIndex <= m_lCellCount Then
        DisplayValue = m_aCell(ColIndex).DisplayValue
    End If
End Property

Public Property Let DisplayValue(ByVal ColIndex As Integer, ByVal sValue As String)
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        m_aCell(ColIndex).DisplayValue = sValue
    End If
End Property

Public Property Get RowType() As jgexRowTypeConstants
Attribute RowType.VB_Description = "Returns a value that represents the type of a JSRowData object."
    RowType = m_eRowType
End Property

Public Property Get RowStyle() As String
Attribute RowStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in all the cells in a row."
    RowStyle = m_sRowStyle
End Property

Public Property Let RowStyle(ByVal sValue As String)
    m_sRowStyle = sValue
End Property

Public Property Get CellStyle(ByVal ColIndex As Integer) As String
Attribute CellStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in a cell."
    If ColIndex >= 1 And ColIndex <= m_lCellCount Then
        CellStyle = m_aCell(ColIndex).CellStyle
    End If
End Property

Public Property Let CellStyle(ByVal ColIndex As Integer, ByVal sValue As String)
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        m_aCell(ColIndex).CellStyle = sValue
    End If
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_Description = "Returns the number of records in a JSRowData object representing a group row."
    If Not m_pModel Is Nothing Then
        pvCheckOwner
        RecordCount = m_pModel.GetRecordCount(m_lRowIndex)
    End If
End Property

Public Property Get RowHeight() As Long
Attribute RowHeight.VB_Description = "Returns/sets the height, in twips, of a row."
    RowHeight = m_lRowHeight
End Property

Public Property Let RowHeight(ByVal lValue As Long)
    m_lRowHeight = lValue
End Property

Public Property Get GroupCaption() As String
Attribute GroupCaption.VB_Description = "Returns/sets the caption of a group row."
    GroupCaption = m_sGroupCaption
End Property

Public Property Let GroupCaption(ByVal sValue As String)
    m_sGroupCaption = sValue
End Property

Public Property Get GroupIconIndex() As Integer
Attribute GroupIconIndex.VB_Description = "Returns/sets the icon of a group row."
    GroupIconIndex = m_lGroupIconIndex
End Property

Public Property Let GroupIconIndex(ByVal nValue As Integer)
    m_lGroupIconIndex = nValue
End Property

Public Property Get CellPicture(ByVal ColIndex As Integer) As Picture
Attribute CellPicture.VB_Description = "Returns/sets the foreground picture of a cell."
    If ColIndex >= 1 And ColIndex <= m_lCellCount Then
        Set CellPicture = m_aCell(ColIndex).CellPicture
    End If
End Property

Public Property Set CellPicture(ByVal ColIndex As Integer, ByVal oValue As Picture)
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        Set m_aCell(ColIndex).CellPicture = oValue
    End If
End Property

Public Property Get PreviewRowVisible() As Boolean
Attribute PreviewRowVisible.VB_Description = "Determines whether a preview row should be displayed in a row."
    PreviewRowVisible = m_bPreviewRowVisible
End Property

Public Property Let PreviewRowVisible(ByVal bValue As Boolean)
    m_bPreviewRowVisible = bValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function GetBookmarks() As Variant
Attribute GetBookmarks.VB_Description = "Returns an array of record bookmarks in a group row."
    If Not m_pModel Is Nothing Then
        pvCheckOwner
        GetBookmarks = m_pModel.GetBookmarks(m_lRowIndex)
    End If
End Function

Public Function GetRowIndexes() As Variant
Attribute GetRowIndexes.VB_Description = "Returns an array of row indexes in a group row."
    If Not m_pModel Is Nothing Then
        pvCheckOwner
        GetRowIndexes = m_pModel.GetRowIndexes(m_lRowIndex)
    End If
End Function

Public Function GetSubTotal(ByVal ColIndex As Integer, ByVal AggregateFunction As jgexAggregateFunctionConstants) As Variant
Attribute GetSubTotal.VB_Description = "Returns the aggregate function result of the set of values contained in a specified column on a group."
    If Not m_pModel Is Nothing Then
        pvCheckOwner
        GetSubTotal = m_pModel.GetSubTotal(m_lRowIndex, ColIndex, AggregateFunction)
    End If
End Function

'= friend ================================================================

'--- view mode, transitional: see the banner at the top of the module
'--- clears the row for reuse: the window hands the same wrapper back to the
'--- model for a different position rather than minting one per row
Friend Sub frReset(ByVal lColCount As Long)
    Set m_pModel = Nothing
    m_lRowIndex = 0
    m_eRowType = jgexRowTypeRecord
    m_lGroupLevel = 0
    m_lVersion = 0
    m_vBookmark = Empty
    m_sRowStyle = vbNullString
    m_sGroupCaption = vbNullString
    m_bGroupPrefixed = False
    m_lGroupIconIndex = 0
    m_bPreviewRowVisible = False
    m_lRowHeight = 0
    m_lCellCount = 0
    Erase m_aCell
    pvEnsureCells lColCount
End Sub

'--- the five the public surface exposes read-only, so only a model can set
'--- them, plus the position the three group forwards are asked about
Friend Sub frInitModel(pModel As IDataModel, ByVal lRowIndex As Long, _
        ByVal eRowType As jgexRowTypeConstants, ByVal lGroupLevel As Long)
    Set m_pModel = pModel
    m_lRowIndex = lRowIndex
    If Not pModel Is Nothing Then
        m_lVersion = pModel.Version
    End If
    m_eRowType = eRowType
    m_lGroupLevel = lGroupLevel
End Sub

Friend Property Get frAllowUpdate() As Boolean
    frAllowUpdate = m_bAllowUpdate
End Property

Friend Property Let frAllowUpdate(ByVal bValue As Boolean)
    Dim lIdx            As Long

    m_bAllowUpdate = bValue
    '--- the writes stand, only the marks go: a row that has stopped being the
    '--- pending one carries its new values like any other, and nothing about
    '--- it is uncommitted any more
    If Not bValue Then
        For lIdx = 1 To m_lCellCount
            m_aCell(lIdx).Dirty = False
        Next
    End If
End Property

Friend Property Get frInitRowIndex() As Long
    '--- the signed index frInitModel stored: positive a record, negative the
    '--- projection slot of a group row -- what the window carries wrappers by
    frInitRowIndex = m_lRowIndex
End Property

Friend Property Get frCellDirty(ByVal lColIndex As Long) As Boolean
    If lColIndex >= 1 And lColIndex <= m_lCellCount Then
        frCellDirty = m_aCell(lColIndex).Dirty
    End If
End Property

Friend Property Get frGroupPrefixed() As Boolean
    frGroupPrefixed = m_bGroupPrefixed
End Property

Friend Property Let frGroupPrefixed(ByVal bValue As Boolean)
    m_bGroupPrefixed = bValue
End Property

Friend Property Get frRowBookmark() As Variant
    AssignVariant frRowBookmark, m_vBookmark
End Property

Friend Property Let frRowBookmark(vValue As Variant)
    AssignVariant m_vBookmark, vValue
End Property

'--- the public Let cannot carry an object, and a cell can hold one
Friend Sub frSetValue(ByVal lColIndex As Long, vValue As Variant)
    pvEnsureCells lColIndex
    If lColIndex >= 1 Then
        AssignVariant m_aCell(lColIndex).Value, vValue
    End If
End Sub

Friend Sub frTerminate()
    Set m_pModel = Nothing
End Sub

'=========================================================================
' Functions
'=========================================================================

Private Sub pvCheckOwner()
    '--- only a group row: a record wrapper stays readable however stale the
    '--- projection has become, which is what the original does. A model that
    '--- leaves Version at 0 never mismatches and so never refuses
    If m_lRowIndex < 0 Then
        If m_pModel.Version <> m_lVersion Then
            Err.Raise ERR_ROWDATA_NUMBER, , ERR_ROWDATA_INVALID
        End If
    End If
End Sub

Private Sub pvEnsureCells(ByVal lColCount As Long)
    If lColCount > m_lCellCount Then
        ReDim Preserve m_aCell(1 To lColCount) As UcsCell
        m_lCellCount = lColCount
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    frTerminate
End Sub
