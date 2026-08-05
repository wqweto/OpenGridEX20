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
' The wrapper has two modes and is meant to end up with one:
'
'   buffer mode  frInitBuffer -- the wrapper carries the row itself. An
'                IDataModel fills it for one display position, RowFormat
'                decorates it, and the grid repaints from it without either
'                running again. This is the mode everything moves to
'   view mode    frInit -- the wrapper is a window onto row storage the
'                control owns, addressed by m_lRowIndex, negative for a
'                group row. Transitional, and what GridEX.GetRowData still
'                hands out
'
' Removing view mode is mechanical by construction: every member below
' opens with a `If m_bView Then ... Exit` prologue and nothing else, so it
' comes out by deleting those prologues, frInit, m_bView and m_lRowIndex.
' Nothing in buffer mode reads either member
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Type UcsCell
    Value                   As Variant
    IconIndex               As Integer
    DisplayValue            As String
    CellStyle               As String
    CellPicture             As Picture
End Type

'--- weak ref (no AddRef) to the owning control, in both modes
Private m_pOwner                    As GridEX
'--- view mode only -- see the banner above
Private m_bView                     As Boolean
Private m_lRowIndex                 As Long
'--- buffer mode. The model reference is strong, which is safe only because
'--- the grid owns the buffers: a model that holds a buffer of its own must
'--- never call frSetRow on it, or the two keep each other alive
Private m_pModel                    As IDataModel
'--- signed like the projection's own: positive a record, negative a group
'--- header or footer
Private m_lBufRowIndex              As Long
'--- the projection generation this wrapper was filled under
Private m_lVersion                  As Long
Private m_eRowType                  As jgexRowTypeConstants
Private m_nGroupLevel               As Integer
Private m_vBookmark                 As Variant
Private m_sRowStyle                 As String
Private m_sGroupCaption             As String
Private m_nGroupIconIndex           As Integer
Private m_bPreviewRowVisible        As Boolean
Private m_lRowHeight                As Long
Private m_nCellCount                As Integer
Private m_aCell()                   As UcsCell

'=========================================================================
' Properties
'=========================================================================

Public Property Get Value(ByVal ColIndex As Integer) As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_Description = "Returns/sets the value of a column."
    If m_bView Then
        If m_pOwner Is Nothing Then
            Err.Raise 91 '--- Object variable or With block variable not set
        End If
        Value = m_pOwner.frRowValue(m_lRowIndex, ColIndex)
        Exit Property
    End If
    If ColIndex >= 1 And ColIndex <= m_nCellCount Then
        AssignVariant Value, m_aCell(ColIndex).Value
    End If
End Property

Public Property Let Value(ByVal ColIndex As Integer, ByVal vntValue As Variant)
    If m_bView Then
        If m_pOwner Is Nothing Then
            Err.Raise 91 '--- Object variable or With block variable not set
        End If
        m_pOwner.frRowValue(m_lRowIndex, ColIndex) = vntValue
        Exit Property
    End If
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        AssignVariant m_aCell(ColIndex).Value, vntValue
    End If
End Property

Public Property Get Bookmark() As Variant
Attribute Bookmark.VB_Description = "Returns the bookmark of the row."
    If m_bView Then
        Bookmark = m_pOwner.frRowBookmark(m_lRowIndex)
        Exit Property
    End If
    AssignVariant Bookmark, m_vBookmark
End Property

Public Property Get ColCount() As Integer
Attribute ColCount.VB_Description = "Returns the number of columns."
    If m_bView Then
        ColCount = m_pOwner.frRowColCount
        Exit Property
    End If
    ColCount = m_nCellCount
End Property

Public Property Get GroupLevel() As Integer
Attribute GroupLevel.VB_Description = "Returns the group level of a row if it is a group row or 0 otherwise."
    If m_bView Then
        GroupLevel = m_pOwner.frRowGroupLevel(m_lRowIndex)
        Exit Property
    End If
    GroupLevel = m_nGroupLevel
End Property

Public Property Get RowIndex() As Long
Attribute RowIndex.VB_Description = "Returns the row index of a row."
    If m_bView Then
        RowIndex = m_pOwner.frRowIndex(m_lRowIndex)
        Exit Property
    End If
    If m_lBufRowIndex > 0 Then
        RowIndex = m_lBufRowIndex
    End If
End Property

Public Property Get IconIndex(ByVal ColIndex As Integer) As Integer
Attribute IconIndex.VB_Description = "Returns/sets the index of the JSGridImage displayed in a cell."
    If m_bView Then
        IconIndex = m_pOwner.frRowIconIndex(m_lRowIndex, ColIndex)
        Exit Property
    End If
    If ColIndex >= 1 And ColIndex <= m_nCellCount Then
        IconIndex = m_aCell(ColIndex).IconIndex
    End If
End Property

Public Property Let IconIndex(ByVal ColIndex As Integer, ByVal nValue As Integer)
    If m_bView Then
        m_pOwner.frRowIconIndex(m_lRowIndex, ColIndex) = nValue
        Exit Property
    End If
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        m_aCell(ColIndex).IconIndex = nValue
    End If
End Property

Public Property Get DisplayValue(ByVal ColIndex As Integer) As String
Attribute DisplayValue.VB_Description = "Returns/sets the display value of a column."
    If m_bView Then
        DisplayValue = m_pOwner.frRowDisplayValue(m_lRowIndex, ColIndex)
        Exit Property
    End If
    If ColIndex >= 1 And ColIndex <= m_nCellCount Then
        DisplayValue = m_aCell(ColIndex).DisplayValue
    End If
End Property

Public Property Let DisplayValue(ByVal ColIndex As Integer, ByVal sValue As String)
    If m_bView Then
        m_pOwner.frRowDisplayValue(m_lRowIndex, ColIndex) = sValue
        Exit Property
    End If
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        m_aCell(ColIndex).DisplayValue = sValue
    End If
End Property

Public Property Get RowType() As jgexRowTypeConstants
Attribute RowType.VB_Description = "Returns a value that represents the type of a JSRowData object."
    If m_bView Then
        RowType = m_pOwner.frRowType(m_lRowIndex)
        Exit Property
    End If
    RowType = m_eRowType
End Property

Public Property Get RowStyle() As String
Attribute RowStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in all the cells in a row."
    If m_bView Then
        RowStyle = m_pOwner.frRowStyle(m_lRowIndex)
        Exit Property
    End If
    RowStyle = m_sRowStyle
End Property

Public Property Let RowStyle(ByVal sValue As String)
    If m_bView Then
        m_pOwner.frRowStyle(m_lRowIndex) = sValue
        Exit Property
    End If
    m_sRowStyle = sValue
End Property

Public Property Get CellStyle(ByVal ColIndex As Integer) As String
Attribute CellStyle.VB_Description = "Returns/sets the name of the JSFormatStyle to be applied in a cell."
    If m_bView Then
        CellStyle = m_pOwner.frRowCellStyle(m_lRowIndex, ColIndex)
        Exit Property
    End If
    If ColIndex >= 1 And ColIndex <= m_nCellCount Then
        CellStyle = m_aCell(ColIndex).CellStyle
    End If
End Property

Public Property Let CellStyle(ByVal ColIndex As Integer, ByVal sValue As String)
    If m_bView Then
        m_pOwner.frRowCellStyle(m_lRowIndex, ColIndex) = sValue
        Exit Property
    End If
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        m_aCell(ColIndex).CellStyle = sValue
    End If
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_Description = "Returns the number of records in a JSRowData object representing a group row."
    If m_bView Then
        RecordCount = m_pOwner.frRowRecordCount(m_lRowIndex)
        Exit Property
    End If
    If Not m_pModel Is Nothing Then
        pvCheckValid
        RecordCount = m_pModel.GetRecordCount(m_lBufRowIndex)
    End If
End Property

Public Property Get RowHeight() As Long
Attribute RowHeight.VB_Description = "Returns/sets the height, in twips, of a row."
    If m_bView Then
        RowHeight = m_pOwner.frRowHeight(m_lRowIndex)
        Exit Property
    End If
    RowHeight = m_lRowHeight
End Property

Public Property Let RowHeight(ByVal lValue As Long)
    If m_bView Then
        m_pOwner.frRowHeight(m_lRowIndex) = lValue
        Exit Property
    End If
    m_lRowHeight = lValue
End Property

Public Property Get GroupCaption() As String
Attribute GroupCaption.VB_Description = "Returns/sets the caption of a group row."
    If m_bView Then
        GroupCaption = m_pOwner.frRowGroupCaption(m_lRowIndex)
        Exit Property
    End If
    GroupCaption = m_sGroupCaption
End Property

Public Property Let GroupCaption(ByVal sValue As String)
    If m_bView Then
        m_pOwner.frRowGroupCaption(m_lRowIndex) = sValue
        Exit Property
    End If
    m_sGroupCaption = sValue
End Property

Public Property Get GroupIconIndex() As Integer
Attribute GroupIconIndex.VB_Description = "Returns/sets the icon of a group row."
    If m_bView Then
        GroupIconIndex = m_pOwner.frRowGroupIconIndex(m_lRowIndex)
        Exit Property
    End If
    GroupIconIndex = m_nGroupIconIndex
End Property

Public Property Let GroupIconIndex(ByVal nValue As Integer)
    If m_bView Then
        m_pOwner.frRowGroupIconIndex(m_lRowIndex) = nValue
        Exit Property
    End If
    m_nGroupIconIndex = nValue
End Property

Public Property Get CellPicture(ByVal ColIndex As Integer) As Picture
Attribute CellPicture.VB_Description = "Returns/sets the foreground picture of a cell."
    If m_bView Then
        Set CellPicture = m_pOwner.frRowCellPicture(m_lRowIndex, ColIndex)
        Exit Property
    End If
    If ColIndex >= 1 And ColIndex <= m_nCellCount Then
        Set CellPicture = m_aCell(ColIndex).CellPicture
    End If
End Property

Public Property Set CellPicture(ByVal ColIndex As Integer, ByVal oValue As Picture)
    If m_bView Then
        Set m_pOwner.frRowCellPicture(m_lRowIndex, ColIndex) = oValue
        Exit Property
    End If
    pvEnsureCells ColIndex
    If ColIndex >= 1 Then
        Set m_aCell(ColIndex).CellPicture = oValue
    End If
End Property

Public Property Get PreviewRowVisible() As Boolean
Attribute PreviewRowVisible.VB_Description = "Determines whether a preview row should be displayed in a row."
    If m_bView Then
        PreviewRowVisible = m_pOwner.frRowPreviewRowVisible(m_lRowIndex)
        Exit Property
    End If
    PreviewRowVisible = m_bPreviewRowVisible
End Property

Public Property Let PreviewRowVisible(ByVal bValue As Boolean)
    If m_bView Then
        m_pOwner.frRowPreviewRowVisible(m_lRowIndex) = bValue
        Exit Property
    End If
    m_bPreviewRowVisible = bValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function GetBookmarks() As Variant
Attribute GetBookmarks.VB_Description = "Returns an array of record bookmarks in a group row."
    If m_bView Then
        Exit Function
    End If
    If Not m_pModel Is Nothing Then
        pvCheckValid
        GetBookmarks = m_pModel.GetBookmarks(m_lBufRowIndex)
    End If
End Function

Public Function GetRowIndexes() As Variant
Attribute GetRowIndexes.VB_Description = "Returns an array of row indexes in a group row."
    If m_bView Then
        Exit Function
    End If
    If Not m_pModel Is Nothing Then
        pvCheckValid
        GetRowIndexes = m_pModel.GetRowIndexes(m_lBufRowIndex)
    End If
End Function

Public Function GetSubTotal(ByVal ColIndex As Integer, ByVal AggregateFunction As jgexAggregateFunctionConstants) As Variant
Attribute GetSubTotal.VB_Description = "Returns the aggregate function result of the set of values contained in a specified column on a group."
    If m_bView Then
        GetSubTotal = m_pOwner.frRowSubTotal(m_lRowIndex, ColIndex, AggregateFunction)
        Exit Function
    End If
    If Not m_pModel Is Nothing Then
        pvCheckValid
        GetSubTotal = m_pModel.GetSubTotal(m_lBufRowIndex, ColIndex, AggregateFunction)
    End If
End Function

'= friend ================================================================

'--- view mode, transitional: see the banner at the top of the module
Friend Sub frInit(ByVal oCtl As GridEX, ByVal lRowIndex As Long)
    Call CopyMemory(m_pOwner, oCtl, PTR_SIZE)
    m_lRowIndex = lRowIndex
    m_bView = True
End Sub

Friend Sub frInitBuffer(ByVal oCtl As GridEX)
    Call CopyMemory(m_pOwner, oCtl, PTR_SIZE)
    m_bView = False
End Sub

'--- clears the row for reuse: the window hands the same wrapper back to the
'--- model for a different position rather than minting one per row
Friend Sub frReset(ByVal nColCount As Integer)
    Set m_pModel = Nothing
    m_lBufRowIndex = 0
    m_eRowType = jgexRowTypeRecord
    m_nGroupLevel = 0
    m_lVersion = 0
    m_vBookmark = Empty
    m_sRowStyle = vbNullString
    m_sGroupCaption = vbNullString
    m_nGroupIconIndex = 0
    m_bPreviewRowVisible = True
    m_lRowHeight = 0
    m_nCellCount = 0
    Erase m_aCell
    pvEnsureCells nColCount
End Sub

'--- the five the public surface exposes read-only, so only a model can set
'--- them, plus the position the three group forwards are asked about
Friend Sub frSetRow(pModel As IDataModel, ByVal lRowIndex As Long, _
        ByVal eRowType As jgexRowTypeConstants, ByVal nGroupLevel As Integer)
    Set m_pModel = pModel
    m_lBufRowIndex = lRowIndex
    If Not pModel Is Nothing Then
        m_lVersion = pModel.Version
    End If
    m_eRowType = eRowType
    m_nGroupLevel = nGroupLevel
End Sub

Friend Property Get frRowBookmark() As Variant
    AssignVariant frRowBookmark, m_vBookmark
End Property

Friend Property Let frRowBookmark(vValue As Variant)
    AssignVariant m_vBookmark, vValue
End Property

'--- the public Let cannot carry an object, and a cell can hold one
Friend Sub frSetValue(ByVal nColIndex As Integer, vValue As Variant)
    pvEnsureCells nColIndex
    If nColIndex >= 1 Then
        AssignVariant m_aCell(nColIndex).Value, vValue
    End If
End Sub

Friend Sub frTerminate()
    Call CopyMemory(m_pOwner, NULL_PTR, PTR_SIZE)
    Set m_pModel = Nothing
End Sub

'=========================================================================
' Functions
'=========================================================================

Private Sub pvCheckValid()
    '--- only a group row: a record wrapper stays readable however stale the
    '--- projection has become, which is what the original does. A model that
    '--- leaves Version at 0 never mismatches and so never refuses
    If m_lBufRowIndex < 0 Then
        If m_pModel.Version <> m_lVersion Then
            Err.Raise ERR_ROWDATA_NUMBER, , ERR_ROWDATA_INVALID
        End If
    End If
End Sub

Private Sub pvEnsureCells(ByVal nColCount As Integer)
    If nColCount > m_nCellCount Then
        ReDim Preserve m_aCell(1 To nColCount) As UcsCell
        m_nCellCount = nColCount
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    '--- zero the raw pointer or VB6 releases a control it never AddRef-ed
    frTerminate
End Sub
