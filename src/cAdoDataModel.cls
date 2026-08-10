VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdoDataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' IDataModel over an ADO Recordset (DataMode = jgexADO)
'
' The recordset IS the cache: no cell value is ever copied into a record.
' What the model keeps per record is the ADO bookmark it was handed on the
' binding walk, and reading a cell means positioning the recordset on that
' bookmark and asking a cached Field for its value. So the recordset has to
' be scrollable and support bookmarks -- a forward-only cursor cannot back a
' grid, which is checked once at bind and raised on
'
' RowIndex is the position of a record in that binding walk and stays put
' for the life of the binding: sorting and grouping run in mdDataModel over
' m_uRowSet and permute RowPosition only. Delete compacts, as the original
' does. The recordset's own order is the RowIndex order, which is why the
' key pass walks it with MoveNext rather than seeking record by record
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cAdoDataModel"

Implements IDataModel

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const GROW_BY                   As Long = 64
Private Const adAffectCurrent           As Long = 1
Private Const adBookmark                As Long = &H2000

Private m_pOwner                    As GridEX
Private m_oRs                       As Object
Private m_lCurRowIndex              As Long
Private m_aField()                  As Object
Private m_lFieldCount               As Long
Private m_uState                    As UcsDataState
Private m_uRowSet                   As UcsRowSet
Private m_uFetch                    As UcsFetchCache

'=========================================================================
' Error management
'=========================================================================

Private Sub RaiseError(sFunction As String)
    PopRaiseError PushError, MODULE_NAME, sFunction
End Sub

Private Sub PrintError(sFunction As String)
    PopPrintError PushError, MODULE_NAME, sFunction
End Sub

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(ByVal oCtl As GridEX)
    Const FUNC_NAME     As String = "frInit"

    On Error GoTo EH
    AssignWeakRef m_pOwner, oCtl
    Set m_uState.Bookmarks = New VBA.Collection
    m_uState.OrderDirty = True
    m_uState.BookmarksDirty = True
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Friend Sub frTerminate()
    AssignWeakRef m_pOwner, Nothing
    Set m_oRs = Nothing
    Set m_uState.Bookmarks = Nothing
    Erase m_uState.Bookmark
    Erase m_aField
    m_uState.ItemCount = 0
    m_uState.BookmarkUBound = 0
    m_lFieldCount = 0
    m_lCurRowIndex = 0
End Sub

'=========================================================================
' Functions
'=========================================================================

'--- storage: this half is what makes the model an ADO one

Private Sub pvBindRecordset()
    Dim lIdx            As Long

    Set m_oRs = m_pOwner.ADORecordset
    m_uState.ItemCount = 0
    m_lCurRowIndex = 0
    Erase m_uState.Bookmark
    m_uState.BookmarkUBound = 0
    If m_oRs Is Nothing Then
        Exit Sub
    End If
    If Not m_oRs.Supports(adBookmark) Then
        Err.Raise vbObjectError, , _
            "The recordset must be scrollable and support bookmarks to be bound to a GridEX control"
    End If
    pvBindFields
    '--- the binding walk is what defines RowIndex: position in the order the
    '--- recordset hands its records over, fixed for the life of the binding
    If Not (m_oRs.BOF And m_oRs.EOF) Then
        m_oRs.MoveFirst
    End If
    Do While Not m_oRs.EOF
        lIdx = lIdx + 1
        DataEnsureBookmarkRoom m_uState, lIdx
        AssignVariant m_uState.Bookmark(lIdx), m_oRs.Bookmark
        m_oRs.MoveNext
    Loop
    m_uState.ItemCount = lIdx
End Sub

Private Sub pvBindFields()
    Dim oCol            As JSColumn

    DataInitFetchCache m_uFetch, m_pOwner
    '--- DataField is the whole of the binding: a column that names no field
    '--- is not bound to the recordset at all and stays Nothing here, which
    '--- reads as Empty. There is deliberately no fall back to the field in
    '--- the same ordinal position -- a calculated or display-only column
    '--- would then silently show whatever field happened to sit there.
    '--- FetchData wins outright, so such a column is not resolved at all
    m_lFieldCount = m_pOwner.Columns.Count
    Erase m_aField
    If m_lFieldCount = 0 Then
        Exit Sub
    End If
    ReDim m_aField(1 To m_lFieldCount) As Object
    For Each oCol In m_pOwner.Columns
        If Not m_uFetch.IsFetchData(oCol.Index) Then
            If LenB(oCol.DataField) <> 0 Then
                Set m_aField(oCol.Index) = pvFindField(oCol.DataField)
            End If
        End If
    Next
End Sub

Private Function pvFindField(sName As String) As Object
    Dim oField          As Object

    '--- walked rather than looked up: the Fields collection raises on a name
    '--- it does not hold, and the control has to stay usable under Break on
    '--- All Errors
    For Each oField In m_oRs.Fields
        If StrComp(oField.Name, sName, vbTextCompare) = 0 Then
            Set pvFindField = oField
            Exit Function
        End If
    Next
End Function

Private Sub pvSeekRecord(ByVal lRowIndex As Long)
    If lRowIndex = m_lCurRowIndex Then
        Exit Sub
    End If
    If lRowIndex < 1 Or lRowIndex > m_uState.ItemCount Then
        Exit Sub
    End If
    m_oRs.Bookmark = m_uState.Bookmark(lRowIndex)
    m_lCurRowIndex = lRowIndex
End Sub

Private Function pvGetCellValue(ByVal lRowIndex As Long, ByVal lColIndex As Long) As Variant
    If lColIndex < 1 Or lColIndex > m_lFieldCount Then
        Exit Function
    End If
    '--- FetchData outranks any DataField the column carries, so it is asked
    '--- first and the recordset is never consulted for such a column
    If m_uFetch.IsFetchData(lColIndex) Then
        AssignVariant pvGetCellValue, DataFetchValue(m_uState, m_uFetch, m_pOwner, lRowIndex, lColIndex)
        Exit Function
    End If
    If m_aField(lColIndex) Is Nothing Then
        Exit Function
    End If
    pvSeekRecord lRowIndex
    If m_lCurRowIndex <> lRowIndex Then
        Exit Function
    End If
    '--- read straight off the cached Field: nothing is copied into the
    '--- record, the recordset is the cache, and the Field already points at
    '--- whichever record the seek above left the cursor on
    AssignVariant pvGetCellValue, m_aField(lColIndex).Value
End Function

Private Sub pvWriteFields(oData As JSRowData)
    Dim lIdx            As Long

    '--- only the cells the buffer marks dirty go back, which is what keeps
    '--- this from fighting optimistic locking on a wide table: a column the
    '--- edit never touched is not this write's to disturb. A FetchData or
    '--- unbound column has nowhere to go and is dropped either way
    For lIdx = 1 To m_lFieldCount
        If Not m_aField(lIdx) Is Nothing Then
            If oData.frCellDirty(lIdx) Then
                m_aField(lIdx).Value = oData.Value(lIdx)
            End If
        End If
    Next
End Sub

Private Sub pvResyncCurrentRow()
    Const FUNC_NAME     As String = "pvResyncCurrentRow"

    '--- a provider that cannot re-read the current record refuses Resync,
    '--- and the refresh degrades to whatever the cursor already holds
    On Error GoTo EH
    m_oRs.Resync adAffectCurrent
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'--- ordering

Private Sub pvEnsureOrder()
    Dim lRow            As Long
    Dim lIdx            As Long
    Dim lCol            As Long
    Dim vValue          As Variant

    If Not m_uState.OrderDirty Then
        Exit Sub
    End If
    m_uState.OrderDirty = False
    m_uRowSet.ItemCount = m_uState.ItemCount
    m_uRowSet.GroupFooterStyle = m_pOwner.GroupFooterStyle
    m_uRowSet.DefaultCollapsed = (m_pOwner.DefaultGroupMode = jgexDGMCollapsed)
    DataReadSortKeys m_uRowSet, m_pOwner
    If m_uRowSet.SortKeyCount > 0 And m_uState.ItemCount > 0 Then
        If Not (m_oRs.BOF And m_oRs.EOF) Then
            m_oRs.MoveFirst
        End If
        '--- walked with MoveNext rather than by seeking each bookmark:
        '--- RowIndex order IS the recordset's own order, so this is one
        '--- sequential pass over the cursor instead of ItemCount random
        '--- repositionings. Filling the keys is the only step of the
        '--- projection needing a cell value, which is why it stays here
        Do While Not m_oRs.EOF And lRow < m_uState.ItemCount
            lRow = lRow + 1
            m_lCurRowIndex = lRow
            For lIdx = 1 To m_uRowSet.SortKeyCount
                lCol = m_uRowSet.SortCol(lIdx)
                If lCol > 0 Then
                    vValue = pvGetCellValue(lRow, lCol)
                    '--- an object in a cell has no ordering, so it sorts blank
                    If Not IsObject(vValue) Then
                        m_uRowSet.Key(lIdx, lRow) = vValue
                    End If
                End If
            Next
            m_oRs.MoveNext
        Loop
        m_lCurRowIndex = 0
    End If
    DataProjectRows m_uRowSet
End Sub

Private Function pvGetGroupCaption(ByVal lGroupIndex As Long, bPrefixed As Boolean) As String
    Dim oCol            As JSColumn
    Dim vValue          As Variant
    Dim sText           As String
    Dim lColIndex       As Long

    '--- formatted on demand from the record that opened the group, so the
    '--- caption follows GroupPrefix, GroupFormat and GroupEmptyStringCaption
    '--- without the projection having to be rebuilt when one of them changes
    lColIndex = m_uRowSet.GroupRow(lGroupIndex).ColIndex
    If lColIndex < 1 Or lColIndex > m_pOwner.Columns.Count Then
        Exit Function
    End If
    Set oCol = m_pOwner.Columns.Item(lColIndex)
    vValue = pvGetCellValue(m_uRowSet.GroupRow(lGroupIndex).FirstRowIndex, lColIndex)
    If Not DataIsBlankValue(vValue) And Not IsObject(vValue) And Not IsArray(vValue) Then
        If LenB(oCol.GroupFormat) <> 0 Then
            sText = Format$(vValue, oCol.GroupFormat)
        Else
            sText = C2Str(vValue)
        End If
    End If
    If LenB(sText) = 0 Then
        sText = oCol.GroupEmptyStringCaption
    End If
    bPrefixed = (LenB(oCol.GroupPrefix) <> 0)
    If bPrefixed Then
        sText = oCol.GroupPrefix & " " & sText
    End If
    pvGetGroupCaption = sText
End Function

Private Function pvBuildGroupValues(ByVal lGroupIndex As Long, ByVal lColIndex As Long, aValues() As Variant) As Long
    Dim lIdx            As Long
    Dim lCount          As Long

    '--- the records a group covers, nested group rows and footers skipped.
    '--- Every value here costs a seek, which is the price of not caching --
    '--- a SQL model answers the same question with one aggregate query
    With m_uRowSet.GroupRow(lGroupIndex)
        ReDim aValues(1 To .RecordCount + 1) As Variant
        For lIdx = .FirstSlot To .LastSlot
            If lIdx >= 1 And lIdx <= m_uRowSet.OrderCount Then
                If m_uRowSet.Order(lIdx) > 0 Then
                    lCount = lCount + 1
                    AssignVariant aValues(lCount), pvGetCellValue(m_uRowSet.Order(lIdx), lColIndex)
                End If
            End If
        Next
    End With
    pvBuildGroupValues = lCount
End Function

'--- bookmarks

'=========================================================================
' Interface IDataModel
'=========================================================================

Private Property Get IDataModel_RowCount() As Long
    Const FUNC_NAME     As String = "IDataModel_RowCount"

    On Error GoTo EH
    pvEnsureOrder
    IDataModel_RowCount = m_uRowSet.VisibleCount
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get IDataModel_ItemCount() As Long
    Const FUNC_NAME     As String = "IDataModel_ItemCount"

    On Error GoTo EH
    IDataModel_ItemCount = m_uState.ItemCount
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get IDataModel_Version() As Long
    Const FUNC_NAME     As String = "IDataModel_Version"

    On Error GoTo EH
    IDataModel_Version = m_uRowSet.Version
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get IDataModel_RowIndex(ByVal lRowPosition As Long) As Long
    Const FUNC_NAME     As String = "IDataModel_RowIndex"

    On Error GoTo EH
    pvEnsureOrder
    IDataModel_RowIndex = DataGetRecordIndex(m_uRowSet, lRowPosition)
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get IDataModel_RowBookmark(ByVal lRowIndex As Long) As Variant
    Const FUNC_NAME     As String = "IDataModel_RowBookmark"

    On Error GoTo EH
    AssignVariant IDataModel_RowBookmark, DataGetRowBookmark(m_uState, lRowIndex)
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Let IDataModel_RowBookmark(ByVal lRowIndex As Long, vValue As Variant)
    Const FUNC_NAME     As String = "IDataModel_RowBookmark"

    On Error GoTo EH
    DataSetRowBookmark m_uState, lRowIndex, vValue
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get IDataModel_RowExpanded(ByVal lRowPosition As Long) As Boolean
    Const FUNC_NAME     As String = "IDataModel_RowExpanded"

    On Error GoTo EH
    pvEnsureOrder
    IDataModel_RowExpanded = DataGetRowExpanded(m_uRowSet, lRowPosition)
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Let IDataModel_RowExpanded(ByVal lRowPosition As Long, ByVal bValue As Boolean)
    Const FUNC_NAME     As String = "IDataModel_RowExpanded"

    On Error GoTo EH
    pvEnsureOrder
    DataSetRowExpanded m_uRowSet, lRowPosition, bValue
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Sub IDataModel_SetAllExpanded(ByVal bValue As Boolean)
    Const FUNC_NAME     As String = "IDataModel_SetAllExpanded"

    On Error GoTo EH
    pvEnsureOrder
    DataSetAllExpanded m_uRowSet, bValue
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub IDataModel_Terminate()
    Const FUNC_NAME     As String = "IDataModel_Terminate"

    On Error GoTo EH
    frTerminate
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub IDataModel_RefreshGroups(ByVal bAllCollapsed As Boolean)
    Const FUNC_NAME     As String = "IDataModel_RefreshGroups"

    On Error GoTo EH
    m_uState.OrderDirty = True
    If bAllCollapsed Then
        pvEnsureOrder
        DataSetAllExpanded m_uRowSet, False
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub IDataModel_RefreshSort()
    Const FUNC_NAME     As String = "IDataModel_RefreshSort"

    On Error GoTo EH
    m_uState.OrderDirty = True
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub IDataModel_RefreshRowIndex(ByVal lRowIndex As Long)
    Const FUNC_NAME     As String = "IDataModel_RefreshRowIndex"

    On Error GoTo EH
    '--- the recordset is the cache, so there is nothing of ours to drop --
    '--- the record just has to be re-read from the source
    If lRowIndex >= 1 And lRowIndex <= m_uState.ItemCount Then
        pvSeekRecord lRowIndex
        pvResyncCurrentRow
        m_lCurRowIndex = 0
        '--- only that record's fetch columns: the rest of the cache is
        '--- still good and refetching it would be gratuitous
        DataInvalidateFetchRow m_uFetch, lRowIndex
        '--- only a value that is a sort or group key can move the record,
        '--- the same guard the update path applies: with no keys the order
        '--- cannot change and a re-sort would be wasted work
        If m_uRowSet.SortKeyCount > 0 Then
            m_uState.OrderDirty = True
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub IDataModel_Refresh()
    Const FUNC_NAME     As String = "IDataModel_Refresh"

    On Error GoTo EH
    '--- a full refresh rebinds: the recordset may be a different object, and
    '--- its order is what defines RowIndex
    pvBindRecordset
    m_uState.OrderDirty = True
    m_uState.BookmarksDirty = True
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function IDataModel_GetRowData(ByVal lRowPosition As Long) As JSRowData
    Const FUNC_NAME     As String = "IDataModel_GetRowData"
    Dim oData           As JSRowData
    Dim lColCount       As Long
    Dim lRowIndex       As Long
    Dim lIdx            As Long
    Dim bPrefixed       As Boolean
    Dim oCol            As JSColumn
    Dim eFunc           As jgexAggregateFunctionConstants
    Dim lCount          As Long
    Dim aValues()       As Variant

    On Error GoTo EH
    pvEnsureOrder
    lColCount = m_pOwner.Columns.Count
    Set oData = New JSRowData
    oData.frReset lColCount
    lRowIndex = DataGetRecordIndex(m_uRowSet, lRowPosition)
    If lRowIndex > 0 Then
        oData.frInitModel Me, lRowIndex, jgexRowTypeRecord, 0
        oData.frRowBookmark = m_uState.Bookmark(lRowIndex)
        '--- one seek for the whole row: pvGetCellValue leaves the cursor where
        '--- it put it, so the remaining columns cost a property call each
        For lIdx = 1 To lColCount
            oData.frSetValue lIdx, pvGetCellValue(lRowIndex, lIdx)
            If lIdx <= m_uFetch.ColCount Then
                If m_uFetch.IsFetchIcon(lIdx) Then
                    oData.IconIndex(lIdx) = m_pOwner.frRaiseFetchIcon(lRowIndex, lIdx, m_uState.Bookmark(lRowIndex))
                End If
            End If
        Next
    ElseIf lRowIndex < 0 Then
        With m_uRowSet.GroupRow(-lRowIndex)
            oData.frInitModel Me, lRowIndex, _
                IIf(.Footer, jgexRowTypeGroupFooter, jgexRowTypeGroupHeader), .Level
        End With
        oData.GroupCaption = pvGetGroupCaption(-lRowIndex, bPrefixed)
        oData.frGroupPrefixed = bPrefixed
        '--- a totals footer carries its columns' aggregates as the cell
        '--- values, so painting one is painting an ordinary row
        If m_uRowSet.GroupRow(-lRowIndex).Footer Then
            If m_uRowSet.GroupFooterStyle = jgexTotalsGroupFooter Then
                For Each oCol In m_pOwner.Columns
                    eFunc = oCol.AggregateFunction
                    If eFunc <> jgexAggregateNone Then
                        lCount = pvBuildGroupValues(-lRowIndex, oCol.Index, aValues)
                        oData.frSetValue oCol.Index, DataAggregateValues(aValues, lCount, eFunc, oCol.SortType)
                    End If
                Next
            End If
        End If
    End If
    Set IDataModel_GetRowData = oData
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub IDataModel_UpdateRowData(oData As JSRowData)
    Const FUNC_NAME     As String = "IDataModel_UpdateRowData"
    Dim lRowIndex       As Long

    On Error GoTo EH
    lRowIndex = oData.RowIndex
    If lRowIndex < 0 Then
        '--- a record that does not exist yet. There is no UnboundAddNew
        '--- here: the cursor mints the bookmark on Update and that is the
        '--- new record's identity
        m_oRs.AddNew
        pvWriteFields oData
        m_oRs.Update
        DataEnsureBookmarkRoom m_uState, m_uState.ItemCount + 1
        m_uState.ItemCount = m_uState.ItemCount + 1
        AssignVariant m_uState.Bookmark(m_uState.ItemCount), m_oRs.Bookmark
        m_lCurRowIndex = m_uState.ItemCount
        '--- the caller reads both off the row it just added. The position is
        '--- left at 0: it is not known until the projection is rebuilt, and
        '--- asking for it here would force a re-sort mid-insert
        oData.frInitModel Me, m_uState.ItemCount, jgexRowTypeRecord, 0
        oData.frRowBookmark = m_uState.Bookmark(m_uState.ItemCount)
        '--- the new record's fetch columns are asked for, and nothing else
        '--- in the cache is disturbed
        DataRefetchRow m_uState, m_uFetch, m_pOwner, m_uState.ItemCount
        m_uState.BookmarksDirty = True
        m_uState.OrderDirty = True
    ElseIf lRowIndex >= 1 And lRowIndex <= m_uState.ItemCount Then
        pvSeekRecord lRowIndex
        pvWriteFields oData
        m_oRs.Update
        '--- the client may compute a fetch column off what just changed, so
        '--- it is asked again -- and that alone can invalidate the order
        DataRefetchRow m_uState, m_uFetch, m_pOwner, lRowIndex
        '--- only a value that is a sort or group key can move the record. With
        '--- no keys at all nothing moves, and reprojecting anyway would bump
        '--- Version and make the grid repopulate its page -- which re-raises
        '--- RowFormat for every visible row where the original raises it once
        '--- for the edited row. Pinned by the event goldens of 058, 068, 071
        If m_uRowSet.SortKeyCount > 0 Then
            m_uState.OrderDirty = True
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub IDataModel_Delete(ByVal lRowPosition As Long)
    Const FUNC_NAME     As String = "IDataModel_Delete"
    Dim lRowIndex       As Long
    Dim lIdx            As Long

    On Error GoTo EH
    pvEnsureOrder
    lRowIndex = IDataModel_RowIndex(lRowPosition)
    If lRowIndex <= 0 Then
        Exit Sub
    End If
    pvSeekRecord lRowIndex
    m_oRs.Delete
    '--- records are indexed by RowIndex itself, so a delete closes the hole
    '--- and everything above it renumbers, as the original does
    For lIdx = lRowIndex To m_uState.ItemCount - 1
        AssignVariant m_uState.Bookmark(lIdx), m_uState.Bookmark(lIdx + 1)
    Next
    m_uState.Bookmark(m_uState.ItemCount) = Empty
    '--- the cache follows the renumbering rather than being thrown away
    DataDeleteFetchRow m_uFetch, lRowIndex, m_uState.ItemCount
    m_uState.ItemCount = m_uState.ItemCount - 1
    m_lCurRowIndex = 0
    m_uState.BookmarksDirty = True
    m_uState.OrderDirty = True
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function IDataModel_GetRowPosition(ByVal lRowIndex As Long) As Long
    Const FUNC_NAME     As String = "IDataModel_GetRowPosition"

    On Error GoTo EH
    pvEnsureOrder
    IDataModel_GetRowPosition = DataGetRowPosition(m_uState, m_uRowSet, lRowIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function IDataModel_GetRowIndex(vBookmark As Variant) As Long
    Const FUNC_NAME     As String = "IDataModel_GetRowIndex"

    On Error GoTo EH
    IDataModel_GetRowIndex = DataGetRowIndex(m_uState, vBookmark)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function IDataModel_GetRecordCount(ByVal lRowIndex As Long) As Long
    Const FUNC_NAME     As String = "IDataModel_GetRecordCount"

    On Error GoTo EH
    pvEnsureOrder
    IDataModel_GetRecordCount = DataGetRecordCount(m_uRowSet, lRowIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function IDataModel_GetSubTotal(ByVal lRowIndex As Long, ByVal lColIndex As Long, ByVal eFunc As jgexAggregateFunctionConstants) As Variant
    Const FUNC_NAME     As String = "IDataModel_GetSubTotal"
    Dim aValues()       As Variant
    Dim lCount          As Long
    Dim eSortType       As jgexSortTypeConstants

    On Error GoTo EH
    pvEnsureOrder
    If lRowIndex < 0 Then
        eSortType = jgexSortTypeString
        If lColIndex >= 1 And lColIndex <= m_pOwner.Columns.Count Then
            eSortType = m_pOwner.Columns.Item(lColIndex).SortType
        End If
        lCount = pvBuildGroupValues(-lRowIndex, lColIndex, aValues)
        IDataModel_GetSubTotal = DataAggregateValues(aValues, lCount, eFunc, eSortType)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function IDataModel_GetBookmarks(ByVal lRowIndex As Long) As Variant
    Const FUNC_NAME     As String = "IDataModel_GetBookmarks"

    On Error GoTo EH
    pvEnsureOrder
    IDataModel_GetBookmarks = DataGetBookmarks(m_uState, m_uRowSet, lRowIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function IDataModel_GetRowIndexes(ByVal lRowIndex As Long) As Variant
    Const FUNC_NAME     As String = "IDataModel_GetRowIndexes"

    On Error GoTo EH
    pvEnsureOrder
    IDataModel_GetRowIndexes = DataGetRowIndexes(m_uRowSet, lRowIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    frTerminate
End Sub
