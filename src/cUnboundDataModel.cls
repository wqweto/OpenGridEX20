VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUnboundDataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' IDataModel over client-supplied rows (DataMode = jgexUnbound)
'
' Records live in m_aRecord indexed by RowIndex itself -- no holes, no slot
' map -- so a Delete compacts the array and renumbers every record above it,
' exactly as the original does (probed: deleting the record holding "bk-3"
' leaves the next read of RowIndex 3 carrying "bk-4"). The bookmark map
' cannot survive that renumbering, so it is marked dirty and rebuilt on the
' next lookup rather than patched entry by entry.
'
' Sorting and grouping run in mdDataModel over m_uRowSet and touch neither
' m_aRecord's indexing nor its contents, which is what keeps a RowIndex
' stable across a re-sort. What is left here is the unbound half: values
' arrive through UnboundReadData once per record and are cached, and every
' write goes back out through UnboundUpdate, UnboundAddNew or UnboundDelete
'
'=========================================================================
Option Explicit
DefObj A-Z

Implements IDataModel

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const GROW_BY               As Long = 64

Private Type UcsRecord
    Bookmark                As Variant
    Fetched                 As Boolean
    CellCount               As Integer
    Values()                As Variant
    IconIndexes()           As Integer
End Type

Private m_pOwner                    As GridEX
Private m_aRecord()                 As UcsRecord
Private m_lItemCount                As Long
Private m_lRecordUBound             As Long
Private m_uRowSet                   As UcsRowSet
Private m_uFetch                    As UcsFetchCache
Private m_oBookmarks                As VBA.Collection
Private m_bBookmarksDirty           As Boolean
Private m_bOrderDirty               As Boolean
Private m_oFetchBuffer              As JSRowData

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(ByVal oCtl As GridEX)
    Call CopyMemory(m_pOwner, oCtl, PTR_SIZE)
    Set m_oBookmarks = New VBA.Collection
    Set m_oFetchBuffer = New JSRowData
    m_oFetchBuffer.frInitBuffer oCtl
    m_bOrderDirty = True
    m_bBookmarksDirty = True
End Sub

Friend Sub frTerminate()
    '--- zero the raw pointer or VB6 releases a control it never AddRef-ed
    Call CopyMemory(m_pOwner, NULL_PTR, PTR_SIZE)
    Set m_oBookmarks = Nothing
    Set m_oFetchBuffer = Nothing
    Erase m_aRecord
    m_lItemCount = 0
    m_lRecordUBound = 0
End Sub

'=========================================================================
' Functions
'=========================================================================

'--- storage: this half is what makes the model an unbound one

Private Sub pvEnsureCapacity(ByVal lCount As Long)
    Dim lIdx            As Long

    If lCount > m_lRecordUBound Then
        m_lRecordUBound = lCount + GROW_BY
        ReDim Preserve m_aRecord(1 To m_lRecordUBound) As UcsRecord
    End If
    For lIdx = m_lItemCount + 1 To lCount
        pvResetRecord lIdx, bDropBookmark:=True
    Next
End Sub

Private Sub pvResetRecord(ByVal lRowIndex As Long, ByVal bDropBookmark As Boolean)
    With m_aRecord(lRowIndex)
        .Fetched = False
        If .CellCount > 0 Then
            ReDim .Values(1 To .CellCount) As Variant
            ReDim .IconIndexes(1 To .CellCount) As Integer
        End If
        '--- probed at both: the original keeps bookmarks across Refetch AND
        '--- across Rebind, so nothing but a Delete or an explicit assignment
        '--- ever drops one
        If bDropBookmark Then
            .Bookmark = Empty
        End If
    End With
End Sub

Private Sub pvEnsureCells(ByVal lRowIndex As Long)
    Dim nColCount       As Integer

    nColCount = m_pOwner.Columns.Count
    With m_aRecord(lRowIndex)
        If .CellCount < nColCount Then
            ReDim Preserve .Values(1 To nColCount) As Variant
            ReDim Preserve .IconIndexes(1 To nColCount) As Integer
            .CellCount = nColCount
        End If
    End With
End Sub

Private Sub pvFetchRecord(ByVal lRowIndex As Long)
    If lRowIndex < 1 Or lRowIndex > m_lItemCount Then
        Exit Sub
    End If
    pvEnsureCells lRowIndex
    If m_aRecord(lRowIndex).Fetched Then
        Exit Sub
    End If
    '--- marked first so a read from inside the client's handler does not
    '--- re-fire the fetch for the same record
    m_aRecord(lRowIndex).Fetched = True
    '--- the buffer is cleared and sized before the client sees it: a handler
    '--- writes its columns as "For i = 1 To Values.ColCount", so an unsized
    '--- one would silently return no data at all. The model is deliberately
    '--- passed as Nothing -- this buffer belongs to the model, and handing it
    '--- a reference back would have the two keep each other alive
    m_oFetchBuffer.frReset m_aRecord(lRowIndex).CellCount
    m_oFetchBuffer.frSetRow Nothing, lRowIndex, jgexRowTypeRecord, 0
    m_oFetchBuffer.frRowBookmark = m_aRecord(lRowIndex).Bookmark
    '--- the control raises it, since that is where the typelib declares it.
    '--- Probed: the original hands the same bookmark in both as the event
    '--- parameter and as Values.Bookmark, and Values.RowIndex reads the
    '--- record's own index
    m_pOwner.frRaiseUnboundReadData lRowIndex, m_aRecord(lRowIndex).Bookmark, m_oFetchBuffer
    pvHarvest lRowIndex, m_oFetchBuffer
End Sub

Private Sub pvHarvest(ByVal lRowIndex As Long, oData As JSRowData)
    Dim nIdx            As Integer

    '--- what the client wrote is copied into the record; only Value and
    '--- IconIndex are kept, since DisplayValue, CellStyle and CellPicture
    '--- belong to RowFormat and are rewritten on every population
    With m_aRecord(lRowIndex)
        For nIdx = 1 To .CellCount
            AssignVariant .Values(nIdx), oData.Value(nIdx)
            .IconIndexes(nIdx) = oData.IconIndex(nIdx)
        Next
    End With
End Sub

Private Function pvCellValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Variant
    If lRowIndex < 1 Or lRowIndex > m_lItemCount Then
        Exit Function
    End If
    '--- FetchData outranks whatever UnboundReadData wrote for the column,
    '--- the same way it outranks DataField in a bound grid
    If nColIndex >= 1 And nColIndex <= m_uFetch.ColCount Then
        If m_uFetch.IsFetchData(nColIndex) Then
            AssignVariant pvCellValue, pvFetchValue(lRowIndex, nColIndex)
            Exit Function
        End If
    End If
    With m_aRecord(lRowIndex)
        If nColIndex >= 1 And nColIndex <= .CellCount Then
            AssignVariant pvCellValue, .Values(nColIndex)
        End If
    End With
End Function

Private Function pvFetchValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Variant
    '--- one cell, on demand. A sort is the only thing that ever asks for all
    '--- of them, and what it leaves behind is a warm cache rather than a
    '--- separate whole-column pass
    DataEnsureFetchRoom m_uFetch, m_lItemCount
    If lRowIndex < 1 Or lRowIndex > m_uFetch.RowCount Then
        Exit Function
    End If
    If Not m_uFetch.Valid(nColIndex, lRowIndex) Then
        '--- marked valid first so a read from inside the client's handler
        '--- does not re-enter the same cell
        m_uFetch.Valid(nColIndex, lRowIndex) = True
        AssignVariant m_uFetch.Value(nColIndex, lRowIndex), _
            m_pOwner.frFireFetchData(lRowIndex, nColIndex, m_aRecord(lRowIndex).Bookmark)
    End If
    AssignVariant pvFetchValue, m_uFetch.Value(nColIndex, lRowIndex)
End Function

Private Sub pvRefetchRow(ByVal lRowIndex As Long)
    Dim nIdx            As Integer
    Dim vOld            As Variant

    '--- a record that changed is asked for again column by column, rather
    '--- than the cache being thrown away and the whole grid refetched
    DataEnsureFetchRoom m_uFetch, m_lItemCount
    If lRowIndex < 1 Or lRowIndex > m_uFetch.RowCount Then
        Exit Sub
    End If
    For nIdx = 1 To m_uFetch.ColCount
        If m_uFetch.IsFetchData(nIdx) Then
            AssignVariant vOld, m_uFetch.Value(nIdx, lRowIndex)
            m_uFetch.Valid(nIdx, lRowIndex) = True
            AssignVariant m_uFetch.Value(nIdx, lRowIndex), _
                m_pOwner.frFireFetchData(lRowIndex, nIdx, m_aRecord(lRowIndex).Bookmark)
            '--- a fetch column can be a sort or a group key, but only a value
            '--- that actually moved can move the record
            If DataCompareValues(vOld, m_uFetch.Value(nIdx, lRowIndex), jgexSortTypeString) <> 0 Then
                m_bOrderDirty = True
            End If
        End If
    Next
End Sub

'--- ordering

Private Sub pvEnsureOrder()
    Dim lRow            As Long
    Dim lIdx            As Long
    Dim nCol            As Integer
    Dim vValue          As Variant

    If Not m_bOrderDirty Then
        Exit Sub
    End If
    m_bOrderDirty = False
    m_uRowSet.ItemCount = m_lItemCount
    m_uRowSet.GroupFooterStyle = m_pOwner.GroupFooterStyle
    m_uRowSet.DefaultCollapsed = (m_pOwner.DefaultGroupMode = jgexDGMCollapsed)
    DataReadSortKeys m_uRowSet, m_pOwner
    '--- filling the keys is the only step of the projection needing a cell
    '--- value, which is why it stays here rather than moving into
    '--- mdDataModel with the rest
    If m_uRowSet.SortKeyCount > 0 Then
        For lRow = 1 To m_lItemCount
            '--- sorting needs every record's key, so the lazy fetch is forced
            pvFetchRecord lRow
            For lIdx = 1 To m_uRowSet.SortKeyCount
                nCol = m_uRowSet.SortCol(lIdx)
                If nCol > 0 Then
                    vValue = pvCellValue(lRow, nCol)
                    '--- an object in a cell has no ordering, so it sorts blank
                    If Not IsObject(vValue) Then
                        m_uRowSet.Key(lIdx, lRow) = vValue
                    End If
                End If
            Next
        Next
    End If
    DataProject m_uRowSet
End Sub

Private Function pvGroupCaption(ByVal lGroupSlot As Long) As String
    Dim oCol            As JSColumn
    Dim vValue          As Variant
    Dim sText           As String
    Dim nColIndex       As Integer

    '--- formatted on demand from the record that opened the group, so the
    '--- caption follows GroupPrefix, GroupFormat and GroupEmptyStringCaption
    '--- without the projection having to be rebuilt when one of them changes.
    '--- GroupFormat re-formats the caption only -- the original still breaks
    '--- groups on the raw value
    nColIndex = m_uRowSet.GroupRow(lGroupSlot).ColIndex
    If nColIndex < 1 Or nColIndex > m_pOwner.Columns.Count Then
        Exit Function
    End If
    Set oCol = m_pOwner.Columns.Item(nColIndex)
    vValue = pvCellValue(m_uRowSet.GroupRow(lGroupSlot).FirstRowIndex, nColIndex)
    If Not DataIsBlank(vValue) And Not IsObject(vValue) And Not IsArray(vValue) Then
        If LenB(oCol.GroupFormat) <> 0 Then
            sText = Format$(vValue, oCol.GroupFormat)
        Else
            sText = C2Str(vValue)
        End If
    End If
    If LenB(sText) = 0 Then
        sText = oCol.GroupEmptyStringCaption
    End If
    If LenB(oCol.GroupPrefix) <> 0 Then
        sText = oCol.GroupPrefix & " " & sText
    End If
    pvGroupCaption = sText
End Function

Private Function pvGroupValues(ByVal lGroupSlot As Long, ByVal nColIndex As Integer, aValues() As Variant) As Long
    Dim lIdx            As Long
    Dim lCount          As Long

    '--- the records a group covers, nested group rows and footers skipped
    With m_uRowSet.GroupRow(lGroupSlot)
        ReDim aValues(1 To .RecordCount + 1) As Variant
        For lIdx = .FirstSlot To .LastSlot
            If lIdx >= 1 And lIdx <= m_uRowSet.OrderCount Then
                If m_uRowSet.Order(lIdx) > 0 Then
                    lCount = lCount + 1
                    AssignVariant aValues(lCount), pvCellValue(m_uRowSet.Order(lIdx), nColIndex)
                End If
            End If
        Next
    End With
    pvGroupValues = lCount
End Function

'--- bookmarks

Private Sub pvEnsureBookmarks()
    Dim lIdx            As Long
    Dim sKey            As String

    If Not m_bBookmarksDirty Then
        Exit Sub
    End If
    m_bBookmarksDirty = False
    Set m_oBookmarks = New VBA.Collection
    '--- walked upwards so the lowest RowIndex wins a duplicate, which is
    '--- what the original resolves to: with "bk-1" on records 1 and 3, all
    '--- of MoveToBookmark, AddBookmark and RefreshRowBookmark answer 1
    For lIdx = 1 To m_lItemCount
        If Not DataIsBlank(m_aRecord(lIdx).Bookmark) Then
            sKey = DataBookmarkKey(m_aRecord(lIdx).Bookmark)
            If Not SearchCollection(m_oBookmarks, sKey) Then
                m_oBookmarks.Add lIdx, sKey
            End If
        End If
    Next
End Sub

'=========================================================================
' Interface IDataModel
'=========================================================================

Private Property Get IDataModel_RowCount() As Long
    pvEnsureOrder
    IDataModel_RowCount = m_uRowSet.VisibleCount
End Property

Private Property Get IDataModel_ItemCount() As Long
    IDataModel_ItemCount = m_lItemCount
End Property

Private Property Get IDataModel_Version() As Long
    IDataModel_Version = m_uRowSet.Version
End Property

Private Property Get IDataModel_RowIndex(ByVal lRowPosition As Long) As Long
    Dim lEntry          As Long

    pvEnsureOrder
    lEntry = DataEntryAt(m_uRowSet, lRowPosition)
    If lEntry > 0 Then
        IDataModel_RowIndex = lEntry
    End If
End Property

Private Property Get IDataModel_RowBookmark(ByVal lRowIndex As Long) As Variant
    If lRowIndex < 1 Or lRowIndex > m_lItemCount Then
        '--- the original sizes its bookmark store to ItemCount and raises
        '--- past the end rather than answering Empty
        Err.Raise vbObjectError, , "Subscript out of range"
    End If
    AssignVariant IDataModel_RowBookmark, m_aRecord(lRowIndex).Bookmark
End Property

Private Property Let IDataModel_RowBookmark(ByVal lRowIndex As Long, vValue As Variant)
    If lRowIndex < 1 Or lRowIndex > m_lItemCount Then
        Err.Raise vbObjectError, , "Subscript out of range"
    End If
    AssignVariant m_aRecord(lRowIndex).Bookmark, vValue
    '--- duplicates are legal, and which record answers depends on every
    '--- other bookmark, so the map is rebuilt rather than patched
    m_bBookmarksDirty = True
End Property

Private Property Get IDataModel_RowExpanded(ByVal lRowPosition As Long) As Boolean
    Dim lEntry          As Long

    pvEnsureOrder
    lEntry = DataEntryAt(m_uRowSet, lRowPosition)
    If lEntry < 0 Then
        IDataModel_RowExpanded = Not m_uRowSet.GroupRow(-lEntry).Collapsed
    End If
End Property

Private Property Let IDataModel_RowExpanded(ByVal lRowPosition As Long, ByVal bValue As Boolean)
    Dim lEntry          As Long

    pvEnsureOrder
    lEntry = DataEntryAt(m_uRowSet, lRowPosition)
    If lEntry < 0 Then
        '--- a footer closes a group rather than opening one, so it has no
        '--- expand box and nothing to toggle. Nested rather than And-ed,
        '--- since VB6 evaluates both operands and the subscript is only
        '--- valid on the negative branch
        If Not m_uRowSet.GroupRow(-lEntry).Footer Then
            If m_uRowSet.GroupRow(-lEntry).Collapsed = bValue Then
                '--- an expand or a collapse only reprojects: the sort order
                '--- underneath stays exactly as it was
                m_uRowSet.GroupRow(-lEntry).Collapsed = Not bValue
                DataBuildVisible m_uRowSet
                DataWritePositions m_uRowSet
            End If
        End If
    End If
End Property

Private Sub IDataModel_Terminate()
    frTerminate
End Sub

Private Sub IDataModel_RefreshGroups()
    m_bOrderDirty = True
End Sub

Private Sub IDataModel_RefreshSort()
    m_bOrderDirty = True
End Sub

Private Sub IDataModel_RefreshRowIndex(ByVal lRowIndex As Long)
    If lRowIndex >= 1 And lRowIndex <= m_lItemCount Then
        pvResetRecord lRowIndex, bDropBookmark:=False
        '--- only that record's fetch columns: the rest of the cache is
        '--- still good and refetching it would be gratuitous
        DataInvalidateFetchRow m_uFetch, lRowIndex
        m_bOrderDirty = True
    End If
End Sub

Private Sub IDataModel_Refresh()
    Dim lIdx            As Long

    m_lItemCount = m_pOwner.ItemCount
    pvEnsureCapacity m_lItemCount
    For lIdx = 1 To m_lItemCount
        pvResetRecord lIdx, bDropBookmark:=False
    Next
    DataInitFetch m_uFetch, m_pOwner
    m_bOrderDirty = True
    m_bBookmarksDirty = True
End Sub

Private Function IDataModel_GetRowData(ByVal lRowPosition As Long) As JSRowData
    Dim oData           As JSRowData
    Dim nColCount       As Integer
    Dim lEntry          As Long
    Dim nIdx            As Integer
    Dim eFunc           As jgexAggregateFunctionConstants
    Dim aValues()       As Variant
    Dim lCount          As Long

    pvEnsureOrder
    nColCount = m_pOwner.Columns.Count
    Set oData = New JSRowData
    oData.frInitBuffer m_pOwner
    oData.frReset nColCount
    lEntry = DataEntryAt(m_uRowSet, lRowPosition)
    If lEntry > 0 Then
        pvFetchRecord lEntry
        oData.frSetRow Me, lEntry, jgexRowTypeRecord, 0
        oData.frRowBookmark = m_aRecord(lEntry).Bookmark
        For nIdx = 1 To nColCount
            '--- through pvCellValue, so a FetchData column comes from the
            '--- client here exactly as it does in a bound grid
            oData.frSetValue nIdx, pvCellValue(lEntry, nIdx)
            If nIdx <= m_uFetch.ColCount And m_uFetch.IsFetchIcon(nIdx) Then
                oData.IconIndex(nIdx) = m_pOwner.frFireFetchIcon(lEntry, nIdx, m_aRecord(lEntry).Bookmark)
            ElseIf nIdx <= m_aRecord(lEntry).CellCount Then
                oData.IconIndex(nIdx) = m_aRecord(lEntry).IconIndexes(nIdx)
            End If
        Next
    ElseIf lEntry < 0 Then
        With m_uRowSet.GroupRow(-lEntry)
            oData.frSetRow Me, lEntry, _
                IIf(.Footer, jgexRowTypeGroupFooter, jgexRowTypeGroupHeader), .Level
        End With
        oData.GroupCaption = pvGroupCaption(-lEntry)
        '--- a totals footer carries its columns' aggregates as the cell
        '--- values, so painting one is painting an ordinary row
        If m_uRowSet.GroupRow(-lEntry).Footer Then
            If m_uRowSet.GroupFooterStyle = jgexTotalsGroupFooter Then
                For nIdx = 1 To nColCount
                    eFunc = m_pOwner.Columns.Item(nIdx).AggregateFunction
                    If eFunc <> jgexAggregateNone Then
                        lCount = pvGroupValues(-lEntry, nIdx, aValues)
                        oData.frSetValue nIdx, DataAggregate(aValues, lCount, eFunc, _
                            m_pOwner.Columns.Item(nIdx).SortType)
                    End If
                Next
            End If
        End If
    End If
    Set IDataModel_GetRowData = oData
End Function

Private Sub IDataModel_UpdateRowData(oData As JSRowData)
    Dim lRowIndex       As Long
    Dim oNewBookmark    As JSRetVariant

    lRowIndex = oData.RowIndex
    If lRowIndex < 0 Then
        '--- a record that does not exist yet. The client is asked for its
        '--- bookmark and nothing else -- probed, NewRowBookmark arrives Empty
        '--- and Values.RowIndex reads ItemCount + 1
        Set oNewBookmark = New JSRetVariant
        m_pOwner.frRaiseUnboundAddNew oNewBookmark, oData
        '--- grown only now the handler has returned: doing it inside the
        '--- event re-enters the add and the original never leaves the new row
        pvEnsureCapacity m_lItemCount + 1
        m_lItemCount = m_lItemCount + 1
        AssignVariant m_aRecord(m_lItemCount).Bookmark, oNewBookmark.Value
        pvEnsureCells m_lItemCount
        pvHarvest m_lItemCount, oData
        m_aRecord(m_lItemCount).Fetched = True
        '--- the caller reads both off the row it just added. The position is
        '--- left at 0: it is not known until the projection is rebuilt, and
        '--- asking for it here would force a re-sort mid-insert. This buffer
        '--- is the grid's, not ours, so it keeps its model reference
        oData.frSetRow Me, m_lItemCount, jgexRowTypeRecord, 0
        oData.frRowBookmark = m_aRecord(m_lItemCount).Bookmark
        pvRefetchRow m_lItemCount
        m_bBookmarksDirty = True
        m_bOrderDirty = True
    ElseIf lRowIndex >= 1 And lRowIndex <= m_lItemCount Then
        pvEnsureCells lRowIndex
        pvHarvest lRowIndex, oData
        m_pOwner.frRaiseUnboundUpdate lRowIndex, m_aRecord(lRowIndex).Bookmark, oData
        '--- the client may compute a fetch column off what just changed, so
        '--- it is asked again -- and that alone can invalidate the order
        pvRefetchRow lRowIndex
        '--- a value that is a sort or group key moves the record, so the
        '--- projection is stale. Whether the original re-sorts here or waits
        '--- for RefreshSort is not probed yet
        m_bOrderDirty = True
    End If
End Sub

Private Sub IDataModel_Delete(ByVal lRowPosition As Long)
    Dim lRowIndex       As Long
    Dim lIdx            As Long

    pvEnsureOrder
    lRowIndex = IDataModel_RowIndex(lRowPosition)
    If lRowIndex = 0 Then
        Exit Sub
    End If
    '--- the client hears which record went, and by which bookmark, before
    '--- the renumbering below makes that index mean something else
    m_pOwner.frRaiseUnboundDelete lRowIndex, m_aRecord(lRowIndex).Bookmark
    '--- records are indexed by RowIndex itself, so a delete closes the hole
    '--- and everything above it renumbers -- which is exactly what the
    '--- original does, and what invalidates the bookmark map wholesale
    For lIdx = lRowIndex To m_lItemCount - 1
        m_aRecord(lIdx) = m_aRecord(lIdx + 1)
    Next
    pvResetRecord m_lItemCount, bDropBookmark:=True
    '--- the cache follows the renumbering rather than being thrown away
    DataDeleteFetchRow m_uFetch, lRowIndex, m_lItemCount
    m_lItemCount = m_lItemCount - 1
    m_bBookmarksDirty = True
    m_bOrderDirty = True
End Sub

Private Function IDataModel_GetRowPosition(ByVal lRowIndex As Long) As Long
    pvEnsureOrder
    If lRowIndex >= 1 And lRowIndex <= m_lItemCount Then
        IDataModel_GetRowPosition = m_uRowSet.RowPosition(lRowIndex)
    End If
End Function

Private Function IDataModel_GetRowIndex(vBookmark As Variant) As Long
    Dim vRetVal         As Variant

    If DataIsBlank(vBookmark) Then
        Exit Function
    End If
    pvEnsureBookmarks
    If SearchCollection(m_oBookmarks, DataBookmarkKey(vBookmark), RetVal:=vRetVal) Then
        IDataModel_GetRowIndex = C2Lng(vRetVal)
    End If
End Function

Private Function IDataModel_GetRecordCount(ByVal lRowIndex As Long) As Long
    pvEnsureOrder
    If lRowIndex < 0 Then
        IDataModel_GetRecordCount = m_uRowSet.GroupRow(-lRowIndex).RecordCount
    End If
End Function

Private Function IDataModel_GetSubTotal(ByVal lRowIndex As Long, ByVal nColIndex As Integer, ByVal eFunc As jgexAggregateFunctionConstants) As Variant
    Dim aValues()       As Variant
    Dim lCount          As Long
    Dim eSortType       As jgexSortTypeConstants

    pvEnsureOrder
    If lRowIndex < 0 Then
        eSortType = jgexSortTypeString
        If nColIndex >= 1 And nColIndex <= m_pOwner.Columns.Count Then
            eSortType = m_pOwner.Columns.Item(nColIndex).SortType
        End If
        lCount = pvGroupValues(-lRowIndex, nColIndex, aValues)
        IDataModel_GetSubTotal = DataAggregate(aValues, lCount, eFunc, eSortType)
    End If
End Function

Private Function IDataModel_GetBookmarks(ByVal lRowIndex As Long) As Variant
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim aRetVal()       As Variant

    pvEnsureOrder
    If lRowIndex >= 0 Then
        Exit Function
    End If
    With m_uRowSet.GroupRow(-lRowIndex)
        ReDim aRetVal(0 To .RecordCount - 1) As Variant
        For lIdx = .FirstSlot To .LastSlot
            If m_uRowSet.Order(lIdx) > 0 Then
                AssignVariant aRetVal(lCount), m_aRecord(m_uRowSet.Order(lIdx)).Bookmark
                lCount = lCount + 1
            End If
        Next
    End With
    IDataModel_GetBookmarks = aRetVal
End Function

Private Function IDataModel_GetRowIndexes(ByVal lRowIndex As Long) As Variant
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim aRetVal()       As Long

    pvEnsureOrder
    If lRowIndex >= 0 Then
        Exit Function
    End If
    With m_uRowSet.GroupRow(-lRowIndex)
        ReDim aRetVal(0 To .RecordCount - 1) As Long
        For lIdx = .FirstSlot To .LastSlot
            If m_uRowSet.Order(lIdx) > 0 Then
                aRetVal(lCount) = m_uRowSet.Order(lIdx)
                lCount = lCount + 1
            End If
        Next
    End With
    IDataModel_GetRowIndexes = aRetVal
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    '--- zero the raw pointer or VB6 releases a control it never AddRef-ed
    frTerminate
End Sub
