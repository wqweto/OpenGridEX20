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

Implements IDataModel

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const GROW_BY               As Long = 64
'--- ADO's own, so the control never references msado15
Private Const adBookmark            As Long = &H2000

Private Type UcsRecord
    Bookmark                As Variant
End Type

Private m_pOwner                    As GridEX
Private m_oRs                       As Object
Private m_aRecord()                 As UcsRecord
Private m_lItemCount                As Long
Private m_lRecordUBound             As Long
Private m_lCurRowIndex              As Long
Private m_aField()                  As Object
Private m_nFieldCount               As Integer
Private m_uRowSet                   As UcsRowSet
Private m_uFetch                    As UcsFetchCache
Private m_oBookmarks                As VBA.Collection
Private m_bBookmarksDirty           As Boolean
Private m_bOrderDirty               As Boolean

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(ByVal oCtl As GridEX)
    Call CopyMemory(m_pOwner, oCtl, PTR_SIZE)
    Set m_oBookmarks = New VBA.Collection
    m_bOrderDirty = True
    m_bBookmarksDirty = True
End Sub

Friend Sub frTerminate()
    '--- zero the raw pointer or VB6 releases a control it never AddRef-ed
    Call CopyMemory(m_pOwner, NULL_PTR, PTR_SIZE)
    Set m_oRs = Nothing
    Set m_oBookmarks = Nothing
    Erase m_aRecord
    Erase m_aField
    m_lItemCount = 0
    m_lRecordUBound = 0
    m_nFieldCount = 0
    m_lCurRowIndex = 0
End Sub

'=========================================================================
' Functions
'=========================================================================

'--- storage: this half is what makes the model an ADO one

Private Sub pvBind()
    Dim lIdx            As Long

    Set m_oRs = m_pOwner.ADORecordset
    m_lItemCount = 0
    m_lCurRowIndex = 0
    Erase m_aRecord
    m_lRecordUBound = 0
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
        pvEnsureCapacity lIdx
        AssignVariant m_aRecord(lIdx).Bookmark, m_oRs.Bookmark
        m_oRs.MoveNext
    Loop
    m_lItemCount = lIdx
End Sub

Private Sub pvBindFields()
    Dim nIdx            As Integer
    Dim oCol            As JSColumn

    DataInitFetch m_uFetch, m_pOwner
    '--- DataField is the whole of the binding: a column that names no field
    '--- is not bound to the recordset at all and stays Nothing here, which
    '--- reads as Empty. There is deliberately no fall back to the field in
    '--- the same ordinal position -- a calculated or display-only column
    '--- would then silently show whatever field happened to sit there.
    '--- FetchData wins outright, so such a column is not resolved at all
    m_nFieldCount = m_pOwner.Columns.Count
    Erase m_aField
    If m_nFieldCount = 0 Then
        Exit Sub
    End If
    ReDim m_aField(1 To m_nFieldCount) As Object
    For nIdx = 1 To m_nFieldCount
        Set oCol = m_pOwner.Columns.Item(nIdx)
        If Not m_uFetch.IsFetchData(nIdx) Then
            If LenB(oCol.DataField) <> 0 Then
                Set m_aField(nIdx) = pvFindField(oCol.DataField)
            End If
        End If
    Next
End Sub

Private Function pvFindField(sName As String) As Object
    Dim lIdx            As Long

    '--- walked rather than looked up: the Fields collection raises on a name
    '--- it does not hold, and the control has to stay usable under Break on
    '--- All Errors
    For lIdx = 0 To m_oRs.Fields.Count - 1
        If StrComp(m_oRs.Fields(lIdx).Name, sName, vbTextCompare) = 0 Then
            Set pvFindField = m_oRs.Fields(lIdx)
            Exit Function
        End If
    Next
End Function

Private Sub pvSeek(ByVal lRowIndex As Long)
    If lRowIndex = m_lCurRowIndex Then
        Exit Sub
    End If
    If lRowIndex < 1 Or lRowIndex > m_lItemCount Then
        Exit Sub
    End If
    m_oRs.Bookmark = m_aRecord(lRowIndex).Bookmark
    m_lCurRowIndex = lRowIndex
End Sub

Private Function pvCellValue(ByVal lRowIndex As Long, ByVal nColIndex As Integer) As Variant
    If nColIndex < 1 Or nColIndex > m_nFieldCount Then
        Exit Function
    End If
    '--- FetchData outranks any DataField the column carries, so it is asked
    '--- first and the recordset is never consulted for such a column
    If m_uFetch.IsFetchData(nColIndex) Then
        AssignVariant pvCellValue, pvFetchValue(lRowIndex, nColIndex)
        Exit Function
    End If
    If m_aField(nColIndex) Is Nothing Then
        Exit Function
    End If
    pvSeek lRowIndex
    If m_lCurRowIndex <> lRowIndex Then
        Exit Function
    End If
    '--- read straight off the cached Field: nothing is copied into the
    '--- record, the recordset is the cache, and the Field already points at
    '--- whichever record the seek above left the cursor on
    AssignVariant pvCellValue, m_aField(nColIndex).Value
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

Private Sub pvWriteFields(oData As JSRowData)
    Dim nIdx            As Integer

    '--- every bound column is written, not only the ones the editor touched:
    '--- JSRowData carries no per-cell dirty flag yet, and adding one is what
    '--- would stop this fighting optimistic locking on a wide table. A
    '--- FetchData or unbound column has nowhere to go and is dropped
    For nIdx = 1 To m_nFieldCount
        If Not m_aField(nIdx) Is Nothing Then
            m_aField(nIdx).Value = oData.Value(nIdx)
        End If
    Next
End Sub

Private Sub pvEnsureCapacity(ByVal lCount As Long)
    If lCount > m_lRecordUBound Then
        m_lRecordUBound = lCount + GROW_BY
        ReDim Preserve m_aRecord(1 To m_lRecordUBound) As UcsRecord
    End If
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
    If m_uRowSet.SortKeyCount > 0 And m_lItemCount > 0 Then
        If Not (m_oRs.BOF And m_oRs.EOF) Then
            m_oRs.MoveFirst
        End If
        '--- walked with MoveNext rather than by seeking each bookmark:
        '--- RowIndex order IS the recordset's own order, so this is one
        '--- sequential pass over the cursor instead of ItemCount random
        '--- repositionings. Filling the keys is the only step of the
        '--- projection needing a cell value, which is why it stays here
        Do While Not m_oRs.EOF And lRow < m_lItemCount
            lRow = lRow + 1
            m_lCurRowIndex = lRow
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
            m_oRs.MoveNext
        Loop
        m_lCurRowIndex = 0
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
    '--- without the projection having to be rebuilt when one of them changes
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

    '--- the records a group covers, nested group rows and footers skipped.
    '--- Every value here costs a seek, which is the price of not caching --
    '--- a SQL model answers the same question with one aggregate query
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
    '--- what the original resolves to
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
        Err.Raise vbObjectError, , "Subscript out of range"
    End If
    AssignVariant IDataModel_RowBookmark, m_aRecord(lRowIndex).Bookmark
End Property

Private Property Let IDataModel_RowBookmark(ByVal lRowIndex As Long, vValue As Variant)
    '--- an ADO bookmark is the cursor's to mint, not the client's: unlike
    '--- unbound mode there is nothing here that assigning one could mean.
    '--- Whether the original raises or swallows it is not probed yet
    Err.Raise vbObjectError, , _
        "RowBookmark is read-only when the control is bound to a recordset"
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

Private Sub IDataModel_RefreshGroups()
    m_bOrderDirty = True
End Sub

Private Sub IDataModel_RefreshSort()
    m_bOrderDirty = True
End Sub

Private Sub IDataModel_RefreshRowIndex(ByVal lRowIndex As Long)
    '--- the recordset is the cache, so there is nothing of ours to drop --
    '--- the record just has to be re-read from the source
    If lRowIndex >= 1 And lRowIndex <= m_lItemCount Then
        pvSeek lRowIndex
        On Error Resume Next
        m_oRs.Resync 1
        On Error GoTo 0
        m_lCurRowIndex = 0
        '--- only that record's fetch columns: the rest of the cache is
        '--- still good and refetching it would be gratuitous
        DataInvalidateFetchRow m_uFetch, lRowIndex
        m_bOrderDirty = True
    End If
End Sub

Private Sub IDataModel_Refresh()
    '--- a full refresh rebinds: the recordset may be a different object, and
    '--- its order is what defines RowIndex
    pvBind
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
        oData.frSetRow Me, lEntry, jgexRowTypeRecord, 0
        oData.frRowBookmark = m_aRecord(lEntry).Bookmark
        '--- one seek for the whole row: pvCellValue leaves the cursor where
        '--- it put it, so the remaining columns cost a property call each
        For nIdx = 1 To nColCount
            oData.frSetValue nIdx, pvCellValue(lEntry, nIdx)
            If nIdx <= m_uFetch.ColCount Then
                If m_uFetch.IsFetchIcon(nIdx) Then
                    oData.IconIndex(nIdx) = m_pOwner.frFireFetchIcon(lEntry, nIdx, m_aRecord(lEntry).Bookmark)
                End If
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

    lRowIndex = oData.RowIndex
    If lRowIndex < 0 Then
        '--- a record that does not exist yet. There is no UnboundAddNew
        '--- here: the cursor mints the bookmark on Update and that is the
        '--- new record's identity
        m_oRs.AddNew
        pvWriteFields oData
        m_oRs.Update
        pvEnsureCapacity m_lItemCount + 1
        m_lItemCount = m_lItemCount + 1
        AssignVariant m_aRecord(m_lItemCount).Bookmark, m_oRs.Bookmark
        m_lCurRowIndex = m_lItemCount
        '--- the caller reads both off the row it just added. The position is
        '--- left at 0: it is not known until the projection is rebuilt, and
        '--- asking for it here would force a re-sort mid-insert
        oData.frSetRow Me, m_lItemCount, jgexRowTypeRecord, 0
        oData.frRowBookmark = m_aRecord(m_lItemCount).Bookmark
        '--- the new record's fetch columns are asked for, and nothing else
        '--- in the cache is disturbed
        pvRefetchRow m_lItemCount
        m_bBookmarksDirty = True
        m_bOrderDirty = True
    ElseIf lRowIndex >= 1 And lRowIndex <= m_lItemCount Then
        pvSeek lRowIndex
        pvWriteFields oData
        m_oRs.Update
        '--- the client may compute a fetch column off what just changed, so
        '--- it is asked again -- and that alone can invalidate the order
        pvRefetchRow lRowIndex
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
    pvSeek lRowIndex
    m_oRs.Delete
    '--- records are indexed by RowIndex itself, so a delete closes the hole
    '--- and everything above it renumbers, as the original does
    For lIdx = lRowIndex To m_lItemCount - 1
        m_aRecord(lIdx) = m_aRecord(lIdx + 1)
    Next
    m_aRecord(m_lItemCount).Bookmark = Empty
    '--- the cache follows the renumbering rather than being thrown away
    DataDeleteFetchRow m_uFetch, lRowIndex, m_lItemCount
    m_lItemCount = m_lItemCount - 1
    m_lCurRowIndex = 0
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
