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
Private Const MODULE_NAME As String = "cUnboundDataModel"

Implements IDataModel

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const GROW_BY                   As Long = 64

Private m_pOwner                    As GridEX
Private m_aRecord()                 As UcsRecord
Private m_lRecordUBound             As Long
Private m_uState                    As UcsDataState
Private m_uRowSet                   As UcsRowSet
Private m_uFetch                    As UcsFetchCache
Private m_oFetchBuffer              As JSRowData

Private Type UcsRecord
    Fetched                 As Boolean
    CellCount               As Long
    Values()                As Variant
    IconIndexes()           As Long
End Type

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
    Set m_oFetchBuffer = New JSRowData
    m_oFetchBuffer.frAllowUpdate = True
    m_uState.OrderDirty = True
    m_uState.BookmarksDirty = True
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Friend Sub frTerminate()
    AssignWeakRef m_pOwner, Nothing
    Set m_uState.Bookmarks = Nothing
    Set m_oFetchBuffer = Nothing
    Erase m_aRecord
    m_uState.ItemCount = 0
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
    DataEnsureBookmarkRoom m_uState, lCount
    For lIdx = m_uState.ItemCount + 1 To lCount
        pvResetRecord lIdx, bDropBookmark:=True
    Next
End Sub

Private Sub pvResetRecord(ByVal lRowIndex As Long, ByVal bDropBookmark As Boolean)
    With m_aRecord(lRowIndex)
        .Fetched = False
        If .CellCount > 0 Then
            ReDim .Values(1 To .CellCount) As Variant
            ReDim .IconIndexes(1 To .CellCount) As Long
        End If
    End With
    '--- probed at both: the original keeps bookmarks across Refetch AND across
    '--- Rebind, so nothing but a Delete or an explicit assignment drops one
    If bDropBookmark Then
        m_uState.Bookmark(lRowIndex) = Empty
    End If
End Sub

Private Sub pvEnsureCells(ByVal lRowIndex As Long)
    Dim lColCount       As Long

    lColCount = m_pOwner.Columns.Count
    With m_aRecord(lRowIndex)
        If .CellCount < lColCount Then
            ReDim Preserve .Values(1 To lColCount) As Variant
            ReDim Preserve .IconIndexes(1 To lColCount) As Long
            .CellCount = lColCount
        End If
    End With
End Sub

Private Sub pvFetchRecord(ByVal lRowIndex As Long)
    If lRowIndex < 1 Or lRowIndex > m_uState.ItemCount Then
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
    m_oFetchBuffer.frInitModel Nothing, lRowIndex, jgexRowTypeRecord, 0
    m_oFetchBuffer.frRowBookmark = m_uState.Bookmark(lRowIndex)
    '--- the control raises it, since that is where the typelib declares it.
    '--- Probed: the original hands the same bookmark in both as the event
    '--- parameter and as Values.Bookmark, and Values.RowIndex reads the
    '--- record's own index
    m_pOwner.frRaiseUnboundReadData lRowIndex, m_uState.Bookmark(lRowIndex), m_oFetchBuffer
    pvHarvest lRowIndex, m_oFetchBuffer
End Sub

Private Sub pvHarvest(ByVal lRowIndex As Long, oData As JSRowData)
    Dim lIdx            As Long

    '--- what the client wrote is copied into the record; only Value and
    '--- IconIndex are kept, since DisplayValue, CellStyle and CellPicture
    '--- belong to RowFormat and are rewritten on every population
    With m_aRecord(lRowIndex)
        For lIdx = 1 To .CellCount
            AssignVariant .Values(lIdx), oData.Value(lIdx)
            .IconIndexes(lIdx) = oData.IconIndex(lIdx)
        Next
    End With
End Sub

Private Function pvCellValue(ByVal lRowIndex As Long, ByVal lColIndex As Long) As Variant
    If lRowIndex < 1 Or lRowIndex > m_uState.ItemCount Then
        Exit Function
    End If
    '--- FetchData outranks whatever UnboundReadData wrote for the column,
    '--- the same way it outranks DataField in a bound grid
    If lColIndex >= 1 And lColIndex <= m_uFetch.ColCount Then
        If m_uFetch.IsFetchData(lColIndex) Then
            AssignVariant pvCellValue, DataFetchValue(m_uState, m_uFetch, m_pOwner, lRowIndex, lColIndex)
            Exit Function
        End If
    End If
    With m_aRecord(lRowIndex)
        If lColIndex >= 1 And lColIndex <= .CellCount Then
            AssignVariant pvCellValue, .Values(lColIndex)
        End If
    End With
End Function

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
    '--- filling the keys is the only step of the projection needing a cell
    '--- value, which is why it stays here rather than moving into
    '--- mdDataModel with the rest
    If m_uRowSet.SortKeyCount > 0 Then
        For lRow = 1 To m_uState.ItemCount
            '--- sorting needs every record's key, so the lazy fetch is forced
            pvFetchRecord lRow
            For lIdx = 1 To m_uRowSet.SortKeyCount
                lCol = m_uRowSet.SortCol(lIdx)
                If lCol > 0 Then
                    vValue = pvCellValue(lRow, lCol)
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

Private Function pvGroupCaption(ByVal lGroupIndex As Long, bPrefixed As Boolean) As String
    Dim oCol            As JSColumn
    Dim vValue          As Variant
    Dim sText           As String
    Dim lColIndex       As Long

    '--- formatted on demand from the record that opened the group, so the
    '--- caption follows GroupPrefix, GroupFormat and GroupEmptyStringCaption
    '--- without the projection having to be rebuilt when one of them changes.
    '--- GroupFormat re-formats the caption only -- the original still breaks
    '--- groups on the raw value
    lColIndex = m_uRowSet.GroupRow(lGroupIndex).ColIndex
    If lColIndex < 1 Or lColIndex > m_pOwner.Columns.Count Then
        Exit Function
    End If
    Set oCol = m_pOwner.Columns.Item(lColIndex)
    vValue = pvCellValue(m_uRowSet.GroupRow(lGroupIndex).FirstRowIndex, lColIndex)
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
    bPrefixed = (LenB(oCol.GroupPrefix) <> 0)
    If bPrefixed Then
        sText = oCol.GroupPrefix & " " & sText
    End If
    pvGroupCaption = sText
End Function

Private Function pvGroupValues(ByVal lGroupIndex As Long, ByVal lColIndex As Long, aValues() As Variant) As Long
    Dim lIdx            As Long
    Dim lCount          As Long

    '--- the records a group covers, nested group rows and footers skipped
    With m_uRowSet.GroupRow(lGroupIndex)
        ReDim aValues(1 To .RecordCount + 1) As Variant
        For lIdx = .FirstSlot To .LastSlot
            If lIdx >= 1 And lIdx <= m_uRowSet.OrderCount Then
                If m_uRowSet.Order(lIdx) > 0 Then
                    lCount = lCount + 1
                    AssignVariant aValues(lCount), pvCellValue(m_uRowSet.Order(lIdx), lColIndex)
                End If
            End If
        Next
    End With
    pvGroupValues = lCount
End Function

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
    IDataModel_RowIndex = DataRecordIndexAt(m_uRowSet, lRowPosition)
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get IDataModel_RowBookmark(ByVal lRowIndex As Long) As Variant
    Const FUNC_NAME     As String = "IDataModel_RowBookmark"

    On Error GoTo EH
    AssignVariant IDataModel_RowBookmark, DataRowBookmark(m_uState, lRowIndex)
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
    IDataModel_RowExpanded = DataRowExpanded(m_uRowSet, lRowPosition)
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
    If lRowIndex >= 1 And lRowIndex <= m_uState.ItemCount Then
        pvResetRecord lRowIndex, bDropBookmark:=False
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
    Dim lIdx            As Long

    On Error GoTo EH
    '--- capacity first, count after: pvEnsureCapacity resets the slots
    '--- between the old count and the new one, dropping the bookmarks a
    '--- shrunk-away record left behind -- with the count already written
    '--- that loop has nothing to walk and a later regrow would resurrect them
    pvEnsureCapacity m_pOwner.ItemCount
    m_uState.ItemCount = m_pOwner.ItemCount
    For lIdx = 1 To m_uState.ItemCount
        pvResetRecord lIdx, bDropBookmark:=False
    Next
    DataInitFetch m_uFetch, m_pOwner
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
    lRowIndex = DataRecordIndexAt(m_uRowSet, lRowPosition)
    If lRowIndex > 0 Then
        pvFetchRecord lRowIndex
        oData.frInitModel Me, lRowIndex, jgexRowTypeRecord, 0
        oData.frRowBookmark = m_uState.Bookmark(lRowIndex)
        For lIdx = 1 To lColCount
            oData.frSetValue lIdx, pvCellValue(lRowIndex, lIdx)
            If lIdx <= m_uFetch.ColCount Then
                If m_uFetch.IsFetchIcon(lIdx) Then
                    oData.IconIndex(lIdx) = m_pOwner.frRaiseFetchIcon(lRowIndex, lIdx, m_uState.Bookmark(lRowIndex))
                ElseIf lIdx <= m_aRecord(lRowIndex).CellCount Then
                    oData.IconIndex(lIdx) = m_aRecord(lRowIndex).IconIndexes(lIdx)
                End If
            ElseIf lIdx <= m_aRecord(lRowIndex).CellCount Then
                oData.IconIndex(lIdx) = m_aRecord(lRowIndex).IconIndexes(lIdx)
            End If
        Next
    ElseIf lRowIndex < 0 Then
        With m_uRowSet.GroupRow(-lRowIndex)
            oData.frInitModel Me, lRowIndex, _
                IIf(.Footer, jgexRowTypeGroupFooter, jgexRowTypeGroupHeader), .Level
        End With
        oData.GroupCaption = pvGroupCaption(-lRowIndex, bPrefixed)
        oData.frGroupPrefixed = bPrefixed
        If m_uRowSet.GroupRow(-lRowIndex).Footer Then
            If m_uRowSet.GroupFooterStyle = jgexTotalsGroupFooter Then
                For Each oCol In m_pOwner.Columns
                    eFunc = oCol.AggregateFunction
                    If eFunc <> jgexAggregateNone Then
                        lCount = pvGroupValues(-lRowIndex, oCol.Index, aValues)
                        oData.frSetValue oCol.Index, DataAggregate(aValues, lCount, eFunc, oCol.SortType)
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
    Dim oNewBookmark    As JSRetVariant

    On Error GoTo EH
    lRowIndex = oData.RowIndex
    If lRowIndex < 0 Or lRowIndex > m_uState.ItemCount Then
        Set oNewBookmark = New JSRetVariant
        m_pOwner.frRaiseUnboundAddNew oNewBookmark, oData
        pvEnsureCapacity m_uState.ItemCount + 1
        m_uState.ItemCount = m_uState.ItemCount + 1
        AssignVariant m_uState.Bookmark(m_uState.ItemCount), oNewBookmark.Value
        pvEnsureCells m_uState.ItemCount
        pvHarvest m_uState.ItemCount, oData
        m_aRecord(m_uState.ItemCount).Fetched = True
        oData.frInitModel Me, m_uState.ItemCount, jgexRowTypeRecord, 0
        oData.frRowBookmark = m_uState.Bookmark(m_uState.ItemCount)
        DataRefetchRow m_uState, m_uFetch, m_pOwner, m_uState.ItemCount
        m_uState.BookmarksDirty = True
        m_uState.OrderDirty = True
    ElseIf lRowIndex >= 1 And lRowIndex <= m_uState.ItemCount Then
        pvEnsureCells lRowIndex
        pvHarvest lRowIndex, oData
        m_pOwner.frRaiseUnboundUpdate lRowIndex, m_uState.Bookmark(lRowIndex), oData
        DataRefetchRow m_uState, m_uFetch, m_pOwner, lRowIndex
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
    '--- probed with Del in row mode: the original deletes the unbound record
    '--- internally and raises no UnboundDelete for it -- the veto pair on the
    '--- control's surface is all the client hears
    '--- records are indexed by RowIndex itself, so a delete closes the hole
    '--- and everything above it renumbers -- which is exactly what the
    '--- original does, and what invalidates the bookmark map wholesale
    For lIdx = lRowIndex To m_uState.ItemCount - 1
        m_aRecord(lIdx) = m_aRecord(lIdx + 1)
    Next
    pvResetRecord m_uState.ItemCount, bDropBookmark:=True
    '--- the cache follows the renumbering rather than being thrown away
    DataDeleteFetchRow m_uFetch, lRowIndex, m_uState.ItemCount
    m_uState.ItemCount = m_uState.ItemCount - 1
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
        lCount = pvGroupValues(-lRowIndex, lColIndex, aValues)
        IDataModel_GetSubTotal = DataAggregate(aValues, lCount, eFunc, eSortType)
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
