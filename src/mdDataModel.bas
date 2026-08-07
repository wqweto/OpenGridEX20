Attribute VB_Name = "mdDataModel"
'=========================================================================
'
' Open GridEX 2000 Control
' The projection engine every IDataModel implementation shares
'
' Sorting, group-row emission, the collapse projection and the position
' write-back are the same work whatever holds the records, so they live
' here over a UcsRowSet the implementation owns. What is left in each
' implementation is the half that is genuinely its own: where a cell value
' comes from, how a record is identified, and how a row is written back.
'
' The one thing the caller must do between reading the sort keys and
' projecting is fill uRowSet.Key -- that is the only step needing a cell
' value, and it is why DataProject cannot simply be handed the control
'
' A SQL model uses none of this: it answers RowCount, the projection and
' the aggregates from ORDER BY, GROUP BY and COUNT(*) instead
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdDataModel"

Public Const ERR_ROWDATA_INVALID    As String = "JSRowData object may have changed. The object is no longer valid."
Public Const ERR_ROWDATA_NUMBER     As Long = vbObjectError + 129
Public Const ERR_VALUE_READONLY     As String = "'Value' property can not be change in this context."
Public Const ERR_VALUE_NUMBER       As Long = vbObjectError

'=========================================================================
' Public types
'=========================================================================

Public Type UcsGroupRow
    Level                   As Long
    ColIndex                As Long
    FirstRowIndex           As Long
    RecordCount             As Long
    Collapsed               As Boolean
    Footer                  As Boolean
    FirstSlot               As Long
    LastSlot                As Long
    RowPosition             As Long
End Type

Public Type UcsRowSet
    Version                 As Long
    ItemCount               As Long
    GroupFooterStyle        As jgexGroupFooterStyleConstants
    DefaultCollapsed        As Boolean
    SortKeyCount            As Long
    GroupColCount           As Long
    SortCol()               As Long
    SortDir()               As Long
    SortType()              As Long
    Key()                   As Variant
    Order()                 As Long
    OrderCount              As Long
    GroupRow()              As UcsGroupRow
    GroupRowCount           As Long
    Visible()               As Long
    VisibleCount            As Long
    RowPosition()           As Long
End Type

Public Type UcsDataState
    ItemCount               As Long
    Bookmark()              As Variant
    BookmarkUBound          As Long
    Bookmarks               As VBA.Collection
    BookmarksDirty          As Boolean
    OrderDirty              As Boolean
End Type

Public Type UcsFetchCache
    ColCount                As Long
    RowCount                As Long
    IsFetchData()           As Boolean
    IsFetchIcon()           As Boolean
    Valid()                 As Boolean
    Value()                 As Variant
End Type

'=========================================================================
' Error management
'=========================================================================

Private Sub RaiseError(sFunction As String)
    PopRaiseError PushError, MODULE_NAME, sFunction
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Sub DataInitFetch(uFetch As UcsFetchCache, oCtl As GridEX)
    Const FUNC_NAME     As String = "DataInitFetch"
    Dim oCol            As JSColumn

    On Error GoTo EH
    '--- the whole cache goes: only a rebind or a change to the columns
    '--- themselves gets to be this expensive
    With uFetch
        .ColCount = oCtl.Columns.Count
        .RowCount = 0
        Erase .Value
        Erase .Valid
        Erase .IsFetchData
        Erase .IsFetchIcon
        If .ColCount = 0 Then
            Exit Sub
        End If
        ReDim .IsFetchData(1 To .ColCount) As Boolean
        ReDim .IsFetchIcon(1 To .ColCount) As Boolean
        '--- asked once here, not once per cell on every paint
        For Each oCol In oCtl.Columns
            .IsFetchData(oCol.Index) = oCol.FetchData
            .IsFetchIcon(oCol.Index) = oCol.FetchIcon
        Next
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'--- grows without losing what is already cached: the row is the last
'--- dimension precisely so ReDim Preserve can do this
Public Sub DataEnsureFetchRoom(uFetch As UcsFetchCache, ByVal lItemCount As Long)
    Const FUNC_NAME     As String = "DataEnsureFetchRoom"

    On Error GoTo EH
    With uFetch
        If .ColCount = 0 Or lItemCount <= .RowCount Then
            Exit Sub
        End If
        ReDim Preserve .Value(1 To .ColCount, 1 To lItemCount) As Variant
        ReDim Preserve .Valid(1 To .ColCount, 1 To lItemCount) As Boolean
        .RowCount = lItemCount
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'--- one record's fetch columns go stale -- an edit, or a Refresh of it
Public Sub DataInvalidateFetchRow(uFetch As UcsFetchCache, ByVal lRowIndex As Long)
    Const FUNC_NAME     As String = "DataInvalidateFetchRow"
    Dim lIdx            As Long

    On Error GoTo EH
    With uFetch
        If lRowIndex < 1 Or lRowIndex > .RowCount Then
            Exit Sub
        End If
        For lIdx = 1 To .ColCount
            .Valid(lIdx, lRowIndex) = False
        Next
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'--- a delete renumbers every record above it, so the cache follows rather
'--- than being thrown away and refetched
Public Sub DataDeleteFetchRow(uFetch As UcsFetchCache, ByVal lRowIndex As Long, ByVal lItemCount As Long)
    Const FUNC_NAME     As String = "DataDeleteFetchRow"
    Dim lIdx            As Long
    Dim lJdx            As Long

    On Error GoTo EH
    With uFetch
        If lRowIndex < 1 Or lRowIndex > .RowCount Then
            Exit Sub
        End If
        For lIdx = lRowIndex To lItemCount - 1
            If lIdx + 1 <= .RowCount Then
                For lJdx = 1 To .ColCount
                    AssignVariant .Value(lJdx, lIdx), .Value(lJdx, lIdx + 1)
                    .Valid(lJdx, lIdx) = .Valid(lJdx, lIdx + 1)
                Next
            End If
        Next
        If lItemCount >= 1 And lItemCount <= .RowCount Then
            For lJdx = 1 To .ColCount
                .Value(lJdx, lItemCount) = Empty
                .Valid(lJdx, lItemCount) = False
            Next
        End If
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub DataReadSortKeys(uRowSet As UcsRowSet, oCtl As GridEX)
    Const FUNC_NAME     As String = "DataReadSortKeys"
    Dim lIdx            As Long
    Dim oGroup          As JSGroup
    Dim oKey            As JSSortKey

    On Error GoTo EH
    '--- grouping sorts too: the group columns lead, in order, and the sort
    '--- keys refine within a group
    With uRowSet
        .GroupColCount = oCtl.Groups.Count
        .SortKeyCount = .GroupColCount + oCtl.SortKeys.Count
        If .SortKeyCount = 0 Then
            Exit Sub
        End If
        ReDim .SortCol(1 To .SortKeyCount) As Long
        ReDim .SortDir(1 To .SortKeyCount) As Long
        ReDim .SortType(1 To .SortKeyCount) As Long
        For Each oGroup In oCtl.Groups
            lIdx = lIdx + 1
            pvReadSortKey uRowSet, oCtl, lIdx, oGroup.ColIndex, oGroup.SortOrder
        Next
        For Each oKey In oCtl.SortKeys
            lIdx = lIdx + 1
            pvReadSortKey uRowSet, oCtl, lIdx, oKey.ColIndex, oKey.SortOrder
        Next
        If .ItemCount > 0 Then
            ReDim .Key(1 To .SortKeyCount, 1 To .ItemCount) As Variant
        End If
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvReadSortKey(uRowSet As UcsRowSet, oCtl As GridEX, ByVal lIdx As Long, ByVal lCol As Long, ByVal lOrder As Long)
    With uRowSet
        If lCol >= 1 And lCol <= oCtl.Columns.Count Then
            .SortCol(lIdx) = lCol
            .SortType(lIdx) = oCtl.Columns.Item(lCol).SortType
        End If
        '--- the enum is +1/-1 already, so the direction multiplies straight
        '--- into the comparison result
        .SortDir(lIdx) = IIf(lOrder = jgexSortDescending, -1, 1)
    End With
End Sub

'--- everything from the seeded order to the finished projection. The caller
'--- has filled Key by now, or has no sort keys at all
Public Sub DataProject(uRowSet As UcsRowSet)
    Const FUNC_NAME     As String = "DataProject"
    Dim lIdx            As Long
    Dim aTemp()         As Long

    On Error GoTo EH
    With uRowSet
        .Version = .Version + 1
        If .ItemCount = 0 Then
            Erase .Order
            Erase .GroupRow
            Erase .RowPosition
            .OrderCount = 0
            .GroupRowCount = 0
        Else
            ReDim .Order(1 To .ItemCount) As Long
            ReDim .RowPosition(1 To .ItemCount) As Long
            '--- seeded in RowIndex order, so the stable merge sort below
            '--- breaks ties the way the records were supplied
            For lIdx = 1 To .ItemCount
                .Order(lIdx) = lIdx
            Next
            .OrderCount = .ItemCount
            If .SortKeyCount > 0 Then
                ReDim aTemp(1 To .ItemCount) As Long
                pvMergeSort uRowSet, .Order, aTemp, 1, .ItemCount
            End If
            If .GroupColCount > 0 Then
                pvBuildGroupRows uRowSet
            Else
                Erase .GroupRow
                .GroupRowCount = 0
            End If
            '--- the keys are only worth their memory while sorting
            Erase .Key
        End If
    End With
    DataBuildVisible uRowSet
    DataWritePositions uRowSet
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'--- an expand or a collapse reprojects with these two and never re-sorts:
'--- the order underneath stays exactly as it was
Public Sub DataBuildVisible(uRowSet As UcsRowSet)
    Const FUNC_NAME     As String = "DataBuildVisible"
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim lHidden         As Long

    On Error GoTo EH
    With uRowSet
        .VisibleCount = 0
        Erase .Visible
        If .OrderCount = 0 Then
            Exit Sub
        End If
        ReDim .Visible(1 To .OrderCount) As Long
        For lIdx = 1 To .OrderCount
            If lHidden > 0 And .Order(lIdx) < 0 Then
                '--- a collapsed group hides its own footer along with its
                '--- records, so only a header at that level -- or anything
                '--- from a level further out, footers included -- brings the
                '--- display back
                If .GroupRow(-.Order(lIdx)).Level < lHidden Then
                    lHidden = 0
                ElseIf .GroupRow(-.Order(lIdx)).Level = lHidden Then
                    If Not .GroupRow(-.Order(lIdx)).Footer Then
                        lHidden = 0
                    End If
                End If
            End If
            If lHidden = 0 Then
                lPos = lPos + 1
                .Visible(lPos) = lIdx
                If .Order(lIdx) < 0 Then
                    If .GroupRow(-.Order(lIdx)).Collapsed Then
                        lHidden = .GroupRow(-.Order(lIdx)).Level
                    End If
                End If
            End If
        Next
        .VisibleCount = lPos
        ReDim Preserve .Visible(1 To lPos) As Long
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub DataWritePositions(uRowSet As UcsRowSet)
    Const FUNC_NAME     As String = "DataWritePositions"
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim lLast           As Long

    On Error GoTo EH
    '--- walks the FULL order, not the visible one, so a record hidden inside
    '--- a collapsed group still gets an answer: the position of the group
    '--- row hiding it, which is the last position actually emitted
    With uRowSet
        For lIdx = 1 To .OrderCount
            If lPos < .VisibleCount Then
                If .Visible(lPos + 1) = lIdx Then
                    lPos = lPos + 1
                    lLast = lPos
                End If
            End If
            If .Order(lIdx) > 0 Then
                .RowPosition(.Order(lIdx)) = lLast
            Else
                .GroupRow(-.Order(lIdx)).RowPosition = lLast
            End If
        Next
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'--- the order entry a display position holds: +n a RowIndex, -n a group row
'--- room for one more record, growing in blocks so an append is not a copy
Public Sub DataEnsureBookmarkRoom(uState As UcsDataState, ByVal lItemCount As Long)
    Const FUNC_NAME     As String = "DataEnsureBookmarkRoom"
    Dim lNeeded         As Long

    On Error GoTo EH
    If lItemCount <= uState.BookmarkUBound Then
        Exit Sub
    End If
    lNeeded = uState.BookmarkUBound
    If lNeeded < 16 Then
        lNeeded = 16
    End If
    Do While lNeeded < lItemCount
        lNeeded = lNeeded * 2
    Loop
    ReDim Preserve uState.Bookmark(1 To lNeeded) As Variant
    uState.BookmarkUBound = lNeeded
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'--- the record a position holds, 0 for a group row
Public Function DataRowIndexAt(uRowSet As UcsRowSet, ByVal lRowPosition As Long) As Long
    Const FUNC_NAME     As String = "DataRowIndexAt"
    Dim lRowIndex       As Long

    On Error GoTo EH
    lRowIndex = DataRecordIndexAt(uRowSet, lRowPosition)
    If lRowIndex > 0 Then
        DataRowIndexAt = lRowIndex
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'--- the original sizes its bookmark store to ItemCount and raises "Subscript
'--- out of range" past the end. This one answers Empty and writes nothing
'--- instead: a row index the caller worked out from a stale count is a
'--- question with no answer rather than a fault, and the paint path asks it
'--- often enough that raising there costs more than it says
Public Function DataRowBookmark(uState As UcsDataState, ByVal lRowIndex As Long) As Variant
    Const FUNC_NAME     As String = "DataRowBookmark"

    On Error GoTo EH
    If lRowIndex < 1 Or lRowIndex > uState.ItemCount Then
        Exit Function
    End If
    AssignVariant DataRowBookmark, uState.Bookmark(lRowIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub DataSetRowBookmark(uState As UcsDataState, ByVal lRowIndex As Long, vValue As Variant)
    Const FUNC_NAME     As String = "DataSetRowBookmark"

    On Error GoTo EH
    If lRowIndex < 1 Or lRowIndex > uState.ItemCount Then
        Exit Sub
    End If
    AssignVariant uState.Bookmark(lRowIndex), vValue
    '--- duplicates are legal, and which record answers depends on every other
    '--- bookmark, so the map is rebuilt rather than patched
    uState.BookmarksDirty = True
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function DataRowExpanded(uRowSet As UcsRowSet, ByVal lRowPosition As Long) As Boolean
    Const FUNC_NAME     As String = "DataRowExpanded"
    Dim lRowIndex       As Long

    On Error GoTo EH
    lRowIndex = DataRecordIndexAt(uRowSet, lRowPosition)
    If lRowIndex < 0 Then
        DataRowExpanded = Not uRowSet.GroupRow(-lRowIndex).Collapsed
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub DataSetRowExpanded(uRowSet As UcsRowSet, ByVal lRowPosition As Long, ByVal bValue As Boolean)
    Const FUNC_NAME     As String = "DataSetRowExpanded"
    Dim lRowIndex       As Long

    On Error GoTo EH
    lRowIndex = DataRecordIndexAt(uRowSet, lRowPosition)
    If lRowIndex < 0 Then
        '--- a footer closes a group rather than opening one, so it has no
        '--- expand box and nothing to toggle. Nested rather than And-ed, since
        '--- VB6 evaluates both operands and the subscript is only valid on the
        '--- negative branch
        If Not uRowSet.GroupRow(-lRowIndex).Footer Then
            If uRowSet.GroupRow(-lRowIndex).Collapsed = bValue Then
                '--- an expand or a collapse only reprojects: the sort order
                '--- underneath stays exactly as it was
                uRowSet.GroupRow(-lRowIndex).Collapsed = Not bValue
                DataBuildVisible uRowSet
                DataWritePositions uRowSet
            End If
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub DataSetAllExpanded(uRowSet As UcsRowSet, ByVal bValue As Boolean)
    Const FUNC_NAME     As String = "DataSetAllExpanded"
    Dim lIdx            As Long

    On Error GoTo EH
    For lIdx = 1 To uRowSet.GroupRowCount
        uRowSet.GroupRow(lIdx).Collapsed = Not bValue
    Next
    '--- reprojects without re-sorting, so Version stays put and a held group
    '--- wrapper survives a collapse -- which is what the original does
    DataBuildVisible uRowSet
    DataWritePositions uRowSet
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function DataGetRowPosition(uState As UcsDataState, uRowSet As UcsRowSet, ByVal lRowIndex As Long) As Long
    Const FUNC_NAME     As String = "DataGetRowPosition"

    On Error GoTo EH
    If lRowIndex >= 1 And lRowIndex <= uState.ItemCount Then
        DataGetRowPosition = uRowSet.RowPosition(lRowIndex)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function DataGetRowIndex(uState As UcsDataState, vBookmark As Variant) As Long
    Const FUNC_NAME     As String = "DataGetRowIndex"
    Dim sKey            As String
    Dim vRetVal         As Variant

    On Error GoTo EH
    If DataIsBlank(vBookmark) Then
        Exit Function
    End If
    sKey = DataBookmarkKey(vBookmark)
    If LenB(sKey) = 0 Then
        Exit Function
    End If
    DataEnsureBookmarks uState
    If SearchCollection(uState.Bookmarks, sKey, RetVal:=vRetVal) Then
        DataGetRowIndex = C2Lng(vRetVal)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function DataGetRecordCount(uRowSet As UcsRowSet, ByVal lRowIndex As Long) As Long
    Const FUNC_NAME     As String = "DataGetRecordCount"

    On Error GoTo EH
    If lRowIndex < 0 Then
        DataGetRecordCount = uRowSet.GroupRow(-lRowIndex).RecordCount
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function DataGetBookmarks(uState As UcsDataState, uRowSet As UcsRowSet, ByVal lRowIndex As Long) As Variant
    Const FUNC_NAME     As String = "DataGetBookmarks"
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim aRetVal()       As Variant

    On Error GoTo EH
    If lRowIndex >= 0 Then
        Exit Function
    End If
    With uRowSet.GroupRow(-lRowIndex)
        ReDim aRetVal(0 To .RecordCount - 1) As Variant
        For lIdx = .FirstSlot To .LastSlot
            If uRowSet.Order(lIdx) > 0 Then
                AssignVariant aRetVal(lCount), uState.Bookmark(uRowSet.Order(lIdx))
                lCount = lCount + 1
            End If
        Next
    End With
    DataGetBookmarks = aRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function DataGetRowIndexes(uRowSet As UcsRowSet, ByVal lRowIndex As Long) As Variant
    Const FUNC_NAME     As String = "DataGetRowIndexes"
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim aRetVal()       As Long

    On Error GoTo EH
    If lRowIndex >= 0 Then
        Exit Function
    End If
    With uRowSet.GroupRow(-lRowIndex)
        ReDim aRetVal(0 To .RecordCount - 1) As Long
        For lIdx = .FirstSlot To .LastSlot
            If uRowSet.Order(lIdx) > 0 Then
                aRetVal(lCount) = uRowSet.Order(lIdx)
                lCount = lCount + 1
            End If
        Next
    End With
    DataGetRowIndexes = aRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub DataEnsureBookmarks(uState As UcsDataState)
    Const FUNC_NAME     As String = "DataEnsureBookmarks"
    Dim lIdx            As Long
    Dim sKey            As String

    On Error GoTo EH
    If Not uState.BookmarksDirty Then
        Exit Sub
    End If
    uState.BookmarksDirty = False
    Set uState.Bookmarks = New VBA.Collection
    '--- walked upwards so the lowest RowIndex wins a duplicate, which is what
    '--- the original resolves to: with "bk-1" on records 1 and 3, all of
    '--- MoveToBookmark, AddBookmark and RefreshRowBookmark answer 1
    For lIdx = 1 To uState.ItemCount
        If Not DataIsBlank(uState.Bookmark(lIdx)) Then
            sKey = DataBookmarkKey(uState.Bookmark(lIdx))
            If LenB(sKey) <> 0 Then
                If Not SearchCollection(uState.Bookmarks, sKey) Then
                    uState.Bookmarks.Add lIdx, sKey
                End If
            End If
        End If
    Next
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'--- one FetchData cell, on demand. A sort is the only thing that ever asks
'--- for all of them, and what it leaves behind is a warm cache rather than a
'--- separate whole-column pass
Public Function DataFetchValue(uState As UcsDataState, uFetch As UcsFetchCache, oCtl As GridEX, _
            ByVal lRowIndex As Long, ByVal lColIndex As Long) As Variant
    Const FUNC_NAME     As String = "DataFetchValue"

    On Error GoTo EH
    DataEnsureFetchRoom uFetch, uState.ItemCount
    If lRowIndex < 1 Or lRowIndex > uFetch.RowCount Then
        Exit Function
    End If
    If Not uFetch.Valid(lColIndex, lRowIndex) Then
        '--- marked valid first so a read from inside the client's handler does
        '--- not re-enter the same cell
        uFetch.Valid(lColIndex, lRowIndex) = True
        AssignVariant uFetch.Value(lColIndex, lRowIndex), _
            oCtl.frRaiseFetchData(lRowIndex, lColIndex, uState.Bookmark(lRowIndex))
    End If
    AssignVariant DataFetchValue, uFetch.Value(lColIndex, lRowIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'--- a record that changed is asked for again column by column, rather than
'--- the cache being thrown away and the whole grid refetched
Public Sub DataRefetchRow(uState As UcsDataState, uFetch As UcsFetchCache, oCtl As GridEX, _
            ByVal lRowIndex As Long)
    Const FUNC_NAME     As String = "DataRefetchRow"
    Dim lIdx            As Long
    Dim vOld            As Variant

    On Error GoTo EH
    DataEnsureFetchRoom uFetch, uState.ItemCount
    If lRowIndex < 1 Or lRowIndex > uFetch.RowCount Then
        Exit Sub
    End If
    For lIdx = 1 To uFetch.ColCount
        If uFetch.IsFetchData(lIdx) Then
            AssignVariant vOld, uFetch.Value(lIdx, lRowIndex)
            uFetch.Valid(lIdx, lRowIndex) = True
            AssignVariant uFetch.Value(lIdx, lRowIndex), _
                oCtl.frRaiseFetchData(lRowIndex, lIdx, uState.Bookmark(lRowIndex))
            '--- a fetch column can be a sort or a group key, but only a value
            '--- that actually moved can move the record
            If DataCompareValues(vOld, uFetch.Value(lIdx, lRowIndex), jgexSortTypeString) <> 0 Then
                uState.OrderDirty = True
            End If
        End If
    Next
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function DataRecordIndexAt(uRowSet As UcsRowSet, ByVal lRowPosition As Long) As Long
    Const FUNC_NAME     As String = "DataRecordIndexAt"

    On Error GoTo EH
    With uRowSet
        If lRowPosition >= 1 And lRowPosition <= .VisibleCount Then
            DataRecordIndexAt = .Order(.Visible(lRowPosition))
        End If
    End With
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function DataIsBlank(vValue As Variant) As Boolean
    Const FUNC_NAME     As String = "DataIsBlank"

    On Error GoTo EH
    Select Case VarType(vValue)
    Case vbEmpty, vbNull, vbError
        DataIsBlank = True
    Case vbString
        DataIsBlank = (LenB(vValue) = 0)
    End Select
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function DataCompareValues(vValue1 As Variant, vValue2 As Variant, ByVal eSortType As jgexSortTypeConstants) As Long
    Const FUNC_NAME     As String = "DataCompareValues"
    Dim bBlank1         As Boolean
    Dim bBlank2         As Boolean

    On Error GoTo EH
    '--- blanks sort first, whatever the type
    bBlank1 = DataIsBlank(vValue1)
    bBlank2 = DataIsBlank(vValue2)
    If bBlank1 Or bBlank2 Then
        DataCompareValues = IIf(bBlank1, 0, 1) - IIf(bBlank2, 0, 1)
        Exit Function
    End If
    Select Case eSortType
    Case jgexSortTypeNumeric
        DataCompareValues = Sgn(C2Dbl(vValue1) - C2Dbl(vValue2))
    Case jgexSortTypeDate, jgexSortTypeDateTime, jgexSortTypeTime
        DataCompareValues = Sgn(C2Date(vValue1) - C2Date(vValue2))
    Case Else
        DataCompareValues = StrComp(C2Str(vValue1), C2Str(vValue2), vbTextCompare)
    End Select
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function DataBookmarkKey(vValue As Variant) As String
    Const FUNC_NAME     As String = "DataBookmarkKey"
    Dim baData()        As Byte

    On Error GoTo EH
    '--- a string never matches a number, but every numeric width does, and
    '--- so do Boolean and Date against their numeric value -- probed against
    '--- the original: True resolves CInt(-1), a Date resolves its CDbl, "7"
    '--- does not resolve 7. An ADO client-side cursor hands its bookmarks
    '--- out as byte arrays, which no scalar coercion can key on, so those
    '--- key on their bytes -- and an array of anything else keys on nothing
    '--- at all, a question with no answer rather than a type mismatch
    If VarType(vValue) = vbArray + vbByte Then
        baData = vValue
        DataBookmarkKey = "B" & ToHex(baData)
    ElseIf IsArray(vValue) Then
        Exit Function
    ElseIf VarType(vValue) = vbString Then
        DataBookmarkKey = "S" & vValue
    Else
        DataBookmarkKey = "#" & C2Dbl(vValue)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'--- one walk over values the caller has already collected for a group's
'--- records serves every function: the counts take any value, the rest only
'--- the ones that are numbers
Public Function DataAggregate(aValues() As Variant, ByVal lValueCount As Long, ByVal eFunc As jgexAggregateFunctionConstants, ByVal eSortType As jgexSortTypeConstants) As Variant
    Const FUNC_NAME     As String = "DataAggregate"
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim dblSum          As Double
    Dim dblSquares      As Double
    Dim dblValue        As Double
    Dim vMin            As Variant
    Dim vMax            As Variant

    On Error GoTo EH
    If eFunc = jgexCount Then
        DataAggregate = lValueCount
        Exit Function
    End If
    For lIdx = 1 To lValueCount
        If Not DataIsBlank(aValues(lIdx)) And Not IsObject(aValues(lIdx)) Then
            lCount = lCount + 1
            If IsEmpty(vMin) Then
                vMin = aValues(lIdx)
                vMax = aValues(lIdx)
            Else
                If DataCompareValues(aValues(lIdx), vMin, eSortType) < 0 Then
                    vMin = aValues(lIdx)
                End If
                If DataCompareValues(aValues(lIdx), vMax, eSortType) > 0 Then
                    vMax = aValues(lIdx)
                End If
            End If
            If IsNumeric(aValues(lIdx)) Then
                dblValue = C2Dbl(aValues(lIdx))
                dblSum = dblSum + dblValue
                dblSquares = dblSquares + dblValue * dblValue
            End If
        End If
    Next
    Select Case eFunc
    Case jgexValueCount
        DataAggregate = lCount
    Case jgexSum
        DataAggregate = dblSum
    Case jgexAvg
        If lCount > 0 Then
            DataAggregate = dblSum / lCount
        End If
    Case jgexMin
        DataAggregate = vMin
    Case jgexMax
        DataAggregate = vMax
    Case jgexStdDev
        '--- population deviation over the values that were numbers
        If lCount > 1 Then
            dblValue = (dblSquares - dblSum * dblSum / lCount) / lCount
            If dblValue > 0 Then
                DataAggregate = Sqr(dblValue)
            Else
                DataAggregate = 0
            End If
        End If
    End Select
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvMergeSort(uRowSet As UcsRowSet, aIdx() As Long, aTemp() As Long, ByVal lFirst As Long, ByVal lLast As Long)
    Dim lMid            As Long
    Dim lLeft           As Long
    Dim lRight          As Long
    Dim lPos            As Long

    If lLast - lFirst < 1 Then
        Exit Sub
    End If
    lMid = (lFirst + lLast) \ 2
    pvMergeSort uRowSet, aIdx, aTemp, lFirst, lMid
    pvMergeSort uRowSet, aIdx, aTemp, lMid + 1, lLast
    lLeft = lFirst
    lRight = lMid + 1
    lPos = lFirst
    Do While lLeft <= lMid And lRight <= lLast
        '--- <= keeps it stable, which is what preserves supply order on ties
        If pvCompareRows(uRowSet, aIdx(lLeft), aIdx(lRight)) <= 0 Then
            aTemp(lPos) = aIdx(lLeft)
            lLeft = lLeft + 1
        Else
            aTemp(lPos) = aIdx(lRight)
            lRight = lRight + 1
        End If
        lPos = lPos + 1
    Loop
    Do While lLeft <= lMid
        aTemp(lPos) = aIdx(lLeft)
        lLeft = lLeft + 1
        lPos = lPos + 1
    Loop
    Do While lRight <= lLast
        aTemp(lPos) = aIdx(lRight)
        lRight = lRight + 1
        lPos = lPos + 1
    Loop
    For lPos = lFirst To lLast
        aIdx(lPos) = aTemp(lPos)
    Next
End Sub

Private Function pvCompareRows(uRowSet As UcsRowSet, ByVal lRow1 As Long, ByVal lRow2 As Long) As Long
    Dim lIdx            As Long

    With uRowSet
        For lIdx = 1 To .SortKeyCount
            pvCompareRows = DataCompareValues(.Key(lIdx, lRow1), .Key(lIdx, lRow2), .SortType(lIdx)) * .SortDir(lIdx)
            If pvCompareRows <> 0 Then
                Exit Function
            End If
        Next
    End With
End Function

Private Sub pvBuildGroupRows(uRowSet As UcsRowSet)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lPos            As Long
    Dim lFrom           As Long
    Dim lRoom           As Long
    Dim aRows()         As Long
    Dim aStart()        As Long

    '--- the sorted rows are walked once and a group row is emitted wherever
    '--- a group key changes, so the display is group row, its records, the
    '--- next group row and so on. Room for a header and a footer per level
    '--- per record, which all-distinct keys would come to
    With uRowSet
        aRows = .Order
        lRoom = .ItemCount + 2 * .ItemCount * .GroupColCount
        ReDim .Order(1 To lRoom) As Long
        ReDim .GroupRow(1 To lRoom) As UcsGroupRow
        .GroupRowCount = lRoom
        ReDim aStart(1 To .GroupColCount) As Long
        For lIdx = 1 To .ItemCount
            '--- the first level whose key changed opens a new group row here
            '--- and at every level below it, so a change high up restarts the
            '--- ones nested inside it
            lFrom = 0
            For lJdx = 1 To .GroupColCount
                If lIdx = 1 Then
                    lFrom = 1
                    Exit For
                End If
                If DataCompareValues(.Key(lJdx, aRows(lIdx)), .Key(lJdx, aRows(lIdx - 1)), .SortType(lJdx)) <> 0 Then
                    lFrom = lJdx
                    Exit For
                End If
            Next
            If lFrom > 0 Then
                pvCloseGroups uRowSet, lFrom, lPos, aStart
                For lJdx = lFrom To .GroupColCount
                    lPos = lPos + 1
                    .Order(lPos) = -lPos
                    With .GroupRow(lPos)
                        .Level = lJdx
                        .ColIndex = uRowSet.SortCol(lJdx)
                        .FirstRowIndex = aRows(lIdx)
                        .RecordCount = 0
                        .Collapsed = uRowSet.DefaultCollapsed
                    End With
                    aStart(lJdx) = lPos
                Next
            End If
            For lJdx = 1 To .GroupColCount
                .GroupRow(aStart(lJdx)).RecordCount = .GroupRow(aStart(lJdx)).RecordCount + 1
            Next
            lPos = lPos + 1
            .Order(lPos) = aRows(lIdx)
        Next
        '--- the last group of every level is closed by running out of records
        pvCloseGroups uRowSet, 1, lPos, aStart
        .OrderCount = lPos
        ReDim Preserve .Order(1 To lPos) As Long
        ReDim Preserve .GroupRow(1 To lPos) As UcsGroupRow
        .GroupRowCount = lPos
    End With
End Sub

Private Sub pvCloseGroups(uRowSet As UcsRowSet, ByVal lFromLevel As Long, lPos As Long, aStart() As Long)
    Dim lJdx            As Long
    Dim lHdr            As Long

    '--- closing a level records the span of rows it covered -- what the
    '--- aggregates are computed over -- and lays down its footer row when
    '--- the control is showing them, innermost level first, so a footer
    '--- lands immediately before the next group row at its own level
    With uRowSet
        For lJdx = .GroupColCount To lFromLevel Step -1
            lHdr = aStart(lJdx)
            If lHdr > 0 Then
                .GroupRow(lHdr).FirstSlot = lHdr + 1
                .GroupRow(lHdr).LastSlot = lPos
                If .GroupFooterStyle <> jgexNoGroupFooter Then
                    lPos = lPos + 1
                    .Order(lPos) = -lPos
                    .GroupRow(lPos) = .GroupRow(lHdr)
                    .GroupRow(lPos).Footer = True
                End If
            End If
        Next
    End With
End Sub
