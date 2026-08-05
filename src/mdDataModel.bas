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

Public Const ERR_ROWDATA_INVALID    As String = "JSRowData object may have changed. The object is no longer valid."
Public Const ERR_ROWDATA_NUMBER     As Long = vbObjectError + 129

'=========================================================================
' Public types
'=========================================================================

Public Type UcsGroupRow
    Level                   As Integer
    ColIndex                As Integer
    '--- the record whose key opened the group: the caption is formatted
    '--- from it on demand rather than stored, since GroupPrefix, GroupFormat
    '--- and GroupEmptyStringCaption are the control's business and can
    '--- change without the projection changing
    FirstRowIndex           As Long
    RecordCount             As Long
    Collapsed               As Boolean
    Footer                  As Boolean
    FirstSlot               As Long
    LastSlot                As Long
    RowPosition             As Long
End Type

Public Type UcsRowSet
    '--- bumped by DataProject alone, so it counts order rebuilds and not
    '--- reprojections: a wrapper filled under one order is refused once the
    '--- order is rebuilt, and survives a collapse untouched. Probed against
    '--- the original, which raises on RecordCount, GetSubTotal, GetBookmarks
    '--- and GetRowIndexes after Delete, RefreshGroups, RefreshSort or Rebind
    '--- but answers them normally across a collapse and expand
    Version                 As Long
    ItemCount               As Long
    GroupFooterStyle        As jgexGroupFooterStyleConstants
    DefaultCollapsed        As Boolean
    '--- sort metadata, read off the control once per projection
    SortKeyCount            As Long
    GroupColCount           As Long
    SortCol()               As Integer
    SortDir()               As Long
    SortType()              As Long
    '--- every key of every record, by (key, RowIndex). Filled by the caller,
    '--- which is the only part of the pass that needs a cell value, and
    '--- dropped as soon as the sort is done
    Key()                   As Variant
    '--- the sorted order: +n a RowIndex, -n an index into GroupRow
    Order()                 As Long
    OrderCount              As Long
    GroupRow()              As UcsGroupRow
    GroupRowCount           As Long
    '--- what is on show: indices into Order, 1..VisibleCount
    Visible()               As Long
    VisibleCount            As Long
    '--- by RowIndex: the position a record shows at, or the position of the
    '--- group row hiding it when its group is collapsed
    RowPosition()           As Long
End Type

'--- a column with FetchData set is supplied by the client through the
'--- FetchData event, whatever else it could have been read from: in a bound
'--- grid that outranks its DataField, and in an unbound one it outranks
'--- whatever UnboundReadData wrote.
'
'--- The cache is per cell, not per column, and fills on demand: a repaint
'--- asks for the screenful it is painting and no more. The only thing that
'--- ever sweeps a whole column is a sort, because a sort genuinely needs
'--- every record's key -- and what it leaves behind is a warm cache rather
'--- than a second pass. Everything narrower invalidates by the row: an edit
'--- or an insert refetches that record's fetch columns, a delete shifts the
'--- rows above it down, and only a rebind or a column change throws the lot
'--- away
Public Type UcsFetchCache
    ColCount                As Integer
    RowCount                As Long
    IsFetchData()           As Boolean
    IsFetchIcon()           As Boolean
    Valid()                 As Boolean
    Value()                 As Variant
End Type

'=========================================================================
' Functions
'=========================================================================

Public Sub DataInitFetch(uFetch As UcsFetchCache, oCtl As GridEX)
    Dim nIdx            As Integer

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
        For nIdx = 1 To .ColCount
            .IsFetchData(nIdx) = oCtl.Columns.Item(nIdx).FetchData
            .IsFetchIcon(nIdx) = oCtl.Columns.Item(nIdx).FetchIcon
        Next
    End With
End Sub

'--- grows without losing what is already cached: the row is the last
'--- dimension precisely so ReDim Preserve can do this
Public Sub DataEnsureFetchRoom(uFetch As UcsFetchCache, ByVal lItemCount As Long)
    With uFetch
        If .ColCount = 0 Or lItemCount <= .RowCount Then
            Exit Sub
        End If
        ReDim Preserve .Value(1 To .ColCount, 1 To lItemCount) As Variant
        ReDim Preserve .Valid(1 To .ColCount, 1 To lItemCount) As Boolean
        .RowCount = lItemCount
    End With
End Sub

'--- one record's fetch columns go stale -- an edit, or a Refresh of it
Public Sub DataInvalidateFetchRow(uFetch As UcsFetchCache, ByVal lRowIndex As Long)
    Dim nIdx            As Integer

    With uFetch
        If lRowIndex < 1 Or lRowIndex > .RowCount Then
            Exit Sub
        End If
        For nIdx = 1 To .ColCount
            .Valid(nIdx, lRowIndex) = False
        Next
    End With
End Sub

'--- a delete renumbers every record above it, so the cache follows rather
'--- than being thrown away and refetched
Public Sub DataDeleteFetchRow(uFetch As UcsFetchCache, ByVal lRowIndex As Long, ByVal lItemCount As Long)
    Dim nIdx            As Integer
    Dim lIdx            As Long

    With uFetch
        If lRowIndex < 1 Or lRowIndex > .RowCount Then
            Exit Sub
        End If
        For lIdx = lRowIndex To lItemCount - 1
            If lIdx + 1 <= .RowCount Then
                For nIdx = 1 To .ColCount
                    AssignVariant .Value(nIdx, lIdx), .Value(nIdx, lIdx + 1)
                    .Valid(nIdx, lIdx) = .Valid(nIdx, lIdx + 1)
                Next
            End If
        Next
        If lItemCount >= 1 And lItemCount <= .RowCount Then
            For nIdx = 1 To .ColCount
                .Value(nIdx, lItemCount) = Empty
                .Valid(nIdx, lItemCount) = False
            Next
        End If
    End With
End Sub

Public Sub DataReadSortKeys(uRowSet As UcsRowSet, oCtl As GridEX)
    Dim lIdx            As Long
    Dim nCol            As Integer
    Dim nOrder          As Integer

    '--- grouping sorts too: the group columns lead, in order, and the sort
    '--- keys refine within a group
    With uRowSet
        .GroupColCount = oCtl.Groups.Count
        .SortKeyCount = .GroupColCount + oCtl.SortKeys.Count
        If .SortKeyCount = 0 Then
            Exit Sub
        End If
        ReDim .SortCol(1 To .SortKeyCount) As Integer
        ReDim .SortDir(1 To .SortKeyCount) As Long
        ReDim .SortType(1 To .SortKeyCount) As Long
        For lIdx = 1 To .SortKeyCount
            If lIdx <= .GroupColCount Then
                nCol = oCtl.Groups.Item(lIdx).ColIndex
                nOrder = oCtl.Groups.Item(lIdx).SortOrder
            Else
                nCol = oCtl.SortKeys.Item(lIdx - .GroupColCount).ColIndex
                nOrder = oCtl.SortKeys.Item(lIdx - .GroupColCount).SortOrder
            End If
            If nCol >= 1 And nCol <= oCtl.Columns.Count Then
                .SortCol(lIdx) = nCol
                .SortType(lIdx) = oCtl.Columns.Item(nCol).SortType
            End If
            '--- the enum is +1/-1 already, so the direction multiplies
            '--- straight into the comparison result
            .SortDir(lIdx) = IIf(nOrder = jgexSortDescending, -1, 1)
        Next
        If .ItemCount > 0 Then
            ReDim .Key(1 To .SortKeyCount, 1 To .ItemCount) As Variant
        End If
    End With
End Sub

'--- everything from the seeded order to the finished projection. The caller
'--- has filled Key by now, or has no sort keys at all
Public Sub DataProject(uRowSet As UcsRowSet)
    Dim lIdx            As Long
    Dim aTemp()         As Long

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
End Sub

'--- an expand or a collapse reprojects with these two and never re-sorts:
'--- the order underneath stays exactly as it was
Public Sub DataBuildVisible(uRowSet As UcsRowSet)
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim lHidden         As Long

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
End Sub

Public Sub DataWritePositions(uRowSet As UcsRowSet)
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim lLast           As Long

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
End Sub

'--- the order entry a display position holds: +n a RowIndex, -n a group row
Public Function DataEntryAt(uRowSet As UcsRowSet, ByVal lRowPosition As Long) As Long
    With uRowSet
        If lRowPosition >= 1 And lRowPosition <= .VisibleCount Then
            DataEntryAt = .Order(.Visible(lRowPosition))
        End If
    End With
End Function

Public Function DataIsBlank(vValue As Variant) As Boolean
    Select Case VarType(vValue)
    Case vbEmpty, vbNull, vbError
        DataIsBlank = True
    Case vbString
        DataIsBlank = (LenB(vValue) = 0)
    End Select
End Function

Public Function DataCompareValues(vValue1 As Variant, vValue2 As Variant, ByVal eSortType As jgexSortTypeConstants) As Long
    Dim bBlank1         As Boolean
    Dim bBlank2         As Boolean

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
    Case Else
        DataCompareValues = StrComp(C2Str(vValue1), C2Str(vValue2), vbTextCompare)
    End Select
End Function

Public Function DataBookmarkKey(vValue As Variant) As String
    Dim baData()        As Byte

    '--- a string never matches a number, but every numeric width does, and
    '--- so do Boolean and Date against their numeric value -- probed against
    '--- the original: True resolves CInt(-1), a Date resolves its CDbl, "7"
    '--- does not resolve 7. An ADO client-side cursor hands its bookmarks
    '--- out as byte arrays, which no scalar coercion can key on, so those
    '--- key on their bytes
    If IsArray(vValue) Then
        baData = vValue
        DataBookmarkKey = "B" & ToHex(baData)
    ElseIf VarType(vValue) = vbString Then
        DataBookmarkKey = "S" & vValue
    Else
        DataBookmarkKey = "#" & C2Dbl(vValue)
    End If
End Function

'--- one walk over values the caller has already collected for a group's
'--- records serves every function: the counts take any value, the rest only
'--- the ones that are numbers
Public Function DataAggregate(aValues() As Variant, ByVal lValueCount As Long, ByVal eFunc As jgexAggregateFunctionConstants, ByVal eSortType As jgexSortTypeConstants) As Variant
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim dblSum          As Double
    Dim dblSquares      As Double
    Dim dblValue        As Double
    Dim vMin            As Variant
    Dim vMax            As Variant

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
