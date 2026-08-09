VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Open GridEX 2000 Control
' The contract between the grid and whatever holds its records
'
' Three address spaces meet here and must never be confused:
'
'   RowPosition   1..RowCount     a row currently on show, group headers and
'                                 footers included, rows inside a collapsed
'                                 group excluded. What the renderer, the
'                                 scrollbar and every public *Position*
'                                 member speak. Derived -- nothing survives
'                                 a re-sort or a collapse
'   RowIndex      1..ItemCount    a record of the source, in the order the
'                                 source supplied it. Where identity lives:
'                                 the current cell and the selection are
'                                 kept on it, and bookmarks are keyed by it.
'                                 Stable across sorting and grouping -- only
'                                 Delete renumbers it, by compaction
'   Bookmark      opaque          the record's own identity as the client
'                                 states it, resolving into RowIndex space
'
' The model owns the projection -- ordering, group rows, collapse state and
' therefore the RowPosition space. The grid owns the viewport and the pixels
' and never sees a record the window is not showing. Columns, SortKeys,
' Groups and ItemCount are read off the weak owner reference each
' implementation takes in its own frInit, so nothing is fed in and nothing
' can go stale
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Properties
'=========================================================================

'--- Rows on show: ItemCount plus the group headers and footers the data
'--- broke into, less whatever a collapsed group hides. The vertical
'--- scrollbar's range, and the domain of every member taking a RowPosition
Public Property Get RowCount() As Long
End Property

'--- Records of the source
Public Property Get ItemCount() As Long
End Property

'--- Counts rebuilds of the projection, so a wrapper can tell whether the
'--- order it was filled under still stands. JSRowData captures it and
'--- refuses the group members on a mismatch. A model that cannot lose a
'--- span this way -- or does not care to -- may leave it at 0, which never
'--- mismatches and so never refuses
Public Property Get Version() As Long
End Property

'--- The signed index a position holds: positive a record's RowIndex,
'--- negative a group row's projection slot, 0 out of range. This and
'--- RowBookmark below are the two queries the original answers publicly,
'--- and they run the same way: from the derived space towards identity,
'--- RowPosition -> RowIndex -> Bookmark -- the public surface masks the
'--- negatives to 0, while the grid's window and selection carry wrappers
'--- and items by them, since neither an expand nor a collapse re-sorts the
'--- order underneath. It reads the projection only and never populates, so
'--- a caller after the index alone does not drag an UnboundReadData behind
'--- it the way GetData would
Public Property Get RowIndex(ByVal lRowPosition As Long) As Long
End Property

'--- Keyed by RowIndex, which is what the original does: probed live,
'--- RowBookmark(ItemCount + 1) raises subscript out of range, the store
'--- survives both Refetch and Rebind, and a Delete compacts it -- deleting
'--- the record holding "bk-3" leaves the next read of RowIndex 3 carrying
'--- "bk-4". Assigning one raises no event, so it cannot be routed through
'--- UpdateData
Public Property Get RowBookmark(ByVal lRowIndex As Long) As Variant
End Property

Public Property Let RowBookmark(ByVal lRowIndex As Long, vValue As Variant)
End Property

Public Property Get RowExpanded(ByVal lRowPosition As Long) As Boolean
End Property

Public Property Let RowExpanded(ByVal lRowPosition As Long, ByVal bValue As Boolean)
End Property

'=========================================================================
' Methods
'=========================================================================

'--- Detaches the model from the control, and is called before the control
'--- releases it. Not the same as letting Class_Terminate do it: a JSRowData
'--- buffer holds its model strongly, so a wrapper a client kept can outlive
'--- the control and keep the model with it. Without this the model's weak
'--- owner pointer would dangle and the next call through that wrapper would
'--- dereference a dead control; after it, the model raises instead
Public Sub Terminate()
End Sub

'--- Everything the model is told, in the shape the control's own methods
'--- come in. Refresh re-reads Columns, SortKeys, Groups and ItemCount off
'--- the owner and rebuilds records and group rows from nothing, which is
'--- what Rebind and Refetch both come to -- bookmarks survive either way,
'--- probed at both
Public Sub Refresh()
End Sub

'--- The projection is stale but the records are not: nothing is refetched,
'--- and GroupFooterStyle is re-read off the owner along with Groups. The
'--- rebuild resets every expand box to DefaultGroupMode; bAllCollapsed
'--- overrides that in the same pass rather than reprojecting a second time
Public Sub RefreshGroups(ByVal bAllCollapsed As Boolean)
End Sub

Public Sub RefreshSort()
End Sub

'--- One record re-read, the projection left alone
Public Sub RefreshRowIndex(ByVal lRowIndex As Long)
End Sub

'--- Mints the wrapper for one display position and fills it: row index, row
'--- type, group level, bookmark, group caption, and every column's Value
'--- and IconIndex. It does not touch DisplayValue, CellStyle or CellPicture
'--- -- those are the client's to write in RowFormat, which the grid raises
'--- once per population and never again on a repaint.
'--- A fresh instance every time, which is what the original does: probed
'--- over 95 populations it handed out 95 distinct wrappers and reused none,
'--- and the one a client keeps goes on reading the row it was filled with
Public Function GetRowData(ByVal lRowPosition As Long) As JSRowData
End Function

'--- Records only. Writes the buffer's values back to the source for the
'--- record its RowIndex names. A negative RowIndex is a record that does
'--- not exist yet: the model appends it, whatever NewRowPos says -- that
'--- places the new-row line, not the record -- then writes the index it was
'--- given and the bookmark it came back with into
'--- the same buffer, so the caller reads both off the row it just added.
'--- Growing ItemCount on that path must wait until the client's handler has
'--- returned -- doing it inside re-enters the add and the original never
'--- leaves the new row
Public Sub UpdateRowData(oData As JSRowData)
End Sub

'--- By position, since the grid deletes what the user is standing on and a
'--- group row is not something it can be standing on for this. Compacts the
'--- records behind it, as the original does unasked
Public Sub Delete(ByVal lRowPosition As Long)
End Sub

'--- Where the current cell and the selection land after a re-sort or a
'--- collapse: a record hidden inside a collapsed group answers with the
'--- group row hiding it, and 0 when it is not in the projection at all
Public Function GetRowPosition(ByVal lRowIndex As Long) As Long
End Function

'--- Resolves a bookmark into the stable space, which is what MoveToBookmark,
'--- RefreshRowBookmark and JSSelectedItems.Add/RemoveBookmark all need. It
'--- answers a RowIndex rather than a position on purpose: a collapsed record
'--- has no position of its own, so resolving straight to one would lose the
'--- record. Compose GetRowPosition over it where a position is wanted.
'--- 0 means unresolved, which is what the original raises "Not a valid
'--- Bookmark." over
Public Function GetRowIndex(vBookmark As Variant) As Long
End Function

'--- every group row at once, which the setter above cannot express: a group
'--- nested inside a collapsed one has no position of its own, so a walk over
'--- positions never reaches it and has to keep re-walking as it uncovers
'--- more. The model has the group rows in hand and touches each exactly once
Public Sub SetAllExpanded(ByVal bValue As Boolean)
End Sub

'--- The JSRowData forwards. All are asked of a group row and answer over
'--- the records it covers, nested group rows skipped. They take the signed
'--- row index the projection itself uses -- positive a record, negative a
'--- group header or footer -- and not a position: a position is re-resolved
'--- against whatever the projection currently says, where the negative
'--- index names the group row the wrapper was actually filled from
Public Function GetRecordCount(ByVal lRowIndex As Long) As Long
End Function

Public Function GetSubTotal(ByVal lRowIndex As Long, ByVal lColIndex As Long, ByVal eFunc As jgexAggregateFunctionConstants) As Variant
End Function

Public Function GetBookmarks(ByVal lRowIndex As Long) As Variant
End Function

Public Function GetRowIndexes(ByVal lRowIndex As Long) As Variant
End Function
