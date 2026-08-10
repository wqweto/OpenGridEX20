# Architecture

How the pieces of Open GridEX 2000 fit together, and why they are shaped the
way they are. `ROADMAP.md` says what is built and in what order; `CHANGELOG.md`
records what changed; this file explains the standing structure.

Sections other than **Data model** and **Painting** are stubs -- the shape of
the thing is already in the code, the writing-down has not happened yet.

## Overview

*Stub.* The control is a VB6 `UserControl` (`GridEX.ctl`) plus a preview
control (`GEXPreview.ctl`) and 23 `JS*` object-model classes, all reproducing
the public surface of the original Janus GridEX 2000 in
`doc/GridEX20.idl`. Compatibility is **source**-level: same member names,
signatures, enum values and event contracts, own GUIDs and ProgIDs.

Worth covering here: the split between the public surface (fixed by the
original) and the private implementation (free), and the rule that anything
private stays off the typelib.

## Data model

### The problem it solves

Until this layer existed, `GridEX.ctl` owned everything: the row cache, the
sort, the group rows, the collapse projection, the aggregates and the
painting. That works for one data source. It does not survive three -- unbound
rows supplied by the client, an ADO recordset, and eventually a SQLite table
where sorting, grouping and counting are the database's job rather than ours.

`IDataModel` is the seam. Below it, an implementation answers questions about
records. Above it, the control lays out pixels and never learns where a value
came from.

### Three address spaces

Everything in this layer hangs on keeping three ways of naming a row apart.
Confusing them is the single most likely source of bugs here.

| | range | survives | who speaks it |
|---|---|---|---|
| **RowPosition** | `1..RowCount` | nothing | the renderer, the scrollbar, every public `*Position` member |
| **RowIndex** | `1..ItemCount` | re-sorting, grouping, collapse | the current cell, the selection, bookmarks, `UnboundReadData` |
| **Bookmark** | opaque | a requery | `MoveToBookmark`, `GetBookmarks`, the unbound events |

RowPosition counts rows *on show* -- group headers and footers included, rows
inside a collapsed group excluded -- so it changes whenever anything is sorted,
grouped or collapsed. RowIndex is where identity lives: it is the order the
source supplied its records in, and **only a delete renumbers it**, by
compaction.

The original works the same way, which is verifiable: sorting a five-record
grid descending permutes the positions while `RowIndex(pos)` reports 5, 4, 3,
2, 1 and `RowBookmark(1)` still answers for the record that was first.

The original also only ever *answers* questions in one direction --
`RowIndex(RowPosition)` and `RowBookmark(RowIndex)`, both running from the
derived space toward identity. The two reverses exist there only as actions
(`MoveToRowIndex`, `MoveToBookmark`, `SelectedItems.AddBookmark`). `IDataModel`
makes them callable, because a private interface can afford to answer what a
public one only performed.

### The contract

`IDataModel` (`src/IDataModel.cls`, `Instancing = Private`, so it never reaches
the typelib):

```
Property Get  RowCount                         rows on show -- scrollbar range
Property Get  ItemCount                        records of the source
Property Get  Version                          bumped by a reprojection alone
Property Get  RowIndex(RowPosition)            signed: negative a group row's
                                               slot; the public surface masks
Property Get/Let RowBookmark(RowIndex)
Property Get/Let RowExpanded(RowPosition)
Sub  SetAllExpanded(bValue)                    every group row at once

Sub  Terminate                                 detach before release
Sub  Refresh                                   re-read everything
Sub  RefreshGroups                             regroup, records untouched
Sub  RefreshSort
Sub  RefreshRowIndex(RowIndex)                 one record re-read
Function GetRowData(RowPosition) As JSRowData   mint and fill a wrapper
Sub  UpdateRowData(oData)                      write back; RowIndex < 0 inserts
Sub  Delete(RowPosition)

Function GetRowPosition(RowIndex)              collapsed -> the group row hiding it
Function GetRowIndex(Bookmark)                 0 when unresolved
Function GetRecordCount(RowIndex)              signed: negative is a group row
Function GetSubTotal(RowIndex, ColIndex, Func)
Function GetBookmarks(RowIndex)
Function GetRowIndexes(RowIndex)
```

Four shaping decisions:

- **Nothing is fed in.** Each implementation takes a weak reference to the
  control in its own `frInit` and reads `Columns`, `SortKeys`, `Groups` and
  `ItemCount` from there when it rebuilds. Feeding them would mean a
  notification path per collection and a second copy that can go stale.
  `GroupFooterStyle` and `DefaultGroupMode` are read the same way, when the
  projection rebuilds.
- **`GetRowIndex` answers a RowIndex, not a position.** A record inside a
  collapsed group has no position of its own, so resolving a bookmark straight
  to one would lose the record. Callers that want a position compose
  `GetRowPosition` over it.
- **The forwards take a signed RowIndex, not a position.** `GetSubTotal` and
  its siblings are asked of a group row, and a position would be re-resolved
  against whatever the projection currently says. Positive is a record,
  negative a group row -- the projection's own addressing.
- **`SetAllExpanded` exists because the by-position setter cannot express it.**
  A group nested inside a collapsed one has no position, so a walk over
  positions never reaches it and has to keep re-walking as it uncovers more.
  The model has the group rows in hand and touches each exactly once.

### Implementations

| | source | identity | cell value |
|---|---|---|---|
| `cUnboundDataModel` | `UnboundReadData`, once per record | client-assigned bookmark | cached in `m_aRecord(i).Values` |
| `cAdoDataModel` | an ADO `Recordset` | the cursor's own bookmark | read live off a cached `Field` |
| *(future)* SQL | a SQLite table | primary key | `SELECT` |

`cUnboundDataModel` caches because the client is only asked once per record.
`cAdoDataModel` does not cache at all -- the recordset **is** the cache. What it
keeps per record is the bookmark it was handed on the binding walk; reading a
cell means positioning the cursor on that bookmark and asking a `Field`. Which
is why it refuses a recordset that does not support bookmarks: without them
there is no identity to return to and no RowIndex worth the name.

A SQL model would use almost none of the shared machinery below -- `RowCount`,
the projection and the aggregates all come from `ORDER BY`, `GROUP BY` and
`COUNT(*)` instead. That is the reason the seam is where it is.

### The shared projection engine

`mdDataModel.bas` holds the work that is identical whatever holds the records:

- `UcsRowSet` -- the projection state: the sort metadata, the key array, the
  order, the group rows, the visible map and the position write-back
- `DataReadSortKeys` -- reads `Groups` then `SortKeys` off the control, group
  columns leading, since grouping sorts too
- `DataProjectRows` -- seed in RowIndex order, stable merge sort, emit group rows,
  build the visible map, write positions
- `DataBuildVisibleRows` / `DataWritePositions` -- the collapse path on their own,
  because expand/collapse **reprojects without re-sorting**
- `DataCompareValues`, `DataIsBlankValue`, `DataGetBookmarkKey`, `DataAggregateValues`

The one step that cannot move is filling `uRowSet.Key`, because that is the
only part of the pass that needs a cell value. Each implementation does it its
own way -- the unbound model forces its lazy fetch, the ADO model walks the
cursor with `MoveNext` since RowIndex order *is* the recordset's order.

Group rows carry their span in order space (`FirstSlot`/`LastSlot`), which is
what the aggregates are computed over, and a footer inherits its header's span
so both total the same records. Captions are **not** stored: a group row keeps
the `RowIndex` that opened it, and the caption is formatted on demand from
`GroupPrefix`, `GroupFormat` and `GroupEmptyStringCaption`. `GroupFormat`
labels a caption without regrouping -- the original breaks groups on the raw
value, so two dates in one month give two identically-captioned groups.

### The window

The grid does not read the model cell by cell while painting. It keeps a
window of `JSRowData` wrappers, one per visible position, minted by
`GetRowData`.
`RowFormat` is raised **once per population**, which is what lets the client
write `DisplayValue`, `CellStyle` and `CellPicture` and have every subsequent
repaint use them without the event firing again.

Buffers are **not reused across populations**: repopulating a position mints a
fresh wrapper. A client holding a `GetRowData` result across a scroll is then
left with a read-only snapshot of the row that used to be there, rather than a
wrapper that silently repoints at a different record. The pending row is the
one exception -- the wrapper carrying uncommitted writes keeps its instance
across a repopulation, or a scroll would drop the edit.

They are deliberately **not** detached. The orphaning discipline elsewhere in
the object model exists because those classes reach the control through a raw,
un-AddRef-ed pointer, and dereferencing one after the control dies takes the
process down. A buffer carries its own data and never dereferences its owner;
the only reference it uses is the model, which is strong. So there is nothing
to protect against, and a stale buffer is harmless.

A buffer is only meaningful in the context it was handed out for. Per-row
presentation members -- `RowHeight`, `RowStyle`, `PreviewRowVisible` -- are
window state, re-established by `RowFormat` on each population; setting one
outside that event is not something the model preserves.

`JSRowData` carries the row itself and nothing else. It holds no reference to
the control at all, which is what makes an orphaned wrapper harmless rather
than a crash waiting to happen.

`Value` is read-only on it. The one place a write is legal is the buffer the
control hands out to be filled during `UnboundReadData`; anywhere else -- a
wrapper from `GetRowData`, a record or a group row alike -- it raises
`vbObjectError` with *"'Value' property can not be change in this context."*,
which is the original's message including its grammar slip. Editing goes
through `GridEX.Value`, which buffers on the current row and commits through
`UpdateRowData` when the row is left. `DisplayValue`, `CellStyle`, `RowStyle`
and the rest stay writable: decorating the row is what `RowFormat` is for.

Because `Bookmark`, `RowType`, `GroupLevel` and `RowIndex` are publicly
read-only, a model fills them through a `Friend` surface -- `frReset`,
`frInitModel`, `frRowBookmark`, `frSetValue`, `frAllowUpdate`. `frSetValue` exists
because the public `Value` Let cannot carry an object and a cell can hold one,
and `frAllowUpdate` is what opens that Let on the one buffer meant to be filled.
`RecordCount` is not filled at all: it forwards to the model, because the
original raises on it once the order has been rebuilt. `frGroupPrefixed` is
filled beside the caption, so the renderer can indent a prefixed group caption
one space less without re-deriving which column the group is on.

`frInitModel` also stamps the wrapper with the projection generation it was
filled under. `RecordCount`, `GetSubTotal`, `GetBookmarks` and
`GetRowIndexes` pass that stamp back, and a group row whose order has since
been rebuilt raises `vbObjectError + 129`, *"JSRowData object may have
changed. The object is no longer valid."* -- the original's own number and
message. Record wrappers never check, and go on reading the row they were
filled with. `UcsRowSet.Version` is bumped by `DataProjectRows` alone, so a
collapse leaves a held header working, which is what the original does.

### FetchData

A column with `JSColumn.FetchData` set is supplied by the client through the
`FetchData` event, and that **outranks whatever else the column could have been
read from** -- its `DataField` in a bound grid, the values `UnboundReadData`
wrote in an unbound one. Both models treat it identically, and every consumer
goes through the same cell accessor, so sorting, grouping, totals and painting
all see the same value.

The cache is per cell and fills on demand: a repaint asks for the screenful it
is painting. The only thing that ever sweeps a whole column is a sort, because
a sort genuinely needs every record's key -- and what it leaves behind is a
warm cache rather than a second pass. Everything narrower invalidates by the
row: an edit or an insert refetches that record's fetch columns and marks the
order stale only if a value actually moved; a delete shifts the rows above it
down; only a rebind or a column change throws the lot away.

### Bookmarks

Bookmarks are keyed by RowIndex and sized to `ItemCount`. Behaviour recorded
from the original rather than assumed:

- **duplicates are legal.** Assigning a bookmark another record already holds
  raises nothing, and every resolver -- `MoveToBookmark`, `AddBookmark`,
  `RefreshRowBookmark` -- answers with the **lowest** RowIndex.
- **an unknown bookmark raises** `vbObjectError + 119`, *"Not a valid
  Bookmark."*, leaving the current row alone. So `GetRowIndex` answering 0 is
  the model's way of saying that.
- **they survive both `Refetch` and `Rebind`.** Nothing but a delete or an
  explicit assignment drops one.
- **a delete compacts the store.** Deleting the record holding `"bk-3"` leaves
  the next read of RowIndex 3 carrying `"bk-4"`.
- **type is part of identity, width is not.** Numeric widths resolve against
  each other, and `Boolean` and `Date` resolve against their numeric value, but
  a string never matches a number. Hence `DataGetBookmarkKey`: `"S" & value` for
  strings, `"#" & C2Dbl(value)` for everything else, and the hex of the bytes
  for the byte-array bookmarks a client-side ADO cursor hands out.

The map is a `VBA.Collection` rebuilt lazily rather than patched, because a
delete renumbers every record above it and there is no cheap way to find the
affected entries.

### Weak references

The control holds its model; the model holds the control **weakly**, as a raw
pointer written with `CopyMemory` and zeroed whenever the holder is detached:
its own `Class_Terminate`, the control's teardown -- and, for collection items,
`Remove` and `Clear`, which detach what they drop on the way out. Every class
in the object model that points back at the control follows the same
discipline, because releasing a control that was never `AddRef`-ed takes the
process down, and a `JSGroup` a client kept across `Groups.Clear` used to be
exactly that crash.

`JSRowData` is the exception that needs none of this: it holds no pointer to
the control, weak or otherwise, so there is nothing to zero and nothing to
dangle. Its only reference is to the model, and that one is strong -- which is
why the control calls `IDataModel.Terminate` before releasing a model rather
than leaving it to `Class_Terminate`. A wrapper a client kept can outlive the
grid and would otherwise keep a model pointing at a dead control.

The corollary for buffers: a buffer the *grid* owns may hold a model
reference, but a buffer the *model* owns must not, or the two keep each other
alive. `cUnboundDataModel`'s shared fetch buffer is passed `Nothing` for
exactly this reason.

### How the grid follows a reprojection

The current row and the selection stay on their *records*, not their positions,
because a re-sort or a collapse moves every position. The control holds the
current row's RowIndex -- the one address space a reprojection cannot move --
and resolves it back through `GetRowPosition`.

Knowing *when* to do that is what `Version` is for. The model bumps it in
`DataProjectRows` and nowhere else, so the control compares it on entry to the
paths that matter and remaps when it moves. A collapse is the exception: it
reprojects through `DataBuildVisibleRows`/`DataWritePositions` without re-sorting,
and deliberately leaves `Version` alone because a held group wrapper has to
survive one. The collapse paths therefore remap directly rather than waiting
for the version watch.

A bind is the other exception. `Rebind` places the marquee itself, so it drops
the held index first -- otherwise the sync that follows would put the row back
on whatever record it was on before the bind.

### Status

**The control draws from the model.** `GridEX.ctl` owns no rows, no order, no
group rows and no aggregates: it asks for a page of `JSRowData` and paints from
that. The projection, the row storage, the sort, `DataAggregateValues` and all 26
`frRow*` accessors are gone, along with `JSRowData`'s view mode -- about 810
lines out of the control.

Known gaps at the time of writing:

- an edit commit suppresses `UnboundUpdate` through a flag on the control
  rather than an argument on `UpdateRowData`. The original reports only
  `AfterUpdate` for a commit, and the interface has no way to say "write this
  without notifying"
- `cAdoDataModel` has no test coverage of any kind: every model test runs
  against the unbound implementation
- `SelectedItems` answers when `MultiSelect` is False where the original raises
  `vbObjectError + 123` -- a divergence predating the data model

## Window structure and subclassing

*Stub.* The control is an outer `UserControl` holding an inner `picGrid`
PictureBox and a band carrying `hsbGrid` and the painted record navigator --
mirroring the original's own window tree. Two windows are subclassed through
the Modern Subclassing Thunk: the grid surface for painting, input and
`WM_VSCROLL`, and the outer control because it is the band's parent. Both
answer `WM_PAINT` and `WM_ERASEBKGND` themselves (see **Painting**).

Worth covering: why `hWnd` returns `picGrid.hWnd`, why the horizontal
scrollbar is a child control while the vertical one is a non-client style, and
the `InitAddressOfMethod` hidden-member trick that keeps the callback off the
typelib.

## Painting

GDI drawing, band by band, each band buffered. Bands from the top: the group-by
box with its chip staircase, the column headers, one band per row, and the
scrollbar band on the outer control carrying the painted record navigator.

### Who owns WM_PAINT

The subclass does, on both windows. VB's `Paint` event is not used: `BeginPaint`
hands out the DC the surface is painted with *and* `rcPaint`, and the second of
those is worth as much as the first -- a band the update region does not reach
is stepped over rather than drawn for GDI to clip away. A drag over the headers
invalidates the strip above the rows and allocates one bitmap; the rows are
never touched.

Owning the message means owning what VB used to do before raising the event:
**it filled the client with `BackColor`**. That fill is where the pixels no
painter covers came from -- the tree indent inside the header band, the band
beside the navigator -- and suppressing `WM_ERASEBKGND` never removed it. Each
band now lays its own background down as its first act.

### The buffer

One bitmap carries the whole paint. `pvInitBuffer` allocates it as tall as the
tallest band that will go through it -- the group-by box, the header strip, or a
row plus the line it shares with the row above -- and `pvFreeBuffer` frees it
once everything has been painted. Between them `pvSetBufferBand` puts it over the
band about to be drawn and `pvFlushBuffer` blits that band out. A client-sized
bitmap held between paints is a lot of memory for a strip of rows; a band-sized
one held for the length of one paint is not, and it saves a create/select/delete
per row. `pvInitBuffer` hands back the caller's own DC if the allocation fails,
so failure degrades to painting direct rather than to not painting.

Painted straight onto the window, a band shows every stage it goes through --
the fill, then what is drawn over it. That is what the eye reads as flicker, and
it is why the unit is the band rather than the whole surface: the bitmap stays
row-sized and a repaint that only dirties one strip only buffers one strip.

`pvSetBufferBand` sets two things per band. `SetViewportOrgEx(0, -lTop)` keeps every
painter addressing the client area, and a clip box the size of the band keeps
what they draw inside it -- a rule the height of a row ends at that row's edges
rather than running the length of the bitmap.

The **first line** of a band is the line a row shares with the row above, where
a group row draws its opening rule; every band covers the rest of itself before
painting anything. That line is never read back off the surface: a band that
carries straight on from the one before it scrolls it up out of the bitmap,
which still holds what the previous band left there. When the previous band was
not painted this pass the line is stale -- and provably outside the update
region, because a clip that reaches the last line of a band reaches the band,
so `BeginPaint`'s own clip drops whatever goes there.

### A row is a band

The rows loop takes the height per row and accumulates the top, so rows of
differing heights need nothing but a taller bitmap. Empty rows are rows in the
same loop. The strip right of the last column belongs to the row beside it and
arrives in the same blit; only the background below the block is left to a plain
`FillRect`.

Two block-wide passes used to follow the rows -- the focus marquee, then the
column rules over it. Both moved into the row, in the order they had, because a
band that is blitted has to arrive finished. The rules still close over the
cells, the horizontal rule and the marquee's dots; a per-row segment reproduces
a block-long line because solid is phase-free, dotted stamps on absolute parity,
and a dashed rule restarts its phase at every row top anyway -- `pvDrawLine` anchors
`pvDrawStampedLine` there.

What reaches outside a row is worth knowing, because each item is a pixel bug
waiting to happen: a group row's opening rule lands on the row above (hence the
seeded line), and the row header's border pair closes one line *below* its row
under vertical-only gridlines -- so the row does not draw it and the block draws
the last one after its rows.

Fills stop one pixel short of a vertical gridline rather than painting over it
and letting the rule be drawn again. Clearing a rule and redrawing it is
invisible in a still image and very visible in motion.

### Pixel-exactness

The rules below were all found by golden diffing rather than reasoning, and each
one is load-bearing:

- text is inset asymmetrically -- two pixels in, three at the clip
- the XOR marquee is a checkerboard anchored **per column**, not per row: a
  single row-wide `DrawFocusRect` only agrees with the original on even column
  boundaries
- dotted rules are stamped pixel by pixel on absolute parity, because GDI's
  cosmetic `PS_DOT` renders 3-on/3-off where the original is 1-on/1-off
- dashed rules restart their phase at every row top, so a row height that is not
  a multiple of six leaves a run of four across the boundary
- `IntersectClipRect` + `DT_NOCLIP` is what makes ClearType fringes match:
  letting `DrawText` clip changes the fringe pixels at the boundary
- the row header's border pair lands by an empirical table, one entry per
  `GridLines` value, and an empty row's cell closes one line short of itself
  whichever mode is set

Chrome is not the grid's palette. The record navigator takes the control's own
background and the system's button text, whatever `BackColorHeader` and
`ForeColorHeader` are set to, and the header band's tree indent shows the
window's background rather than any colour property -- both probed by setting a
distinct colour on every candidate and reading the original's pixels.

## Metrics and DPI

*Stub.* Twips-facing properties stored internally in pixels and snapped on
set, exactly as the original does. Heights derived from font metrics rather
than literals, so they hold at 96, 120 and 144 dpi.

Worth covering: the `tmHeight + 3` / `tmHeight + 6` rules, the OLE font-size
snapping that makes a requested 8.25 realize as 8 at 144dpi, and the handful of
metrics that are genuinely fixed pixel counts at every scale.

## Input and editing

*Stub.* Keyboard, mouse and scrolling arrive through the subclass. Editing
runs an in-place native `EDIT` created per session, with a pending row buffer
between the editor and storage.

The header carries three gestures off the same press, told apart on the button
*up*: a divider grabbed first is a resize, a pointer that has left the system
double-click box around the press is a move, and anything else is the click that
sorts. A move dragged past either edge scrolls the strip a column at a time on a
timer -- a pointer that has stopped moving sends nothing -- and the drop stays
cancelled while it is out there.

Worth covering: the event contracts recorded from the original (the update
trio, two-level Escape, where `SelectionChange` sits relative to
`RowColChange`), why the pending buffer stays in the control rather than moving
into the data model, and the `ColResize` contract -- raised once on the release,
with `Column.Width` still the width the drag started from.

## Object model

*Stub.* 23 `JS*` classes reproducing the original's collections and value
objects, with `Friend` wiring and weak owner references throughout.

Worth covering: the collection idiom (keyed, reindexing, lazy-created owned
objects), the protected built-in `FormatStyles`, and the `JSRet*` carriers.

## Persistence and property pages

*Stub.* Not implemented. `WriteProperties`/`ReadProperties` routed through the
same friend setters as the runtime path, plus `.pag` property pages.

## Printing

*Stub.* Not implemented. `JSPrinterProperties`, pagination, `PrintGrid`, and
the `PrintPreview` handshake with `GEXPreview`.

## Build and compatibility

*Stub.* `src/make.bat` builds the OCX under `CompatibleMode=2` against
`OpenGridEX20.cmp`. `tools/CompareIdl.ps1` diffs an OleView dump of the built
control against `doc/GridEX20.idl` and must report zero differences;
`tools/DumpSurface.vbs` does the same job through TLI.

Worth covering: why private classes and `Friend` members are free, why
`Public Enum` in a `.bas` is not, and the append-only rule on public members.

## Testing

*Stub.* Four harnesses:

- `test/ModelTests` -- object model, collections, events, snapshot round-trips
- `test/VisualDiff` -- pixel and event-log diffing against goldens recorded
  from the original control: 83 scenarios at 96, 120 and 144 dpi
- `test/Samples` -- the ported original samples under a smoke runner
- `test/Snapshot` -- smoke: both controls dump their object model to JSON

Worth covering: the record/verify/selftest cycle, why goldens live per dpi,
what the `.events.txt` logs deliberately leave out, and the practice of probing
the original for behaviour rather than guessing it.
