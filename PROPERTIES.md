# GridEX property inventory

Every public property of the control and its object model -- **288** of them
across `GridEX`, `GEXPreview` and the 23 `JS*` classes -- read out of the
original's type library (`doc\GridEX20.idl`), so the list is the surface itself
rather than whatever happens to be implemented.

The document is in three parts:

1. **Properties that paint** -- the 126 that change pixels, each with its
   verification status against the original. This is the M7 gate and the bulk of
   the file.
2. **The non-paint matrix** -- the other 162 in the same shape: data, binding,
   editor, sorting, printing and drag/drop members that never reach the
   renderer, each with the milestone that owes it behaviour.
3. **The object model** -- one table per class, every property, including those
   already carried by the paint matrix.

## Status at a glance

| status | count | meaning |
|---|---:|---|
| paint | 126 | changes pixels -- part 1 carries its verification state |
| consumed | 71 | the engine reads it: it drives data, layout or events |
| storage | 85 | stored, returned and round-tripped; nothing reads it yet |
| derived | 6 | no member of its own -- computed on the fly, or parameterized |
| **total** | **288** | |

**storage is not a stub.** Every one of those members holds its value, survives
`ImportObject`/`SnapshotToJson` and is proven by the corpus round-trip in
`test\ModelTests` -- what is missing is behaviour behind it, and that behaviour
belongs to a milestone the roadmap names. The line this table draws is between a
property the control *acts on* and one it merely *remembers*.

## Part 1 -- Properties that paint

Every property that affects what the control paints -- the direct `GridEX`
members plus the `JSColumn` and `JSFormatStyle` sub-properties -- in one
alphabetical list, whether or not the renderer reads it yet.

A property belongs here when it changes what pixels appear in the client area.
Pure data and behaviour members (`DataMode`, `Connect`, `AllowDelete`,
`JSColumn.DataField`, `JSColumn.Key`, ...) live in parts 2 and 3 instead.

### The verification bar

A property counts as **verified** only when the golden corpus renders it at **two
or more distinct values** and every one is pixel-identical to the original. That
bar matters: a property can be correctly wired into the paint path and still be
completely unproven, either because no scenario sets it or because the scenario
that sets it cannot render the affected element. `GridLinesColor` is the worked
example -- consumed by `pvLine` since M3c and set to `0x0000FF` by
`002-gridlines-dots-colors`, yet that scenario declares **no rows**, so no data
gridline is ever drawn and the golden's full colour histogram
(`808080`, `F0F0F0`, `A0A0A0`, `FFFFFF`, `696969`, `000000`) holds no blue pixel.

| status | meaning |
|---|---|
| **verified** | consumed by the paint path, >= 2 distinct values rendered and pixel-matched |
| **weak** | consumed and rendered, but only one value (often just the default) is covered |
| **unverified** | consumed by the paint path, no scenario renders a non-default value |
| **partial** | partially consumed -- some of its visual effect is missing |
| **not impl** | the paint path never reads it; the property only stores its value |
| n/a | state the renderer keeps but never paints -- listed so its absence is on record |

### Summary of part 1

| | count | share |
|---|---:|---:|
| verified | 49 | 39% |
| weak | 0 | 0% |
| unverified | 0 | 0% |
| partial | 0 | 0% |
| n/a (state only) | 1 | 1% |
| **not implemented** | **76** | **60%** |
| **total** | **126** | |

**50 of 126 (40%) are read by the paint path**; the other **76 (60%) are not
implemented** -- they store and return their value, and round-trip through the
snapshot corpus, but the renderer never looks at them.

**Every property the renderer reads is proven** against the original at two or
more distinct values. Two of them cannot be proven by a picture and are pinned
by `ModelTests` instead: `ContinuousScroll` only shows up mid-drag, and `Redraw`
suppresses painting rather than changing it. `Col` is state with no pixels of
its own -- it is listed to record that, not as a gap. The 76 that remain are
unimplemented, each owned by a later milestone.

Getting there took eight scenarios and turned up **eleven** real defects. Each had
survived because the corpus could not see it: solid gridlines hide both dotted
rules, a black default `ForeColor` makes a wrong XOR mask invisible, and no
scenario had ever set `GridLines` to anything but its default.

Not-implemented properties by owning milestone:

| milestone | count |
|---|---:|
| M10 styling extras | 38 |
| M4 sorting/grouping | 16 |
| M6 card view | 12 |
| M5 editing | 5 |
| **unowned** | **5** |

M3d has none left: `Col`, `ColumnAutoResize`, `ContinuousScroll`, `FrozenColumns`
and `Redraw` closed the milestone.

**M10 is by far the largest single block of unimplemented painting** -- 38
properties, mostly the `JSFormatStyle` font and picture families -- which is a
much bigger milestone than its one-line roadmap entry suggests.

Card view was the other gap this matrix exposed and now has its own milestone
(M6), placed before the paint gate because it *adds* twelve paint properties.
Five remain unowned: the drag/resize affordances (`AllowColumnDrag`,
`AllowRowSizing`, `DetectRowDrag`, `JSColumn.AllowSizing`) and `Options`.

### The matrix

Scenario names refer to `test\VisualDiff\scenarios\NNN-*.json`, numbered in
creation order; all 34 are verified at 96 and 120 dpi, and the first 28 at 144 as
well. The `Commit` column records where the property started affecting pixels,
not where it was first stored -- every member was declared in the M2 storage
commit, so that hash carries no information about painting.

| Property | Type | Status | Milestone | Commit | Test |
|---|---|---|---|---|---|
| `AllowAddNew` | `Boolean` | **not impl** | M5 | -- | -- |
| `AllowCardSizing` | `Boolean` | **not impl** | M6 | -- | -- |
| `AllowColumnDrag` | `Boolean` | **not impl** | unowned | -- | -- |
| `AllowRowSizing` | `Boolean` | **not impl** | unowned | -- | -- |
| `AutomaticArrange` | `Boolean` | **not impl** | M6 | -- | -- |
| `BackColor` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (`0xC0FFC0`) vs default |
| `BackColorBkg` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (`0x404040`), `002-gridlines-dots-colors` (`0x808080`) |
| `BackColorGBBox` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x604020`) vs default |
| `BackColorHeader` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x8000FF`) vs default |
| `BackColorInfoText` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x30A0C0`) vs default |
| `BackColorRowGroup` | `OLE_COLOR` | **not impl** | M4 | -- | -- |
| `BorderStyle` | `jgexBorderStyleConstants` | **not impl** | M10 | -- | -- |
| `CardBorders` | `Boolean` | **not impl** | M6 | -- | -- |
| `CardCaptionPrefix` | `String` | **not impl** | M6 | -- | -- |
| `CardSpacing` | `Long` | **not impl** | M6 | -- | -- |
| `CardWidth` | `Long` | **not impl** | M6 | -- | -- |
| `Col` | `Integer` | n/a | M3d | -- | `ModelTests` clamps it to the column range |
| `ColumnAutoResize` | `Boolean` | verified | M3d | -- | `032-column-autoresize` (True, 900/1500/600tw -> 113/189/76px) vs default False |
| `ColumnHeaderFont` | `Font` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `009-font-large`, `020-font-tahoma`, `019-font-segoeui` + MS Sans Serif default |
| `ColumnHeaderHeight` | `Long` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `003-headers-flat` sets 400tw; font-derived defaults elsewhere |
| `ColumnHeaders` | `Boolean` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `004-no-headers-no-rowheaders` (False) vs default True |
| `Columns` | `JSColumns` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | all 30 scenarios -- see the `JSColumn.*` rows |
| `ContinuousScroll` | `Boolean` | verified | M3d | -- | `ModelTests` at both values -- a thumb drag has no static picture to golden |
| `DefaultColumnWidth` | `Long` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `024-default-column-width` (2100tw -> 140px columns) vs default |
| `DefaultGroupMode` | `jgexDefaultGroupModeConstants` | **not impl** | M4 | -- | -- |
| `DetectRowDrag` | `Boolean` | **not impl** | unowned | -- | -- |
| `EditMode` | `jgexEditModeConstants` | **not impl** | M5 | -- | -- |
| `EmptyRows` | `Boolean` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `008-empty-rows` (True) vs default False |
| `Enabled` | `Boolean` | **not impl** | M10 | -- | -- |
| `FirstItem` | `Long` | verified | M3d | [`e95c15e`](../../commit/e95c15e1b1d35f09c9bbe557228ddc1329ed6c6f) | `016-scrolled` sets 4 in its `post` block; 1 elsewhere |
| `FmtConditions` | `JSFmtConditions` | **not impl** | M10 | -- | -- |
| `Font` | `Font` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `009-font-large`, `020-font-tahoma`, `019-font-segoeui` + default |
| `ForeColor` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (`0xFF0000`) vs default -- also drives the marquee XOR mask |
| `ForeColorHeader` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x00FFFF`) vs default |
| `ForeColorInfoText` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x0020FF`) vs default |
| `ForeColorRowGroup` | `OLE_COLOR` | **not impl** | M4 | -- | -- |
| `FormatStyles` | `JSFormatStyles` | verified | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `025-formatstyles-selection` overrides the `SelectedRow` system style |
| `FrozenColumns` | `Integer` | verified | M3d | -- | `031-frozen-columns` (2, pinned under `LeftCol` = 4) vs 0 elsewhere |
| `GridImages` | `JSGridImages` | **not impl** | M10 | -- | -- |
| `GridLines` | `jgexGridLinesConstants` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | all four modes: default Both, plus `027-gridlines-none`, `028-gridlines-vertical`, `026-gridlines-horizontal` |
| `GridLinesColor` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (`0x0000FF`) vs default |
| `GridLineStyle` | `jgexGridLineStyleConstants` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (dots), `028-gridlines-vertical` (dashes), solid default |
| `GroupByBoxInfoText` | `String` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` custom text vs default |
| `GroupByBoxVisible` | `Boolean` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `010-headers-noborder`, `011-headers-single3d`, `012-headers-singleflat` vs default |
| `GroupFooterStyle` | `jgexGroupFooterStyleConstants` | **not impl** | M4 | -- | -- |
| `Groups` | `JSGroups` | **not impl** | M4 | -- | -- |
| `HeaderStyle` | `jgexHeaderStyleConstants` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `003-headers-flat`, `010-headers-noborder`, `011-headers-single3d`, `012-headers-singleflat` (4 values) |
| `HideSelection` | `Boolean` | **not impl** | M10 | -- | -- |
| `ImageHeight` | `Integer` | **not impl** | M10 | -- | -- |
| `ImageWidth` | `Integer` | **not impl** | M10 | -- | -- |
| `ItemCount` | `Long` | verified | M3a | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | row counts 2/3/5/6/14/16 across the corpus |
| `JSColumn.AggregateFunction` | `jgexAggregateFunctionConstants` | **not impl** | M4 | -- | -- |
| `JSColumn.AllowSizing` | `Boolean` | **not impl** | unowned | -- | -- |
| `JSColumn.ButtonStyle` | `jgexButtonStyleConstants` | **not impl** | M5 | -- | -- |
| `JSColumn.Caption` | `String` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | distinct captions in every scenario |
| `JSColumn.CardCaption` | `Boolean` | **not impl** | M6 | -- | -- |
| `JSColumn.CardIcon` | `Boolean` | **not impl** | M6 | -- | -- |
| `JSColumn.CellStyle` | `String` | **not impl** | M10 | -- | -- |
| `JSColumn.ColPosition` | `Integer` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `023-column-order` reorders 3 columns to 1/2/3 from declaration order |
| `JSColumn.ColumnType` | `jgexColumnTypeConstants` | **not impl** | M10 | -- | -- |
| `JSColumn.DefaultIcon` | `Integer` | **not impl** | M10 | -- | -- |
| `JSColumn.EditType` | `jgexEditTypeConstants` | **not impl** | M5 | -- | -- |
| `JSColumn.Format` | `String` | **not impl** | M10 | -- | -- |
| `JSColumn.GroupEmptyStringCaption` | `String` | **not impl** | M4 | -- | -- |
| `JSColumn.GroupFormat` | `String` | **not impl** | M4 | -- | -- |
| `JSColumn.GroupPrefix` | `String` | **not impl** | M4 | -- | -- |
| `JSColumn.HeaderAlignment` | `jgexAlignmentConstants` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `003-headers-flat` sets 2; default 0 elsewhere |
| `JSColumn.HeaderIcon` | `Integer` | **not impl** | M10 | -- | -- |
| `JSColumn.HeaderStyle` | `String` | **not impl** | M10 | -- | -- |
| `JSColumn.HeaderToolTip` | `String` | **not impl** | M10 | -- | -- |
| `JSColumn.IsGrouped` | `Boolean` | **not impl** | M4 | -- | -- |
| `JSColumn.MaxRowsInCardView` | `Long` | **not impl** | M6 | -- | -- |
| `JSColumn.MinRowsInCardView` | `Long` | **not impl** | M6 | -- | -- |
| `JSColumn.ShowCaptionInCardView` | `Boolean` | **not impl** | M6 | -- | -- |
| `JSColumn.SortOrder` | `jgexSortOrderConstants` | **not impl** | M4 | -- | -- |
| `JSColumn.SortType` | `jgexSortTypeConstants` | **not impl** | M4 | -- | -- |
| `JSColumn.TextAlignment` | `jgexAlignmentConstants` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | 1 in `006-unbound-rows`, 2 in `019-font-segoeui`, default 0 elsewhere |
| `JSColumn.TotalRowFormat` | `String` | **not impl** | M4 | -- | -- |
| `JSColumn.TotalRowPrefix` | `String` | **not impl** | M4 | -- | -- |
| `JSColumn.Visible` | `Boolean` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `013-hidden-column` (False) vs default True |
| `JSColumn.Width` | `Long` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | 900/1200/1400/1500/1800tw across the corpus |
| `JSColumn.WordWrap` | `Boolean` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.BackColor` | `OLE_COLOR` | verified | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `025-formatstyles-selection` (`0x004080`) vs the system default |
| `JSFormatStyle.BackgroundPicture` | `Picture` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.DrawModeBackGroundPicture` | `jgexDrawModePictureBackgroundConstants` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.FontBold` | `Boolean` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.FontCharset` | `Integer` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.FontItalic` | `Boolean` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.FontName` | `String` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.FontSize` | `Currency` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.FontStrikeThru` | `Boolean` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.FontUnderline` | `Boolean` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.ForeColor` | `OLE_COLOR` | verified | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `025-formatstyles-selection` (`0x80FFFF`) vs the system default |
| `JSFormatStyle.Picture` | `Picture` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.PictureDrawMode` | `jgexPictureDrawModeConstants` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.PictureHorzAlignment` | `jgexHorzPictureAlignmentConstants` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.PictureVertAlignment` | `jgexVertPictureAlignmentConstants` | **not impl** | M10 | -- | -- |
| `JSFormatStyle.TextAlignment` | `jgexAlignmentConstants` | **not impl** | M10 | -- | -- |
| `LeftCol` | `Integer` | verified | M3d | [`e95c15e`](../../commit/e95c15e1b1d35f09c9bbe557228ddc1329ed6c6f) | `029-hscrolled` (3), `034-hscroll-middle` (2), `031-frozen-columns` (4); column 1 elsewhere |
| `MaskColor` | `OLE_COLOR` | **not impl** | M10 | -- | -- |
| `MultiSelect` | `jgexMultiSelectConstants` | verified | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `017-multiselect` (contiguous 2-4) and `028-gridlines-vertical` (disjoint 1/3/6) |
| `NewRowPos` | `jgexNewRowPositionConstants` | **not impl** | M5 | -- | -- |
| `Options` | `Long` | **not impl** | unowned | -- | -- |
| `PreviewColumn` | `Integer` | **not impl** | M10 | -- | -- |
| `PreviewRowIndent` | `Long` | **not impl** | M10 | -- | -- |
| `PreviewRowLines` | `Integer` | **not impl** | M10 | -- | -- |
| `RecordNavigator` | `Boolean` | verified | M3d | [`f9acd79`](../../commit/f9acd792b194f87374c89c6c2af0789785ec5e8b) | `030-record-navigator` (True) vs default False |
| `RecordNavigatorString` | `String` | verified | `pvNavLayout` | M3d | [`f9acd79`](../../commit/f9acd792b194f87374c89c6c2af0789785ec5e8b) | `033-record-navigator-string` (`Row:\|from a total of`) vs the default `Record:\|of` |
| `Redraw` | `Boolean` | verified | M3d | -- | `ModelTests`: changes under False paint once when it goes back True |
| `Row` | `Long` | verified | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `026-gridlines-horizontal` sets `Row` = 3 via `post`; row 1 elsewhere |
| `RowColorEven` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `006-unbound-rows` and `027-gridlines-none` (`0xE0E0FF`) |
| `RowColorOdd` | `OLE_COLOR` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `006-unbound-rows` and `027-gridlines-none` (`0xFFE0E0`) |
| `RowCount` | `Long` | verified | M3a | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | read-only view of `ItemCount`; the dark last-row gridline depends on it |
| `RowExpanded` | `Boolean` | **not impl** | M4 | -- | -- |
| `RowHeaders` | `Boolean` | verified | M3c | [`54c6101`](../../commit/54c6101c20939be5df126be4f4b84e1c541926bb) | `014-rowheaders`, `017-multiselect`, `019-font-segoeui` vs `004-no-headers-no-rowheaders` |
| `RowHeight` | `Long` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | never set explicitly -- exercised indirectly at 19/22/24/28/32px via the font scenarios x 3 dpi |
| `RowSelected` | `Boolean` | verified | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `017-multiselect` (contiguous) and `028-gridlines-vertical` (disjoint) |
| `ScrollToolTipColumn` | `Integer` | **not impl** | M10 | -- | -- |
| `ScrollToolTips` | `Boolean` | **not impl** | M10 | -- | -- |
| `SelectedItems` | `JSSelectedItems` | verified | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `017-multiselect` (contiguous) and `028-gridlines-vertical` (disjoint) |
| `SelectionStyle` | `jgexSelectionStyleConstants` | **not impl** | M10 | -- | -- |
| `ShowEmptyFields` | `Boolean` | **not impl** | M10 | -- | -- |
| `ShowToolTips` | `Boolean` | **not impl** | M10 | -- | -- |
| `SortKeys` | `JSSortKeys` | **not impl** | M4 | -- | -- |
| `UseEvenOddColor` | `Boolean` | verified | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `006-unbound-rows` (True) vs default False |
| `View` | `jgexViewConstants` | **not impl** | M6 | -- | -- |

### Closing the gaps

The unverified block is **cleared**. Five scenarios did it, and each defect they
found had been invisible for a structural reason worth remembering: a rule can
only be wrong in a way the corpus can see.

| scenario | closed | found |
|---|---|---|
| `022-colors-rows` | `BackColor`, `ForeColor`, `GridLinesColor`, `GridLineStyle`, `BackColorBkg` | dotted-pen duty cycle, dotted-gap colour, marquee XOR mask |
| `021-colors-chrome` | the 6 header / group-by-box / info-text colours | -- |
| `025-formatstyles-selection` | `FormatStyles`, `JSFormatStyle.BackColor`, `JSFormatStyle.ForeColor` | -- |
| `023-column-order` | `JSColumn.ColPosition` | invalid `ColPosition` blanks the grid |
| `024-default-column-width` | `DefaultColumnWidth` | new columns ignored it, defaulting to 1000 **px** |

The three `gridlines-*` scenarios each carry 5-6 properties rather than one, so
clearing the weak block cost three captures instead of seven -- recording is the
slow part of the loop, so scenarios are packed deliberately.

The corpus is recorded and verified at **96 and 120 dpi** throughout. `golden/144`
covers the first 28 scenarios only: everything from `029-hscrolled` on was
recorded after the machine left 150%, so the scrollbar band, the navigator,
frozen columns and column auto-resize have no 150% golden yet. That pair is also the third data point the
navigator's box-width term (`tmHeight + BandH + 16`) still wants -- two scales
cannot separate a constant from a metric-derived one.

Note on goldens and the system accent: it changed twice during this work
(`0x0078D7` -> `0x0078D4` -> back), and every selected row carries it. A golden
set recorded in an earlier session can therefore fail on colour alone while the
geometry is perfect -- the tell is a diff whose only pairs are the accent and its
XOR complement. Re-record rather than debug. `golden/144` currently holds the
older accent and wants a re-record next time the machine is at 150%.

M3d closed with four more scenarios (`031`-`034`), and each turned up a rule that
no amount of reading the docs would have produced:

- **`ColumnAutoResize` scales proportionally, not equally.** 900/1500/600tw over a
  378px client comes out 113/189/76px -- the boundaries land on the rounded
  running total, so the last column absorbs the rounding.
- **`FrozenColumns` set before the first layout is dropped by the original.** The
  same scenario applied through the `post` block pins the columns as documented;
  applied through `props` (before Show) it does nothing at all. The scenario uses
  `post` because that is the behaviour worth matching.
- **`LeftCol` clamps at the last full page**, `colCount - fit + 1`, counted over
  the *scrollable* strip: with two frozen columns eating 200 of 378px, `LeftCol`
  = 4 stands, without them it clamps to 3.
- **The horizontal thumb sits a pixel off if VB6 is left to place it.** VB6 maps a
  scrollbar's `Min`..`Max` onto 0..32767 before handing it to Windows, and the
  thumb computed from that lands one pixel left of the original's at any middle
  position (ends match, which is why this survived until a scenario scrolled to
  the middle). `GetScrollInfo` on the original's bar shows it keeps the column
  numbers themselves -- `min=3 max=5 page=1 pos=4` for `031` -- so the fix is a
  `SetScrollInfo` over VB6's mapping. Reading the original's window back beats
  inferring geometry from pixels, and `scrollinfo` mode in the harness now does
  exactly that for any scenario.

M7 is complete. Horizontal scrolling landed with it, since `LeftCol` was the one
property the renderer only partly consumed. The scrollbar band that came out of
it brought two more paint properties into the matrix: `RecordNavigator`, verified
against the original at both values, and `RecordNavigatorString`, which is the
only weak entry left -- the corpus renders it at its default `Record:|of` and
nothing else, so the prefix/middle text metrics that drive the whole navigator
layout are unproven at any other string.

Both harness scripts take an optional scenario mask -- `make.bat 029*`,
`record.bat *gridlines-*` -- which cuts a single-scenario check to about a second
against a couple of minutes for the corpus. The mask is matched with VB `Like`
against the whole file name, so it has to allow for the `NNN-` prefix. Recording
is the slow half of the loop, so scenarios are also packed 5-6 properties each.

One rule is recorded as an empirical table rather than a derivation: where the
row header's border pair lands per `GridLines` value (`Vertical` +1,
`Horizontal` -1, `Both`/`None` unchanged). Four goldens agree, but no mechanical
reason was found for the single-direction modes pulling in opposite directions,
so treat it as measured, not understood.

### Maintenance

Update this file in the same commit that changes a property's status. Parts 2
and 3 are mechanical -- the property list comes from `doc\GridEX20.idl`, type and
access from the same declaration, and the status from whether anything in `src`
reads the member behind the accessor -- so a property appears here as soon as it
exists in the type library, implemented or not.

Two traps when regenerating it mechanically: paint routines read some values
through their **public property getter** rather than the member (`RowCount`,
`SelectedItems`), and some getters **compute** instead of returning a member
(`RowHeight`, `ColumnHeaderHeight`, `DefaultColumnWidth`). A member-only scan
misses both groups. Scenario coverage likewise lives in **two** places -- the
`props` block and the `post` block -- and `FirstItem` is only ever set in `post`.

## Part 2 -- The non-paint matrix

The other 162 properties in the same shape as the paint matrix: everything the
control and its object model expose that never reaches the renderer -- the data
source and binding families, the cell editor, sorting and grouping state,
printing, drag/drop and layout persistence.

The bar is different here because pixels are not the evidence. **consumed** means
something in `src` reads the value back and acts on it, **storage** means it is
held, returned and round-tripped but nothing reads it yet, **derived** means there
is no member behind it at all -- computed on demand, or parameterized like
`RowSelected(RowPosition)`. That comes to 71 consumed, 85 storage, 6 derived.

`Milestone` is filled only for `storage` rows, since those are the ones still
owed behaviour -- M4 sorting/grouping (5), M5 editing (15), M8 ADO binding (17),
M10 styling (14), M11 printing (33), M12 long tail (1). `Commit` points at the
commit that introduced the routine doing the consuming, and `Test` at what covers
the property today: `snapshot round-trip` means the corpus proves the value
survives import and export, which is all a storage property can be asked to
prove.

| Property | Type | Status | Milestone | Commit | Test |
|---|---|---|---|---|---|
| `ActAsDropDown` | `Boolean` | consumed | -- | -- | snapshot round-trip |
| `ADORecordset` | `Object` | storage | M8 | -- | -- |
| `AllowDelete` | `Boolean` | storage | M5 | -- | `Samples` frmUnboundArray |
| `AllowEdit` | `Boolean` | storage | M5 | -- | snapshot round-trip |
| `AutomaticSort` | `Boolean` | storage | M4 | -- | snapshot round-trip |
| `BoundColumnIndex` | `Variant` | consumed | -- | -- | -- |
| `CalendarNoneText` | `String` | storage | M5 | -- | snapshot round-trip |
| `CalendarTodayText` | `String` | storage | M5 | -- | snapshot round-trip |
| `Connect` | `String` | storage | M8 | -- | snapshot round-trip |
| `CursorLocation` | `jgexCursorLocationConstants` | storage | M8 | -- | snapshot round-trip |
| `DatabaseName` | `String` | storage | M8 | -- | snapshot round-trip |
| `DataChanged` | `Boolean` | storage | M8 | -- | -- |
| `DataMode` | `jgexDataModeConstants` | consumed | -- | [`577e39b`](../../commit/577e39b83d084cbd5105745ef66db1942755c4fe) | `ModelTests` pvTestUnbound |
| `ErrorText` | `String` | storage | M8 | -- | -- |
| `Exclusive` | `Boolean` | storage | M8 | -- | snapshot round-trip |
| `FullyLoaded` | `Boolean` | storage | M8 | -- | -- |
| `GEXPreview.BackColor` | `OLE_COLOR` | consumed | -- | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `ModelTests` pvTestFormatStyles |
| `GEXPreview.CloseButtonText` | `String` | storage | M11 | -- | `ModelTests` pvTestRoundTrip |
| `GEXPreview.CurrentPage` | `Long` | storage | M11 | -- | -- |
| `GEXPreview.hWnd` | `Long` | computed | -- | -- | `ModelTests` pvTestScroll |
| `GEXPreview.PageSetupText` | `String` | storage | M11 | -- | -- |
| `GEXPreview.PrintText` | `String` | storage | M11 | -- | -- |
| `GEXPreview.ToolbarFont` | `Font` | storage | M11 | -- | `ModelTests` pvTestRoundTrip |
| `GEXPreview.ToolbarVisible` | `Boolean` | storage | M11 | -- | `ModelTests` pvTestRoundTrip |
| `GEXPreview.TotalPages` | `Long` | storage | M11 | -- | -- |
| `GEXPreview.Zoom` | `jgexZoomConstants` | storage | M11 | -- | `ModelTests` pvTestRoundTrip |
| `HoldSortSettings` | `Boolean` | consumed | -- | [`577e39b`](../../commit/577e39b83d084cbd5105745ef66db1942755c4fe) | snapshot round-trip |
| `hWnd` | `Long` | computed | -- | -- | `ModelTests` pvTestScroll |
| `hWndEdit` | `Long` | computed | -- | -- | -- |
| `JSColumn.DataChanged` | `Boolean` | storage | M8 | -- | -- |
| `JSColumn.DataField` | `String` | storage | M8 | -- | snapshot round-trip |
| `JSColumn.DefaultValue` | `Variant` | storage | M5 | -- | -- |
| `JSColumn.DropDownControl` | `Object` | storage | M5 | -- | -- |
| `JSColumn.FetchData` | `Boolean` | storage | M10 | -- | snapshot round-trip |
| `JSColumn.FetchIcon` | `Boolean` | storage | M10 | -- | snapshot round-trip |
| `JSColumn.HasValueList` | `Boolean` | consumed | -- | -- | `ModelTests` pvTestValueList |
| `JSColumn.Index` | `Integer` | consumed | -- | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `ModelTests` pvTestColumns |
| `JSColumn.Key` | `String` | storage | M8 | -- | -- |
| `JSColumn.MaxLength` | `Long` | storage | M5 | -- | snapshot round-trip |
| `JSColumn.NullBehavior` | `jgexNullBehaviorConstants` | storage | M5 | -- | snapshot round-trip |
| `JSColumn.ReplaceValues` | `Boolean` | storage | M5 | -- | snapshot round-trip |
| `JSColumn.Selectable` | `Boolean` | storage | M5 | -- | snapshot round-trip |
| `JSColumn.Tag` | `String` | storage | M8 | -- | snapshot round-trip |
| `JSColumn.ValueList` | `JSValueList` | consumed | -- | -- | `ModelTests` pvTestValueList |
| `JSColumns.Count` | `Integer` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSColumns.Item(vntIndexKey)` | `JSColumn` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSDataObject.Files` | `JSDataObjectFiles` | consumed | -- | [`0fb8743`](../../commit/0fb87434c0852fb35e02628f70786f1111de0626) | -- |
| `JSDataObjectFiles.Count` | `Long` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSDataObjectFiles.Item(Index)` | `String` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSFmtCondition.ColIndex` | `Integer` | storage | M10 | -- | `ModelTests` pvTestSortKeysGroups |
| `JSFmtCondition.FormatStyle` | `JSFormatStyle` | storage | M10 | -- | `Samples` frmUnbound2 |
| `JSFmtCondition.Index` | `Integer` | consumed | -- | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `ModelTests` pvTestColumns |
| `JSFmtCondition.Key` | `String` | storage | M10 | -- | -- |
| `JSFmtCondition.Operator` | `jgexConditionOperatorConstants` | storage | M10 | -- | `ModelTests` pvTestFmtConditions |
| `JSFmtCondition.Value1` | `Variant` | storage | M10 | -- | `ModelTests` pvTestFmtConditions |
| `JSFmtCondition.Value2` | `Variant` | storage | M10 | -- | -- |
| `JSFmtConditions.ApplyGroupCondition` | `Boolean` | storage | M10 | -- | `Samples` frmUnbound2 |
| `JSFmtConditions.Count` | `Long` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSFmtConditions.GroupCondition` | `JSFmtCondition` | storage | M10 | -- | `ModelTests` pvTestFmtConditions |
| `JSFmtConditions.GroupConditionCountTitle` | `String` | storage | M10 | -- | `Samples` frmUnbound2 |
| `JSFmtConditions.Item(Index)` | `JSFmtCondition` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSFmtConditions.ShowGroupConditionCount` | `Boolean` | storage | M10 | -- | `Samples` frmUnbound2 |
| `JSFormatStyle.Name` | `String` | consumed | -- | -- | `ModelTests` pvTestFormatStyles |
| `JSFormatStyles.Count` | `Long` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSFormatStyles.Item(Index)` | `JSFormatStyle` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSGridImage.Index` | `Integer` | consumed | -- | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `ModelTests` pvTestColumns |
| `JSGridImage.Picture` | `Picture` | storage | M10 | -- | -- |
| `JSGridImages.Count` | `Integer` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSGridImages.hImageList` | `Long` | storage | M10 | -- | -- |
| `JSGridImages.Item(Index)` | `JSGridImage` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSGroup.ColIndex` | `Integer` | storage | M4 | -- | `ModelTests` pvTestSortKeysGroups |
| `JSGroup.Index` | `Integer` | consumed | -- | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `ModelTests` pvTestColumns |
| `JSGroup.SortOrder` | `jgexSortOrderConstants` | storage | M4 | -- | `Samples` frmUnbound1 |
| `JSGroups.Count` | `Long` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSGroups.Item(Index)` | `JSGroup` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSPrinterProperties.BottomMargin` | `Long` | consumed | -- | -- | snapshot round-trip |
| `JSPrinterProperties.CardColumnsPerPage` | `Long` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.ClientHeight` | `Long` | consumed | -- | -- | -- |
| `JSPrinterProperties.ClientWidth` | `Long` | consumed | -- | -- | `ModelTests` pvTestPrinterProperties |
| `JSPrinterProperties.Collate` | `Boolean` | storage | M11 | -- | -- |
| `JSPrinterProperties.ColorMode` | `jgexPPColorModeConstants` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.Copies` | `Integer` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.DeviceName` | `String` | computed | -- | -- | -- |
| `JSPrinterProperties.DocumentName` | `String` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.DriverName` | `String` | computed | -- | -- | -- |
| `JSPrinterProperties.FitColumns` | `Boolean` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.FooterDistance` | `Long` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.FooterString(Position)` | `String` | storage | M11 | -- | `ModelTests` pvTestPrinterProperties |
| `JSPrinterProperties.HeaderDistance` | `Long` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.HeaderString(Position)` | `String` | storage | M11 | -- | `ModelTests` pvTestPrinterProperties |
| `JSPrinterProperties.LeftMargin` | `Long` | consumed | -- | -- | `ModelTests` pvTestPrinterProperties |
| `JSPrinterProperties.MeasurementUnits` | `jgexMeasurementUnitsConstants` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.Orientation` | `jgexPPOrientationConstants` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.PageFooterFont` | `Font` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.PageHeaderFont` | `Font` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.PageSetupCanceled` | `Boolean` | storage | M11 | -- | -- |
| `JSPrinterProperties.PaperBin` | `jgexPPPaperBinConstants` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.PaperHeight` | `Long` | consumed | -- | -- | snapshot round-trip |
| `JSPrinterProperties.PaperSize` | `Integer` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.PaperWidth` | `Long` | consumed | -- | -- | `ModelTests` pvTestPrinterProperties |
| `JSPrinterProperties.PrintPreviewRows` | `Boolean` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.PrintProgressDialog` | `Boolean` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.PrintQuality` | `Integer` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.RepeatFrozenCols` | `Boolean` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.RepeatHeaders` | `Boolean` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.RightMargin` | `Long` | consumed | -- | -- | `ModelTests` pvTestPrinterProperties |
| `JSPrinterProperties.TopMargin` | `Long` | consumed | -- | -- | snapshot round-trip |
| `JSPrinterProperties.TranslateColors` | `Boolean` | storage | M11 | -- | snapshot round-trip |
| `JSPrinterProperties.TransparentBackground` | `Boolean` | storage | M11 | -- | snapshot round-trip |
| `JSRetBoolean.Value` | `Boolean` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRetInteger.Value` | `Integer` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRetVariant.Value` | `Variant` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.Bookmark` | `Variant` | consumed | -- | -- | -- |
| `JSRowData.CellPicture(ColIndex)` | `Picture` | consumed | -- | -- | -- |
| `JSRowData.CellStyle(ColIndex)` | `String` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.ColCount` | `Integer` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.DisplayValue(ColIndex)` | `String` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.GroupCaption` | `String` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.GroupIconIndex` | `Integer` | consumed | -- | -- | -- |
| `JSRowData.GroupLevel` | `Integer` | consumed | -- | -- | -- |
| `JSRowData.IconIndex(ColIndex)` | `Integer` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.PreviewRowVisible` | `Boolean` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.RecordCount` | `Long` | consumed | -- | -- | `Samples` frmUnbound1 |
| `JSRowData.RowHeight` | `Long` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.RowIndex` | `Long` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.RowStyle` | `String` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.RowType` | `jgexRowTypeConstants` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSRowData.Value(ColIndex)` | `Variant` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSSelectedItem.Bookmark` | `Variant` | consumed | -- | -- | -- |
| `JSSelectedItem.RowIndex` | `Long` | consumed | -- | [`0fb8743`](../../commit/0fb87434c0852fb35e02628f70786f1111de0626) | `ModelTests` pvTestRowData |
| `JSSelectedItem.RowPosition` | `Long` | consumed | -- | -- | -- |
| `JSSelectedItem.RowType` | `jgexRowTypeConstants` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSSelectedItems.Count` | `Long` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSSelectedItems.Item(Index)` | `JSSelectedItem` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSSortKey.ColIndex` | `Integer` | storage | M4 | -- | `ModelTests` pvTestSortKeysGroups |
| `JSSortKey.Index` | `Integer` | consumed | -- | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `ModelTests` pvTestColumns |
| `JSSortKey.SortOrder` | `jgexSortOrderConstants` | storage | M4 | -- | `Samples` frmUnbound1 |
| `JSSortKeys.Count` | `Long` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSSortKeys.Item(Index)` | `JSSortKey` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSValueItem.IconIndex` | `Integer` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSValueItem.Index` | `Long` | consumed | -- | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `ModelTests` pvTestColumns |
| `JSValueItem.Text` | `String` | storage | M5 | -- | `ModelTests` pvTestValueList |
| `JSValueItem.Value` | `Variant` | consumed | -- | -- | `ModelTests` pvTestRowData |
| `JSValueItem.Visible` | `Boolean` | consumed | -- | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `ModelTests` pvTestValueList |
| `JSValueList.Count` | `Long` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSValueList.Item(Index)` | `JSValueItem` | consumed | -- | -- | `ModelTests` pvTestColumns |
| `JSValueList.ItemByValue(Value)` | `JSValueItem` | consumed | -- | -- | `ModelTests` pvTestValueList |
| `JSValueList.VisibleCount` | `Long` | consumed | -- | -- | `ModelTests` pvTestValueList |
| `LockType` | `jgexLockTypeConstants` | storage | M8 | -- | snapshot round-trip |
| `OLEDropMode` | `jgexOleDropModeConstants` | storage | M12 | -- | snapshot round-trip |
| `PrinterProperties` | `JSPrinterProperties` | storage | M11 | -- | `ModelTests` pvTestPrinterProperties |
| `ReadOnly` | `Boolean` | storage | M8 | -- | snapshot round-trip |
| `Recordset` | `Object` | storage | M8 | -- | -- |
| `RecordsetType` | `jgexRecordsetTypeConstants` | storage | M8 | -- | snapshot round-trip |
| `RecordSource` | `String` | storage | M8 | -- | snapshot round-trip |
| `ReplaceColumnIndex` | `Variant` | consumed | -- | -- | -- |
| `RowBookmark(RowIndex)` | `Variant` | consumed | -- | -- | `ModelTests` pvTestUnbound |
| `SelLength` | `Long` | storage | M5 | -- | -- |
| `SelStart` | `Long` | storage | M5 | -- | -- |
| `SelText` | `String` | storage | M5 | -- | -- |
| `TabKeyBehavior` | `jgexTabKeyBehaviorConstants` | storage | M5 | -- | snapshot round-trip |
| `Value(ColIndex)` | `Variant` | computed | -- | -- | `ModelTests` pvTestRowData |

## Part 3 -- The object model

One table per class, in dependency order, with the original's own description
of each member. Properties already carried by part 1 keep their row here, marked
`paint`, so every class reads as a complete list.

### `JSColumn`

A column: layout, data binding, editing and card behaviour -- 46 properties (31 paint, 3 consumed, 12 storage).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Key` | `String` | read-only | storage | Returns/sets a string that uniquely identifies a column in a collection. |
| `HeaderAlignment` | `jgexAlignmentConstants` | read/write | [paint](#the-matrix) | Returns/sets a value that determines the alignment of text in the column's header. |
| `HeaderIcon` | `Integer` | read/write | [paint](#the-matrix) | Returns/sets the index of the JSGridImage displayed in the column's header. |
| `DefaultIcon` | `Integer` | read/write | [paint](#the-matrix) | Returns/sets the index of the JSGridImage displayed in the column's cells. |
| `Index` | `Integer` | read-only | consumed | Returns a value that represents the index of the JSColumn in the JSColumns collection. |
| `ColPosition` | `Integer` | read/write | [paint](#the-matrix) | Returns/sets the position of the column. |
| `ValueList` | `JSValueList` | read-only | consumed | Returns the JSValueList collection for the column. |
| `Visible` | `Boolean` | read/write | [paint](#the-matrix) | Returns/sets a value indicating whether an object is visible or hidden. |
| `Width` | `Long` | read/write | [paint](#the-matrix) | Returns/sets the width of a column. |
| `DataField` | `String` | read/write | storage | Returns/sets a value that represents a field in the Recordset. |
| `Caption` | `String` | read/write | [paint](#the-matrix) |  |
| `AllowSizing` | `Boolean` | read/write | [paint](#the-matrix) | Returns/sets a value indicating Determines whether a user can resize the column. |
| `ColumnType` | `jgexColumnTypeConstants` | read/write | [paint](#the-matrix) | Returns/sets a value that indicates how column's contents are displayed. |
| `CardCaption` | `Boolean` | read/write | [paint](#the-matrix) | Determines whether the column's value will be displayed in the card caption. |
| `EditType` | `jgexEditTypeConstants` | read/write | [paint](#the-matrix) | Returns/sets a value that indicates how column is edited by the user. |
| `HasValueList` | `Boolean` | read/write | consumed | Controls whether the column has a list of values. |
| `SortOrder` | `jgexSortOrderConstants` | read-only | [paint](#the-matrix) | Returns a value that represents the sort applied to a column. |
| `SortType` | `jgexSortTypeConstants` | read/write | [paint](#the-matrix) | Determines the way that values in the column are sorted and/or grouped. |
| `ShowCaptionInCardView` | `Boolean` | read/write | [paint](#the-matrix) | Controls whether the caption of a column is displayed before its value in the card body. |
| `IsGrouped` | `Boolean` | read-only | [paint](#the-matrix) | Returns a value that indicates whether the Column is grouped. |
| `Format` | `String` | read/write | [paint](#the-matrix) | Returns/sets a value indicating the format string for the Column. |
| `TextAlignment` | `jgexAlignmentConstants` | read/write | [paint](#the-matrix) | Returns/sets a value that determines the alignment of text in the column. |
| `GroupFormat` | `String` | read/write | [paint](#the-matrix) | Returns/sets a value indicating the format string for the value of the group rows when the control is grouped by the column. |
| `GroupPrefix` | `String` | read/write | [paint](#the-matrix) | Returns/sets the string displayed in a group row before the group value. |
| `GroupEmptyStringCaption` | `String` | read/write | [paint](#the-matrix) | Returns/sets the string displayed in a group row when the value grouped is an empty string. |
| `CardIcon` | `Boolean` | read/write | [paint](#the-matrix) | Determines whether the column's icon will be displayed in the card caption. |
| `FetchIcon` | `Boolean` | read/write | storage | Returns/sets whether the FetchIcon event will be fired for the column. |
| `FetchData` | `Boolean` | read/write | storage | Returns/sets whether the FetchData event will be fired for the column. |
| `Tag` | `String` | read/write | storage | Returns/sets an expression that stores any extra data needed for your program. |
| `DataChanged` | `Boolean` | read/write | storage | Returns/sets a value indicating that data in the column has changed by some process other than by retrieving data from the current record. |
| `ReplaceValues` | `Boolean` | read/write | storage | Returns/sets whether the column values are replaced for using the ValueList for the column. |
| `DefaultValue` | `Variant` | read/write | storage | Returns/sets the default value for the column in a new record. |
| `MaxLength` | `Long` | read/write | storage | Returns/sets the maximum number of characters that can be entered in a field. |
| `NullBehavior` | `jgexNullBehaviorConstants` | read/write | storage | Determines how empty strings are written in the database. |
| `ButtonStyle` | `jgexButtonStyleConstants` | read/write | [paint](#the-matrix) | Determines if a column will show a button when the user enters edit mode. |
| `Selectable` | `Boolean` | read/write | storage | Controls whether cell can be selected. |
| `CellStyle` | `String` | read/write | [paint](#the-matrix) | Returns/sets the name of the JSFormatStyle to be applied in all the cells in a column. |
| `WordWrap` | `Boolean` | read/write | [paint](#the-matrix) | determines whether text in a cell is wordwrapped. |
| `AggregateFunction` | `jgexAggregateFunctionConstants` | read/write | [paint](#the-matrix) | Returns/sets the aggregate function to be shown in group footers for a column. |
| `MinRowsInCardView` | `Long` | read/write | [paint](#the-matrix) | Returns or sets the minimum number of rows displayed in a word wrap column. |
| `MaxRowsInCardView` | `Long` | read/write | [paint](#the-matrix) | Returns or sets the maximum number of rows displayed in a word wrap column. |
| `DropDownControl` | `Object` | read/write | storage | Returns/sets the GridEX control that acts as a drop down list for the column. |
| `HeaderToolTip` | `String` | read/write | [paint](#the-matrix) | Returns/sets the text displayed in a tool tip when users moves the mouse pointer over a column header. |
| `HeaderStyle` | `String` | read/write | [paint](#the-matrix) | Returns/sets the name of the JSFormatStyle to be applied in the column header. |
| `TotalRowFormat` | `String` | read/write | [paint](#the-matrix) | Returns/sets the format string for the aggregate function result of a column in the group footer row. |
| `TotalRowPrefix` | `String` | read/write | [paint](#the-matrix) | Returns/sets the string displayed in a group footer row before the aggregate function result. |

### `JSColumns`

The column collection -- 2 properties (2 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Item(vntIndexKey)` | `JSColumn` | read-only | consumed | Returns a specific JSColumn either by index or by key. |
| `Count` | `Integer` | read-only | consumed | Returns the number of objects in a collection. |

### `JSValueList`

The value list of a column -- 4 properties (4 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Count` | `Long` | read-only | consumed | Returns the number ofJS ValueItem objects in a JSValueList. |
| `VisibleCount` | `Long` | read-only | consumed | Returns the number of visible value items. |
| `Item(Index)` | `JSValueItem` | read-only | consumed | Returns a specific JSValueItem of the JSValueList. |
| `ItemByValue(Value)` | `JSValueItem` | read-only | consumed | Returns a specific JSValueItem of the JSValueList. |

### `JSValueItem`

One entry of a value list -- 5 properties (4 consumed, 1 storage).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Index` | `Long` | read-only | consumed | Returns a value that represents the index of an object in a collection. |
| `Text` | `String` | read/write | storage | Returns/sets the text that is linked to the Value. |
| `Value` | `Variant` | read/write | consumed | Returns/sets the value for the object. |
| `IconIndex` | `Integer` | read/write | consumed | Returns/sets the index of the JSGridImage that is linked to the value. |
| `Visible` | `Boolean` | read/write | consumed | Determines whether a JSValueItem is visible or hidden in a drop-down list. |

### `JSFormatStyle`

A named format style -- 17 properties (16 paint, 1 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `BackColor` | `OLE_COLOR` | read/write | [paint](#the-matrix) | Returns/sets the background color. |
| `Name` | `String` | read-only | consumed | Returns a string that uniquely identifies a member in a JSFormatStyles collection. |
| `ForeColor` | `OLE_COLOR` | read/write | [paint](#the-matrix) | Returns/sets the foreground color. |
| `FontBold` | `Boolean` | read/write | [paint](#the-matrix) | Returns/sets font bold style. |
| `FontItalic` | `Boolean` | read/write | [paint](#the-matrix) | Returns/sets font italic style. |
| `FontUnderline` | `Boolean` | read/write | [paint](#the-matrix) | Returns/sets font underline style. |
| `FontStrikeThru` | `Boolean` | read/write | [paint](#the-matrix) | Returns/sets font strikethrough style. |
| `FontSize` | `Currency` | read/write | [paint](#the-matrix) | Returns/sets the size of the font. |
| `FontName` | `String` | read/write | [paint](#the-matrix) | Returns/sets the typeface name of the font. |
| `TextAlignment` | `jgexAlignmentConstants` | read/write | [paint](#the-matrix) | Determines the alignment of text in a JSFormatStyle object. |
| `BackgroundPicture` | `Picture` | read/write | [paint](#the-matrix) | Returns/sets the background picture. |
| `Picture` | `Picture` | read/write | [paint](#the-matrix) | Returns/sets a picture of a JSFormatStyle object. |
| `FontCharset` | `Integer` | read/write | [paint](#the-matrix) | Returns/sets the character set used in the font. |
| `DrawModeBackGroundPicture` | `jgexDrawModePictureBackgroundConstants` | read/write | [paint](#the-matrix) | Determines the way a background picture is drawn. |
| `PictureHorzAlignment` | `jgexHorzPictureAlignmentConstants` | read/write | [paint](#the-matrix) | Determines the horizontal alignment of a picture in a JSFormatStyle object. |
| `PictureVertAlignment` | `jgexVertPictureAlignmentConstants` | read/write | [paint](#the-matrix) | Determines the vertical alignment of a picture in a JSFormatStyle object. |
| `PictureDrawMode` | `jgexPictureDrawModeConstants` | read/write | [paint](#the-matrix) | Determines how GridEX control draws cells with pictures. |

### `JSFormatStyles`

The format style collection -- 2 properties (2 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Count` | `Long` | read-only | consumed | Returns the number of objects in a collection. |
| `Item(Index)` | `JSFormatStyle` | read-only | consumed | Returns a specific JSFormatStyle of the collection. |

### `JSFmtCondition`

One conditional format -- 7 properties (1 consumed, 6 storage).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `FormatStyle` | `JSFormatStyle` | read-only | storage | Returns the JSFormatStyle object for the JSFmtCondition. |
| `ColIndex` | `Integer` | read/write | storage | Returns/sets the index of the column  to be compared. |
| `Operator` | `jgexConditionOperatorConstants` | read/write | storage | Returns/sets the operator used for comparison. |
| `Value1` | `Variant` | read/write | storage | Returns/sets the value that is compared with column values. |
| `Value2` | `Variant` | read/write | storage | Returns/sets the value that is compared with column values. |
| `Index` | `Integer` | read-only | consumed | Returns a value that represents the index of an object in a collection. |
| `Key` | `String` | read-only | storage | Returns/sets a string that uniquely identifies a member in a collection. |

### `JSFmtConditions`

The conditional format collection -- 6 properties (2 consumed, 4 storage).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `GroupCondition` | `JSFmtCondition` | read-only | storage | Returns a JSFmtCondition object that is used through the group rows. |
| `Item(Index)` | `JSFmtCondition` | read-only | consumed | Returns a specific JSFmtCondition of  the collection either by index or by key. |
| `Count` | `Long` | read-only | consumed | Returns the number of objects in a collection. |
| `ApplyGroupCondition` | `Boolean` | read/write | storage | Controls whether the GroupCondition property is valid. |
| `GroupConditionCountTitle` | `String` | read/write | storage | Returns/sets the text displayed in a group row when one or more rows in the group meet the criteria specified in the JSGroupCondition object's properties. |
| `ShowGroupConditionCount` | `Boolean` | read/write | storage | Controls whether the count of rows, that meet the criteria in GroupCondition property settings, will be displayed in a group row. |

### `JSGroup`

One grouping level -- 3 properties (1 consumed, 2 storage).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Index` | `Integer` | read-only | consumed | Returns a value that represents the index of the JSGroup in the JSGroups collection. |
| `SortOrder` | `jgexSortOrderConstants` | read/write | storage |  |
| `ColIndex` | `Integer` | read/write | storage | Returns/sets the index of the grouped column. |

### `JSGroups`

The grouping collection -- 2 properties (2 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Item(Index)` | `JSGroup` | read-only | consumed | Returns a specific JSGroup of the Collection. |
| `Count` | `Long` | read-only | consumed | Returns the number of objects in a collection. |

### `JSSortKey`

One sort key -- 3 properties (1 consumed, 2 storage).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Index` | `Integer` | read-only | consumed | Returns a value that represents the index of the JSSortKey in the JSSortKeys collection. |
| `ColIndex` | `Integer` | read/write | storage | Returns/sets the index of the sorted column. |
| `SortOrder` | `jgexSortOrderConstants` | read/write | storage | Returns/sets a value that represents the sort order. |

### `JSSortKeys`

The sort key collection -- 2 properties (2 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Item(Index)` | `JSSortKey` | read-only | consumed | Returns a specific JSSortKey of the Collection. |
| `Count` | `Long` | read-only | consumed | Returns the number of objects in a collection. |

### `JSGridImage`

One image in the control image list -- 2 properties (1 consumed, 1 storage).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Picture` | `Picture` | read-only | storage | Returns or sets a graphic to be displayed. |
| `Index` | `Integer` | read-only | consumed | Returns a value that represents the index of an object in a collection. |

### `JSGridImages`

The image collection -- 3 properties (2 consumed, 1 storage).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `hImageList` | `Long` | read/write | storage | Returns/sets the handle of an external ImageList. |
| `Item(Index)` | `JSGridImage` | read-only | consumed | Returns a specific JSGridImage of the JSGridImages Collection. |
| `Count` | `Integer` | read-only | consumed | Returns the number of objects in a collection. |

### `JSRowData`

The values of one row, handed to the unbound events -- 16 properties (16 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Bookmark` | `Variant` | read-only | consumed | Returns the bookmark of the row. |
| `ColCount` | `Integer` | read-only | consumed | Returns the number of columns. |
| `GroupLevel` | `Integer` | read-only | consumed | Returns the group level of a row if it is a group row or 0 otherwise. |
| `RowIndex` | `Long` | read-only | consumed | Returns the row index of a row. |
| `Value(ColIndex)` | `Variant` | read/write | consumed | Returns/sets the value of a column. |
| `IconIndex(ColIndex)` | `Integer` | read/write | consumed | Returns/sets the index of the JSGridImage displayed in a cell. |
| `DisplayValue(ColIndex)` | `String` | read/write | consumed | Returns/sets the display value of a column. |
| `RowType` | `jgexRowTypeConstants` | read-only | consumed | Returns a value that represents the type of a JSRowData object. |
| `RowStyle` | `String` | read/write | consumed | Returns/sets the name of the JSFormatStyle to be applied in all the cells in a row. |
| `CellStyle(ColIndex)` | `String` | read/write | consumed | Returns/sets the name of the JSFormatStyle to be applied in a cell. |
| `RecordCount` | `Long` | read-only | consumed | Returns the number of records in a JSRowData object representing a group row. |
| `RowHeight` | `Long` | read/write | consumed | Returns/sets the height, in twips, of a row. |
| `GroupCaption` | `String` | read/write | consumed | Returns/sets the caption of a group row. |
| `GroupIconIndex` | `Integer` | read/write | consumed | Returns/sets the icon of a group row. |
| `CellPicture(ColIndex)` | `Picture` | read/write | consumed | Returns/sets the foreground picture of a cell. |
| `PreviewRowVisible` | `Boolean` | read/write | consumed | Determines whether a preview row should be displayed in a row. |

### `JSSelectedItem`

One entry of the selection -- 4 properties (4 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Bookmark` | `Variant` | read-only | consumed | Returns the bookmark of a selected row. |
| `RowIndex` | `Long` | read-only | consumed | Returns the original index of a selected row. |
| `RowPosition` | `Long` | read-only | consumed | Returns the current position of a selected row. |
| `RowType` | `jgexRowTypeConstants` | read-only | consumed | Returns the type of a selected row. |

### `JSSelectedItems`

The selection collection -- 2 properties (2 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Count` | `Long` | read-only | consumed | Returns the number of objects in a collection. |
| `Item(Index)` | `JSSelectedItem` | read-only | consumed | Returns a specific JSSelectedItem of the collection. |

### `JSPrinterProperties`

The printer setup carried by the control -- 34 properties (8 consumed, 24 storage, 2 derived).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Copies` | `Integer` | read/write | storage | Returns/sets the number of copies to be printed. |
| `Collate` | `Boolean` | read/write | storage | Determines whether collation should be used when printing multiple copies. |
| `PrintQuality` | `Integer` | read/write | storage | Returns/sets the printer resolution. |
| `ColorMode` | `jgexPPColorModeConstants` | read/write | storage | Returns/sets a value that determines whether a document should be printed in color or monochrome. |
| `DeviceName` | `String` | read-only | derived | Returns the name of the device a driver supports. |
| `DriverName` | `String` | read-only | derived | Returns the name of the printer's driver. |
| `Orientation` | `jgexPPOrientationConstants` | read/write | storage | Determines whether documents are printed in portrait or landscape mode. |
| `PaperBin` | `jgexPPPaperBinConstants` | read/write | storage | Returns/sets a value indicating the default paper bin on the printer from which paper is fed when printing. |
| `PaperSize` | `Integer` | read/write | storage | Returns/sets the paper size for the current printer. |
| `PaperWidth` | `Long` | read/write | consumed | Returns/sets the physical width, in twips, of the paper set up for the printing device. |
| `PaperHeight` | `Long` | read/write | consumed | Returns/sets the physical height, in twips, of the paper set up for the printing device. |
| `LeftMargin` | `Long` | read/write | consumed | Returns/sets the width, in twips, of the left margin. |
| `TopMargin` | `Long` | read/write | consumed | Returns/sets the height, in twips, of the top margin. |
| `RightMargin` | `Long` | read/write | consumed | Returns/sets the width, in twips, of the right margin. |
| `BottomMargin` | `Long` | read/write | consumed | Returns/sets the height, in twips, of the bottom margin. |
| `ClientWidth` | `Long` | read-only | consumed | Returns the width, in twips, of the printable area in a page. |
| `ClientHeight` | `Long` | read-only | consumed | Returns the height, in twips, of the printable area in a page. |
| `RepeatHeaders` | `Boolean` | read/write | storage | Determines whether column headers should appear on each page. |
| `FitColumns` | `Boolean` | read/write | storage | DEtermines whether the GridEX should scale printed output to fit all visible columns in a page. |
| `TranslateColors` | `Boolean` | read/write | storage | Determines whether system colors should be translated to black and white colors. |
| `HeaderDistance` | `Long` | read/write | storage | Returns/sets the distance, in twips, from the top edge of the paper to the top edge of the header. |
| `FooterDistance` | `Long` | read/write | storage | Returns/sets the distance, in twips, from the bottom edge of the paper to the top edge of the footer. |
| `HeaderString(Position)` | `String` | read/write | storage | Returns/sets the string displayed in the page header. |
| `FooterString(Position)` | `String` | read/write | storage | Returns/sets the string displayed in the page footer |
| `DocumentName` | `String` | read/write | storage | Returns/sets the name of the document to be printed. |
| `RepeatFrozenCols` | `Boolean` | read/write | storage | Determines whether frozen columns should appear on each page. |
| `PrintPreviewRows` | `Boolean` | read/write | storage | Determines whether the control prints preview rows or not. |
| `CardColumnsPerPage` | `Long` | read/write | storage | Returns/sets the number of card columns to be printed in a page. |
| `PageHeaderFont` | `Font` | read/write | storage | Returns/sets a Font object used to draw page headers in a document. |
| `PageFooterFont` | `Font` | read/write | storage | Returns/sets a Font object used to draw page footers in a document. |
| `PrintProgressDialog` | `Boolean` | read/write | storage | Determines whether print progress dialog should be displayed when GridEX control is printing a document. |
| `TransparentBackground` | `Boolean` | read/write | storage | Determines whether background color should be used to when printing. |
| `MeasurementUnits` | `jgexMeasurementUnitsConstants` | read/write | storage | Returns/sets measurement units to be used in the Page Setup dialog. |
| `PageSetupCanceled` | `Boolean` | read-only | storage | Returns True if user pressed cancel in the Page Setup dialog. |

### `JSDataObject`

The OLE drag/drop payload -- 1 properties (1 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Files` | `JSDataObjectFiles` | read-only | consumed | Returns a JSDataObjectFiles collection, which in turn contains a list of all filenames used by a JSDataObject object. |

### `JSDataObjectFiles`

The file list of an OLE drag/drop payload -- 2 properties (2 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Count` | `Long` | read-only | consumed | Returns the number of file names in the collection. |
| `Item(Index)` | `String` | read-only | consumed | Returns an specific file name. |

### `JSRetBoolean`

The by-reference Boolean the cancellable events pass -- 1 properties (1 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Value` | `Boolean` | read/write | consumed |  |

### `JSRetInteger`

The by-reference Integer the events pass -- 1 properties (1 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Value` | `Integer` | read/write | consumed |  |

### `JSRetVariant`

The by-reference Variant the events pass -- 1 properties (1 consumed).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `Value` | `Variant` | read/write | consumed |  |

### `GEXPreview`

The print preview control -- 10 properties (1 consumed, 8 storage, 1 derived).

| Property | Type | Access | Status | Description |
|---|---|---|---|---|
| `TotalPages` | `Long` | read-only | storage | Returns the number of pages in a document. |
| `CurrentPage` | `Long` | read/write | storage | Returns or sets the page displayed. |
| `ToolbarFont` | `Font` | read/write | storage | Returns/sets a Font object used in the toolbar. |
| `Zoom` | `jgexZoomConstants` | read/write | storage | Determines how GEXPreview control should display pages. |
| `ToolbarVisible` | `Boolean` | read/write | storage | Determines whether the tool bar is displayed. |
| `PageSetupText` | `String` | read/write | storage | Returns/sets the text displayed in the <Page Setup> button. |
| `PrintText` | `String` | read/write | storage | Returns/sets the text displayed in the <Print> button. |
| `CloseButtonText` | `String` | read/write | storage | Returns/sets the text displayed in the <Close> button. |
| `hWnd` | `Long` | read-only | derived | Returns the handle of a GEXPreview control. |
| `BackColor` | `OLE_COLOR` | read/write | consumed | Retusns/sets the background color of the control. |
