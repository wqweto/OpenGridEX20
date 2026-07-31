# GridEX paint property matrix

Every property that affects what the control paints -- the direct `GridEX`
members plus the `JSColumn` and `JSFormatStyle` sub-properties -- in one
alphabetical list, whether or not the renderer reads it yet.

A property belongs here when it changes what pixels appear in the client area.
Pure data and behaviour members (`DataMode`, `Connect`, `AllowDelete`,
`JSColumn.DataField`, `JSColumn.Key`, ...) are excluded even where implemented.

## The verification bar

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

## Summary

| | count | share |
|---|---:|---:|
| verified | 44 | 35% |
| weak | 1 | 1% |
| unverified | 0 | 0% |
| partial | 0 | 0% |
| **not implemented** | **81** | **64%** |
| **total** | **126** | |

**45 of 126 (36%) are read by the paint path**; the other **81 (64%) are not
implemented** -- they store and return their value, and round-trip through the
snapshot corpus, but the renderer never looks at them.

**Every property the renderer reads is proven** against the original at two or
more distinct values, except `RecordNavigatorString`, which renders only at its
default. The 81 that remain are unimplemented, each owned by a later milestone.

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
| M3d (remainder) | 5 |
| M5 editing | 5 |
| **unowned** | **5** |

**M10 is by far the largest single block of unimplemented painting** -- 38
properties, mostly the `JSFormatStyle` font and picture families -- which is a
much bigger milestone than its one-line roadmap entry suggests.

Card view was the other gap this matrix exposed and now has its own milestone
(M6), placed before the paint gate because it *adds* twelve paint properties.
Five remain unowned: the drag/resize affordances (`AllowColumnDrag`,
`AllowRowSizing`, `DetectRowDrag`, `JSColumn.AllowSizing`) and `Options`.

## The matrix

Scenario names refer to `test\VisualDiff\scenarios\NNN-*.json`, numbered in
creation order; all 30 are verified at 96 and 120 dpi, and the first 28 at 144 as
well. The `Commit` column records where the property started affecting pixels,
not where it was first stored -- every member was declared in the M2 storage
commit, so that hash carries no information about painting.

| Property | Type | Status | Paint routine | Milestone | Commit | Test |
|---|---|---|---|---|---|---|
| `AllowAddNew` | `Boolean` | **not impl** | -- | M5 | -- | -- |
| `AllowCardSizing` | `Boolean` | **not impl** | -- | M6 | -- | -- |
| `AllowColumnDrag` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `AllowRowSizing` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `AutomaticArrange` | `Boolean` | **not impl** | -- | M6 | -- | -- |
| `BackColor` | `OLE_COLOR` | verified | `pvPaintDataRow`, `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (`0xC0FFC0`) vs default |
| `BackColorBkg` | `OLE_COLOR` | verified | `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (`0x404040`), `002-gridlines-dots-colors` (`0x808080`) |
| `BackColorGBBox` | `OLE_COLOR` | verified | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x604020`) vs default |
| `BackColorHeader` | `OLE_COLOR` | verified | `pvPaintHeaderCell`, `pvPaintRowHeader`, `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x8000FF`) vs default |
| `BackColorInfoText` | `OLE_COLOR` | verified | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x30A0C0`) vs default |
| `BackColorRowGroup` | `OLE_COLOR` | **not impl** | -- | M4 | -- | -- |
| `BorderStyle` | `jgexBorderStyleConstants` | **not impl** | -- | M10 | -- | -- |
| `CardBorders` | `Boolean` | **not impl** | -- | M6 | -- | -- |
| `CardCaptionPrefix` | `String` | **not impl** | -- | M6 | -- | -- |
| `CardSpacing` | `Long` | **not impl** | -- | M6 | -- | -- |
| `CardWidth` | `Long` | **not impl** | -- | M6 | -- | -- |
| `Col` | `Integer` | **not impl** | -- | M3d | -- | -- |
| `ColumnAutoResize` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `ColumnHeaderFont` | `Font` | verified | `pvPaintHeaders` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `009-font-large`, `020-font-tahoma`, `019-font-segoeui` + MS Sans Serif default |
| `ColumnHeaderHeight` | `Long` | verified | `pvPaintHeaders`, `pvPaint` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `003-headers-flat` sets 400tw; font-derived defaults elsewhere |
| `ColumnHeaders` | `Boolean` | verified | `pvPaint` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `004-no-headers-no-rowheaders` (False) vs default True |
| `Columns` | `JSColumns` | verified | `pvPaintHeaders`, `pvPaintDataRow`, `pvPaintRows`, `pvPaintRowMarquee` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | all 30 scenarios -- see the `JSColumn.*` rows |
| `ContinuousScroll` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `DefaultColumnWidth` | `Long` | verified | via `JSColumns.Add` at init | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `024-default-column-width` (2100tw -> 140px columns) vs default |
| `DefaultGroupMode` | `jgexDefaultGroupModeConstants` | **not impl** | -- | M4 | -- | -- |
| `DetectRowDrag` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `EditMode` | `jgexEditModeConstants` | **not impl** | -- | M5 | -- | -- |
| `EmptyRows` | `Boolean` | verified | `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `008-empty-rows` (True) vs default False |
| `Enabled` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `FirstItem` | `Long` | verified | `pvPaintRows`, `pvUpdateScrollBars` | M3d | [`e95c15e`](../../commit/e95c15e1b1d35f09c9bbe557228ddc1329ed6c6f) | `016-scrolled` sets 4 in its `post` block; 1 elsewhere |
| `FmtConditions` | `JSFmtConditions` | **not impl** | -- | M10 | -- | -- |
| `Font` | `Font` | verified | `pvPaintRows`, `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `009-font-large`, `020-font-tahoma`, `019-font-segoeui` + default |
| `ForeColor` | `OLE_COLOR` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (`0xFF0000`) vs default -- also drives the marquee XOR mask |
| `ForeColorHeader` | `OLE_COLOR` | verified | `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x00FFFF`) vs default |
| `ForeColorInfoText` | `OLE_COLOR` | verified | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` (`0x0020FF`) vs default |
| `ForeColorRowGroup` | `OLE_COLOR` | **not impl** | -- | M4 | -- | -- |
| `FormatStyles` | `JSFormatStyles` | verified | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `025-formatstyles-selection` overrides the `SelectedRow` system style |
| `FrozenColumns` | `Integer` | **not impl** | -- | M3d | -- | -- |
| `GridImages` | `JSGridImages` | **not impl** | -- | M10 | -- | -- |
| `GridLines` | `jgexGridLinesConstants` | verified | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | all four modes: default Both, plus `027-gridlines-none`, `028-gridlines-vertical`, `026-gridlines-horizontal` |
| `GridLinesColor` | `OLE_COLOR` | verified | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (`0x0000FF`) vs default |
| `GridLineStyle` | `jgexGridLineStyleConstants` | verified | `pvPenStyle` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `022-colors-rows` (dots), `028-gridlines-vertical` (dashes), solid default |
| `GroupByBoxInfoText` | `String` | verified | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `021-colors-chrome` custom text vs default |
| `GroupByBoxVisible` | `Boolean` | verified | `pvPaint` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `010-headers-noborder`, `011-headers-single3d`, `012-headers-singleflat` vs default |
| `GroupFooterStyle` | `jgexGroupFooterStyleConstants` | **not impl** | -- | M4 | -- | -- |
| `Groups` | `JSGroups` | **not impl** | -- | M4 | -- | -- |
| `HeaderStyle` | `jgexHeaderStyleConstants` | verified | `pvPaintHeaders`, `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `003-headers-flat`, `010-headers-noborder`, `011-headers-single3d`, `012-headers-singleflat` (4 values) |
| `HideSelection` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `ImageHeight` | `Integer` | **not impl** | -- | M10 | -- | -- |
| `ImageWidth` | `Integer` | **not impl** | -- | M10 | -- | -- |
| `ItemCount` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow` (via `RowCount`) | M3a | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | row counts 2/3/5/6/14/16 across the corpus |
| `JSColumn.AggregateFunction` | `jgexAggregateFunctionConstants` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.AllowSizing` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `JSColumn.ButtonStyle` | `jgexButtonStyleConstants` | **not impl** | -- | M5 | -- | -- |
| `JSColumn.Caption` | `String` | verified | `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | distinct captions in every scenario |
| `JSColumn.CardCaption` | `Boolean` | **not impl** | -- | M6 | -- | -- |
| `JSColumn.CardIcon` | `Boolean` | **not impl** | -- | M6 | -- | -- |
| `JSColumn.CellStyle` | `String` | **not impl** | -- | M10 | -- | -- |
| `JSColumn.ColPosition` | `Integer` | verified | `pvPaintRows` (via `ItemByPosition`) | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `023-column-order` reorders 3 columns to 1/2/3 from declaration order |
| `JSColumn.ColumnType` | `jgexColumnTypeConstants` | **not impl** | -- | M10 | -- | -- |
| `JSColumn.DefaultIcon` | `Integer` | **not impl** | -- | M10 | -- | -- |
| `JSColumn.EditType` | `jgexEditTypeConstants` | **not impl** | -- | M5 | -- | -- |
| `JSColumn.Format` | `String` | **not impl** | -- | M10 | -- | -- |
| `JSColumn.GroupEmptyStringCaption` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.GroupFormat` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.GroupPrefix` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.HeaderAlignment` | `jgexAlignmentConstants` | verified | `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `003-headers-flat` sets 2; default 0 elsewhere |
| `JSColumn.HeaderIcon` | `Integer` | **not impl** | -- | M10 | -- | -- |
| `JSColumn.HeaderStyle` | `String` | **not impl** | -- | M10 | -- | -- |
| `JSColumn.HeaderToolTip` | `String` | **not impl** | -- | M10 | -- | -- |
| `JSColumn.IsGrouped` | `Boolean` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.MaxRowsInCardView` | `Long` | **not impl** | -- | M6 | -- | -- |
| `JSColumn.MinRowsInCardView` | `Long` | **not impl** | -- | M6 | -- | -- |
| `JSColumn.ShowCaptionInCardView` | `Boolean` | **not impl** | -- | M6 | -- | -- |
| `JSColumn.SortOrder` | `jgexSortOrderConstants` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.SortType` | `jgexSortTypeConstants` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.TextAlignment` | `jgexAlignmentConstants` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | 1 in `006-unbound-rows`, 2 in `019-font-segoeui`, default 0 elsewhere |
| `JSColumn.TotalRowFormat` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.TotalRowPrefix` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.Visible` | `Boolean` | verified | `pvPaintRows`, `pvPaintDataRow`, `pvPaintHeaders` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `013-hidden-column` (False) vs default True |
| `JSColumn.Width` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow`, `pvPaintHeaders` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | 900/1200/1400/1500/1800tw across the corpus |
| `JSColumn.WordWrap` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.BackColor` | `OLE_COLOR` | verified | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `025-formatstyles-selection` (`0x004080`) vs the system default |
| `JSFormatStyle.BackgroundPicture` | `Picture` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.DrawModeBackGroundPicture` | `jgexDrawModePictureBackgroundConstants` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.FontBold` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.FontCharset` | `Integer` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.FontItalic` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.FontName` | `String` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.FontSize` | `Currency` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.FontStrikeThru` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.FontUnderline` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.ForeColor` | `OLE_COLOR` | verified | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `025-formatstyles-selection` (`0x80FFFF`) vs the system default |
| `JSFormatStyle.Picture` | `Picture` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.PictureDrawMode` | `jgexPictureDrawModeConstants` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.PictureHorzAlignment` | `jgexHorzPictureAlignmentConstants` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.PictureVertAlignment` | `jgexVertPictureAlignmentConstants` | **not impl** | -- | M10 | -- | -- |
| `JSFormatStyle.TextAlignment` | `jgexAlignmentConstants` | **not impl** | -- | M10 | -- | -- |
| `LeftCol` | `Integer` | verified | `pvPaintRows`, `pvPaintDataRow`, `pvPaintHeaders`, `pvUpdateScrollBars` | M3d | [`e95c15e`](../../commit/e95c15e1b1d35f09c9bbe557228ddc1329ed6c6f) | `029-hscrolled` scrolls to column 3; column 1 elsewhere |
| `MaskColor` | `OLE_COLOR` | **not impl** | -- | M10 | -- | -- |
| `MultiSelect` | `jgexMultiSelectConstants` | verified | input path, rendered via `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `017-multiselect` (contiguous 2-4) and `028-gridlines-vertical` (disjoint 1/3/6) |
| `NewRowPos` | `jgexNewRowPositionConstants` | **not impl** | -- | M5 | -- | -- |
| `Options` | `Long` | **not impl** | -- | unowned | -- | -- |
| `PreviewColumn` | `Integer` | **not impl** | -- | M10 | -- | -- |
| `PreviewRowIndent` | `Long` | **not impl** | -- | M10 | -- | -- |
| `PreviewRowLines` | `Integer` | **not impl** | -- | M10 | -- | -- |
| `RecordNavigator` | `Boolean` | verified | `pvPaintNavigator` | M3d | [`f9acd79`](../../commit/f9acd792b194f87374c89c6c2af0789785ec5e8b) | `030-record-navigator` (True) vs default False |
| `RecordNavigatorString` | `String` | weak | `pvNavLayout` | M3d | [`f9acd79`](../../commit/f9acd792b194f87374c89c6c2af0789785ec5e8b) | rendered at its default `Record:|of` only |
| `Redraw` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `Row` | `Long` | verified | `pvPaintDataRow`, `pvPaintRows` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `026-gridlines-horizontal` sets `Row` = 3 via `post`; row 1 elsewhere |
| `RowColorEven` | `OLE_COLOR` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `006-unbound-rows` and `027-gridlines-none` (`0xE0E0FF`) |
| `RowColorOdd` | `OLE_COLOR` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `006-unbound-rows` and `027-gridlines-none` (`0xFFE0E0`) |
| `RowCount` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow`, `pvUpdateScrollBars` | M3a | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | read-only view of `ItemCount`; the dark last-row gridline depends on it |
| `RowExpanded` | `Boolean` | **not impl** | -- | M4 | -- | -- |
| `RowHeaders` | `Boolean` | verified | `pvPaintHeaders`, `pvPaintRows` | M3c | [`54c6101`](../../commit/54c6101c20939be5df126be4f4b84e1c541926bb) | `014-rowheaders`, `017-multiselect`, `019-font-segoeui` vs `004-no-headers-no-rowheaders` |
| `RowHeight` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | never set explicitly -- exercised indirectly at 19/22/24/28/32px via the font scenarios x 3 dpi |
| `RowSelected` | `Boolean` | verified | `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `017-multiselect` (contiguous) and `028-gridlines-vertical` (disjoint) |
| `ScrollToolTipColumn` | `Integer` | **not impl** | -- | M10 | -- | -- |
| `ScrollToolTips` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `SelectedItems` | `JSSelectedItems` | verified | `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `017-multiselect` (contiguous) and `028-gridlines-vertical` (disjoint) |
| `SelectionStyle` | `jgexSelectionStyleConstants` | **not impl** | -- | M10 | -- | -- |
| `ShowEmptyFields` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `ShowToolTips` | `Boolean` | **not impl** | -- | M10 | -- | -- |
| `SortKeys` | `JSSortKeys` | **not impl** | -- | M4 | -- | -- |
| `UseEvenOddColor` | `Boolean` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `006-unbound-rows` (True) vs default False |
| `View` | `jgexViewConstants` | **not impl** | -- | M6 | -- | -- |

## Closing the gaps

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
covers the first 28 scenarios only: `029-hscrolled` and `030-record-navigator`
were recorded after the machine left 150%, so the scrollbar band and the
navigator have no 150% golden yet. That pair is also the third data point the
navigator's box-width term (`tmHeight + BandH + 16`) still wants -- two scales
cannot separate a constant from a metric-derived one.

Note on goldens and the system accent: it changed twice during this work
(`0x0078D7` -> `0x0078D4` -> back), and every selected row carries it. A golden
set recorded in an earlier session can therefore fail on colour alone while the
geometry is perfect -- the tell is a diff whose only pairs are the accent and its
XOR complement. Re-record rather than debug. `golden/144` currently holds the
older accent and wants a re-record next time the machine is at 150%.

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

## Maintenance

Update this file in the same commit that changes a property's status.

Two traps when regenerating it mechanically: paint routines read some values
through their **public property getter** rather than the member (`RowCount`,
`SelectedItems`), and some getters **compute** instead of returning a member
(`RowHeight`, `ColumnHeaderHeight`, `DefaultColumnWidth`). A member-only scan
misses both groups. Scenario coverage likewise lives in **two** places -- the
`props` block and the `post` block -- and `FirstItem` is only ever set in `post`.
