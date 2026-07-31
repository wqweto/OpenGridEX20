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
`gridlines-dots-colors`, yet that scenario declares **no rows**, so no data
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
| verified | 42 | 33% |
| weak | 0 | 0% |
| unverified | 0 | 0% |
| partial | 1 | 1% |
| **not implemented** | **83** | **66%** |
| **total** | **126** | |

**43 of 126 (34%) are read by the paint path**; the other **83 (66%) are not
implemented** -- they store and return their value, and round-trip through the
snapshot corpus, but the renderer never looks at them.

**Every property the renderer reads is now proven** against the original at two or
more distinct values, except `LeftCol`, which is only partly consumed because
horizontal scrolling is unimplemented.

Getting there took eight scenarios and turned up **eleven** real defects. Each had
survived because the corpus could not see it: solid gridlines hide both dotted
rules, a black default `ForeColor` makes a wrong XOR mask invisible, and no
scenario had ever set `GridLines` to anything but its default.

Not-implemented properties by owning milestone:

| milestone | count |
|---|---:|
| M9 styling extras | 38 |
| **unowned** | **17** |
| M4 sorting/grouping | 16 |
| M3d (remainder) | 7 |
| M5 editing | 5 |

Two things this ranking makes obvious. **M9 is by far the largest single block of
unimplemented painting** -- 38 properties, mostly the `JSFormatStyle` font and
picture families -- which is a much bigger milestone than its one-line roadmap
entry suggests. And the `unowned` group is a real roadmap gap: **card view has no
milestone**. M3 declares it out of scope and nothing later picks it up, yet
`View`, `CardBorders`, `CardCaptionPrefix`, `CardWidth`, `CardSpacing` and the
`JSColumn` card members are public surface and appear in the sample corpus. The
drag/resize affordances (`AllowColumnDrag`, `AllowRowSizing`, `DetectRowDrag`,
`JSColumn.AllowSizing`) and `Options` have no home either.

## The matrix

Scenario names refer to `test\VisualDiff\scenarios\*.json`; each is verified at
96, 120 and 144 dpi. The `Commit` column records where the property started
affecting pixels, not where it was first stored -- every member was declared in
the M2 storage commit, so that hash carries no information about painting.

| Property | Type | Status | Paint routine | Milestone | Commit | Test |
|---|---|---|---|---|---|---|
| `AllowAddNew` | `Boolean` | **not impl** | -- | M5 | -- | -- |
| `AllowCardSizing` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `AllowColumnDrag` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `AllowRowSizing` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `AutomaticArrange` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `BackColor` | `OLE_COLOR` | verified | `pvPaintDataRow`, `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-rows` (`0xC0FFC0`) vs default |
| `BackColorBkg` | `OLE_COLOR` | verified | `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-rows` (`0x404040`), `gridlines-dots-colors` (`0x808080`) |
| `BackColorGBBox` | `OLE_COLOR` | verified | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-chrome` (`0x604020`) vs default |
| `BackColorHeader` | `OLE_COLOR` | verified | `pvPaintHeaderCell`, `pvPaintRowHeader`, `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-chrome` (`0x8000FF`) vs default |
| `BackColorInfoText` | `OLE_COLOR` | verified | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-chrome` (`0x30A0C0`) vs default |
| `BackColorRowGroup` | `OLE_COLOR` | **not impl** | -- | M4 | -- | -- |
| `BorderStyle` | `jgexBorderStyleConstants` | **not impl** | -- | M9 | -- | -- |
| `CardBorders` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `CardCaptionPrefix` | `String` | **not impl** | -- | unowned | -- | -- |
| `CardSpacing` | `Long` | **not impl** | -- | unowned | -- | -- |
| `CardWidth` | `Long` | **not impl** | -- | unowned | -- | -- |
| `Col` | `Integer` | **not impl** | -- | M3d | -- | -- |
| `ColumnAutoResize` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `ColumnHeaderFont` | `Font` | verified | `pvPaintHeaders` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `font-large`, `font-tahoma`, `font-segoeui` + MS Sans Serif default |
| `ColumnHeaderHeight` | `Long` | verified | `pvPaintHeaders`, `pvPaint` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `headers-flat` sets 400tw; font-derived defaults elsewhere |
| `ColumnHeaders` | `Boolean` | verified | `pvPaint` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `no-headers-no-rowheaders` (False) vs default True |
| `Columns` | `JSColumns` | verified | `pvPaintHeaders`, `pvPaintDataRow`, `pvPaintRows`, `pvPaintRowMarquee` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | all 20 scenarios -- see the `JSColumn.*` rows |
| `ContinuousScroll` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `DefaultColumnWidth` | `Long` | verified | via `JSColumns.Add` at init | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `default-column-width` (2100tw -> 140px columns) vs default |
| `DefaultGroupMode` | `jgexDefaultGroupModeConstants` | **not impl** | -- | M4 | -- | -- |
| `DetectRowDrag` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `EditMode` | `jgexEditModeConstants` | **not impl** | -- | M5 | -- | -- |
| `EmptyRows` | `Boolean` | verified | `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `empty-rows` (True) vs default False |
| `Enabled` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `FirstItem` | `Long` | verified | `pvPaintRows`, `pvUpdateScrollBars` | M3d | [`e95c15e`](../../commit/e95c15e1b1d35f09c9bbe557228ddc1329ed6c6f) | `scrolled` sets 4 in its `post` block; 1 elsewhere |
| `FmtConditions` | `JSFmtConditions` | **not impl** | -- | M9 | -- | -- |
| `Font` | `Font` | verified | `pvPaintRows`, `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `font-large`, `font-tahoma`, `font-segoeui` + default |
| `ForeColor` | `OLE_COLOR` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-rows` (`0xFF0000`) vs default -- also drives the marquee XOR mask |
| `ForeColorHeader` | `OLE_COLOR` | verified | `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-chrome` (`0x00FFFF`) vs default |
| `ForeColorInfoText` | `OLE_COLOR` | verified | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-chrome` (`0x0020FF`) vs default |
| `ForeColorRowGroup` | `OLE_COLOR` | **not impl** | -- | M4 | -- | -- |
| `FormatStyles` | `JSFormatStyles` | verified | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `formatstyles-selection` overrides the `SelectedRow` system style |
| `FrozenColumns` | `Integer` | **not impl** | -- | M3d | -- | -- |
| `GridImages` | `JSGridImages` | **not impl** | -- | M9 | -- | -- |
| `GridLines` | `jgexGridLinesConstants` | verified | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | all four modes: default Both, plus `gridlines-none`, `gridlines-vertical`, `gridlines-horizontal` |
| `GridLinesColor` | `OLE_COLOR` | verified | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-rows` (`0x0000FF`) vs default |
| `GridLineStyle` | `jgexGridLineStyleConstants` | verified | `pvPenStyle` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-rows` (dots), `gridlines-vertical` (dashes), solid default |
| `GroupByBoxInfoText` | `String` | verified | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `colors-chrome` custom text vs default |
| `GroupByBoxVisible` | `Boolean` | verified | `pvPaint` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `headers-noborder`, `headers-single3d`, `headers-singleflat` vs default |
| `GroupFooterStyle` | `jgexGroupFooterStyleConstants` | **not impl** | -- | M4 | -- | -- |
| `Groups` | `JSGroups` | **not impl** | -- | M4 | -- | -- |
| `HeaderStyle` | `jgexHeaderStyleConstants` | verified | `pvPaintHeaders`, `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `headers-flat`, `headers-noborder`, `headers-single3d`, `headers-singleflat` (4 values) |
| `HideSelection` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `ImageHeight` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `ImageWidth` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `ItemCount` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow` (via `RowCount`) | M3a | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | row counts 2/3/5/6/14/16 across the corpus |
| `JSColumn.AggregateFunction` | `jgexAggregateFunctionConstants` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.AllowSizing` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `JSColumn.ButtonStyle` | `jgexButtonStyleConstants` | **not impl** | -- | M5 | -- | -- |
| `JSColumn.Caption` | `String` | verified | `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | distinct captions in every scenario |
| `JSColumn.CardCaption` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `JSColumn.CardIcon` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `JSColumn.CellStyle` | `String` | **not impl** | -- | M9 | -- | -- |
| `JSColumn.ColPosition` | `Integer` | verified | `pvPaintRows` (via `ItemByPosition`) | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `column-order` reorders 3 columns to 1/2/3 from declaration order |
| `JSColumn.ColumnType` | `jgexColumnTypeConstants` | **not impl** | -- | M9 | -- | -- |
| `JSColumn.DefaultIcon` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `JSColumn.EditType` | `jgexEditTypeConstants` | **not impl** | -- | M5 | -- | -- |
| `JSColumn.Format` | `String` | **not impl** | -- | M9 | -- | -- |
| `JSColumn.GroupEmptyStringCaption` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.GroupFormat` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.GroupPrefix` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.HeaderAlignment` | `jgexAlignmentConstants` | verified | `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `headers-flat` sets 2; default 0 elsewhere |
| `JSColumn.HeaderIcon` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `JSColumn.HeaderStyle` | `String` | **not impl** | -- | M9 | -- | -- |
| `JSColumn.HeaderToolTip` | `String` | **not impl** | -- | M9 | -- | -- |
| `JSColumn.IsGrouped` | `Boolean` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.MaxRowsInCardView` | `Long` | **not impl** | -- | unowned | -- | -- |
| `JSColumn.MinRowsInCardView` | `Long` | **not impl** | -- | unowned | -- | -- |
| `JSColumn.ShowCaptionInCardView` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `JSColumn.SortOrder` | `jgexSortOrderConstants` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.SortType` | `jgexSortTypeConstants` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.TextAlignment` | `jgexAlignmentConstants` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | 1 in `unbound-rows`, 2 in `font-segoeui`, default 0 elsewhere |
| `JSColumn.TotalRowFormat` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.TotalRowPrefix` | `String` | **not impl** | -- | M4 | -- | -- |
| `JSColumn.Visible` | `Boolean` | verified | `pvPaintRows`, `pvPaintDataRow`, `pvPaintHeaders` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `hidden-column` (False) vs default True |
| `JSColumn.Width` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow`, `pvPaintHeaders` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | 900/1200/1400/1500/1800tw across the corpus |
| `JSColumn.WordWrap` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.BackColor` | `OLE_COLOR` | verified | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `formatstyles-selection` (`0x004080`) vs the system default |
| `JSFormatStyle.BackgroundPicture` | `Picture` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.DrawModeBackGroundPicture` | `jgexDrawModePictureBackgroundConstants` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontBold` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontCharset` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontItalic` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontName` | `String` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontSize` | `Currency` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontStrikeThru` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontUnderline` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.ForeColor` | `OLE_COLOR` | verified | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `formatstyles-selection` (`0x80FFFF`) vs the system default |
| `JSFormatStyle.Picture` | `Picture` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.PictureDrawMode` | `jgexPictureDrawModeConstants` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.PictureHorzAlignment` | `jgexHorzPictureAlignmentConstants` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.PictureVertAlignment` | `jgexVertPictureAlignmentConstants` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.TextAlignment` | `jgexAlignmentConstants` | **not impl** | -- | M9 | -- | -- |
| `LeftCol` | `Integer` | **partial** | `pvUpdateScrollBars` only | M3d | [`e95c15e`](../../commit/e95c15e1b1d35f09c9bbe557228ddc1329ed6c6f) | sets the horizontal thumb; the paint path ignores it, so columns do not scroll |
| `MaskColor` | `OLE_COLOR` | **not impl** | -- | M9 | -- | -- |
| `MultiSelect` | `jgexMultiSelectConstants` | verified | input path, rendered via `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `multiselect` (contiguous 2-4) and `gridlines-vertical` (disjoint 1/3/6) |
| `NewRowPos` | `jgexNewRowPositionConstants` | **not impl** | -- | M5 | -- | -- |
| `Options` | `Long` | **not impl** | -- | unowned | -- | -- |
| `PreviewColumn` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `PreviewRowIndent` | `Long` | **not impl** | -- | M9 | -- | -- |
| `PreviewRowLines` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `RecordNavigator` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `RecordNavigatorString` | `String` | **not impl** | -- | M3d | -- | -- |
| `Redraw` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `Row` | `Long` | verified | `pvPaintDataRow`, `pvPaintRows` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `gridlines-horizontal` sets `Row` = 3 via `post`; row 1 elsewhere |
| `RowColorEven` | `OLE_COLOR` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `unbound-rows` and `gridlines-none` (`0xE0E0FF`) |
| `RowColorOdd` | `OLE_COLOR` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `unbound-rows` and `gridlines-none` (`0xFFE0E0`) |
| `RowCount` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow`, `pvUpdateScrollBars` | M3a | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | read-only view of `ItemCount`; the dark last-row gridline depends on it |
| `RowExpanded` | `Boolean` | **not impl** | -- | M4 | -- | -- |
| `RowHeaders` | `Boolean` | verified | `pvPaintHeaders`, `pvPaintRows` | M3c | [`54c6101`](../../commit/54c6101c20939be5df126be4f4b84e1c541926bb) | `rowheaders`, `multiselect`, `font-segoeui` vs `no-headers-no-rowheaders` |
| `RowHeight` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | never set explicitly -- exercised indirectly at 19/22/24/28/32px via the font scenarios x 3 dpi |
| `RowSelected` | `Boolean` | verified | `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `multiselect` (contiguous) and `gridlines-vertical` (disjoint) |
| `ScrollToolTipColumn` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `ScrollToolTips` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `SelectedItems` | `JSSelectedItems` | verified | `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `multiselect` (contiguous) and `gridlines-vertical` (disjoint) |
| `SelectionStyle` | `jgexSelectionStyleConstants` | **not impl** | -- | M9 | -- | -- |
| `ShowEmptyFields` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `ShowToolTips` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `SortKeys` | `JSSortKeys` | **not impl** | -- | M4 | -- | -- |
| `UseEvenOddColor` | `Boolean` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `unbound-rows` (True) vs default False |
| `View` | `jgexViewConstants` | **not impl** | -- | unowned | -- | -- |

## Closing the gaps

The unverified block is **cleared**. Five scenarios did it, and each defect they
found had been invisible for a structural reason worth remembering: a rule can
only be wrong in a way the corpus can see.

| scenario | closed | found |
|---|---|---|
| `colors-rows` | `BackColor`, `ForeColor`, `GridLinesColor`, `GridLineStyle`, `BackColorBkg` | dotted-pen duty cycle, dotted-gap colour, marquee XOR mask |
| `colors-chrome` | the 6 header / group-by-box / info-text colours | -- |
| `formatstyles-selection` | `FormatStyles`, `JSFormatStyle.BackColor`, `JSFormatStyle.ForeColor` | -- |
| `column-order` | `JSColumn.ColPosition` | invalid `ColPosition` blanks the grid |
| `default-column-width` | `DefaultColumnWidth` | new columns ignored it, defaulting to 1000 **px** |

The three `gridlines-*` scenarios each carry 5-6 properties rather than one, so
clearing the weak block cost three captures instead of seven -- recording is the
slow part of the loop, so scenarios are packed deliberately.

The corpus is recorded and verified at **96, 120 and 144 dpi**.

Note on goldens and the system accent: it changed twice during this work
(`0x0078D7` -> `0x0078D4` -> back), and every selected row carries it. A golden
set recorded in an earlier session can therefore fail on colour alone while the
geometry is perfect -- the tell is a diff whose only pairs are the accent and its
XOR complement. Re-record rather than debug. `golden/144` currently holds the
older accent and wants a re-record next time the machine is at 150%.

What remains for M6: **`LeftCol`** stays `partial` until horizontal scrolling
lands (M3d remainder).

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
