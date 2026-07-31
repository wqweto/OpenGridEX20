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
| verified | 19 | 15% |
| weak | 8 | 6% |
| unverified | 15 | 12% |
| partial | 1 | 1% |
| **not implemented** | **83** | **66%** |
| **total** | **126** | |

**43 of 126 (34%) are read by the paint path**; the other **83 (66%) are not
implemented** -- they store and return their value, and round-trip through the
snapshot corpus, but the renderer never looks at them.

Of the 43 the renderer does read, only **19 (15% of the surface) are proven**
against the original at more than one value. **15 have no covering test at all**,
and 8 more are exercised at a single value.

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
| `BackColor` | `OLE_COLOR` | **unverified** | `pvPaintDataRow`, `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | set by `gridlines-dots-colors`, which has no rows -- `0xC0FFC0` is absent from the golden |
| `BackColorBkg` | `OLE_COLOR` | weak | `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `gridlines-dots-colors` -- renders, one value only |
| `BackColorGBBox` | `OLE_COLOR` | **unverified** | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario sets it |
| `BackColorHeader` | `OLE_COLOR` | **unverified** | `pvPaintHeaderCell`, `pvPaintRowHeader`, `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario sets it |
| `BackColorInfoText` | `OLE_COLOR` | **unverified** | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario sets it |
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
| `DefaultColumnWidth` | `Long` | **unverified** | via `JSColumns.Add` at init | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario sets it |
| `DefaultGroupMode` | `jgexDefaultGroupModeConstants` | **not impl** | -- | M4 | -- | -- |
| `DetectRowDrag` | `Boolean` | **not impl** | -- | unowned | -- | -- |
| `EditMode` | `jgexEditModeConstants` | **not impl** | -- | M5 | -- | -- |
| `EmptyRows` | `Boolean` | verified | `pvPaintRows` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `empty-rows` (True) vs default False |
| `Enabled` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `FirstItem` | `Long` | verified | `pvPaintRows`, `pvUpdateScrollBars` | M3d | [`e95c15e`](../../commit/e95c15e1b1d35f09c9bbe557228ddc1329ed6c6f) | `scrolled` sets 4 in its `post` block; 1 elsewhere |
| `FmtConditions` | `JSFmtConditions` | **not impl** | -- | M9 | -- | -- |
| `Font` | `Font` | verified | `pvPaintRows`, `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `font-large`, `font-tahoma`, `font-segoeui` + default |
| `ForeColor` | `OLE_COLOR` | **unverified** | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario sets it |
| `ForeColorHeader` | `OLE_COLOR` | **unverified** | `pvPaintHeaderCell` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario sets it |
| `ForeColorInfoText` | `OLE_COLOR` | **unverified** | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario sets it |
| `ForeColorRowGroup` | `OLE_COLOR` | **not impl** | -- | M4 | -- | -- |
| `FormatStyles` | `JSFormatStyles` | **unverified** | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | only the system styles at their defaults -- see the `JSFormatStyle.*` rows |
| `FrozenColumns` | `Integer` | **not impl** | -- | M3d | -- | -- |
| `GridImages` | `JSGridImages` | **not impl** | -- | M9 | -- | -- |
| `GridLines` | `jgexGridLinesConstants` | weak | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | only ever set to `jgexGLBoth`, which is the default -- the other three modes are unrendered |
| `GridLinesColor` | `OLE_COLOR` | **unverified** | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | same scenario, same reason -- no blue pixel in the golden |
| `GridLineStyle` | `jgexGridLineStyleConstants` | **unverified** | `pvPenStyle` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `gridlines-dots-colors` has no rows, so dotted data gridlines are never drawn |
| `GroupByBoxInfoText` | `String` | **unverified** | `pvPaintGroupByBox` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario sets it |
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
| `JSColumn.ColPosition` | `Integer` | **unverified** | `pvPaintRows` (via `ItemByPosition`) | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | no scenario reorders columns |
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
| `JSFormatStyle.BackColor` | `OLE_COLOR` | **unverified** | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | only the `SelectedRow` system default is rendered |
| `JSFormatStyle.BackgroundPicture` | `Picture` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.DrawModeBackGroundPicture` | `jgexDrawModePictureBackgroundConstants` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontBold` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontCharset` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontItalic` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontName` | `String` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontSize` | `Currency` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontStrikeThru` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.FontUnderline` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.ForeColor` | `OLE_COLOR` | **unverified** | `pvSelColors` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | only the `SelectedRow` system default is rendered |
| `JSFormatStyle.Picture` | `Picture` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.PictureDrawMode` | `jgexPictureDrawModeConstants` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.PictureHorzAlignment` | `jgexHorzPictureAlignmentConstants` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.PictureVertAlignment` | `jgexVertPictureAlignmentConstants` | **not impl** | -- | M9 | -- | -- |
| `JSFormatStyle.TextAlignment` | `jgexAlignmentConstants` | **not impl** | -- | M9 | -- | -- |
| `LeftCol` | `Integer` | **partial** | `pvUpdateScrollBars` only | M3d | [`e95c15e`](../../commit/e95c15e1b1d35f09c9bbe557228ddc1329ed6c6f) | sets the horizontal thumb; the paint path ignores it, so columns do not scroll |
| `MaskColor` | `OLE_COLOR` | **not impl** | -- | M9 | -- | -- |
| `MultiSelect` | `jgexMultiSelectConstants` | weak | input path, rendered via `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `multiselect` (True) vs default False |
| `NewRowPos` | `jgexNewRowPositionConstants` | **not impl** | -- | M5 | -- | -- |
| `Options` | `Long` | **not impl** | -- | unowned | -- | -- |
| `PreviewColumn` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `PreviewRowIndent` | `Long` | **not impl** | -- | M9 | -- | -- |
| `PreviewRowLines` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `RecordNavigator` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `RecordNavigatorString` | `String` | **not impl** | -- | M3d | -- | -- |
| `Redraw` | `Boolean` | **not impl** | -- | M3d | -- | -- |
| `Row` | `Long` | weak | `pvPaintDataRow`, `pvPaintRows` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `multiselect` + ModelTests keynav; no scenario sets `Row` directly |
| `RowColorEven` | `OLE_COLOR` | weak | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `unbound-rows` -- one value only |
| `RowColorOdd` | `OLE_COLOR` | weak | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `unbound-rows` -- one value only |
| `RowCount` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow`, `pvUpdateScrollBars` | M3a | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | read-only view of `ItemCount`; the dark last-row gridline depends on it |
| `RowExpanded` | `Boolean` | **not impl** | -- | M4 | -- | -- |
| `RowHeaders` | `Boolean` | verified | `pvPaintHeaders`, `pvPaintRows` | M3c | [`54c6101`](../../commit/54c6101c20939be5df126be4f4b84e1c541926bb) | `rowheaders`, `multiselect`, `font-segoeui` vs `no-headers-no-rowheaders` |
| `RowHeight` | `Long` | verified | `pvPaintRows`, `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | never set explicitly -- exercised indirectly at 19/22/24/28/32px via the font scenarios x 3 dpi |
| `RowSelected` | `Boolean` | weak | `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `multiselect` -- one selection shape only |
| `ScrollToolTipColumn` | `Integer` | **not impl** | -- | M9 | -- | -- |
| `ScrollToolTips` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `SelectedItems` | `JSSelectedItems` | weak | `pvIsRowSelected` | M3d | [`a5d9590`](../../commit/a5d959050e1911c9b78ce1a64c2743bdba83d47c) | `multiselect` -- one selection shape only |
| `SelectionStyle` | `jgexSelectionStyleConstants` | **not impl** | -- | M9 | -- | -- |
| `ShowEmptyFields` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `ShowToolTips` | `Boolean` | **not impl** | -- | M9 | -- | -- |
| `SortKeys` | `JSSortKeys` | **not impl** | -- | M4 | -- | -- |
| `UseEvenOddColor` | `Boolean` | verified | `pvPaintDataRow` | M3c | [`db8b3ba`](../../commit/db8b3ba335f1e7f03c8022049ba9e8fd6d1dbed9) | `unbound-rows` (True) vs default False |
| `View` | `jgexViewConstants` | **not impl** | -- | unowned | -- | -- |

## Closing the gaps

The **unverified** block is cheap to clear because the harness already supports
everything needed -- it is scenario authoring, not control work:

1. **Give `gridlines-dots-colors` rows.** One edit fixes four entries at once
   (`GridLinesColor`, `GridLineStyle`, `GridLines`, `BackColor`), none of which
   can be proven while the scenario renders no data area.
2. **Add a `colors-*` pair** setting `ForeColor`, `BackColorHeader`,
   `ForeColorHeader`, `BackColorGBBox`, `BackColorInfoText`, `ForeColorInfoText`
   and `GroupByBoxInfoText` to distinctive non-system values, in two variants so
   each gets two rendered values.
3. **Add `formatstyles-selection`** overriding the `SelectedRow` system style,
   which is the only way `pvSelColors` -- and with it `JSFormatStyle.BackColor`
   and `JSFormatStyle.ForeColor` -- is exercised beyond defaults.
4. **Add `column-order`** setting `JSColumn.ColPosition`, the one paint-consumed
   column member with no coverage.
5. **Vary the weak ones**: a second value for `BackColorBkg`, `RowColorEven`,
   `RowColorOdd`, `SelectedItems`/`RowSelected`, and the three non-default
   `GridLines` modes.

Anything that then disagrees with the original becomes a control fix, which is
the point of doing this before more surface is built on top.

## Maintenance

Update this file in the same commit that changes a property's status.

Two traps when regenerating it mechanically: paint routines read some values
through their **public property getter** rather than the member (`RowCount`,
`SelectedItems`), and some getters **compute** instead of returning a member
(`RowHeight`, `ColumnHeaderHeight`, `DefaultColumnWidth`). A member-only scan
misses both groups. Scenario coverage likewise lives in **two** places -- the
`props` block and the `post` block -- and `FirstItem` is only ever set in `post`.
