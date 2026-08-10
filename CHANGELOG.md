# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added

- Stubbed complete public API from the original `GridEX20.ocx` type library (`doc/GridEX20.idl`): `GridEX` and `GEXPreview` controls, 24 dependent classes, 44 enums, 58 events and all methods/properties incl. help strings
- `OpenGEXHelper` helper type library (`typelib/OpenGEXHelper.odl`, compiled with `mktyplib`) exposing the `IObjectSafety` interface
- `IObjectSafety` implementation on both controls reporting safe for scripting/initialization
- `src/make.bat` build script which restores CRLF line endings and compiles `OpenGridEX20.vbp` with VB6
- Project README with repository layout and build instructions
- `ROADMAP.md` with implementation milestones M0-M11 and testing strategy
- `tools/DumpSurface.vbs` typelib surface dumper (TLI based) for source-compatibility checks against `doc/GridEX20.idl`
- `tools/CompareIdl.ps1` source-compatibility gate: canonical diff of `doc/OpenGridEX20.idl` (OleView dump of the built OCX) against `doc/GridEX20.idl`; currently passes with zero differences
- Snapshot engine (M1): `tools/GenProfiles.ps1` generates `tools/common/mdProfiles.bas` (267 property profiles parsed from the stubs incl. runtime-only/collection classification), `tools/common/mdSnapshot.bas` profile-driven late-bound object model walker building `mdJson.bas` documents, JSON schema frozen in `tools/common/SCHEMA.md` (all JSON goes through `mdJson.bas` using its canonical idioms; `$schema`/`$errors` meta keys addressed via JSONPath dot form `$.$schema`)
- `test/Snapshot` smoke test hosting both controls and dumping their object model to JSON; builds and passes against the stub OCX

### Added (Documentation)

- `doc/Help`: original `gridex2000.chm` converted to markdown (32 files: topics combined per TOC section -- Properties/Methods/Events per control, one file per `JS*` object/collection, all code samples in `Examples.md`, object model diagrams in `images/`); `index.md` reproduces the CHM table of contents and all 2900+ cross-topic links (See Also, Applies To, Example, image maps) resolve to `file.md#anchor` form

### Fixed

- `doc/Help`: repaired 37 syntax lines across 16 files where CHM conversion collapsed bold/italic markers into `***`/`****` runs that rendered as literal asterisks (e.g. `**GetData** **(***format* **As Integer)**`); parentheses normalized to the plain-paren convention used by the rest of the docs, plus the fused `RowIndex` syntax line in `Methods.md` split and its parameter italicized
- `doc/Help`: repaired 10 more lines where bold markers were fused to adjacent words (e.g. `**For DAO DataMode:**The`, `The**GridEX**`, `**SaveLayout**method`), an apostrophe trapped inside bold (`**GridEX'**s` -> `**GridEX**'s`) and a fused sentence break in `Examples.md`
- `mdJson.bas`: `JsonDump` now emits `{}` for `Nothing` objects (respecting `CompoundChars`) instead of an empty string that produced malformed JSON when a `Nothing` value was stored inside a document

### Added (M2 -- object model)

- Implemented state storage and collection semantics in all 23 `JS*` classes (keyed collections with reindexing, lazy-created owned objects, `Friend` wiring via `frInit`/`frSet*`) and member-backed properties in both controls with sensible defaults; public typelib surface unchanged (binary compatibility preserved)
- `src/mdGlobals.bas` with non-raising `SearchCollection` over the raw `IVBCollection` vtable added to `typelib/OpenGEXHelper.odl` (keeps `Break on All Errors` debugging usable)
- `JSRowData` reworked as a virtual view over control-internal row storage: the control keeps per-row data in a private `UcsRowData` UDT array (with a nested per-cell `UcsCellData` array -- value/icon/display value/cell style/cell picture -- lazily sized to column count) with a lazily created `JSRowData` wrapper per row (`GetRowData`), while `JSRowData` holds only a weak reference to the owning control -- a typed `m_pOwner As GridEX` member written raw via `CopyMemory` (no refcount cycle) -- and delegates every property to `frRow*` friend accessors; `UserControl_Terminate` detaches outstanding wrappers (`frTerm`) and `Class_Terminate` zeroes the raw pointer so VB6 never releases a control it did not AddRef; the default `Value` member raises error 91 when the wrapper is orphaned; `mdGlobals.bas` gained the shared `CopyMemory` declare with `PTR_SIZE`/`NULL_PTR` consts and a `LongPtr` shim enum for future twinBASIC x64 builds
- `test/ModelTests`: 19 new JSRowData assertions (59 total) -- lazy wrapper caching, virtual storage visible through twin references, cell array growth on late `Columns.Add`, and a weak-ref proof using a disposable `frmWeak` host: unloading the host detaches the wrapper (orphaned default-member access raises error 91) which a strong reference would have prevented
- `tools/common/mdImport.bas`: profile-driven snapshot import (per-collection `Add` mappings, `C2Str`/`pvAssignVariant` Variant handling, contextual error re-raise)
- `test/ModelTests`: 40 assertions over collection semantics plus GridEX/GEXPreview export->import->export round-trips -- all passing losslessly
- Coding style additions applied project-wide: `DefObj A-Z`, section separator banners
- `tools/OpenGEXAddin`: export-only VB6 add-in (`IDTExtensibility2`, no UI) that snapshots every `GridEX20.*`/`OpenGridEX20.*` control on the active project's form designers via the shared engine, augmented with a `raw` section of flattened `.frm` propbag keys (`IntProp*`, `Column(n)` frx refs, `MethodHoldFields`, ...); activates only when `OPENGEX_SNAPSHOT_DIR` is set so it stays dormant in normal IDE sessions
- `tools/OpenGEXAddin/export.ps1` batch driver: flips the add-in to load-on-startup in `vbaddin.ini` (both `%WINDIR%` and VirtualStore copies), launches the VB6 IDE per original Janus sample project, waits for the done-marker and kills the IDE; two frx-less damaged samples (Unbound Array/UDTs) fall back to raw-only capture
- `test/snapshots`: design-time snapshot corpus recorded from the original control -- 28 JSONs from 20 sample projects (ADO 1-5, Advanced, Combo DropDowns, Custom Edit, Icon columns, OLE Drag, Preview Rows, 3 Print/PrintPreview, Unbound 1/2/Array/UDTs/Collection, FormatStyles)
- M2 exit reached: corpus round-trip test imports each snapshot into fresh control instances and diffs the re-export -- all 26 full snapshots round-trip losslessly (86 assertions total) after canonicalization that drops props the original could not read at design time (its `$errors`) and quantizes expected font sizes through a real `StdFont` (OLE `StdFont` snaps bitmap-font sizes, e.g. MS Sans Serif 7.8 reads back 8.25, while the original's own font object preserves 7.8)
- Model fidelity fixes surfaced by the corpus: `JSColumn.ValueList` returns `Nothing` until `HasValueList` is set (matches original), `JSPrinterProperties.PageHeaderFont`/`PageFooterFont` lazily create default fonts instead of returning `Nothing`, and `ImportObject` now imports in two passes (scalars first) so gate props like `HasValueList` are in effect before dependent compound props
- `tools/common/mdUtils.bas`: shared `C2Obj`/`C2Dbl`/`C2Lng` non-failing Variant coercions (modeled after `C2Str`, no `On Error Resume Next` helpers in consumers) plus `AssignVariant`/`ReadTextFile`/`WriteTextFile`/`EnumFiles` replacing per-project private copies in ModelTests, Snapshot and the add-in

### Added (M3a -- unbound data pipeline)

- `ROADMAP.md`: M3 split into M3a (unbound data pipeline, no pixels), M3b (VisualDiff harness + golden recorder), M3c (static table painting) and M3d (scrolling + input); card view explicitly out of M3 scope
- Unbound fetch pipeline: in `DataMode = jgexUnbound` any cell accessor lazily fires `UnboundReadData(RowIndex, Bookmark, Values)` once per row -- the `Values` buffer is the row's own `JSRowData` wrapper so the handler writes straight into control storage; re-entrant reads inside the handler do not re-fire (`Fetched` flag set first)
- Cache invalidation: `Rebind` full-resets rows (values, bookmarks, row props) and `Refetch` resets values only (bookmarks survive); both honor the `HoldSortSettings` property (default True) as default for their optional parameter, clearing `SortKeys`/`Groups` when not held; `RefreshRowIndex`/`RefreshRowBookmark` mark a single row for lazy refetch
- `RowBookmark(rowindex)` get/let backed by row storage and passed to `UnboundReadData`; `RowIndex(position)` returns the 1:1 original index (0 for group rows), `IsGroupItem` checks the stored `RowType`
- Navigation events: `Row`/`Col` lets fire `RowColChange(LastRow, LastCol)` with the previous position on actual change only; `FirstItem` fires `FirstItemChange`
- `test/ModelTests`: 24 new assertions (110 total) driving a fresh control on the `frmWeak` host (which doubles as event sink writing an ordered `EventLog`): fetch-once per row, lazy refetch per row/bookmark, bookmark round-trip into the event, `Refetch` vs `Rebind` reset scope, sort-hold semantics and exact navigation event sequences

### Added (M3b -- VisualDiff harness + golden corpus)

- `test/VisualDiff`: pixel-diff harness with no compile-time OCX reference -- the control under test (original `GridEX20.GridEX` or our `OpenGridEX20.GridEX`) is created at runtime via `Licenses.Add`/`Controls.Add`, scenario props applied through the shared import engine, unbound rows fed late-bound through `VBControlExtender.ObjectEvent`, and the control window blitted from its *own client DC at (0,0)* to 24bpp DIB (`GetClientRect`-sized) -- no screen coordinate mapping, so DPI virtualization on high-DPI displays (e.g. 120%) cannot offset or rescale the capture; capture-until-stable loop defeats paint races
- JSON scenarios (`scenarios/*.json`, `opengex-visualdiff/1`): window size, `before-show`/`after-show` apply ordering, snapshot-style `props` and optional `unbound.rows` data; initial set of 6 (defaults, flat headers, no headers/row headers, gridline style + colors, even/odd unbound rows, after-show unbound)
- Modes: `record` (golden corpus from the original control), `selftest` (fresh original capture vs golden must diff to zero -- passes 6/6), `verify` (our control vs golden, gated by `make.bat`); `record.bat` = record + selftest; diff failures report pixel count + bounding box and save `output\<name>.actual.bmp`
- M3b exit reached: goldens recorded (`golden/*.png` committed via GDI+ `SaveBitmapAsPng`/`LoadPng`), harness self-validates; `verify` fails all by design -- our control paints nothing yet, which is the M3c red baseline
- Data scenarios record correctly against the original: it cannot switch `DataMode`/`ItemCount` at runtime, so the harness feeds it a fabricated in-memory ADO recordset post-show with `DataField` mappings (`F1..Fn`) wired onto the imported columns; `HoldFields` is called after prop import *and* re-armed around the recordset assignment because the original resets the imported column layout whenever its display (re)initializes over a recordset -- the samples persist exactly this as `MethodHoldFields`; our control is fed through its unbound pipeline instead
- `dump` mode snapshots the runtime-created control state to `output\<name>.dump.json` for divergence diagnosis; `mdImport.bas` gained an `OPENGEX_IMPORT_TRACE`-gated per-prop trace; capture runs against the control window's own client DC, actual-vs-golden artifacts land in `output\`
- Recon findings for M3c defaults parity (probe vs our stubs): original defaults are `DataMode=jgexDAO`, `GroupByBoxVisible=True`, `RowHeaders=False`, `DefaultColumnWidth=1500`, `RowHeight=285`, `BackColorBkg=vbWindowBackground`, `HeaderStyle=0`, plus two default 1500tw columns on a fresh control
- `mdTest.bas`: `TestSkip` for graceful "RESULT: PASSED (0 tests, skipped)" on machines without the original control license

### Added (M3c -- defaults parity)

- Control defaults aligned with the original (verified by diffing a runtime-created original's full snapshot against ours): `DataMode=jgexDAO`, `RecordsetType=jgexRSDAODynaset`, `CursorLocation=jgexUseServer`, `BorderStyle=jgexFixed`, `HideSelection=jgexHideSelection`, `GroupByBoxVisible=True`, `RowHeaders=False`, `AutomaticSort/ShowToolTips/ScrollToolTips/HoldSortSettings=False`, `DefaultColumnWidth=1500`, `RowHeight=285`, `CardWidth=3750`, `CardSpacing=180`, `PreviewRowIndent=600`, `PreviewRowLines=0`, `BackColorBkg=vbWindowBackground`, `BackColorGBBox/GridLinesColor/ForeColorInfoText=vb3DShadow`, `RowColorEven=&HC1D7B0`, `RowColorOdd=&HBFFFFF`, `RecordNavigatorString="Record:|of"`
- Six built-in `FormatStyles` created on init (Default/OddRow/EvenRow/RowGroup/PreviewRow/SelectedRow with original colors); `JSFormatStyle.FontCharset` defaults to the system locale charset via a probing `StdFont`; `JSFmtConditions` defaults `GroupConditionCountTitle="Items"`/`ShowGroupConditionCount=True`; `JSPrinterProperties` defaults `ColorMode=jgexPPCMMonochrome`, `PrintQuality=-3`, `CardColumnsPerPage=2`, paper size/extent 0 and `RepeatFrozenCols=False` (fixed values -- empirically the original does *not* read the default printer DEVMODE)
- `UserControl_InitProperties` creates the original's two default 1500tw columns on freshly placed controls (persisted/sited controls are unaffected)
- `BoundColumnIndex`/`ReplaceColumnIndex` raise error 393 unless `ActAsDropDown` is set (matches original; they land in snapshot `$errors` identically); corpus canonicalization now strips `$errors`-listed props from *both* sides symmetrically
- UserControl designer: 3D client-edge border + pixel `ScaleMode` so the client area matches the original's 396x256 in a 400x260 site

### Added (M3c -- static table painting)

- GDI painting pipeline in `GridEX.ctl` (`UserControl_Paint`, double-buffer-free direct DC drawing with shared declares in `mdGlobals.bas`): group-by box (shadow background + info box sized by text extent with the info text at fixed inset), Double3D column headers (full-width top highlight and bottom shadow/dark lines with per-cell white left edge and shadow+dark right edges overwriting them at boundaries, filler cell to the right edge, captions per `HeaderAlignment`), data rows (selection via system highlight, even/odd colors, per-column `TextAlignment`, `DisplayValue` over `Value`), horizontal/vertical gridlines in `GridLineStyle` pen styles with a dark line under the last row, `BackColorBkg` filler areas and the XOR focus marquee (`DrawFocusRect` against a `BackColor` DC background, drawn under the vertical gridlines)
- Pixel-exact behaviors discovered and encoded: text insets are asymmetric (captions at x+2 with a 2px vertical drop, cell text centered over `x+2..x+w-4`), and `Rebind` positions the current cell at (1,1)
- Twips-facing metric props are stored internally in *pixels* and snap to the nearest pixel on runtime set exactly like the original (probed live: 400->405, 290->285, 1790->1785, 1793->1800, 187->180): `RowHeight`, `ColumnHeaderHeight`, `JSColumn.Width`, `DefaultColumnWidth`, `CardWidth`, `CardSpacing`; `PreviewRowIndent` stays raw twips and `ImageWidth`/`ImageHeight` are native pixels (also probed); public getters convert back via `Screen.TwipsPerPixel*`
- All 7 VisualDiff scenarios now verify pixel-identical to the goldens recorded from the original control (group-by box, headers incl. 27px flat variant, gridline styles/colors, unbound data rows with selection and focus, before/after-show application) -- M3c static painting exit reached for the initial corpus
- Corpus canonicalization additions: expected `RowHeight`/`ColumnHeaderHeight`/`Width`/`DefaultColumnWidth`/`CardWidth`/`CardSpacing` quantized like the runtime Lets do (design-time propbag values such as 288/228 are not pixel-multiples at 96dpi)
- Corpus grown to 14 scenarios, all verifying pixel-identical: header styles `jgexHSNoBorder` (flat fill only), `jgexHSSingleFlat` (1px dark frame + full-height separators) and `jgexHSSingle3D` (1px raised white/shadow), `EmptyRows` continuation grid to the bottom edge, row headers (18px Double3D corner + per-row cells with the black current-row triangle at x+6..11), hidden columns, large anti-aliased fonts and a group-by-box-less layout; text now draws with an OPAQUE background against the known fill color so glyph anti-aliasing blends byte-identical to the original (TRANSPARENT blending differed by one unit on AA edges); `partial-rows` deferred to M3d (the original shows a scrollbar, shrinking its client area)
- Default heights are font-dependent like the original (probed live with 5 fonts): row height defaults to `tmHeight + 3` of `Font` floored at 19px and an explicit `RowHeight` survives later font changes, while `ColumnHeaderHeight` is always recalculated to `tmHeight + 6` of `ColumnHeaderFont` on its font's change; implemented via `WithEvents StdFont` members reacting to `FontChanged` with a shared GDI `FontTextHeight` helper in `mdGlobals.bas`, refresh routed through a window-safe `pvInvalidate`; `ImportObject` now applies fonts before scalars so imported explicit heights override the recalc (mirrors propbag load order)

### Added (M3d -- scrolling, first slice)

- Vertical scrollbar as a real non-client `WS_VSCROLL` style toggled when rows overflow the data area (client shrinks exactly like the original), with `SetScrollInfo` range/page/pos kept in sync; recursion-guarded against the style-change resize feedback
- Painting starts at `FirstItem` (partial last row clipped naturally); `FirstItem` let clamps to valid range, fires `FirstItemChange` and repaints; `Rebind` homes it to 1
- VisualDiff `post` scenario section applies runtime props after the data feed (used for `FirstItem`); corpus at 16 scenarios all pixel-identical incl. restored `partial-rows` (scrollbar client shrink) and `scrolled` (`FirstItem=4`) -- both passed on first verify
- Interactive scrolling via the Modern Subclassing Thunk (`src/mdModernSubclassing.bas`, wqweto MST): the control IDE-safely subclasses its own window and routes `WM_VSCROLL` to `pvOnVScroll` (line up/down, page up/down from `GetScrollInfo` page size, thumb position/track from `SIF_TRACKPOS`) which drives `FirstItem`; the callback `ControlSubclassProc` is a hidden (`VB_MemberFlags = "40"`) member reached through `InitAddressOfMethod`, kept off the public typelib surface (gate still zero-diff); subclass set up in `UserControl_InitProperties`/`ReadProperties`, torn down via `TerminateSubclassingThunk` in `UserControl_Terminate`
- `test/ModelTests`: 6 new assertions (118 total) driving real `WM_VSCROLL` messages (`SB_LINEUP`/`SB_LINEDOWN`/`SB_PAGEDOWN`) through the subclassed window proc and asserting the resulting `FirstItem` steps and `FirstItemChange` event sequence
- Keyboard navigation routed through the subclass on `WM_KEYDOWN`: arrows move the current cell (`Row`/`Col`), Page Up/Down step by the visible-row count, Home/End jump to the first/last row, each calling `EnsureVisible` to auto-scroll; every keydown raises the public `KeyDown` event with a `GetKeyState`-derived shift mask first; `Col` set past the visible column count clamps to the last (per docs), `EnsureVisible(row)` scrolls `FirstItem` minimally to reveal a row, and shared geometry helpers (`pvTopHeight`/`pvVisibleRows`/`pvVisibleColCount`) back both scrolling and navigation
- `test/ModelTests`: 12 new assertions (130 total) sending real `WM_KEYDOWN` (arrows/Home/End/PageDown) through the subclassed proc and asserting current-cell moves, last-column clamp and auto-scroll into view
- Mouse handling through the subclass: `WM_LBUTTONDOWN` raises `MouseDown` then hit-tests the click -- a data cell sets the current `Row`/`Col`, a column header raises `ColumnHeaderClick(column)`; `WM_LBUTTONUP` raises `MouseUp` then `Click`, `WM_LBUTTONDBLCLK` raises `DblClick`; coordinates decoded from `lParam` as signed words via `CopyMemory`, shift mask from the `MK_*` wParam bits (+ `GetKeyState` for Alt); `pvColAtX` walks visible column widths (honouring the row-header offset) to map an x to a column position and object
- `test/ModelTests`: 5 new assertions (135 total) sending real `WM_LBUTTONDOWN`/`WM_LBUTTONUP` at computed cell/header coordinates and asserting the current cell moves, the `Click` event and the `ColumnHeaderClick` caption
- Row selection model: `RowSelected(pos)` get/let, the `SelectedItems` collection as the source of truth, and `SelectionChange` firing on change; click and keyboard navigation drive selection through a shared `pvNavigate` -- single-select follows the current row, multi-select toggles on Ctrl+click and range-selects on Shift+click/Shift+arrow from an anchor; `Rebind` selects the first row (matching the original); `EnsureVisible(row)` implemented and `Col` clamps to the last visible column; selected rows paint from the `SelectedRow` FormatStyle (system-highlight fallback) instead of hardcoded colors
- FormatStyles emulate the original's protected built-ins (verified by probing the original live): the 6 system styles are created up front with SelectedRow first, flagged via a `Friend frIsSystem` on `JSFormatStyle`, cannot be removed (`Remove` raises 380 like the original) and survive `Clear` (only user styles drop); the snapshot importer upserts styles by name so the corpus round-trip stays lossless despite the protected `Clear`
- `test/ModelTests`: 8 selection + 6 rewritten FormatStyles assertions (146 total) covering select-on-bind, `RowSelected` get/let with `SelectionChange`, system-style removal protection and `Clear`-keeps-system
- Record-selector painting verified against the original: the current row is always shown selected (highlighted) and the row-header ▶ arrow marks the current row only (not merely-selected rows); new `multiselect` VisualDiff golden (row headers, three selected rows plus the current row) recorded from the original and matched pixel-for-pixel via a scenario `select` list that drives `SelectedItems.Add`
- `KeyPress` (`WM_CHAR`) and `MouseMove` (`WM_MOUSEMOVE`) events raised through the subclass; a left-button drag extends a multi-select range from the mouse-down anchor to the row under the cursor
- `test/ModelTests`: 4 new assertions (150 total) for `WM_CHAR` -> `KeyPress` and left-drag range selection
- Horizontal scrollbar: `WS_HSCROLL` toggled (alongside the interacting `WS_VSCROLL`) when the visible columns overflow the client width, with `SetScrollInfo` range/page/pos; both scrollbars now appear together and shrink the client in both dimensions exactly like the original. Overflow painting matched to the original: the focus marquee clips its right edge to the visible width and a `vbButtonFace` separator line is drawn along the bottom client row under the horizontal scrollbar. New `both-scrollbars` VisualDiff golden (5 columns over 16 rows) recorded from the original and matched pixel-for-pixel (18 goldens total)

### Added (M3 -- multi-DPI golden verification)

- `test/VisualDiff` is now DPI aware: `VisualDiff.res` embeds a minimal `dpiAware` manifest (built from `VisualDiff.manifest`/`.rc` by `makeres.bat`; deliberately no Common-Controls v6 dependency so DPI awareness is the only behavior it changes) and goldens live in `golden\<dpi>\`. The same executable covers both modes -- `__COMPAT_LAYER=DPIUNAWARE` forces the OS-virtualized 96dpi path, native runs at the real system DPI -- so `make.bat` verifies and `record.bat` records at both, and a third scale needs no code change
- Golden corpus doubled to 40 images: all 20 scenarios recorded from the original at 96dpi and 120dpi (125% scaling), including two new font scenarios (`font-tahoma`, `font-segoeui`) that exercise TrueType/ClearType metrics against the bitmap MS Sans Serif default
- DPI fidelity fixes surfaced by the 120dpi goldens, each verified against the original at both scales:
  - `RowHeight`/`ColumnHeaderHeight` are derived from their font at init instead of being seeded with the 96dpi literals (the original keeps a 19px row at both scales but grows the header 19px -> 22px, which is exactly `tmHeight + 6`); `DefaultColumnWidth`, `CardWidth` and `CardSpacing` are confirmed fixed pixel counts at both scales
  - scrollbar reservation uses `GetSystemMetrics(SM_CXVSCROLL/SM_CYHSCROLL)` (17px at 96dpi, 21px at 120dpi) rather than a hardcoded 17
  - header captions are `DT_VCENTER`-ed over the full cell height (identical at 96dpi, a pixel higher at 120dpi) and data cell text gets one more pixel of positioning rect so centered text lands on the original's pixel
  - the focus marquee is drawn explicitly as a XOR checkerboard anchored *per column* -- inverted when `(x - column left) + (y - row top)` is odd -- replacing a single row-wide `DrawFocusRect`, which only looked correct while every column boundary and the row height happened to be even (true at 96dpi, false at 120dpi)
  - cell text is positioned in its inset rect but clipped to the whole cell (`IntersectClipRect` + `DT_NOCLIP`), so ClearType fringes bleed into the inset exactly as the original's do while long text still stops at the cell edge
- Corpus extended to 150% scaling (144dpi): all 20 scenarios verify at 96dpi and 144dpi. Two harness bugs had to be fixed before the original could even be recorded there, and both were harness-side, not control-side:
  - the host form's ambient font is inherited by a control added with `Controls.Add`, and the VB6 default MS Sans Serif collapses the *original's* own layout at 144dpi; the form is now Tahoma, with `frmHost.pvSetDefaultFonts` explicitly assigning the grid's `Font`/`ColumnHeaderFont` after `Controls.Add` so the scenario baseline stays independent of the ambient font
  - the test baseline font is now `MS Sans Serif 8` rather than 8.25: OLE snaps a requested size to the nearest one a raster font actually has when it realizes the `hFont`, which at 144dpi turns 8.25 into 8 (and Tahoma 8.25 into 8.5) *inside the original's own font object* -- the original's dump reports `Font.Size = 8` at 144dpi but 8.25 at 96dpi. Both controls were handed the same 8.25 and resolved it differently, which is what made every derived metric come out exactly 1px too large at 144dpi (row 23 vs 22, header 26 vs 25, group-by box 40 vs 39). Asking for a size the font actually has removes the ambiguity and the existing `tmHeight + 3` / `tmHeight + 6` formulas then reproduce the original at all three scales
- `mdMain.bas` gained a `metrics` mode (`VisualDiff.exe metrics`) that dumps `GetTextMetrics` for the corpus fonts at the current DPI, incl. the requested -> realized size, which is how the snapping above was pinned down
- 144dpi fidelity fix: text of the *current* row clips inside the focus marquee, so a column running past the client edge stops one pixel short of the marquee's right border instead of painting through it (only observable when a cell is clipped by the client edge, i.e. at 144dpi in `font-segoeui`)
- `mdMain.bas` probes control availability by creating it on a throwaway host instead of calling `Licenses.Add`, which fails now that the harness carries a compile-time reference to the original OCX (`RemoveUnusedControlInfo=0` keeps its design license in the executable)
- `test/VisualDiff/scenarios/*.json` reformatted human-readable (4-space indent, leaf objects and unbound rows kept inline) with colors as `0x`-prefixed hex literals
- Font defaults are no longer hardcoded anywhere: `mdGlobals.NewStdFont` is the single factory that mints one, pinning the typeface only and leaving the point size to `StdFont` itself, and both controls inherit the container's font in `UserControl_InitProperties`/`ReadProperties` via `mdGlobals.CloneFont` (one `IFont.Clone` call, so every attribute is copied and the ambient font is never aliased). The `Initialize` event still seeds the members from the factory so they are never `Nothing` before the container is known. This is what the original does -- the reason a scenario renders differently under an MS Sans Serif host form than under a Tahoma one -- and it retires the `8.25` literals that had been copied into `GridEX`, `GEXPreview`, `JSPrinterProperties` and `JSFormatStyle` along with the now-redundant `pvDefaultFont`/`pvDefaultCharset` helpers
- Re-recorded `golden/96` after Windows updates KB5100998/KB5101650 changed the system accent from `0x0078D7` to `0x0078D4`; every differing pixel was that substitution or its XOR complement (`0xFF8728` -> `0xFF872B`), so the change is a palette shift, not a layout one

### Added (M6 -- paint property verification)

- `colors-rows` and `colors-chrome` scenarios close 11 entries in the paint matrix (verified now 30 of 126, unverified down to 5): row/background/gridline colours with a dotted style over real rows, and the header/group-by-box/info-text colours. Added as new scenarios rather than by giving `gridlines-dots-colors` rows, so the existing 144dpi goldens stay valid on a machine that cannot currently record them
- Three rendering defects found by those scenarios, none of which any earlier scenario could have surfaced -- every existing one uses solid gridlines, which cover the gridline row completely, and the only dotted scenario had no rows:
  - dotted gridlines used GDI's cosmetic `PS_DOT`, which renders 3-on/3-off; the original is 1-on/1-off on the *checkerboard parity* `(x + y)`, the same absolute anchoring the focus marquee uses. `pvLine` now routes `PS_DOT` to `pvDottedLine`, which stamps explicit pixels -- the phase error was invisible on the vertical gridlines (both fell on odd x) and only showed on the horizontal ones, where two of the three lines came out inverted
  - a dotted gridline's gaps show `BackColorBkg`, not the row colour underneath, so `pvDottedLine` lays the whole run down in the background colour before stamping the dots
  - the focus marquee inverts *every* border pixel, alternating the XOR mask between white and `ForeColor` like a two-colour pattern brush, rather than inverting every other pixel with white. At the default black `ForeColor` the second pass is a no-op, which is exactly why 21 scenarios agreed with the original while the rule was wrong
- `formatstyles-selection`, `column-order` and `default-column-width` scenarios clear the rest of the unverified block: **35 of 126 properties verified, 0 unverified**, 7 weak, 83 not implemented. Two further defects found:
  - a column added without an explicit `Width` kept `JSColumn`'s own default of `1000` -- and since widths are stored in *pixels*, that was 1000px per column, wide enough to force a horizontal scrollbar and shrink the client by 17px. The original gives a new column the control's `DefaultColumnWidth`, so `JSColumns` now carries that default (`frDefaultWidth`, pushed from the `DefaultColumnWidth` setter) and applies it in `Add`
  - an out-of-range `ColPosition` renders an empty grid, because `ItemByPosition` finds no column at the expected slot, where the original ignores the value and keeps declaration order. Left unfixed and recorded in `PAINT-PROPERTIES.md` -- the corpus now uses valid positions, so nothing covers it yet

- Three `gridlines-*` scenarios (None/Vertical/Horizontal, each also carrying row colours, a disjoint multi-selection or an explicit current `Row`) clear the last weak entries: **42 of 126 verified, 0 weak, 0 unverified**, leaving only `LeftCol` partial and 83 unimplemented. Six more defects, all of them invisible while every scenario used the default `GridLines`:
  - `GridLineStyle = Dashes` used GDI's `PS_DASH`; the original draws 3-on/3-off **re-anchored at every row top**, so a row height that is not a multiple of 6 (19px at 96dpi) leaves a four-pixel run across each row boundary
  - cell text and the focus marquee centre over the full row height when no horizontal gridline is drawn -- we always reserved its pixel line
  - the marquee's right edge, and the horizontal gridline, each run one pixel further right when no vertical gridline claims the block's last column
  - the row header's border pair lands on a different line per `GridLines` value; recorded in `PAINT-PROPERTIES.md` as an empirical table since no mechanical rule explains why the single-direction modes shift it in opposite directions
  - what shows below a short row header is the grid `BackColor`, not the row's own colour and not the control background
  - a dashed gridline's phase keeps running past the last row instead of restarting, so the block's closing line sits at offset `RowHeight` in the 6px cycle: `19 mod 6 = 1` draws it at 96dpi, `22 mod 6 = 4` leaves it blank at 144dpi. Only visible because the corpus spans both scales -- the reset-every-row version looked correct at 96dpi alone
  - assigning `Row` from outside collapses the selection onto that row and repaints; navigation and drag now go through a private `pvSetRow` that leaves the selection alone, which is what keeps `SelectionChange` ordering and the drag anchor intact

### Added (M3d -- horizontal scrolling)

- Columns scroll horizontally: painting and hit-testing start at `pvFirstCol` instead of column 1, `WM_HSCROLL` is handled (line/page/thumb) through `pvOnHScroll`, and `LeftCol` repaints, refreshes the scrollbar and raises the previously-declared-but-never-fired `LeftColChange`. Scrolling moves whole columns, as the original does -- the `hscrolled` golden shows column 3 flush against the row header, with no partial column
- The bottom separator strip is gauged on the width of *every* column rather than those visible from `LeftCol`: the horizontal scrollbar is what puts the strip there, and it depends on the total
- `LeftCol` was the last property the renderer only partly consumed, so the paint matrix is now **43 of 126 verified with nothing weak, unverified or partial** -- M6 complete
- Scenarios carrying a `post` block settle before capture (`pvSettle`, ten 1ms pumps): the original repaints some runtime property changes off a timer, so the capture-until-stable loop could take two identical shots before it ever fired. That was the intermittent selftest failure -- the original diffed against a golden recorded seconds earlier, one capture holding a selected row the other did not
- Scenarios renamed `NNN-name.json` in creation order, with every golden PNG renamed to match -- 30 scenarios across three dpi folders. Masks now need the prefix or a wildcard (`make.bat 030*`), since the harness matches with VB `Like`
- `make.bat` and `record.bat` take an optional scenario mask (`make.bat hscrolled`, `record.bat gridlines-*`), which the harness already supported but the scripts never passed on. A single-scenario check runs in about a second against a couple of minutes for the corpus

### Added (M3d -- scrollbar band and record navigator)

- The grid paints onto a constituent `picGrid` PictureBox sized to everything above the scrollbar band, mirroring the original's structure -- its window tree is an outer UserControl holding an inner grid window plus a PictureBox band with an `HScrollBar` and a `TextBox`. `WS_VSCROLL` moves to `picGrid`, so the vertical bar stops where the band starts and its thumb geometry matches without any arithmetic; the `hWnd` property now returns `picGrid.hWnd`, which is what the original exposes. `pvClientBottom` disappeared with it -- that helper only existed to compensate for a child scrollbar not shrinking the client
- Horizontal scrolling uses a child `HScrollBar` (`hsbGrid`) in the band, not `WS_HSCROLL`: a non-client bar always spans the whole edge, so it could never share the strip with the navigator. `TabStop = False` keeps it out of the tab order, matching the original, whose scrollbar carries no `WS_TABSTOP`
- Two windows are subclassed now: the grid surface for painting, input and `WM_VSCROLL`, and the outer control because it is the band's parent and so receives `WM_CTLCOLORSCROLLBAR` -- which is what makes the shaft dither match
- The record navigator is painted on the outer control (`pvPaintNavigator`), with no child controls: buttons, arrow glyphs, the first/last bars, the sunken record box and its number -- so there is no text box and no way to type a record number, and nothing in it takes focus. First/prev/next/last respond to clicks, hit-tested in `pvOnNavigatorClick` against the same layout used to paint them; the outer control repaints alongside the grid so the number and the greyed pair keep up. Six ModelTests assertions drive the buttons through `WM_LBUTTONDOWN` on the parent window, including the clamp at the first row and a click that lands between buttons. Layout is derived from metrics rather than pixels, so it holds at every scale: buttons `BandH +/- 1`, glyph half-height `BandH \ 4` (5x9 at 96dpi, 6x11 at 120dpi), bar `BandH \ 4` / `BandH \ 7` / `BandH \ 5`, box `tmHeight + BandH + 16`, labels centred on the font box
- Three details of the original that no amount of reasoning would have produced: its record box is a `TextBox` *taller than the band and clipped by it*; only the inner `<` button greys at the first row, never `|<`; and the disabled glyph is the triangle drawn twice with the highlight offset (+1,+1), only its uncovered edge column showing
- `pvVisibleColsInWidth` counts columns that fit **strictly** within the available width -- a column ending exactly on the edge does not count. Found with a 13-scenario sweep of column widths: `3 x 132 = 396` against a 396px client behaves as 2 visible, not 3, and the same boundary at 120dpi (`3 x 125` against ~375px) was breaking `both-scrollbars` there. The thumb is one per scroll position, `colCount - visible + 1`
- `LeftCol` and `Row` repaint when assigned; `LeftColChange` fires for the first time since it was declared in M2

### Added (M3d -- generated sample setup)

- `tools\GenSample.ps1` turns a snapshot recorded from the original into plain VB6 that configures our control, so a Janus sample ports by re-pointing the reference and calling the generated `Sub`. It mirrors `mdImport.bas` assignment for assignment: scalars, enums as constant names, fonts, `Add` arguments per collection, and -- recursively -- read-only sub-objects such as `PrinterProperties` or a condition's `FormatStyle`, including parameterized `HeaderString(1..3)`. Collections carrying persistable properties of their own snapshot as an object with an `items` array rather than a bare array, which is what `FmtConditions` is; the generator handles both shapes
- `pvTestGeneratedSetup` round-trips the three generated modules (`Unbound-1`, `Unbound-2`, `Unbound-Collection`) against the snapshots they came from: 159 model tests pass. `GridImages` is out of scope on both sides -- pictures cannot be written as code literals, a ported sample keeps them in its own `.frx`

### Added (M4 -- sorting)

- `SortKeys` sorts the rows: a stable merge sort over an index map, so equal keys
  keep the order the client app supplied. Everything that speaks in display
  positions goes through the map -- `RowIndex` answers what data a position
  holds, and the current row and the selection are remapped after each re-sort
  because the original keeps both on the data they were on, not on their old
  positions. `JSColumn.SortType` picks the comparison; blanks sort first
- The header sort arrow, verified at two DPIs and three header heights: a fixed
  8x7 engraved triangle -- the original draws the same pixels at 120dpi as at 96
  -- sitting on the caption's baseline rather than centred in the band, which is
  what `036-sort-tall-header` (a 600 twip header) exists to pin down. Descending
  mirrors ascending except that its flat top edge is shadow where the ascending
  base is highlight
- Scenarios `035-sorted-column`, `036-sort-tall-header`, `037-sorted-descending`
  and `pvTestSorting` in `ModelTests`; `SortKeys`, `JSColumn.SortOrder` and
  `JSColumn.SortType` move from not implemented to verified

### Added (M4 -- grouping)

- `Groups` groups the rows: the grouped columns lead the sort keys, then a group
  row is emitted for every level whose key changed, so a change high up restarts
  the levels nested inside it. Group rows carry a caption, an expand box and the
  record count of their own level; records are indented one level width (16px)
  per group with a gridline ruling every level boundary they sit behind
- The group-by box shows a chip per level instead of the info text: a raised
  button carrying the column caption and the same sort arrow its header shows,
  stepping right and half a header row down per level, joined by an elbow that
  drops out of the chip above and meets the next one a line below its top edge
- Both placements that would otherwise hardcode pixels are read off the font, so
  they hold at any dpi: the group caption starts a space width plus the usual
  two pixel text margin past the expand box (5, 7, 9 pixels at 96dpi for the
  8.25/12/16pt strikes and 5, 8, 10 at 120dpi -- `tmAveCharWidth` matches the
  96dpi column by coincidence and is 1-2px wide at 120), and the chip staircase
  steps by half the column header height
- A grouped column's header carries the same sort arrow an explicit sort key
  gets, since grouping sorts by the column too
- Scenarios `038-grouped-one-level`, `039-grouped-large-font`,
  `040-group-caption-width`, `041-grouped-two-levels`, `042-sorted-two-keys`,
  `043-group-colors` and `044-grouped-huge-font`, all passing at both DPIs;
  `Groups`, `BackColorRowGroup` and `ForeColorRowGroup` move from not
  implemented to verified
- Known divergence: within a group the original leaves records in an unstable
  order (interleaved input comes back scrambled by index position), we keep the
  supplied order; the goldens are recorded from contiguous input where both
  agree

### Added (M4 -- expand/collapse)

- `RowExpanded` collapses and expands a group row, with `CollapseAll`/`ExpandAll`
  and `DefaultGroupMode` (`jgexDGMCollapsed`) driving the same state. Collapsing
  reprojects rather than re-sorts: the sort order underneath is built once into
  `m_aOrder`, and a second `m_aVisible` map holds just the rows on show, so
  everything under a collapsed row drops out until a group row at that level or
  above shows up again. A display position now resolves through the map, and a
  row hidden inside a collapsed group answers with the group row that hides it,
  which is where the current row and the selection land when their data
  collapses away
- `JSRowData.RowIndex` reports the index the owner maps a stored row to, so a
  wrapper held across a collapse keeps answering for the record it wraps; an
  orphaned wrapper raises error 91 for it as it already did for `Value`, since
  dereferencing the zeroed weak pointer raises that by itself
- The current-row marquee on a group row spans the block as a single run: it
  starts at the very left edge, tree indent included, no column rule breaks its
  XOR checkerboard, and with no vertical gridline to yield to it runs the full
  width out -- one pixel further right than a data row's
- Scenarios `045-group-collapsed`, `046-group-collapsed-one`,
  `047-group-collapsed-nested` and `048-group-default-collapsed`, all passing at
  both DPIs; the harness gained a `calls` block for the methods and indexed
  properties a scenario needs after its data is in (`CollapseAll`,
  `RowExpanded(2) = False`), neither of which fits the plain `props`/`post`
  shape. `RowExpanded` and `DefaultGroupMode` move from not implemented to
  verified

### Added (M4 -- automatic sort)

- `AutomaticSort` sorts on a column header click, doing what the original
  documents client code used to write in the `ColumnHeaderClick` and
  `GroupByBoxHeaderClick` events: a grouped column flips the sort order of its
  group, any other one becomes the only sort key, ascending unless it already
  was ascending -- so a header never cycles back to unsorted. The event goes out
  before the control sorts, and under the default `False` a click still only
  raises it
- `GroupByBoxHeaderClick` fires: the chips were painted but never hit-tested, so
  the event was dead. `JSGroup` now keeps the rectangle the layout gave its chip,
  and painting and hit-testing both read it instead of each walking the
  staircase. A chip stands for its column, so clicking one sorts through the
  same path a click on that column's header takes
- `JSColumn.IsGrouped` answers off the `Groups` collection through a weak owner
  reference, the same one every other part of the object model uses to point
  back; it is what routes a header click to the group rather than the keys
- Group row captions honour `JSColumn.GroupPrefix`, `GroupFormat` and
  `GroupEmptyStringCaption` (defaulting to `(none)` as the original does, which
  our `JSColumn` was missing). A prefix is joined to the value by a space, and
  the caption then starts one space earlier, so the value lands exactly where an
  unprefixed one does: the caption is really `prefix & " " & value` drawn two
  pixels past the expand box, and what looked like a space-width margin for an
  unprefixed level is that same leading space with an empty prefix
- `GroupFormat` labels the caption and nothing else. The help reads as though it
  groups too ("in order to group records on a Month-Year basis"), but the
  original still breaks groups on the raw value: two dates in one month give two
  groups whose captions read alike, which `052-group-format` records
- `AutomaticSort` and `JSColumn.IsGrouped` move to consumed, covered by
  `pvTestAutomaticSort` in `ModelTests`; `GroupPrefix`, `GroupFormat` and
  `GroupEmptyStringCaption` move to verified against scenarios `049-group-prefix`,
  `050-group-empty-caption`, `051-group-empty-default` and `052-group-format`,
  recorded from the original at both DPIs

### Added (M5 -- event log recorder and the in-place text editor)

- VisualDiff records what a control *raised*, not only what it painted: a
  scenario gains an `input` block (clicks, keys, typed text) and the harness
  writes an `.events.txt` beside each golden PNG, which verify diffs line by
  line and reports the first disagreement. Event order is the contract editing
  lives or dies by, and no picture can hold it. Two things are left out of a
  logged event: parameters carrying an object, since the `JSRet*` and
  `JSRowData` carriers say nothing by identity and the test asserts their
  contents through the control instead, and `Shift`, since a synthetic
  `WM_KEYDOWN` carries no modifier state and both controls read it off the live
  keyboard -- a modifier held while the corpus runs would otherwise show up as a
  mismatch that says nothing about either of them
- The in-place text editor: `EditType = jgexEditTextBox` opens a native EDIT
  over the cell, one created per session and destroyed with it, because the
  styles a column asks for are fixed when the window is made -- `ES_LEFT`/
  `ES_CENTER`/`ES_RIGHT` from `TextAlignment`, `ES_MULTILINE` or `ES_AUTOHSCROLL`
  from `WordWrap`, `EM_LIMITTEXT` from `MaxLength`. It sits one pixel inside the
  cell so the row's selection colour and marquee still show around it, and
  `EM_SETMARGINS` gives its text the two pixel margin a painted cell has (sent
  after the font, which resets it)
- The recorded contracts it satisfies: a click raises `RowColChange` and
  `BeforeColEdit` before the client sees `MouseDown`; each keystroke raises
  `KeyDown`/`KeyPress`/`Change`/`KeyUp`; Enter commits through `BeforeColUpdate`,
  `AfterColUpdate`, `AfterColEdit`, `BeforeUpdate`, `RowFormat`, `AfterUpdate`,
  `RowFormat`, `SelectionChange`, `RowColChange`; Escape cancels through
  `RowFormat` and `AfterColEdit`. Scenarios `057-edit-textbox`,
  `058-edit-commit-enter`, `059-edit-cancel-escape` and `061-edit-tall-row`
  pass pixels and event log against the original at both DPIs
- The editor's text lands where the painted cell's does, which is what places
  the window: an EDIT draws at the top of its client, so the window itself is
  put where the text centres to -- `(rowContentHeight - tmHeight) \ 2`. A row
  keeps its 19 pixels at either dpi while the font grows, and that alone is the
  difference between a two pixel inset at 96 and one at 120;
  `061-edit-tall-row` holds the same rule at a 40 pixel row
- It shows a caret because it holds the focus, which needs no help from the
  control beyond not losing it: the grid's own click handling takes the focus
  for itself, so the default runs first and the editor takes it back after

### Fixed (M5 -- what the event log caught)

- `MouseDown`/`MouseUp`/`MouseMove` reported pixels where the original reports
  container units: a click 38 pixels in comes out as 570 twips
- `SelectionChange` fired even when the selection had not changed, and fired
  after `RowColChange` where the original raises it before
- A fresh bind left the current column at 1. `Col = 0` means the whole row is
  selected, which is the state the original starts in -- its first
  `RowColChange` reports `LastCol=0` for that reason. It also decides a paint
  rule: inside a selected row the current cell is drawn unselected, inset by
  one pixel so the row's colour still shows in the marquee band
- Three of those four were M3d-era divergences that no pixel golden could see

### Added (M5 -- the checkbox editor and wrapping cells)

- A `jgexCheckBox` column draws its state instead of its text: a fixed 11x12
  box, white inside with one grey line around it and a shade darker on the cell
  the marquee is on. The tick is the same seven runs of pixels at either dpi but
  sits further into the box as the screen scales, two pixels in at 96 and three
  at 120 -- the font does not move it, since a 12pt cell font at 96 leaves it
  exactly where an 8.25pt one does (`063-checkbox-large-font` is the scenario
  that settled that), and neither `DrawFrameControl` nor a plain dpi ratio
  reproduces it, so the two offsets are the ones the recordings show
- `EditType = jgexEditCheckBox` flips the value on a click level with the box
  and says one `Change` about it, with no editor window and no update trio.
  Only the vertical band counts, which is the whole of why the same point
  toggles at 96dpi and not at 120: the taller bands there put the row lower and
  leave the point above the box
- A `WordWrap` column wraps its painted text from the top of the cell rather
  than centring one line in it, and its editor takes the whole cell, since the
  room it has is what decides how many lines the text breaks into. The click
  that opens an editor carries on into it, so the caret lands on the character
  under the point -- which is what decides where typing goes and, in a wrapping
  cell, which line the view sits on
- Scenarios `060`-`064` recorded from the original: both checkbox states, a
  large font, a click straight onto the box, and a wrapping editor in a 40 pixel
  row

### Fixed (M5 -- what the checkbox found)

- `Col` raised `RowColChange` without ever invalidating. The current cell is
  drawn out of the selected row it sits in, so moving between columns is a
  repaint and not only an event -- reachable by arrow key, not just by the
  scenario that exposed it, where the capture kept showing a pre-click paint
- A wrapping editor was created without `ES_AUTOVSCROLL`, so it refused text it
  could not fit and the refusal came back as a phantom `KeyPress(8)` after every
  character typed
- The white the current cell keeps inside a selected row insets only where the
  marquee runs, which is the row's own edge -- not the boundary between two
  cells, which no marquee follows

### Fixed (horizontal scrollbar)

- Dragging the horizontal thumb always landed on the first column. The bar's
  VB6 side -- `Min`, `Max` and `Value` -- was never set at all, so `Value` sat
  on its 0 and the handler computed `LeftCol` as `0 + frozen + 1` whatever the
  thumb did. The control writes the column numbers straight onto the window
  with `SetScrollInfo`, over VB6's `Min..Max` mapping, so the position is read
  back off the window the same way: `nTrackPos` mid-drag, `nPos` once it is
  released. VB6's own three are set to those same numbers as well, so whichever
  of the two a drag is reported through says the same thing
- No automated test covers the drag: a scrollbar runs a modal tracking loop on
  button down, so a synthetic click deadlocks the harness, and VB6 ignores a
  posted `WM_HSCROLL` -- its `Change`/`Scroll` only fire for real mouse input.
  The write side is covered, by `scrollinfo` mode and the `029`-`034` goldens

### Added (M5 -- the pending row buffer and `Value`)

- `Value(colindex)` was an empty pair of stubs -- reads gave `Empty` and writes
  were dropped. It is implemented, and implementing it meant building what sits
  under it: the row being edited buffers its cells rather than writing straight
  through. A column commit and an assignment to `Value` both land in that
  buffer, the grid paints from it, leaving the row writes it to storage behind
  `BeforeUpdate`/`RowFormat`/`AfterUpdate`/`RowFormat`, and Escape drops it
- Escape is two-level, which is what proved the buffer exists: the first press
  cancels the cell, the second cancels the row and reverts every column
  buffered since the row became current. Editing three columns and pressing
  Escape once leaves the first two edits standing and undoes the third;
  pressing it twice undoes all three
- `RowFormat` on a cancelled cell only fires when that cancel ends the row's
  edit -- with other columns still buffered the row stays dirty and keeps it
- Six scenarios record the whole model from the original, pixels and event log:
  `069`-`071` for a buffered `Value` showing, being cancelled and being
  committed by a row move, `072`-`074` for the editor's own commits under one
  and two Escapes. `ModelTests` pvTestEditLeaveCell covers what they cannot --
  that storage is untouched until the row is left, which a repaint from the
  buffer looks identical to
- `AssignVariant` moved into `mdGlobals.bas`: copying a cell value needs `Set`
  for an object and refuses it for anything else, and three places now do it

### Fixed (M5 -- leaving a cell)

- `hWndEdit` was an empty property that always answered 0, while the handle it
  is meant to return sat in a member beside it. It returns the editor's window
  now, and 0 when no cell is being edited -- which is what the original does,
  read off its own object model with a cell open and with none

- The in-place editor survived a move of the current cell: it stayed on screen
  over its old cell and what had been typed was never committed. It now shows
  on the current cell and nowhere else -- every path that moves either
  coordinate closes it and commits, whether the move comes from a click, a key
  or an assignment to `Row`/`Col`
- The commit is split the way the original splits it, which two new scenarios
  pin down: leaving the *column* raises `BeforeColUpdate`/`AfterColUpdate`/
  `AfterColEdit` and nothing else, while leaving the *row* adds
  `BeforeUpdate`/`RowFormat`/`AfterUpdate`/`RowFormat` behind them. That also
  fixes where the commit sits in the sequence -- it has to run before
  `SelectionChange`, which is why it hangs off `pvNavigate` rather than the
  row setter it first went into. Escape still discards, unchanged
- `067-edit-commit-on-move` and `068-edit-commit-on-row-move` record both cases
  from the original, pixels and event log. `ModelTests` pvTestEditLeaveCell
  covers what a golden cannot: that the typed text reaches the row's storage
  rather than only the painted cell, and that the editor window is destroyed

### Fixed (invalidation)

- `ItemCount` marked the sort order stale but never repainted, so rows appearing
  or going only showed up when something else caused a paint. The count drives
  the scrollbar range as well as the contents, so it now invalidates like every
  other change (the `Redraw = False` batch guard still holds it back). No golden
  could have caught this: the visual harness always follows `ItemCount` with
  `Rebind`, whose own invalidate covered for it, and the original control
  refuses an `ItemCount` change at runtime at all, so there is nothing to record
  a golden from. `ModelTests` pvTestItemCountInvalidate pins it instead, on the
  `WS_VSCROLL` style the same paint pass puts on the control's window --
  reverting the one-word fix turns that assertion red
- `pvUpdateScrollBars` decided differently depending on whether it had already
  run. It measured the surface from `picGrid`, which the previous pass had
  already shortened by the band and narrowed by the vertical bar, and then
  charged for both again: a second pass invented a horizontal bar for columns
  that fit and took its height out of the vertical page, leaving the thumb a row
  short. It now adds back whatever the last layout took before deciding. The
  band is one strip the navigator and the horizontal bar share, so it costs its
  height once rather than once per occupant -- which was wrong for a grid with
  both, and only stayed invisible because nothing made that routine run twice
  until `ItemCount` above started invalidating (`018-both-scrollbars` caught it)

### Fixed (object model)

- `JSColumn.SortOrder` always answered `jgexSortNone`. It returned a member of
  its own that nothing ever wrote -- `frSetSortOrder` had no callers -- while
  the header arrow was painted from a private helper reading the real sort
  state, so the property and the pixels had drifted apart without either being
  obviously wrong. It is read-only in the original, so it is now a view over
  that same state: the helper is a `Friend` and the property delegates to it,
  the way `IsGrouped` already delegated to `frColIsGrouped`. Grouping counts
  towards it -- dumping the original's own object model for
  `038-grouped-one-level` shows a grouped column reporting the group's order
  with no `SortKeys` at all -- which is the rule the arrow already painted by.
  Eight assertions in `ModelTests` pvTestAutomaticSort now read it back off
  keys and groups, including the transitions the old code could not have failed

### Fixed (M5 -- 150% scaling)

The corpus was re-recorded from the original at 144dpi and ran 61 of 70 there.
Every divergence was a metric calibrated against two scales that a third pulled
apart, and each is now a rule that holds at all three rather than a third
constant:

- The group-by box grows a staircase step per level past the first and
  `pvTopHeight` added a flat 14 pixels for it. What the box does not take is
  what the rows get, so a grouped grid that would otherwise just fit gained a
  vertical scrollbar it should not have had
- The record navigator's box holds seven characters and a margin,
  `7 * tmAveCharWidth + 11` -- 46px at 96dpi, 53 at 120 and 67 at 144 -- which
  is why 5, 50 and 500 records all render identically. It had been a constant
- A group chip's elbow meets the next chip on that chip's text baseline, so it
  follows the ascent rather than the whole font box. The two part company only
  once the descent grows: 2px at 96 and 120dpi, 4 at 144
- A grouped block that overflows comes up scrolled to its bottom, reporting
  `FirstItem = 2` straight after a rebind where an ungrouped block of the same
  height stays at the top. `018-both-scrollbars` overflows at 144dpi and stays
  put, which is what makes this grouping-specific rather than a rule about
  overflow. Collapsing a group now clamps `FirstItem` so it cannot point past
  the last screenful
- The navigator's arrow and bar centre a row above the button's own middle,
  which shows only once the band height turns odd: 247 either way at 96 and
  120dpi, 242 rather than 243 at 144. The bar is exactly as tall as the arrow
  beside it and shares that centre -- 9, 11 and 13 rows -- where it had been
  inset from the button's top and bottom independently, and it starts
  `(BandH + 2) \ 4` in from the edge, 4, 5 and 7 pixels
- The record number sits at the top of the box's interior. Centring it in what
  is left of the band agrees at 96 and 120dpi and puts it a pixel low at 144
- `pvSetHScrollInfo` counted the columns fitting after the frozen block without
  the clamp its VB6-side twin already had, so a column too wide for the room
  left put the last valid `LeftCol` one past the final column and the thumb came
  out a step too small -- 1500 twip columns take 150px at 144dpi with 72 left
- The checkbox's 11x12 box does not scale but the tick in it is the system's own
  and does, so at 144dpi only the mark's top four rows are inside it, both arms
  already there and the vertex below. It is not the 96dpi shape scaled -- that
  one's top row is a single pixel on the right
- A click toggles a checkbox wherever in the cell it lands. The vertical band
  the toggle required was read off a 120dpi recording of a scenario that has
  since moved its click, and both current recordings toggle from a point a row
  above the box
- The in-place editor is two pixels narrower: it ends where the painted cell's
  text is clipped, at `lX + lW - 3`. Measured against the original's own window
  -- 96x37 against our 98x37 at 96dpi, 146x57 against 148x57 at 144 -- and it
  decides where a wrapping cell breaks its words

`pvDumpWindows` now reports a window's text, selection, line count and first
visible line, and `windows-ours` runs the mode against this control, so the
original's editor can be read directly rather than inferred from a picture of
it. That is what settled the editor width, and what narrowed the one below.

The 120dpi goldens for `057`-`064` were stale here: those scenarios were made
dpi-robust (group-by box off, clicks inside row 1 at every scale) and only 96
and 144 had been re-recorded. Re-recorded at 125% below.

### Fixed (M5 -- 125% scaling)

The corpus was re-recorded from the original at 120dpi and ran 65 of 70 there.
Every divergence was again a metric that two scales agreed on and a third pulled
apart -- and three of the four turn out to be the same thing, a quantity the
original rounds where we truncated:

- The record navigator's box is the width of `"9999999"` plus the 2px client
  border on each side -- 42, 49 and 63 measured, so 46, 53 and 67 -- and not
  `7 * tmAveCharWidth + 11`. That formula held at 96 and 144dpi by coincidence,
  since the average character is 5 and 8 there; at 120 it is 7 rather than the 6
  the constant assumed, so the box came out 60 against the original's own 53px
  TextBox and shoved the navigator and the scrollbar beside it to the right
- A group chip's elbow has nothing to do with the font: it meets the next chip
  three quarters of the way down that chip, less a pixel. Measured over seven
  chip heights -- 19, 22, 26, 25, 30, 31 and 37px, three faces at two or three
  dpi -- the leg lands at 13, 16, 18, 18, 22, 22 and 27. Three of those are
  exactly `.5` and all three go the way `CLng` rounds, to even, which is why no
  `\` form of it fits: 15.5 has to give 16 while 18.5 has to give 18. The
  ascent-based rule it replaces fit six of the seven
- The group-by box measures its staircase whole and rounds once, rather than
  accumulating a truncated half-chip per level. At a 25px chip the original's
  box is 52 high -- `CLng` of 37.5 -- where per-level truncation gives 51 and
  shifts every band below it up a pixel; 19 and 31px chips agree with both, at
  42 and 60. The chip tops themselves do truncate, 9, 12 and 15, so the box and
  `pvChipStagger` cannot share the one expression
- `pvTextWidth` measures with `GetTextExtentPoint32`, which is what VB6's own
  `TextWidth` calls, instead of `DrawText` with `DT_CALCRECT`. The two agree on
  every string in the corpus in MS Sans Serif and Tahoma and part company on
  TrueType, where `CALCRECT` drops the last glyph's overhang: "Region" in Segoe
  UI 14 measures 58 that way against the 59 the original lays its chip out with

`061-edit-tall-row`, red since the 144dpi pass, is fixed and the divisor guessed
at there was the wrong question. Probing the original with the editor open and
clicks at three heights -- the recording that entry asked for -- shows only the
first line ever answers: a click on line 0 gives the character under it at both
dpi, while anything below that line takes the caret to the end of the text. It
hit-tests a wrapping cell the way it does an unwrapped one. The 96dpi click sits
20px down a 13px line and so falls past it; the same click at 120 sits 15px down
a 16px line and is still on line 0, which is why one scale looked correct.

### Added (M5 -- two-level grouping at large fonts)

- `065-grouped-two-levels-tahoma` and `066-grouped-two-levels-segoeui`: the
  corpus had two-level grouping only in the default font, which is one chip
  height per dpi and cannot tell a rule from a fit. Tahoma 12 and Segoe UI 14
  add four more chip heights, and they are what settled the elbow above and
  turned up both the group-by box rounding and the text-measurement difference

All 72 verify at 144dpi as well, on goldens recorded before any of the above,
so the three rules were re-derived without loosening the scale they came from.
The two new scenarios were recorded at 144 afterwards and matched on the first
verify: chips of 35 and 44px, elbows at 25 and 32, boxes of 66 and 80 -- nine
distinct chip heights now (19, 22, 25, 26, 30, 31, 35, 37, 44) with no
exception. 65's box is the one that pays: `CLng` of 52.5 is 66, where rounding
a half away from zero would have given 67.

### Added (M4 -- aggregates and group footers)

- `GroupFooterStyle` closes every group with a footer row: `jgexCaptionGroupFooter`
  repeats the group caption, `jgexTotalsGroupFooter` reads the columns' aggregates
  across the block. Closing a level is what records the span of rows it covered,
  so the spans exist whether or not footers are shown, and a footer carries no
  expand box -- it closes a group rather than opening one, which is also why the
  rule above it starts where the records start rather than at the group's own edge
- `JSRowData.GetSubTotal` computes all eight functions (`jgexCount`, `jgexSum`,
  `jgexAvg`, `jgexMin`, `jgexMax`, `jgexStdDev`, `jgexValueCount`) over the
  records a group row stands for, header or footer alike, skipping the nested
  group rows that sit inside the span. It is the same routine the totals footer
  paints from, driven per column by `JSColumn.AggregateFunction` and formatted by
  `TotalRowPrefix`/`TotalRowFormat`; totals sit on the cell origins a record uses,
  so they line up under the values they total
- A collapsed group hides its own footer along with its records, while a footer
  from a level further out still ends the hiding
- `GroupFooterStyle`, `JSColumn.AggregateFunction`, `TotalRowFormat` and
  `TotalRowPrefix` move to verified against scenarios `053-group-footer-caption`,
  `054-group-footer-totals`, `055-group-footer-prefix-format` and
  `056-group-footer-aggregates`, recorded from the original at both DPIs;
  `pvTestGroupFooter` in `ModelTests` covers `GetSubTotal` including the standard
  deviation and value count no scenario renders
- `RefreshSort` was an empty stub, which the Advanced Sample's sort dialog calls
  after rebuilding the keys; it now brings the rebuild forward from the next
  paint exactly as `RefreshGroups` does
- The Advanced Sample's M4 parts are covered by `pvTestAdvancedSampleParts` in
  `ModelTests`, which runs the API sequence each of its three dialogs uses --
  `frmSort` clearing and re-adding `SortKeys` then `RefreshSort`, `frmGroupBy`
  doing the same over `Groups` (including the all-collapsed overload), and
  `frmSummary` reading the keys back and totalling a column. The sample itself
  is not ported: its data is ADO and it opens on card view and print preview,
  which are M8, M6 and M11
- **M4 is complete**: no property it owns is left unimplemented or merely stored
  in `PROPERTIES.md`, and the corpus stands at 56 scenarios passing at 96 and
  120 dpi

### Fixed (M4 -- automatic sort)

- The click handler sized the group-by box without the room the painter adds for
  the chip staircase, so with two or more group levels every band below it was
  hit-tested too high and a click landed on the wrong row

### Fixed

- Every class holding the control through a weak reference now zeroes it in
  `Class_Terminate`, as `JSRowData` already did: `JSGroup`, `JSSortKey`,
  `JSGroups`, `JSSortKeys` and `JSColumns`/`JSColumn`. None of them AddRef the
  control, so releasing one that was still attached -- which `Groups.Clear` is
  the first caller to do while the control is alive -- had VB6 release a control
  it never AddRef-ed and took the process down. `Clear` and `Remove` both drop
  items that way, so it was reachable from ordinary client code
- A group row's `JSRowData` wrappers are detached before the array holding them
  is discarded, by `Erase`, by `ReDim`, or at control teardown. The order array
  and the group row array are discarded on different paths -- a sort with no
  grouping keeps the first and erases the second -- so `UserControl_Terminate`
  walking group rows up to `m_lOrderCount` indexed an unallocated array and took
  the process down with it, which is what made every `SortKeys` scenario hang
  behind an invisible error dialog on the isolated test desktop
- `003-headers-flat` set `HeaderStyle` to 0, the default, so despite its name it
  had been painting `jgexHSDouble3D` since M3c and the flat style was covered
  only at the default header height. It now sets 2 and carries rows; a sweep of
  all four values over a 400 twip header found no other disagreement
- A sort key change repainted from inside `Rebind`, which clears the keys halfway
  through its own reset: the partial repaint toggled `WS_VSCROLL` twice in quick
  succession and left the themed vertical scrollbar with a stale thumb, which
  `018-both-scrollbars` caught. Rebind now paints once, at the end
- `FontTextHeight`/`FontTextAscent` collapse into `FontTextMetrics`, which takes
  the font as `IFont` (no cast at the call site) and an optional `hDC`, so a
  caller already painting reuses its own DC instead of borrowing the screen's

### Changed

- `PAINT-PROPERTIES.md` renamed `PROPERTIES.md` and widened from the paint matrix
  to the full property inventory: all **288** public properties of `GridEX`,
  `GEXPreview` and the 23 `JS*` classes, read out of `doc\GridEX20.idl` so the
  list is the surface itself rather than whatever happens to be implemented.
  Part 1 is the paint matrix unchanged; part 2 adds the 33 `GridEX` data,
  binding and editor members that never reach the renderer; part 3 gives every
  class its own table, `JSColumn`'s 46 and `JSFmtCondition`'s 7 included, with
  properties already in the matrix marked `paint` so each class reads as a
  complete list. Status is evidence-based like the matrix: `consumed` when
  something in `src` reads the member behind the accessor, `storage` when it is
  only stored and round-tripped, `derived` when there is no member at all --
  which is also how the parameterized properties (`RowSelected(RowPosition)`,
  `HeaderString(Index)`, `Value(ColIndex)`) finally got written down; the
  snapshot engine skips them, so they had been missing from every inventory
- Part 2 is a second matrix in the same shape as the paint one -- `Property`,
  `Type`, `Status`, `Milestone`, `Commit`, `Test` -- over all 162 properties that
  never reach the renderer. `Milestone` is filled only for the 85 `storage` rows,
  the ones still owed behaviour (M11 printing 33, M8 binding 17, M5 editing 15,
  M10 styling 14, M4 sorting 5, M12 1), `Commit` points at the commit that
  introduced the consuming routine, and `Test` at what covers the property today.
  Both matrices dropped the `Paint routine` column

### Added (M3d complete -- the last five properties and the ported samples)

- `test\Samples`: the five original unbound samples (Unbound 1, 2, Array, Array
  UDTs, Collection) ported and running under a smoke runner that loads each form,
  lets it paint and asserts the grid took the sample's data. The port is the
  recipe the tooling was built for: re-point the reference, drop the grid's
  design-time property bag from the `.frm`, call the `Sub` generated from its
  snapshot by `tools\GenSample.ps1`. Unbound 1 and 2 read Products from
  `doc\Samples\JSNWind.mdb` through DAO exactly as the originals do; the other
  three build their data in code. Array and Array UDTs ship without their `.frx`,
  so their snapshot is raw-only and their design-time `DataMode` had to come
  across as code -- the one thing the generator could not supply
- `FrozenColumns`, `ColumnAutoResize`, `Col`, `ContinuousScroll` and `Redraw`, the
  five properties M3d still owed. Column layout now runs through one order array
  (`pvColOrder`: frozen block, then the scrollable rest from `LeftCol`) and one
  width function (`pvColWidth`), instead of six loops each walking the columns
  their own way
- Scenarios `031-frozen-columns`, `032-column-autoresize`,
  `033-record-navigator-string` and `034-hscroll-middle`; `ModelTests` gained
  `pvTestScrollProps` for the two properties no static picture can pin down.
  `RecordNavigatorString` moves from weak to verified, so every property the
  renderer reads is now proven at two or more values -- 49 of 126, and M3d has no
  unimplemented paint properties left

### Fixed

- The horizontal thumb sat one pixel left of the original's at any middle scroll
  position, ends matching, which is why nothing caught it until a scenario
  scrolled to the middle. VB6 maps a scrollbar's `Min`..`Max` onto 0..32767
  before handing it to Windows; `GetScrollInfo` on the original's bar shows it
  keeps the column numbers themselves (`min=3 max=5 page=1 pos=4`), so the fix is
  a `SetScrollInfo` over VB6's mapping. The harness grew a `scrollinfo` mode that
  reads range and thumb rect off either control -- reading the original's window
  back beats inferring its geometry from pixels
- `LeftCol` never clamped: it now stops at the last full page, `colCount - fit +
  1` counted over the scrollable strip, which is what the original does -- with
  two frozen columns taking 200 of 378px `LeftCol` = 4 stands, without them it
  clamps to 3
- The grid's bottom separator strip appeared only when the horizontal scrollbar
  did; the original puts it there whenever the band takes the last client row,
  including a record navigator with no scrollbar beside it
- `record.bat`/`make.bat` ran both DPI passes at system DPI: `Start-Process` does
  not pick up a `__COMPAT_LAYER` inherited from cmd, so the watchdog added in
  `37d573b` had silently turned "PASSED at both DPIs" into the same pass twice.
  Set inside the launching PowerShell it works, and the 96dpi goldens verify again

### Changed (test harness layout)

- Everything a ModelTests run produces now lands in `test\ModelTests\output\`:
  the results file, the round-trip dumps, the per-snapshot `.expected`/`.actual`
  diffs, the error log and the VB6 compiler log. The project folder had grown 51
  generated files around 8 sources. `OutputFile()` in `mdUtils.bas` builds the
  path and creates the folder, so every writer goes through one place, and
  `.gitignore` drops four patterns for one
- Snapshots renamed `NNN-<sample>_<form>_<control>.json`, matching the scenario
  convention. `NNN` is the sample's position in `export.ps1`'s walk, passed to
  the add-in as `OPENGEX_PREFIX`, so a re-export reproduces the names rather
  than dropping the numbering. The walk sorts ordinally now -- a culture-aware
  sort could renumber the corpus on another machine

### Fixed (test harness)

- Every event handler in the test projects, the add-in and both controls logs through `LogError` instead of a bare `Debug.Print`, which is invisible in a compiled exe: the message goes to `output\errors.log` next to the runner (`%TEMP%` for the OCX), timestamped, and the logger disables itself rather than fail. Handlers pass `Erl`, so a procedure carrying VB6 line numbers pinpoints the failing line -- which is how the hang below was tracked down. `frmHost`'s local `pvLogError` folded into it
- ModelTests hung instead of failing: `pvTestGeneratedSetup` errored past its `Unload oForm`, and a host form left loaded keeps the message loop alive long after the results file is written. The test cleans up through `QH` now, and `Form_Load` unloads whatever is left in `Forms` before unloading itself. All three `make.bat` scripts run the exe under a watchdog (120s model, 300s visual) so a wedged run can never hang the loop again
- `Set oCtl = oForm.GridEX1.Object` raised *Type mismatch* -- the extender's `.Object` hands back a plain `IDispatch` that will not QI to the control class, while the extender itself converts. `LogError` also resets `Err`, so the handler now reads `Err.Number`/`Err.Description` before logging

### Changed

- `PAINT-PROPERTIES.md`: alphabetical matrix of all 126 properties that affect painting -- the direct `GridEX` members plus the `JSColumn` and `JSFormatStyle` sub-properties -- with type, status, consuming paint routine, owning milestone, the commit where it started affecting pixels and the covering test. The bar for **verified** is two or more distinct values rendered and pixel-matched against the original: 43 properties are read by the paint path, only 19 clear that bar, 15 have no covering test at all and 83 are not implemented. `GridLinesColor` is the worked example: consumed by `pvLine` since M3c and set to `0x0000FF` by `gridlines-dots-colors`, yet that scenario declares no rows, so no data gridline is drawn and the golden holds no blue pixel
- `ROADMAP.md`: new **M6 -- Paint property matrix** gating the milestone formerly numbered M6 (ADO binding, now M7; persistence M8, styling M9, printing M10, long tail M11), so binding is built on a verified renderer. It also records a two gaps the audit exposed: card view (`View`, `CardBorders`, `CardCaptionPrefix`, `CardWidth`, `CardSpacing` and the `JSColumn` card members) plus the drag/resize affordances have no owning milestone at all (17 properties), and M9 carries 38 unimplemented paint properties behind a one-line roadmap entry
- README refreshed to current M2 state: milestone status, source-compatibility scope note, layout table covering `tools`/`test`/`doc/Help` and a Testing section for `test\ModelTests\make.bat`
- First clean VB6 build of the stub OCX; regenerated `OpenGridEX20.cmp` binary-compatibility baseline from the complete stubbed API surface (26 coclasses, 44 enums, 58 + 6 events, dispids and enum values verified against `doc/GridEX20.idl`)

### Added (data model -- `IDataModel` and the unbound and ADO implementations)

The grid's rows, its sort, its group rows and its aggregates move behind one
interface, so the same renderer can be fed by client-supplied rows, an ADO
recordset or -- later -- a SQLite table that does the sorting, grouping and
counting itself. **Nothing is wired into `GridEX.ctl` yet**: the control still
owns its own storage and projection, and the classes below compile, verify
against the corpus and are reached by nothing. Cutting the control over is the
next piece of work.

- `src/IDataModel.cls`, `Instancing = Private` so it never reaches the typelib
  (`DumpSurface` still reports 26 coclasses). Three address spaces meet on it
  and are kept apart deliberately: **RowPosition** `1..RowCount`, rows on show
  with group headers and footers in and collapsed records out; **RowIndex**
  `1..ItemCount`, the source's own order, where identity lives and which only a
  `Delete` renumbers; and **Bookmark**, resolving into RowIndex space. The
  original only ever *answers* in one direction -- `RowIndex(RowPosition)` and
  `RowBookmark(RowIndex)`, both running from the derived space toward identity
  -- and exposes the two reverses solely as actions (`MoveToRowIndex`,
  `MoveToBookmark`, `SelectedItems.AddBookmark`); a private interface can
  afford to make them callable
- `src/cUnboundDataModel.cls`: records indexed by RowIndex itself, no holes and
  no slot map, so a `Delete` compacts the array and renumbers above it. Values
  arrive through `UnboundReadData` once per record; writes go back out through
  `UnboundUpdate`, `UnboundAddNew` and `UnboundDelete`
- `src/cAdoDataModel.cls`: the recordset **is** the cache -- no cell value is
  copied. Per record it keeps only the ADO bookmark from the binding walk, and
  a cell read positions the cursor there and asks a `Field` object cached from
  `JSColumn.DataField`. A recordset that does not support bookmarks is refused
  at bind rather than failing in a paint, and the key pass walks the cursor
  with `MoveNext` since RowIndex order *is* the recordset's order
- `src/mdDataModel.bas`: the projection engine both share -- `UcsRowSet`,
  `DataReadSortKeys`, `DataProject` (seed, stable merge sort, group-row
  emission, visible map, position write-back), `DataBuildVisible` and
  `DataWritePositions` as the collapse path on their own, plus
  `DataCompareValues`, `DataBookmarkKey` and `DataAggregate`. The one step that
  cannot move is filling `uRowSet.Key`, the only part needing a cell value
- `JSRowData` gains a **buffer mode** beside its existing view mode: the
  wrapper carries the row itself rather than viewing control storage. Every
  member opens with an `If m_bView Then ... Exit` prologue and nothing else, so
  removing view mode is deleting those prologues, `frInit`, `m_bView` and
  `m_lRowIndex`. One signed index addresses both kinds of row -- positive a
  record, negative a group row -- matching what view mode already did
- `GridEX.ctl` gains six `Friend` raisers so a model can reach the events the
  typelib puts on the control: `frRaiseUnboundReadData`, `frRaiseUnboundAddNew`,
  `frRaiseUnboundUpdate`, `frRaiseUnboundDelete`, and `frFireFetchData` /
  `frFireFetchIcon`, which mint the `JSRet*` carrier and return the plain value
  so the carrier never leaves the control
- `JSColumn.FetchData` outranks whatever else a column could be read from --
  its `DataField` in a bound grid, `UnboundReadData`'s values in an unbound one
  -- identically in both models, and every consumer goes through the same cell
  accessor so sorting, grouping, totals and painting see the same value. The
  cache is per cell and fills on demand; the only thing that sweeps a whole
  column is a sort, which genuinely needs every key and leaves a warm cache
  behind. An edit refetches that record's fetch columns and marks the order
  stale only if a value actually moved; a delete shifts the rows above it down
- `mdGlobals.bas`: `C2Str` and `C2Lng` (the `C2Dbl` shape, `VariantChangeType`
  with `VARIANT_ALPHABOOL`, empty rather than raising), an `ArrPtr` declare, and
  `ToHex` fixed to work from `LBound` and to answer `""` for an unallocated
  array instead of raising on `UBound`
- `ARCHITECTURE.md`: the data model written up in full, with stubs for the
  window/subclassing, painting, DPI, input, object model, persistence, printing,
  build and testing sections

### Added (probe findings -- recorded from the original control)

Twelve probe passes drove the design above; the harness is a reduced copy of
the original *Unbound Collection* sample, kept outside the repo. The original
refuses a runtime `DataMode` change, so a design-time control is the only way
to drive an unbound grid at all -- which is why the probe is a sample form
rather than a runtime-created control the way `VisualDiff` does it.

- **Bookmarks** come from exactly one place, `GridEX.RowBookmark(RowIndex) = v`
  -- `JSRowData.Bookmark` is propget-only, and `Values.Bookmark` inside an event
  is the same value surfaced twice. They **survive both `Refetch` and `Rebind`**
  (our `pvResetRow(bFullReset:=True)` clears them, which is a live divergence),
  the store is sized to `ItemCount` and raises past the end, and a `Delete`
  **compacts** it: deleting the record holding `"bk-3"` leaves the next read of
  RowIndex 3 carrying `"bk-4"`
- **Duplicates are legal** and every resolver -- `MoveToBookmark`,
  `AddBookmark`, `RefreshRowBookmark` -- answers with the lowest RowIndex. An
  unknown bookmark raises `vbObjectError + 119`, *"Not a valid Bookmark."*,
  leaving the current row alone
- **Type is part of a bookmark's identity, width is not**: numeric widths
  resolve against each other, `Boolean` and `Date` resolve against their numeric
  value, a string never matches a number. Hence `DataBookmarkKey` keying `"S" &`
  for strings and `"#" & C2Dbl` for the rest
- **RowIndex is stable across a re-sort** -- positions permute while
  `RowIndex(pos)` reports the original supply order and `RowBookmark(1)` still
  answers for the record that was first
- **`JSRowData` instances are minted per population, never cached or pooled**:
  95 `RowFormat` fires produced 95 distinct instances with zero reuse, held
  alive in a collection so address recycling could not explain it.
  `GetRowData(pos)` hands back the current population's wrapper, and a wrapper a
  client keeps goes on reading the row it was filled with
- **`RowFormat` fires per population, not per paint**: a bare `WM_PAINT` fires
  none, while `Refresh` and a scroll away-and-back each re-fire for every
  visible row. What the handler writes does not survive the round trip
- **A held group header is invalidated by an order rebuild, not by a
  reprojection**: `Delete`, `RefreshGroups`, `RefreshSort` and `Rebind` each make
  `RecordCount`, `GetSubTotal`, `GetBookmarks` and `GetRowIndexes` raise
  `vbObjectError + 129`, *"JSRowData object may have changed. The object is no
  longer valid."*, while a collapse and expand leaves all four answering
  normally. `RowType`, `GroupLevel` and `GroupCaption` keep answering either
  way. Reproduced by a generation counter bumped in `DataProject` alone, which a
  wrapper captures and passes back -- a model that cannot lose a span this way
  may leave `Version` at 0, which never mismatches
- **`SelectedItems` raises** `vbObjectError + 123` unless `MultiSelect` is True;
  ours returns the collection unconditionally, a second live divergence
- **The control decrements `ItemCount` itself** on `Delete`, so neither the
  model nor client code may write it on that path

### Changed

- `CLAUDE.md`: order local variable declarations by how early each is used in
  the procedure, and never pass `Source` to `Err.Raise`
- Every procedure in the `Control events` section of `GridEX.ctl` now traps and
  reports through `PrintError` instead of letting the error out. An unhandled
  error in an event crosses the ActiveX boundary into the container, where VB6
  and the IDE both show it as the *host's* error and can take the host down
  with it; the base class events (`Initialize`, `Terminate`) are left alone, as
  a failure there is a failure to construct or tear down and must not be
  swallowed
- `PopPrintError` passed `LocalErr(2)`, the description, where `GetErrorSource`
  takes `ErrLine As Long`, so reporting any error whose description was not
  numeric raised type mismatch *inside the reporter* and threw the original
  error back out of the handler that was trying to contain it. It now passes
  `LocalErr(3)`, the `Erl` that `PushError` was already capturing and dropping,
  and reports through `LogError` so a compiled run -- where `Debug.Print` goes
  nowhere -- still leaves a trace
- Every `LogError` callsite across the OCX, both preview controls, all four
  test harnesses and the export add-in now goes through a per-module
  `PrintError(sFunction)` helper, which each module carries in its own
  `Error management` section beside a `MODULE_NAME` const -- the shape
  `GridEX.ctl` and `mdJson.bas` already used. A handler is one line again, and
  the message it produces gained the module name and the error number the
  hand-built strings mostly lacked
- `tools\common\mdUtils.bas` gained `PushError`, `PopPrintError` and
  `GetErrorSource` so the test and tool projects report the same way the OCX
  does. The split matters: `PushError` is passed as an *argument*, so `Erl` is
  read in the frame that failed -- read inside the reporter it would be the
  reporter's own line, which is what the `#Else` branch of `mdJson.bas`
  quietly loses
- Nine handlers read `Err.Number`/`Err.Description` *after* reporting, to build
  an assert message, a `.err` file or an error summary. Executing any
  `On Error` statement clears `Err` and the reporter has one, so every one of
  those had been logging `&H0` with an empty description since it was written.
  They now capture into a local first and report second

### Changed (the grid draws from the data model)

`GridEX.ctl` no longer owns rows, order, group rows or aggregates. It asks an
`IDataModel` for a page of `JSRowData` and paints from that, which is what the
data model was written for. The control lost ~810 lines.

- **Gone from the control**: `m_aRows`, `m_aOrder`, `m_aGroupRow`, `m_aVisible`
  and the sort-key decoration arrays; `pvBuildOrder`, `pvBuildGroupRows`,
  `pvCloseGroups`, `pvBuildVisible`, `pvDecorate`, `pvMergeSort`,
  `pvInsertionSort`, `pvCompareRows`, `pvCompareValues`, `pvSlot`, `pvSlotPos`,
  `pvDataRow`, `pvDataRowPos`, `pvGroupSlot`, `pvEnsureRow`, `pvFetchRow`,
  `pvResetRow(s)`, `pvAggregate`, `pvGroupCaption`, and all 26 `frRow*`
  accessors. `UcsCellData`, `UcsRowData` and `UcsGroupRow` went with them
- **The page buffer**: `pvPopulateWindow` mints one wrapper per visible
  position and raises `RowFormat` on each, once. Repaints walk the buffers, so
  a bare `WM_PAINT` raises nothing -- which is what the original does
- **`JSRowData` lost view mode**: 29 `If m_bView Then ... Exit` prologues,
  `frInit`, `frInitBuffer`, `m_bView`, `m_lRowIndex` and the weak owner
  reference are gone. A wrapper is now a snapshot that never reads back
  through the control, so one a client keeps cannot be orphaned by the grid
  going away -- it answers for the row it was filled with, as the original's
  does
- **The current row and the selection** follow their records across a
  reprojection by holding a RowIndex -- the one address space a reprojection
  cannot move -- and resolving it back through `GetRowPosition` when the
  model's `Version` moves. A collapse deliberately does not bump `Version`
  (held group wrappers must survive one), so it remaps directly instead
- **Divergence closed**: bookmarks now survive `Rebind`, as probed on the
  original. The test that asserted the old behaviour was inverted

### Fixed

- `JSRowData.Value` is read-only outside the fill the control asks for, and
  raises `vbObjectError` with *"'Value' property can not be change in this
  context."* -- the original's own wording, including the grammar slip. Probed:
  the original refuses the write on a record, on a fresh wrapper and on a group
  row alike, while `Values.Value(n) = ...` inside `UnboundReadData` still works
- An edit commit no longer reprojects when there are no sort keys. It used to
  mark the order stale unconditionally -- a `TODO` in the model admitted this
  was unprobed -- which bumped `Version` and made the grid repopulate its page,
  re-raising `RowFormat` for every visible row where the original raises it
  once for the edited row. The event goldens of 058, 068 and 071 pin it
- An edit commit no longer raises `UnboundUpdate`. The original puts the row in
  its own store and reports only `AfterUpdate`
- `pvCreateDataModel` moved to `UserControl_Initialize`. It ran from
  `InitProperties`/`ReadProperties`, but `Initialize` already invalidates
  through the font handlers, so every paint before it raised error 91 into a
  handler that swallowed it -- 408 times in one test run -- and the grid
  silently stopped updating its scrollbars
- The checkbox column paints the pending buffer, so a toggle shows immediately
  rather than one commit behind

### Changed (Integer -> Long internally)

Every internal column index, count and loop counter is a `Long`. The public
surface keeps the original's `Integer` signatures to the letter -- events,
properties and methods are what a client compiles against -- so the change
stops at the boundary and the typelib is untouched.

- **Converted**: the projection's `UcsGroupRow.Level`/`.ColIndex`,
  `UcsRowSet.SortCol`, `UcsFetchCache.ColCount`, `UcsRecord.CellCount` and its
  icon array, `JSRowData`'s cell store and its `fr*` fillers, both models'
  `pvCellValue`/`pvGroupCaption`/`pvGroupValues`/`IDataModel.GetSubTotal`, and
  the control's `m_nCol`, `m_nLeftCol`, `m_nFrozenColumns`,
  `m_nPreviewRowLines`, `m_nEditCol` with every private and friend procedure
  they reach. Renamed with the type: `nIdx` -> `lIdx`, `nColIndex` ->
  `lColIndex`, `m_nCol` -> `m_lCol`
- **Left as `Integer` on purpose**: everything the public interface fixes (a
  `Public Event` parameter is passed `ByRef`, so the local feeding
  `RaiseEvent KeyDown` has to match); `pvShiftState`, `pvMouseShift` and
  `pvMouseButton`, which exist to fill those parameters; `pvLoWord`/`pvHiWord`,
  whose `nWord` is a two-byte `CopyMemory` target; the `JS*` members backing an
  `Integer` property; and the API-mandated ones in `mdGlobals.bas`
  (`GetKeyState`'s `SHORT`, `TEXTMETRIC`'s char fields, the `FreeFile` number)
- Five `Dim nIdx As Integer` declarations left unused by the data-model move
  went with the sweep

### Changed (the current row's JSRowData is the pending row)

The control keeps one wrapper for the current row -- `m_oRowData`, the page's
own instance of it -- and that wrapper is also the edit buffer. It already
carries every column, which is what `UpdateRowData` writes back from, so a
commit hands it to the model as it stands. Four pieces of state went away with
the `m_aPendValue`/`m_aPendDirty` pair it replaced: `m_bPendAny`, `m_lPendRow`,
the separate buffer object, and `m_bDataChanged`.

- `JSRowData` grew the per-cell dirty flag it had been missing. `UcsCell.Dirty`
  is set by a write through the public `Value`, not by the fill a model does
  through `frSetValue`, so what it marks is a cell modified since the wrapper
  was filled. `frCellDirty(lColIndex)` reads it back, bounds-checked
- `frAllowUpdate` is the pending state, not a flag beside it: the row is
  writable for exactly as long as it is the pending row, and `DataChanged`
  answers off it -- so a single object holds what the control paints, what the
  client reads through `Value`, and whether any of it is uncommitted. Outside
  its own edit the wrapper refuses a write the way the original does
- Nothing overlays a buffer over the row any more, so `pvPendSet`/`pvPendGet`
  are gone: `Value` reads and writes the wrapper, the checkbox paints its
  value, and the editor's commit is a write like any other. The dirty mark is
  what a cell being painted still asks about -- an uncommitted write shows what
  was written rather than the `DisplayValue` the row was decorated with before
  it -- and clearing `frAllowUpdate` clears the marks with it, since a row that
  has stopped being the pending one has nothing uncommitted about it. The
  values that were written stand
- A commit keeps the wrapper it wrote from, decoration and all. A cancel cannot
  un-write it, since what the cells were written into is the row the page
  paints, so `pvCancelRow` re-reads the row from the model into the page's slot
  and `RowFormat` re-decorates the one row -- which is what the goldens of the
  two-level Escape pin
- `pvSyncRowData` re-takes the wrapper wherever `m_lRow` is assigned and
  wherever the page replaces an instance, and it is where a row is left, so it
  is where the buffer goes through: it commits when the instance it is letting
  go of is not the one it takes. That covers the paths that move the row
  without `pvLeaveCell` -- a `Rebind`, or the remap a reprojection does
- A page rebuilt under a pending row hands the same instance back rather than
  re-reading that row, so a scroll mid-edit is not a leave and cannot drop what
  the wrapper holds. `pvPendCommit` also refuses to re-enter: a client handler
  of the events `UpdateRowData` raises can move the row, which comes back
  through the sync with the write half-done
- A cancel can no longer revert by dropping a buffer, since what it would drop
  is the row the page paints. `pvRefreshWindowRow` re-reads that one slot from
  the model and `RowFormat` re-decorates it -- one row, one event, which is
  what the goldens of the two-level Escape pin
- The wrapper follows the row, so moving the current row onto a record nothing
  has read yet reads it. On screen that is invisible -- painting the row reads
  the record anyway and the model answers the second ask from its cache -- but
  `ModelTests` never paints, and its nav event order records the fetch now
- The dirty flag is also what `cAdoDataModel.pvWriteFields` needs to stop
  writing every bound column on an update -- noted there as missing since the
  model was written, and still to do

### Fixed (RowFormat went to the wrong row)

The row flush raised `RowFormat` twice for the row it had just committed. The
original raises the second one for the row being *landed on*, not a second time
for the one being left -- which is why the corpus never caught it: the harness
skipped object-valued event parameters, and `RowFormat`'s only parameter is the
`JSRowData`, so every one of them logged as a bare `RowFormat()` and a repeat of
one row read the same as a sweep of two.

- `pvFormatEvent` logs `RowFormat(RowIndex=n)` -- the one carrier whose identity
  is behaviour rather than plumbing. The six goldens carrying a `RowFormat` line
  were re-recorded from the original at 96 and 120dpi and differ by exactly that
  annotation, pixels untouched. 058 and 059 were re-recorded at 144dpi with the
  rest of that scale (below)
- Scenario 075 moves the current row with no edit anywhere near it, and the
  original still raises `RowFormat` for the row landed on, ahead of
  `SelectionChange` and `RowColChange`. So it belongs to becoming current, not
  to the commit: `pvNavigate` raises it when the row changes, and a column move
  raises nothing -- 067 has no `RowFormat` at all
- `pvCommitRow` (was `pvFlushRow`, with `pvPendCommit` folded into it) raises
  the other one, for the row it just wrote, between `BeforeUpdate` and
  `AfterUpdate`
- The row half moved out of `pvEndEdit`'s "the text changed" branch to the two
  places that leave a row. It used to be skipped when an editor was closed on
  the same text it opened with, which lost a column edited earlier in the row
- `pvLeaveCell` went with it: `pvEndEdit` already returns when no editor is
  open, so what was left was the two calls its callers can make themselves.
  `pvEndEdit` takes `Optional bCancel` now, so the four ordinary callers pass
  nothing and only Escape says what it is

### Fixed (the bars, the editor and the paint)

- The horizontal bar is driven by the `WM_HSCROLL` its own window sends to the
  parent, decoded like the vertical one has always been -- line, page, thumb
  and end. It used to come through VB's `Change`/`Scroll` events, which report
  VB's `Value` rather than the column numbers written onto the window, so a
  drag that ended anywhere reported the first column. VB's `Min`/`Max`/`Value`
  are no longer written at all: two sets of numbers on one window is what made
  the two disagree
- `LeftCol` and `FirstItem` are re-clamped wherever the room changes, not just
  where they are assigned: widening a grid parked at the right end, or making
  one taller parked at the bottom, brings the columns and rows hidden off the
  other side back into view instead of painting blank where they were. The two
  are not symmetric -- a client's `FirstItem` past the end sticks where every
  row fits, which the nav event order pins, while `LeftCol` goes back to the
  first column
- The in-place editor follows its cell. It was placed once, at creation, and
  left there: scrolling moved the grid out from under it. `pvLayoutEditor` runs
  on every repaint and moves it, clips it where its column is half off the
  right edge, and puts it away when the cell is off the view altogether --
  which needed `pvCellRect` keyed by the column itself, since a column's place
  among the visible ones moves with `LeftCol`
- A cell is scrolled all the way into view before its editor opens on it, and
  the click is carried along with the scroll so the caret still lands where it
  was put
- The editor goes away while more than one row is selected, and a block of rows
  paints as one: no current cell is lifted out of it
- Ctrl+click selects and nothing else -- no editor opens under it -- and the
  click that does open one keeps its drag, since the button is still down over
  the grid and the first move would otherwise drag a selection out of the cell
  being edited
- `WM_ERASEBKGND` is answered by the grid: `pvPaint` covers the whole surface,
  so erasing first only paints the background twice
- Painting goes through `InvalidateRect`/`UpdateWindow` rather than VB's
  `Refresh`. With `HasDC = False` on both surfaces, `hDC` answers for the paint
  in hand and nothing else, so the two places that measure text for a hit test
  ask for a DC of their own -- and the band, which fills the full width, puts
  the scrollbar back over itself afterwards

### Changed (naming and shape)

- `LoWord`/`HiWord` are unsigned, as the macros they mirror are, and the signed
  pair is `GetXLParam`/`GetYLParam` after the `windowsx.h` names -- a scroll
  position counts, a mouse coordinate measures, and only one of them may come
  back negative. `MakeDWord` and a `Clamp` sit beside them in `mdGlobals.bas`
- `frRaiseFetchData`/`frRaiseFetchIcon` join the four `frRaiseUnbound*` raisers
  under one prefix; `frSortChanged` is `frNotifySortChanged`, which is what it
  does -- a collection telling the control, not an event going out
- `IDataModel.RefreshGroups` takes `bAllCollapsed`, so the rebuild and the
  collapse are one pass rather than a reprojection each
- `pvInvalidate` takes `SkipScroll`, passed by the eleven callers that move the
  marquee, the selection or a cell's contents -- none of which can change what
  a scrollbar is
- Dead state removed: `m_bHScroll` and `m_bSortDirty` were written and never
  read. `m_bBand` is `m_bBandVisible`, `m_hScrollH` is `m_hWndHScroll`, and the
  member block is property-backing storage first, then the control's own state

### Changed (weak references and the horizontal page)

- The eight classes that point back at their owner set and clear that pointer
  through `mdGlobals.AssignWeakRef` rather than each writing it raw. Under the IDE
  it takes a real reference instead, so Stop never finds a member holding a
  pointer nothing AddRef-ed -- a leak while debugging in place of a crash
- A horizontal page is what the strip can hold rather than a fixed number of
  columns: paging right starts the view on the column that was cut off at the
  right edge, paging left leaves the column the view started on cut off there
  instead. It had been reading `nPage` back off the bar, which the range in
  column numbers keeps at 1, so a page moved exactly one column
- `pvVisibleColsInWidth` is `pvColsFitFrom`, beside a new `pvColsFitBefore`
  counting the other way, which is what the two page directions ask for. Both
  answer at least one, so a page always moves

### Added (columns under the mouse: sizing, auto-size and moving)

- A column is resized by dragging the divider on its header, for as long as the
  button is down. The divider takes the press before the header under it does,
  or every resize would sort the column it started on, and `AllowSizing` False
  leaves the press to the header. The grid takes the capture and `WM_CANCELMODE`
  -- a message box going up under the drag -- puts the width back without asking
- `ColResize` is raised once, when the button comes up, with `Column.Width`
  still the width the drag started from, and the new width applied only if the
  client does not cancel: probed off the original, which reports the same way.
  Ten pixels is as narrow as a drag can make a column, at 96dpi and at 120dpi
  alike -- scenarios `076`/`077`, recorded from the original at both
- Double-clicking a divider sizes the column to what it holds, and so does
  `JSColumn.AutoSize`: the widest of the values on the page in the data font,
  plus the two pixels a cell starts its text at and the three it clips at, and
  the caption in the header's own font with the room that one leaves it.
  Scrolled-away values are not measured -- scenarios `078`/`079`/`080`
- A header dragged onto another column moves it there, raising `BeforeColMove`
  with the new position and a cancel, then `AfterColMove`. A drop outside the
  client area, or back on the column it started from, moves nothing.
  `JSColumn.ColPosition` set by a client renumbers through the same path, so
  neither route leaves a gap in the positions
- The header being dragged paints inverted, a three-pixel red rule marks the
  boundary the column would land on if it were dropped now, and the pointer
  carries the four-way mover. The mover and the sizer are put up from the drag
  itself: a capture keeps `WM_SETCURSOR` from ever arriving
- `ColumnHeaderClick` and the automatic sort moved from the press to the
  release. A press is still a click for as long as the pointer stays inside the
  double-click box around it -- `SM_CXDOUBLECLK`/`SM_CYDOUBLECLK`, the slack the
  shell gives a second click -- and once it leaves, the press is a move, which
  takes the header click with it
- Ctrl+click on the current row takes it out of the selection: it keeps the
  focus rect and loses the selection colour, which the original does and a
  current row painted selected unconditionally could not

### Fixed (flicker)

- The band erased itself before the navigator painted it. `pvInvalidate` asked
  for an erase on the outer control, and that window's branch of the subclass
  had no `WM_ERASEBKGND` case at all, so VB filled the band and
  `UserControl_Paint` painted it again -- two paints per update. Both ends are
  fixed. The grid surface had answered the message all along and in practice
  never sees one: every invalidation on it asks for no erase
- A column dragged over the headers repaints the strip above the rows --
  `pvInvalidateHeaders`, group-by box included -- and only when the header under
  the pointer changes, rather than the whole grid on every mouse move
- A cell fills its own rectangle and stops one pixel short of its right edge
  wherever a vertical gridline stands. The row had been filled in one go across
  the whole block, so every rule was cleared and drawn again on each repaint,
  which is what made the column boundaries shimmer

### Changed (the control paints its own WM_PAINT, one band at a time)

- `WM_PAINT` is answered in the subclass rather than through VB's `Paint`
  event, on the grid surface and on the outer control both. `BeginPaint` hands
  out the DC and, with it, `rcPaint`: a band the update region does not reach
  is stepped over instead of being drawn for GDI to clip away, so a drag over
  the headers allocates one bitmap and touches no rows
- Every band paints into an off-screen bitmap and arrives in one blit --
  `pvBufferBegin`/`pvBufferEnd`, with the DC and the bitmap living no longer
  than the band they carry. The group-by box, the header band, the navigator
  strip and every row are separate bands, which keeps the bitmap the size of a
  row rather than the client. Painted straight onto the window a band shows
  every stage it goes through -- the fill, then what is drawn over it -- and
  that is what the eye reads as flicker
- The rows loop takes the height per row and accumulates the top, so rows that
  differ in height need nothing more than a taller bitmap. Empty rows are rows
  in the same loop, and the strip right of the last column belongs to the row
  it sits beside; the background below the block is what is left to a plain
  `FillRect`
- The marquee and the column rules moved into the row from the block-wide
  passes that used to follow it, in the order they had -- the rules close over
  the cells, the horizontal rule and the marquee's dots alike. Solid, dotted
  and dashed rules all reproduce a block-long line one row at a time
- VB filled the client with `BackColor` before raising `Paint`, which is where
  the pixels no painter covers came from -- the tree indent inside the header
  band, the band beside the navigator. Each band lays that background down for
  itself now, and only the first line of a band is seeded from the surface: the
  line a row shares with the row above, where a group row draws its opening rule
- The row header's border pair closes one line below its row under
  vertical-only gridlines, which is the next row's line rather than its own, so
  the block draws the last one after its rows

### Fixed (chrome the original does not colour)

- The record navigator takes the control's own background and the system's
  button text: its band, its button faces, its text and the separator strip
  along the grid's bottom edge had been drawn from `BackColorHeader` and
  `ForeColorHeader`, which looked right only while those properties kept their
  defaults. `021-colors-chrome` now sets a colour on every one of them and the
  navigator with it, which is what pins this
- A header dragged past either edge scrolls the strip a column at a time, which
  is how a column reaches a position that is not in view. A timer carries the
  scrolling on while the pointer stays out there, since a pointer that has
  stopped moving sends nothing more. The drop stays cancelled for as long as it
  is outside -- there is nowhere out there to drop on -- so the gesture is out,
  wait, back in, release
- Empty rows carry the record selector on past the last record, chrome cell by
  chrome cell, as the original does -- and with empty rows switched on those
  cells run under the records too, so a record's cell closes over one that
  reaches a line further down. An empty row's cell closes one line short of
  itself whichever gridlines are set, where a record's follows the mode
  (`081`-`083`, recorded from the original at both DPIs)

### Fixed (the 144dpi pass)

- The view follows the current row when the group set changes: `FirstItem`
  becomes the current row's position, clamped at the last full page. Probed on
  the original -- `Row` 4 gives 4 and `Row` 9 gives 6 against a five-row view,
  which is neither ensure-visible nor a preserved screen offset. Only 144dpi
  shows it, since nothing in the corpus overflows the view at 96 or 120
- Seven event goldens at 144dpi and one at 96 had recorded a gesture that does
  not land at that scale -- the divider `076`-`080` press for is 50 pixels to
  the right, and the third click of `073`/`074` falls on the open editor. What
  the original logs there is `DragDetect` answering synthetic input, so those
  scales carry no log for them and are asked about their pixels alone, as
  `golden\120` already was for `076`-`080`
- `Samples.vbp` had carried `Startup="frmUnbound1"` since `1a680f6`, where the
  IDE rewrote it: the smoke runner showed the first sample and sat there until
  the watchdog killed it, so none of the five had been checked since. The same
  save put three debug edits into `frmUnbound1.frm`, also reverted
- `ColMove: and again while it stays out there` expected a second scroll out of
  a second identical `WM_MOUSEMOVE`, which raises none -- the timer carries it,
  and a test that never pumps cannot see it fire. It sends the tick itself now
- The visual suite's watchdog was 20 seconds, which the corpus had grown past at
  125%: a run was killed partway and read as a failure. Two minutes

### Changed (one bitmap for the paint, band by band)

- The band buffer is allocated once per paint rather than once per band, as tall
  as the tallest band that goes through it. A page of rows used to create,
  select, blit and delete a bitmap apiece
- The line a row shares with the row above is scrolled up out of the previous
  band rather than read back off the surface. Where that band was skipped the
  line is stale and provably outside the update region, so `BeginPaint`'s clip
  drops it
- Each band paints inside a clip box its own size, so a rule the height of a row
  ends at the row's edges rather than running the length of the bitmap
- `pvPaintGroupChips` and `pvChipStagger` folded into their callers, and a
  group-by box chip inverts with the column it stands for while that column is
  being dragged, as its header does

### Fixed (dashed rules, and what a gap in one shows)

`026-gridlines-horizontal` was re-recorded with `GridLineStyle` on dashes, which
the corpus had never carried -- `028` was the only dashed scenario and it sets
vertical lines only. One capture showed three things.

- GDI's cosmetic `PS_DASH` is 18 on and 6 off where the original's rule is 3 and
  3, the same trap `PS_DOT` set, so dashes are stamped pixel by pixel
- The phase is anchored on the right end of the run, not its start: the block
  begins at x=20 at 96 and at 120dpi alike and the first dash lands at 22 and at
  23, which only the far end accounts for
- A row's colour stops one line short of the rule closing it -- the gaps under
  the current row show the grid background, not the selection. `pvRowContentH`
  already named that line as the gridline's for text and the marquee; the cell
  fill now agrees. No golden could have caught it, since solid and dotted rules
  both cover the line
- The rule closing a row is laid down with each cell rather than in one run over
  the finished row, and `pvDottedLine`/`pvDashedLine` are one `pvStampedLine`:
  `lOn` marks out of every `lPeriod`, counted from `lAnchor`. Parity ignores
  sign, so a dotted run's `(x + y) Mod 2` is the same test with the cross-axis
  coordinate for an anchor. `pvLine` picks the anchor, which is the one thing
  the two patterns disagree about

### Added (a way to test the buffered paint from a remote session)

- The remote-session check in `pvBufferOpen` sits behind `#If FORCE_BUFFER = 0`,
  and `src\make.bat buffer` passes `/d FORCE_BUFFER=1`. On RDP the buffer is off
  and the corpus says nothing about the bitmap path at all
- The comparison is the point: `#If Not FORCE_BUFFER` reads correctly and is
  wrong, because `Not` is bitwise -- `Not 1` is `-2`, which the preprocessor
  takes for true, so the override did nothing
- Verified with two mutations the corpus has to catch, the band blit put a pixel
  low and the shared line carried from the wrong row. Both are caught, so the
  bitmap and the scroll are on the path the paint takes

### Added (a header dragged into the group by box)

- A header let go in the group-by box groups the grid by that column, which is
  the gesture the box's own info text asks for. The original names the operation
  `jgexGroupInsert` and gates it on `BeforeGroupChange`, handing over a `JSGroup`
  that is not in the collection yet with the position it would take;
  `AfterGroupChange` follows once it is in. Neither event had been raised by
  anything until now
- Three red pixels stand where the drop would land -- on the right border of the
  last chip, or where the first one would start with nothing in the box yet --
  the same mark the header row puts on a boundary a column would move to
- A drop appends. Where between two chips the pointer has to be to land between
  them is not probed, and neither is anything the original paints in the box
  mid-drag: the harness releases at the end of every drag it drives, so a drag
  in flight is the one thing the corpus cannot capture

### Changed (tests say where they click in twips)

- `ModelTests` drives its mouse messages through `TwipsDWord`, so a press lands
  on the same part of the grid whatever the screen is: with two 1500 twip
  columns, 750 is the middle of the first and 2250 the middle of the second.
  The pixel counts they replace only meant that at one scale, and said nothing
  about the widths the same test had just set
- The navigator clicks stay in pixels, and say so: those buttons are laid out
  off the band height and the text extent, so where they land does not scale

### Changed (mdDataModel answers for its own errors)

- The module had no error handling at all. Every public entry point now carries
  `On Error GoTo EH` and `RaiseError`, which annotates the error with the
  procedure it came from and re-raises it -- the same shape `mdJson` uses
- A row index outside `1..ItemCount` is no longer a fault. `DataRowBookmark`
  answers `Empty` and `DataSetRowBookmark` writes nothing, where the original
  raises "Subscript out of range" -- a deliberate divergence, since the paint
  path asks that question off counts that may have moved under it

### Added (the group by box takes drops, and chips travel)

- A chip can be picked up and put down: dropped on another chip it takes that
  place in the order (`jgexGroupMove`), dropped anywhere off the box it stops
  being a level at all (`jgexGroupDelete`), and a header under it says where its
  column goes on the way out. `BeforeGroupDrag` gates the gesture, which is what
  that event is for -- it had never been raised by anything
- Which half of the target the pointer is over decides which side of it the drop
  goes, for chips and for column headers alike, and the red mark stands on that
  side. A drop that would put the thing back where it already is -- the near
  half of either neighbour, or the target being the dragged thing itself --
  marks nothing and does nothing
- `GroupByBoxHeaderClick` is raised on the release, the way `ColumnHeaderClick`
  is: until the button comes up a press on a chip is still a move waiting to
  happen, and the old press-time click meant a chip had already sorted by the
  time it was dropped somewhere else

### Fixed (what counts as a drop target)

- `pvColAtX` answers for any y it is handed, so the header row is bounded at
  both ends now. A chip taken off the top of the control used to resolve
  whatever column its x crossed and carry that column along with the delete;
  above the control is off it, not in the band the coordinates would put it in
- The drop is read off the release rather than off the last move before it.
  Which half of the target the pointer is over cannot be decided by anything
  else, and letting go past the client edge is a cancel rather than a drop
  wherever the pointer was last seen inside
- `pvMoveGroup` took the shift for the removed level off twice -- once in
  `pvDropPosition` and again after `m_oGroups.Remove` -- so a chip dragged
  rightwards landed a place short, which with two chips is where it started.
  Only the leftward reorder had a test, and leftward never subtracts

### Fixed (a horizontally scrolled block)

- `Col` numbers the visible columns from the first one, not from `LeftCol`.
  `pvColAtX`, `pvColByPosition` and the paint loop each counted over the render
  order, whose scrollable part starts at `LeftCol`, so once the block was
  scrolled the current cell matched no column and was painted as part of the
  selected row instead of being lifted out of it -- and a click reported a `Col`
  that named a different cell than it had before the scroll. New
  `pvVisiblePosition` counts from position 1; at `LeftCol = 1` the two
  numberings agree, which is why only the narrow 144dpi goldens caught it
- `BeforeColEdit` is raised from where the cell has landed. `pvBeginEdit`
  scrolled the cell into view after asking the client for its veto, so the
  original's `LeftColChange` came first and ours came second

### Fixed (harness)

- Scenario mouse coordinates are twips and are converted at send time
  (`TwipsDWord`), so a scenario points at the same place at every DPI instead of
  at whatever a fixed pixel lands on. Eight event logs had recorded the drift
- `CaptureWindowClient` hides the caret before it captures. A blinking caret is
  in the shot or not depending on when the capture falls, for the original as
  much as for us, and it had been read as a rendering difference

### Changed (where a drop lands)

- The group by box is cut at the chips' left edges and nowhere else. It used to
  hit-test as the staircase of rectangles it is drawn as, so the drop target
  depended on which step the pointer was level with; now y is not asked at all,
  the run before the first chip belongs to it as much as the run past the last
  belongs to the last, and an empty box takes a first level anywhere along it
- Everything below the box answers for the header it stands in the column of,
  the rows as much as the header row itself -- a column being carried about is
  somewhere in the order the whole way down. Off the control is still off it,
  and there the pointer walks the block sideways instead

### Fixed (the end of the horizontal scroll)

- `LeftCol` stops where the last column reaches the right border, which is
  counted back from that column rather than forward from the first. How many
  fit from the left says nothing about how many fit at the end once the widths
  differ: with 100, 100 and 300 in a 350 pixel strip the forward count stopped
  the scroll at the second column and painted the third fifty pixels past the
  border with the thumb already at its stop. The scrollbar's own maximum and
  the `LeftCol` setter had each worked it out separately and now share
  `pvMaxLeftCol`

### Fixed (two review sweeps, the control)

- A plain click no longer leaves the pressed header or chip inverted: every
  path that clears the press without a drag's own repaint -- the click branch
  of the release, a cancel on a press, a refused drag -- invalidates the bands
- `BeforeColumnDrag` is raised as a header press turns into a drag, the gate
  the typelib always declared and nothing ever raised; a veto ends the gesture
- `EnsureVisible` honours its `Col` argument now that the machinery exists,
  `Row` clamps to the block the way `Col` clamps to its columns, and an
  out-of-range position can no longer plant a phantom selected item
- Runtime property changes repaint: 28 visual setters invalidate and
  `GroupFooterStyle` refreshes the groups its rows come from -- before this a
  `GridLines` change at runtime showed nothing until something else painted
- The width math measures from `pvBlockLeft`, so the tree indent counts:
  the scrollbar appears when grouping pushes the columns past the edge,
  auto-resize fits beside the indent, and the last magic `18`s are gone
- `pvMaxLeftCol` answers in column-position space, so a hidden column no
  longer shears the scrollbar range against `LeftCol`
- A group level whose column is gone has no chip and takes no drop, and
  `pvEndEdit` writes the committed text into the row it was typed into even
  when a `BeforeColUpdate` handler moves the current row

### Fixed (two review sweeps, the models)

- The three date sort types compare chronologically through the new `C2Date`
  instead of as the text a date renders at the locale -- sorting, group
  boundaries and Min/Max all took the fix at once
- What a client keeps outlives what it was taken from: `Remove` and `Clear` on
  the columns, groups and sort keys detach the items they drop, so a held one
  answers for itself instead of dereferencing a freed control; removing a
  column also closes the hole in the position space and notifies the control
- `JSSelectedItems` resolves whichever of a row's three names the caller
  hands its public Adds into the other two, and `RemoveBookmark` removes by
  key -- `=` on an ADO byte-array bookmark was a type mismatch
- A record cut off by a smaller `ItemCount` stops existing: growing back
  exposes a fresh row rather than resurrecting the dead one's bookmark
- `pvWriteFields` writes only the cells the buffer marks dirty, which is what
  keeps a commit from fighting optimistic locking on a wide table
- `RefreshRowIndex` re-sorts only when sort keys exist, the same guard the
  update path applies; ADO's `Resync` moved behind a named constant and a
  proper handler
- Both models carry the `mdDataModel` error convention now: `RaiseError`
  chains on every `IDataModel` member and `frInit`, a logging `PrintError` on
  `frTerminate` where raising out of teardown would crash
- `MakeDWord` survives a sign-bit high word, `AssignWeakRef` is the Sub it
  behaved as, a non-byte array keys no bookmark instead of raising, and
  `JSRowData.PreviewRowVisible` defaults False

### Changed (walking collections)

- Indexed `For lIdx = 1 To Count` walks over `VBA.Collection`s are `For Each`
  throughout src -- `Item(i)` walks to position i, so the old loops were
  quadratic. Position-space walks through `ItemByPosition` stay indexed, since
  their iteration order is the point

### Fixed (harness)

- `AssertEquals` fails one assert on a value it cannot compare -- a Null or an
  array used to raise through `Form_Load` and silently skip every test after
  the bad one
- GDI+ starts once per process instead of a startup/shutdown pair per PNG,
  and two stray `- Copy` files left `test\Samples`

### Changed (documents)

- `ARCHITECTURE`, `ROADMAP`, `PROPERTIES` and `README` caught up with the
  code: the closed dirty-flag gap, the detach discipline, the three-scale
  corpus, the M10/unowned property counts reconciled at 36/3, and the group-by
  drop work in the M5 summary

### Added (M5 -- the keyboard)

- `mdIPAO.bas` hooks `IOleInPlaceActiveObject.TranslateAccelerator` through a
  vtable thunk: the container translates accelerators before the window ever
  sees `WM_KEYDOWN`, which is where a host steals arrows and tabs --
  `frBeforeTranslateAccel` claims the grid's navigation keys and every key
  the in-place editor works with (the EDIT is a plain child the container has
  never heard of, so an unclaimed arrow becomes dialog navigation and Enter
  or Escape fall to a Default or Cancel button), then hand-delivers each to
  the window it was going to -- the window path stays the one place a key is
  raised and dispatched, and posted keys walk the same road as real input
- The keys the original taught by probing: a single-line editor gives up the
  vertical arrows anywhere and Left at the very start, a wrapping one keeps
  every arrow to walk its lines except Right at the very end of the text,
  which leaves either kind; the caret is asked at key-down time through
  `EM_GETSEL`, and a boundary arrow with nowhere to go is a no-op that keeps
  the editor open
- Column currency walks selectable columns only -- arrows skip a
  `Selectable=False` column, the mouse refuses it, programmatic `Col` on it
  clears to no column at all -- and the arrows keep walking with editing off;
  a keyboard move scrolls a half-shown target column into view, the move
  announced from where the cell was (`RowColChange` ahead of
  `LeftColChange`, pinned in 088)
- Tab heeds `TabKeyBehavior` in the accelerator alone -- probed, the window
  path walks columns for it whatever the property says, which is exactly the
  split the hook architecture wants -- and a cell tabbed into opens its
  editor selected whole; F2 opens on the current cell (selected whole, ours
  on purpose where the original appends); a printable character on the
  surface opens the editor the same way, so what is typed replaces the cell;
  Enter with no editor steps the row like a down arrow, and in the editor
  commits and steps ahead of any Default button
- Del clears the cell where a column is current -- the editor opens selected
  whole and the selection is cut, outranking `AllowDelete` -- and in the
  column-less row mode deletes the selected rows through the now-real
  `Delete` method: one vetoable `BeforeDelete`, a vetoable `BeforeDeleteEX`
  per record answering with the row index when the client never bookmarked
  it, the currency re-landed from nowhere (`RowColChange` carries
  `LastRow=0`) and `AfterDelete` closing out; probed, the unbound record goes
  quietly -- no `UnboundDelete` is raised for it (093)
- Escape is a ladder, every rung claimed ahead of the container: mid-drag it
  is `WM_CANCELMODE` by hand abandoning the resize or the drag whole, over
  an editor it cancels the cell, and the second press drops every column the
  row has buffered
- Group rows by mouse and keyboard: the expand box toggles on the press
  after the marquee lands (084-087, with the caption double-click toggling
  anywhere on the row and its release folding into one `Click`); Left
  collapses an expanded header quietly -- a collapse exposes nothing, so no
  `RowFormat` -- Right expands a collapsed one in place and steps into the
  first child of an expanded one (094-096); Space toggles the header too,
  a deliberate step past the original, which ignores it
- `AllowEdit` matches the original's model, probed: it defaults True, a
  click on an editable cell opens an edit session even where `EditType`
  asks for no window -- `BeforeColEdit` on the click, `AfterColEdit` when
  the currency leaves, nothing to type into between -- and with it off
  there is no column currency at all: clicks land the row whole, the
  horizontal keys go dead and F2 opens nothing (090-092)
- Rows are swept into a drag-selection only by a press that landed on one:
  a button already down when the pointer arrives, or one that went down on
  the chrome or past the rows, selects nothing as it moves
- Ctrl+Tab and Ctrl+Shift+Tab tab out even where plain Tab walks columns:
  the chord goes back to the container's `IOleControlSite.TranslateAccelerator`
  stripped to the plain (or shifted) Tab the host never translates chorded,
  falling back to the parent's site for a control on a nested container
- The strip left of the cells selects the row whole -- the row headers when
  they show, a probed 5 pixel band when they do not -- with the currency
  riding the usual row move before the column clears with a `RowColChange`
  of its own; the pointer over the strip becomes the original's
  right-pointing arrow, ours mirrored from the stock arrow on the fly
  (mask, colour plane and hotspot) rather than shipped as a resource
- Selection painting knows whether the control holds the focus, probed with
  a real `WM_KILLFOCUS` at each `HideSelection` mode: the default trades the
  fill for a thin highlight outline one short of the block rule, inactive
  keeps a button-face fill, normal does not notice -- and the marquee hides
  with the focus. The look is focused until a real `WM_KILLFOCUS` arrives,
  which is how a control that never held focus paints the full selection,
  and focus crossing to the in-place editor does not count as leaving
- Focus leaving the editor for another control tears the box down silently,
  probed: the text stays buffered in the row, the cell repaints under the
  outline, and none of the edit events say a word

### Fixed (what the keyboard work turned up)

- `cUnboundDataModel.GetRowData` indexed the fetch map past its bounds for
  any column added after the bind -- VB6's `And` does not short-circuit --
  and the sweep that followed nested the only two live hazards of that
  shape, both reading `m_hWndEdit` in the same breath that created it
- A control shorter than the scrollbar band asked `picGrid.Move` for a
  negative height: error 380 out of `pvLayoutGrid`, swallowed at the paint
  boundary but leaving every later scrollbar update silently skipped;
  `Clamp` grew proper optional bounds and stands guard there
- `IDataModel.RowIndex` returns signed indexes (negative names a group
  row's projection slot, stable across expand/collapse), which let the
  window carry, the selection and the current-row hold share one identity
  through reprojection -- `GetSignedRowIndex` folded away
- The scenario corpus stands at 100 (089 retired with the F2 decision, its
  golden pinned the original's append), recorded at 96 and 120dpi with
  084-101 queued for goldens at the next 144dpi session; the harness gained
  real focus-in/focus-out input actions for the unfocused looks

### Added (M5 -- what a cell reads as)

- `JSRowData.DisplayValue` is written for every column before `RowFormat`
  sees the row -- probed, the original hands the client a buffer whose plain
  columns already carry their value as text -- so a handler reads what the
  cell will say and is free to overwrite it, and the painter takes whatever
  is left there
- `JSColumn.Format` decides what a cell reads as: the value through the
  VB format string, a value it cannot format passing through as it stands,
  a blank staying blank (103). It is presentation only -- probed with a
  window dump, the editor over a cell painting `1.50` holds `1.5`, because
  what is edited is the value underneath (104)
- `JSColumn.ValueList` replaces what it carries: a cell reads as the
  matching item's text and, probed, as nothing at all when the list has no
  item for the value -- not as the value itself (102). Here the text *is*
  what is edited, so the editor opens on it (105) and a commit writes back
  the value the text stands for; a text the list does not carry is refused
  where it stands, after the update pair has gone out
- `JSColumn.ReplaceValues` gates all of that, as the docs say: with it off
  the list is drop-down entries and nothing more (102 carries a replacing
  column beside a non-replacing one over the same list)

### Fixed

- `JSColumn.ReplaceValues` defaulted False where the original defaults True
  -- every one of the 250 columns in the recorded snapshot corpus carries
  it True, and the round-trip could not see it because a snapshot imports
  the value explicitly
- `JSColumn.NullBehavior` stays storage, now on the record as such: probed,
  it changes neither what is painted nor what an unbound cell stores --
  under `jgexNBAutomatic` a cleared cell leaves an empty string behind
  rather than a Null -- since what it governs is how an empty cell reaches
  a database. The harness event log spells `Null` out now, which is the
  only reason that reads as an answer rather than as a blank

### Added (M5 -- the new row and the wheel)

- `AllowAddNew`/`NewRowPos`: the phantom record row, addressed as `Row = -1`
  above the records under a five-pixel divider (106) or as `RowCount + 1`
  below them (107); it lives outside `RowCount` and no event acknowledges
  it until something is typed. Grouping hides the bottom flavour only
  (108); the top flavour indents with the groups, keeps the marquee
  walking the indent as its own dither segment, and -- its selection lost
  to the regroup -- wears the unfocused one-pixel outline instead of the
  selection fill (109). The divider dresses after `HeaderStyle`, probed
  per style at both DPIs: the double-3D default, flat single line (110)
  and single-3D (111)
- The new row edits: a click or a typed character opens the editor over a
  fresh pending buffer, `RowFormat` seeing the prospective index; the
  commit is an insertion -- `IDataModel.UpdateRowData` past `ItemCount`
  raises `UnboundAddNew`, harvests the buffer into new storage and
  re-initializes the wrapper onto the record it became. Leaving the row
  with data changed commits it, by click or by keyboard; `Escape` drops
  the buffer and re-arms; scrolling keeps the editor glued to the band
- `Enter` on a modified new row, probed: the original commits, touches
  down on the record it just became and comes straight back to a fresh
  new row -- `Row` ends at `-1`, the data-entry loop. Pinned in
  ModelTests; not in the corpus, because the original defers its
  `UnboundAddNew` past the capture window where ours raises it
  synchronously, where a client needs it
- The vertical scrollbar counts the extra row and surrenders the band's
  height, so the range stays honest with the phantom on screen
- The mouse wheel scrolls: three rows a notch vertically, a column a
  notch with `Shift` held; the single-line editor forwards the wheel to
  the grid so scrolling does not stop at the caret. The delta comes out
  of `wParam` through `HiWordInt` -- the masked-division spelling loses
  the sign when the low word is set
- `src/reg.bat`: unregisters the stale typelib key and re-registers the
  built OCX in one step

### Fixed (what the new row turned up)

- `pvPaintDataRow` painted the new row's other columns empty: the window
  buffer has no row at the phantom index, so the painter falls back to
  the pending buffer for the current new row
- The initial bind with `NewRowPos = jgexTop` set `Row = -1` with no
  pending buffer behind it, so the first edit wrote into `Nothing` and
  painted nothing; the bind now arms the buffer the way navigation does
- `pvPaintNewRowDivider` stopped at the last column, leaving the band's
  right remainder undrawn past the block

### Changed (procedures grouped by trade and named by what they do)

- The private block of `GridEX.ctl` reorganized into twelve groups under
  one-line headers -- painting, editing, input, navigation and selection,
  the new row, columns under the mouse, grouping, data model, layout and
  scrolling, hit testing, the record navigator, drawing primitives --
  entry points first, helpers trailing; verified a pure move (the line
  multiset differs by exactly the headers)
- Procedure naming convention adopted and written into `CLAUDE.md`:
  Verb+[Adj+]Noun, with `Get` for cheap values, `Count`/`Calc`/`Build`/
  `Measure` where work happens, `Hit` for point lookups, `Handle` for
  message reactions, `Track`/`End` for drag pairs, `Is`/`Has`/`Needs`/
  `Uses` for predicates
- 104 private procedures renamed in `GridEX.ctl` to match (787 call
  sites): the `pvOn*` handlers became `pvHandle*`, the noun-only getters
  gained their verbs, the `pvEdit*`/`pvBuffer*` noun-first families
  flipped, the point lookups joined `pvHit*` -- and the long-standing
  `pvBufferTermiante` typo died as `pvFreeBuffer`
- 26 more across the models and `JS*` classes (`pvHarvest` ->
  `pvHarvestRecord`, `pvReindex` -> `pvReindexItems`, `pvNotify` ->
  `pvNotifyOwner`, `pvDefaultFont` -> `pvCreateDefaultFont` -- it
  allocates), and nine `Data*` publics in `mdDataModel.bas`
  (`DataProject` -> `DataProjectRows`, `DataAggregate` ->
  `DataAggregateValues`, the noun-only getters -> `DataGet*`);
  `mdIPAO.bas` and `mdModernSubclassing.bas` keep their upstream names
- `ARCHITECTURE.md`/`PROPERTIES.md` references updated, including the
  stale `pvBufferOpen`/`pvBufferClose` aliases and a `pvAggregate` that
  had long since become the public `DataAggregateValues`
