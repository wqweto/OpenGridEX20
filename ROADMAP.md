# Roadmap

Agreed scope: **source-compatible** implementation -- code written against the
original control compiles and behaves the same after re-referencing
`OpenGridEX20` (same member names, signatures, enum values, event contracts,
default members). Binary compatibility with the original (its GUIDs/ProgIDs,
its `.frx` propbag format, drop-in OCX replacement) is a non-goal. ADO and
Unbound data modes only (no DAO binding). Own GUIDs/ProgIDs also mean the
original and this control coexist on one machine -- required for twin testing.
Samples get ported to `OpenGridEX20` references.

## M0 -- Baseline build (infrastructure)

- Compile `typelib\OpenGEXHelper.tlb` and the stub OCX; fix whatever the stubs broke
- Regenerate `OpenGridEX20.cmp` from the freshly built stub OCX so the complete API
  surface becomes the binary-compatibility baseline; keep `CompatibleMode=2` and
  only append members from here on
- Typelib surface check: OleView dump of built OCX diffed against `doc\GridEX20.idl`
  (names/signatures; dispids exempt)

Exit: `src\make.bat` builds and registers clean, surface diff shows no drift.

## M1 -- Tooling: snapshot engine + JSON schema

- Shared snapshot engine as plain modules (`CallByName` property walker, TLI
  reflection check, JSON I/O, Font/Picture serializers, raw `.frm` key capture)
  + GridEX/GEXPreview profile generated from the stubs; consumed by three thin
  shells: the twin test harness, a JSON -> `Form_Load` code generator (sample
  porting, aligned with source compatibility), and a one-command export add-in
- Freeze the JSON snapshot schema (also used later by tests)
- `tools\OpenGEXAddin`: export-only VB6 add-in wrapping the engine; batch-export
  all sample grids to `test\snapshots\` as `NNN-<sample>_<form>_<control>.json`,
  where `NNN` is the sample's position in `export.ps1`'s ordinal walk, so a
  re-export reproduces the names (needs original control installed; not on the
  critical path -- may slide to just before sample-based exits)

Exit: engine + schema frozen; sample form grid state captured as JSON incl.
`raw` section; unknown propbag keys (`IntProp*`, `MethodHoldFields`, ...)
inventoried.

## M2 -- Object model

- Implement state storage in all `JS*` classes and control-level properties;
  `frInit`-style owner wiring, one internal setter path per property
- Add-in import command (JSON -> our object model via `CallByName`)
- `test\ModelTests.vbp` runner: collection semantics (Add/Item/Remove/Clear/
  NewEnum/keys), `JSRet*` defaults, 1-based `JSRowData` values, import->export
  round-trip diff over all sample snapshots

Exit: sample JSON round-trips losslessly through our object model.

## M3 -- Unbound engine + table view (split into a-d)

### M3a -- Unbound data pipeline (no pixels)

- `DataMode`/`ItemCount` wiring and the `UnboundReadData`/`UnboundAddNew`/
  `UnboundDelete`/`UnboundUpdate` events feeding the control row cache
  (`m_aRows` + `JSRowData` wrappers from M2); cache invalidation via
  `Rebind`/`Refetch`/`Refresh`
- Row/col navigation state (`Row`, `Col`, `FirstItem`, `RowCount`, bookmarks)
  with correct event ordering (`RowColChange`, `FirstItemChange`,
  `SelectionChange`)

Exit: scripted unbound scenarios in `test\ModelTests` populate and read back
the cache with event sequences matching the original control.

### M3b -- VisualDiff harness + golden recorder (tooling only)

- `test\VisualDiff.vbp` harness built against the *original* control first:
  fixed-size twin form, runtime-scripted scenarios (JSON), BitBlt capture to
  DIB, pixel diff with cluster report, sentinel A/B footprint tests,
  simultaneous-sentinel paint map, record vs verify modes, before-show/
  after-show application ordering modes; MS Sans Serif for AA-free runs
- Record initial golden corpus from the original control

Exit: goldens recorded for the initial scenario set; harness self-validates
(original vs original captures diff to zero).

### M3c -- Static table painting

- Window plumbing, double-buffered painting: background, borders, column
  headers, gridlines and line styles, cell text from the row cache, even/odd
  colors, static selection rendering -- developed golden-first, scenario by
  scenario against the M3b corpus

Exit: static paint scenarios pass pixel diff against the golden corpus.

### M3d -- Scrolling + input (done)

- Scrollbars, `ContinuousScroll`, keyboard/mouse navigation, live selection,
  core events (`Click`, `DblClick`, `RowColChange`, `SelectionChange`, key
  events), `LeftCol`/`FirstItem` tracking
- Closed with `Col`, `ColumnAutoResize`, `ContinuousScroll`, `FrozenColumns` and
  `Redraw`; scenarios `031`-`034` cover the three that paint, `ModelTests` the
  two that only change behaviour

Exit (met): all five unbound samples are ported under `test\Samples` and run in
the smoke runner -- Unbound 1 and 2 read Products through DAO as the originals
do, the other three build their data in code -- and the golden corpus passes at
96 and 120 dpi. (Table view only -- card view intentionally out of M3 scope, and
sorting, grouping and editing stay with M4/M5, so the samples' header-click and
edit paths are inert for now.)

## M4 -- Sorting and grouping (done)

- `SortKeys`/`SortType` comparers, header-click auto sort, `Groups`, group rows,
  expand/collapse, group-by box, aggregates and group footers

Exit (met): scenarios `035`-`056` cover sorting, grouping, expand/collapse, the
group caption properties and both footer styles with their aggregates, all
passing at 96 and 120 dpi; `PROPERTIES.md` has no M4 property left unimplemented
or merely stored. The Advanced Sample's M4 parts are covered as the API sequence
its `frmSort`, `frmGroupBy` and `frmSummary` dialogs run (`ModelTests`), the
sample itself waiting on ADO binding, card view and printing.

Known divergence recorded from the original: `JSColumn.GroupFormat` labels a
group caption but does not group by it, despite the help reading otherwise.

## M5 -- Editing

- `EditType` editors (TextBox, CheckBox, DropDown, Calendar, Combo), custom edit
  event trio, `ValueList` replacement, `JSRetBoolean` cancel semantics,
  edit-related events ordering recorded from original

Exit: Custom Edit, Combo DropDowns, ADO3-style value list behavior (unbound data).

## M6 -- Card view

The one part of the public surface no milestone owned: M3 declared card view out
of scope and nothing later picked it up, yet `View`, `CardBorders`,
`CardCaptionPrefix`, `CardWidth`, `CardSpacing` and the `JSColumn` card members
are all public and appear in the sample corpus. Twelve paint properties in total.

- `View` switching between table and card layout, card painting: borders
  (`CardBorders`), caption prefix (`CardCaptionPrefix`), and the fixed
  `CardWidth`/`CardSpacing` pixel metrics confirmed at M3 to be DPI-independent
- per-column card behaviour: `JSColumn.CardCaption`, `CardIcon`,
  `ShowCaptionInCardView`, `MinRowsInCardView`, `MaxRowsInCardView`
- `AutomaticArrange` and `AllowCardSizing`
- developed golden-first exactly like the table view, recorded from the original
  and verified at 96/120/144 dpi

Comes before the paint property matrix on purpose: card view *adds* twelve paint
properties, so gating on "nothing unverified" only means something once they
exist.

Exit: card scenarios pass pixel diff at all recorded scales; no `unowned` card
entries left in `PROPERTIES.md`.

## M7 -- Paint property matrix

Gate before any data-binding work: every property that affects painting is proven
against the original at two or more distinct values, so that binding is built on
a verified renderer rather than on properties that merely compile.

- `PROPERTIES.md` is the tracking document -- type, status, paint routine,
  owning milestone, implementing commit and covering test for each of the 126
  paint-affecting properties (`GridEX` plus the `JSColumn`/`JSFormatStyle`
  sub-properties); kept current in the same commit that changes a status. It
  inventories the whole public property surface around that matrix -- all 288
  properties read out of `doc\GridEX20.idl`, class by class -- so a member that
  no milestone owns cannot hide by simply not being written down
- Baseline at the time of writing: 43 of 126 are read by the paint path, only 19
  are proven at more than one value, 15 have no covering test at all
- Close the unverified block by authoring scenarios, not control code (rows for
  `gridlines-dots-colors`, a two-palette `colors-*` pair, `formatstyles-selection`,
  `column-order`); fix whatever disagrees with the original
- Give the weak entries a second rendered value, incl. the three non-default
  `GridLines` modes
- Card view is owned by M6 now. Five properties are still unowned: the
  drag/resize affordances (`AllowColumnDrag`, `AllowRowSizing`, `DetectRowDrag`,
  `JSColumn.AllowSizing`) and `Options`
- Re-scope M10: the matrix puts 38 unimplemented paint properties there, mostly
  the `JSFormatStyle` font and picture families, which is far more than its
  one-line entry implies

Exit: no `unverified` rows left in `PROPERTIES.md` for properties whose
milestone has shipped; card view has an owning milestone.

## M8 -- ADO binding

- Late-bound ADO (`As Object`, no msado15 reference), `ADORecordset` get/set,
  `DatabaseName`/`RecordSource`/`CursorLocation`/`LockType`, bookmarks,
  AddNew/Delete/Update, `Rebind`/`Refetch`; `DataMode = jgexDAO` and DAO-only
  members raise trappable "not supported" errors but still store values

Exit: ADO1-5 samples and NorthWind run against `JSNwind.mdb`.

## M9 -- Persistence and property pages

- Our `WriteProperties`/`ReadProperties` routed through the same friend setters;
  add-in import + IDE save converts sample forms to our `.frm`/`.frx`
- Equivalence test: persisted-form snapshot == runtime-scripted snapshot
- Property pages (`.pag`) for the design-time dialog the samples reference

Exit: all ported sample forms load in the IDE with correct design-time state.

## M10 -- Styling extras

- Full `JSFormatStyle` application, `JSFmtConditions` conditional formatting,
  preview rows, tooltips, scroll tips

Exit: Using FormatStyles, ADO5 conditional formatting, Preview Rows samples.

## M11 -- Printing

- `JSPrinterProperties`, pagination, `PrintGrid`, `GEXPreview` rendering/toolbar/
  zoom, `PrintPreview` handshake between controls

Exit: the three Print & PrintPreview samples.

## M12 -- Long tail (optional, in rough order of value)

- OLE drag & drop (`JSDataObject`), `SaveLayout`/`LoadLayout`/`LayoutString`
- RDS / DataEnvironment / HTML samples: likely skip (obsolete)

## Cross-cutting rules

- Keep `CHANGELOG.md` current per milestone
- Binary compatibility locked from M2 onward; never reorder/remove public members
- Every batch tool runs from a `make.bat`-style script with a meaningful exit code
- Original OCX needed only on the machine that records goldens/exports snapshots,
  never committed to the repo
