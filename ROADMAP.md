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
  all sample grids to `test\snapshots\` (needs original control installed;
  not on the critical path -- may slide to just before sample-based exits)

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

### M3d -- Scrolling + input

- Scrollbars, `ContinuousScroll`, keyboard/mouse navigation, live selection,
  core events (`Click`, `DblClick`, `RowColChange`, `SelectionChange`, key
  events), `LeftCol`/`FirstItem` tracking

Exit: Unbound 1/2/Array/UDTs/Collection samples run; paint-map scenario passes.
(Table view only -- card view intentionally out of M3 scope.)

## M4 -- Sorting and grouping

- `SortKeys`/`SortType` comparers, header-click auto sort, `Groups`, group rows,
  expand/collapse, group-by box, aggregates and group footers

Exit: sorting/grouping visual scenarios + relevant Advanced Sample parts.

## M5 -- Editing

- `EditType` editors (TextBox, CheckBox, DropDown, Calendar, Combo), custom edit
  event trio, `ValueList` replacement, `JSRetBoolean` cancel semantics,
  edit-related events ordering recorded from original

Exit: Custom Edit, Combo DropDowns, ADO3-style value list behavior (unbound data).

## M6 -- ADO binding

- Late-bound ADO (`As Object`, no msado15 reference), `ADORecordset` get/set,
  `DatabaseName`/`RecordSource`/`CursorLocation`/`LockType`, bookmarks,
  AddNew/Delete/Update, `Rebind`/`Refetch`; `DataMode = jgexDAO` and DAO-only
  members raise trappable "not supported" errors but still store values

Exit: ADO1-5 samples and NorthWind run against `JSNwind.mdb`.

## M7 -- Persistence and property pages

- Our `WriteProperties`/`ReadProperties` routed through the same friend setters;
  add-in import + IDE save converts sample forms to our `.frm`/`.frx`
- Equivalence test: persisted-form snapshot == runtime-scripted snapshot
- Property pages (`.pag`) for the design-time dialog the samples reference

Exit: all ported sample forms load in the IDE with correct design-time state.

## M8 -- Styling extras

- Full `JSFormatStyle` application, `JSFmtConditions` conditional formatting,
  preview rows, tooltips, scroll tips

Exit: Using FormatStyles, ADO5 conditional formatting, Preview Rows samples.

## M9 -- Printing

- `JSPrinterProperties`, pagination, `PrintGrid`, `GEXPreview` rendering/toolbar/
  zoom, `PrintPreview` handshake between controls

Exit: the three Print & PrintPreview samples.

## M10 -- Long tail (optional, in rough order of value)

- OLE drag & drop (`JSDataObject`), `SaveLayout`/`LoadLayout`/`LayoutString`
- RDS / DataEnvironment / HTML samples: likely skip (obsolete)

## Cross-cutting rules

- Keep `CHANGELOG.md` current per milestone
- Binary compatibility locked from M2 onward; never reorder/remove public members
- Every batch tool runs from a `make.bat`-style script with a meaningful exit code
- Original OCX needed only on the machine that records goldens/exports snapshots,
  never committed to the repo
