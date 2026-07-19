# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added

- Stubbed complete public API from the original `GridEX20.ocx` type library (`doc/GridEX20.idl`): `GridEX` and `GEXPreview` controls, 24 dependent classes, 44 enums, 58 events and all methods/properties incl. help strings
- `OpenGEXHelper` helper type library (`typelib/OpenGEXHelper.odl`, compiled with `mktyplib`) exposing the `IObjectSafety` interface
- `IObjectSafety` implementation on both controls reporting safe for scripting/initialization
- `src/make.bat` build script which restores CRLF line endings and compiles `OpenGridEX20.vbp` with VB6
- Project README with repository layout and build instructions
- `ROADMAP.md` with implementation milestones M0-M10 and testing strategy
- `tools/DumpSurface.vbs` typelib surface dumper (TLI based) for source-compatibility checks against `doc/GridEX20.idl`
- `tools/CompareIdl.ps1` source-compatibility gate: canonical diff of `doc/OpenGridEX20.idl` (OleView dump of the built OCX) against `doc/GridEX20.idl`; currently passes with zero differences
- Snapshot engine (M1): `tools/GenProfiles.ps1` generates `tools/common/mdProfiles.bas` (267 property profiles parsed from the stubs incl. runtime-only/collection classification), `tools/common/mdSnapshot.bas` profile-driven late-bound object model walker building `mdJson.bas` documents, JSON schema frozen in `tools/common/SCHEMA.md` (all JSON goes through `mdJson.bas` using its canonical idioms; `$schema`/`$errors` meta keys addressed via JSONPath dot form `$.$schema`)
- `test/Snapshot` smoke test hosting both controls and dumping their object model to JSON; builds and passes against the stub OCX

### Added (Documentation)

- `doc/Help`: original `gridex2000.chm` converted to markdown (32 files: topics combined per TOC section -- Properties/Methods/Events per control, one file per `JS*` object/collection, all code samples in `Examples.md`, object model diagrams in `images/`); `index.md` reproduces the CHM table of contents and all 2900+ cross-topic links (See Also, Applies To, Example, image maps) resolve to `file.md#anchor` form

### Fixed

- `doc/Help`: repaired 37 syntax lines across 16 files where CHM conversion collapsed bold/italic markers into `***`/`****` runs that rendered as literal asterisks (e.g. `**GetData** **(***format* **As Integer)**`); parentheses normalized to the plain-paren convention used by the rest of the docs, plus the fused `RowIndex` syntax line in `Methods.md` split and its parameter italicized
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

### Changed

- First clean VB6 build of the stub OCX; regenerated `OpenGridEX20.cmp` binary-compatibility baseline from the complete stubbed API surface (26 coclasses, 44 enums, 58 + 6 events, dispids and enum values verified against `doc/GridEX20.idl`)
