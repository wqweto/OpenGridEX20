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

### Fixed

- `mdJson.bas`: `JsonDump` now emits `{}` for `Nothing` objects (respecting `CompoundChars`) instead of an empty string that produced malformed JSON when a `Nothing` value was stored inside a document

### Changed

- First clean VB6 build of the stub OCX; regenerated `OpenGridEX20.cmp` binary-compatibility baseline from the complete stubbed API surface (26 coclasses, 44 enums, 58 + 6 events, dispids and enum values verified against `doc/GridEX20.idl`)
