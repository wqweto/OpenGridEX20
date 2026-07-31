# Open GridEX 2000 Control

Open reimplementation of Janus GridEX 2000b Control (DAO 3.6 & ADO 2.x)

## Status

Work in progress -- see [ROADMAP.md](ROADMAP.md) for milestones and [CHANGELOG.md](CHANGELOG.md) for details. Completed so far:

- **M0 -- Baseline build:** the complete public API of the original `GridEX20.ocx` -- both controls, 24 dependent classes, 44 enums, 58 events and all methods/properties incl. help strings -- is stubbed from the original type library, builds clean and passes the typelib surface diff against the original with zero differences.
- **M1 -- Snapshot tooling:** a profile-driven object model walker exports/imports control state as JSON; an export-only VB6 add-in recorded a design-time snapshot corpus (28 JSONs from 20 original Janus sample projects) using the original control.
- **M2 -- Object model:** state storage and collection semantics implemented in all 23 `JS*` classes and both controls; the entire snapshot corpus round-trips losslessly through the object model.
- **M3 -- Unbound engine + table view:** the unbound fetch pipeline (`UnboundReadData` and friends), static painting, scrolling, keyboard/mouse navigation and selection. Rendering is verified pixel-for-pixel against the original: a 30-scenario golden corpus recorded from the original control passes at 96, 120 and 144 dpi. Horizontal scrolling uses a child `HScrollBar` in a band below the grid, and the record navigator is painted there beside it.
- **M7 -- Paint property matrix:** every property that affects painting is catalogued in [PAINT-PROPERTIES.md](PAINT-PROPERTIES.md) with its status, owning milestone and covering test. 44 of 126 are proven against the original at two or more distinct values; the rest are unimplemented and each owned by a later milestone.

Still open: card view (M6), sorting/grouping (M4), editing (M5), ADO binding (M8) and everything after.

The implementation is *source-compatible*: code written against the original control compiles and behaves the same after re-referencing `OpenGridEX20`. The control has its own GUIDs/ProgIDs, so it coexists with the original on one machine (required for twin testing). ADO and Unbound data modes only -- DAO binding is out of scope.

## Repository Layout

| Folder | Contents |
| ------ | -------- |
| `src` | VB6 ActiveX control project (`OpenGridEX20.vbp`) with `GridEX` and `GEXPreview` user-controls and the `JSColumn(s)`, `JSFormatStyle(s)`, `JSRowData`, etc. dependent classes |
| `doc` | Reference IDL dumps of the original control's type libraries (`GridEX20.idl`, `JSIOSafe.idl`), the OleView dump of the built OCX (`OpenGridEX20.idl`) and `Help` -- the original `gridex2000.chm` converted to markdown |
| `typelib` | `OpenGEXHelper` helper type library with the `IObjectSafety` interface used by the controls |
| `tools` | `CompareIdl.ps1` source-compatibility gate, `GenProfiles.ps1` profile generator, `common` shared snapshot engine modules and `OpenGEXAddin` export add-in for recording snapshots from the original control |
| `test` | `ModelTests` object model and snapshot round-trip runner, `VisualDiff` pixel-diff harness with its scenario/golden corpus, `Snapshot` smoke test and `snapshots` design-time corpus recorded from the original Janus samples |

## Building

Compile the helper type library first (requires `mktyplib.exe`), then the control (requires VB6):

    cd typelib
    make.bat
    cd ..\src
    make.bat

`src\make.bat` also normalizes line endings back to CRLF in case git checkout mangled them, so the sources load correctly in the VB6 IDE.

## Testing

    cd test\ModelTests
    make.bat

Builds and runs the test executable: collection semantics, `JSRowData` virtual storage/weak-reference behavior and lossless import->export round-trips of the recorded snapshot corpus against the registered `OpenGridEX20.ocx`.

Rendering is checked separately against golden images recorded from the original control:

    cd test\VisualDiff
    make.bat              REM build and verify every scenario
    make.bat 030*         REM one scenario, about a second
    record.bat            REM re-record the goldens from the original

Both scripts take an optional scenario mask -- the harness matches with VB `Like`, so use `030*` or `*navigator`, not a bare name. Every run covers two scales from one executable: `__COMPAT_LAYER=DPIUNAWARE` forces the virtualized 96dpi path and the embedded manifest runs at the real system DPI, so goldens live in `golden\<dpi>\`. Recording needs the original OCX registered; verifying does not.

Scenarios are `NNN-name.json` in creation order, and each golden PNG is named to match.

## Notes

- Both `GridEX` and `GEXPreview` implement `IObjectSafety` via the helper type library and report safe for scripting/initialization, matching the original control.
- Collection classes expose hidden `NewEnum` (DISPID -4) for `For Each` support and default `Item` members (DISPID 0), as in the original type library.

## License

[MIT](LICENSE)
