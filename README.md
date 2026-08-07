# Open GridEX 2000 Control

Open reimplementation of Janus GridEX 2000b Control (DAO 3.6 & ADO 2.x)

## Status

Work in progress -- see [ROADMAP.md](ROADMAP.md) for milestones and [CHANGELOG.md](CHANGELOG.md) for details. Completed so far:

- **M0 -- Baseline build:** the complete public API of the original `GridEX20.ocx` -- both controls, 24 dependent classes, 44 enums, 58 events and all methods/properties incl. help strings -- is stubbed from the original type library, builds clean and passes the typelib surface diff against the original with zero differences.
- **M1 -- Snapshot tooling:** a profile-driven object model walker exports/imports control state as JSON; an export-only VB6 add-in recorded a design-time snapshot corpus (28 `NNN-<sample>_<form>_<control>.json` files from 20 original Janus sample projects) using the original control.
- **M2 -- Object model:** state storage and collection semantics implemented in all 23 `JS*` classes and both controls; the entire snapshot corpus round-trips losslessly through the object model.
- **M3 -- Unbound engine + table view:** the unbound fetch pipeline (`UnboundReadData` and friends), static painting, scrolling, keyboard/mouse navigation, selection, frozen columns and column auto-resize. Horizontal scrolling uses a child `HScrollBar` in a band below the grid with the record navigator painted beside it, and all five original unbound samples are ported and run under `test\Samples`.
- **M4 -- Sorting and grouping:** automatic sort with a stable merge sort that remaps `Row` and the selection through the new order, the group-by box and its chip staircase, group captions and expand/collapse, aggregates and group footers.
- **M5 -- Editing (in progress):** the in-place text editor, the checkbox editor and wrapping cells, plus an event log recorder so a scenario asserts the event sequence it fires alongside its pixels. Column headers carry their three mouse gestures -- resize by the divider, auto-size on a double-click, and a drag that repositions the column, scrolling the strip when it leaves the view -- each told from the others on the button up. The group-by box takes drops too: a header dragged into it groups, chips reorder along the box or leave it -- deleting their level and repositioning their column -- with a red mark standing where the drop would land and `BeforeColumnDrag`/`BeforeGroupDrag` gating the gestures. Dropdown, combo and calendar editors are still open.
- **M7 -- Property inventory:** [PROPERTIES.md](PROPERTIES.md) catalogues all 288 public properties of the control and its object model, read out of the original's type library rather than out of the implementation. Two matrices share the same columns: the 126 that paint, of which 65 are proven against the original at two or more distinct values and 2 more over the part of their range that exists, and the 162 that do not, of which 71 are consumed by the engine and 85 are storage waiting on a named milestone. A third part lists every class -- `JSColumn`'s 46 members, `JSFmtCondition`'s 7 -- with the original's own descriptions.

Rendering is verified pixel-for-pixel against the original: an **83-scenario golden corpus** recorded from the original control passes at **96, 120 and 144 dpi**. Three scales is the working minimum -- two cannot separate a constant from a metric-derived term, and several rules that looked settled on a pair turned out to be wrong on the third.

The control paints itself band by band into an off-screen bitmap and answers `WM_PAINT` in its own subclass, so a repaint updates only the bands the update region reaches and never shows an intermediate state.

Still open: card view (M6), ADO binding (M8) and everything after.

The implementation is *source-compatible*: code written against the original control compiles and behaves the same after re-referencing `OpenGridEX20`. The control has its own GUIDs/ProgIDs, so it coexists with the original on one machine (required for twin testing). ADO and Unbound data modes only -- DAO binding is out of scope.

## Repository Layout

| Folder | Contents |
| ------ | -------- |
| `src` | VB6 ActiveX control project (`OpenGridEX20.vbp`) with `GridEX` and `GEXPreview` user-controls and the `JSColumn(s)`, `JSFormatStyle(s)`, `JSRowData`, etc. dependent classes |
| `doc` | Reference IDL dumps of the original control's type libraries (`GridEX20.idl`, `JSIOSafe.idl`), the OleView dump of the built OCX (`OpenGridEX20.idl`), `Help` -- the original `gridex2000.chm` converted to markdown -- and `Samples`, the original Janus sample projects with their `JSNWind.mdb` |
| `typelib` | `OpenGEXHelper` helper type library with the `IObjectSafety` interface used by the controls |
| `tools` | `CompareIdl.ps1` source-compatibility gate, `GenProfiles.ps1` profile generator, `GenSample.ps1` snapshot-to-VB6 generator used to port samples, `common` shared snapshot engine modules and `OpenGEXAddin` export add-in for recording snapshots from the original control |
| `test` | `ModelTests` object model and snapshot round-trip runner, `VisualDiff` pixel-diff harness with its scenario/golden corpus, `Samples` the ported Janus unbound samples with a smoke runner, `Snapshot` smoke test and `snapshots` design-time corpus recorded from the original Janus samples |

## Documents

| File | Contents |
| ---- | -------- |
| [ARCHITECTURE.md](ARCHITECTURE.md) | How the pieces fit together and why they are shaped that way -- the data model seam and its three address spaces, and the painting path |
| [ROADMAP.md](ROADMAP.md) | Milestones M0-M12 with their exit criteria, and the scope decisions behind them |
| [PROPERTIES.md](PROPERTIES.md) | Every public property of the control and its object model, with implementation status, owning milestone and covering test |
| [CHANGELOG.md](CHANGELOG.md) | What landed in each milestone, including the behaviours of the original that had to be reverse-engineered |

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

Both scripts take an optional scenario mask -- the harness matches with VB `Like`, so use `030*` or `*navigator`, not a bare name. All three suites launch their executable on a desktop of its own through the shared `test\run.ps1` (`CreateDesktopW` + `STARTUPINFO.lpDesktop`), so a run does not flicker windows across whatever else the machine is doing; the windows are still real and still paint, so the captures are unchanged. Pass `same` -- `make.bat 038* same` for the visual suite, `make.bat same` for the other two -- to keep the run on the interactive desktop and watch it. Every run covers two scales from one executable: `__COMPAT_LAYER=DPIUNAWARE` forces the virtualized 96dpi path and the embedded manifest runs at the real system DPI, so goldens live in `golden\<dpi>\`. The third scale comes from changing the system DPI and re-running; all three carry the whole corpus and pass at each. A gesture written in pixels does not always land at every scale -- a divider that sits under the press at 96dpi is 50 pixels further right at 144 -- so a scenario whose input misses its target at one scale carries no event log there and is asked about its pixels alone. Recording needs the original OCX registered; verifying does not.

Scenarios are `NNN-name.json` in creation order, and each golden PNG is named to match. The harness also answers questions about either control directly, which beats inferring geometry from pixels:

    VisualDiff.exe windows 030*      REM window tree, classes, rects and styles
    VisualDiff.exe windows-ours 061* REM the same tree for this control, to diff against it
    VisualDiff.exe scrollinfo 031*   REM scrollbar range and thumb rect of the original
    VisualDiff.exe metrics           REM font metrics at the current DPI

`windows` reports each window's text, selection, line count and first visible line, so the original's in-place editor can be read directly instead of inferred from a picture of it -- which is how the editor's width and its click hit-test were settled.

The five original unbound samples are ported and run as a third suite:

    cd test\Samples
    make.bat

Each sample was ported by re-pointing the reference at `OpenGridEX20`, dropping the grid's design-time property bag and calling the setup `Sub` generated from its recorded snapshot by `tools\GenSample.ps1` -- which is what source compatibility is supposed to buy. The runner loads every form, lets it paint and asserts the grid took the sample's data. Unbound 1 and 2 read Products from `doc\Samples\JSNWind.mdb` through DAO, exactly as the originals do.

Everything a run writes -- results, dumps, error log, compiler log -- goes to `output\` next to the runner.

## Notes

- Both `GridEX` and `GEXPreview` implement `IObjectSafety` via the helper type library and report safe for scripting/initialization, matching the original control.
- Collection classes expose hidden `NewEnum` (DISPID -4) for `For Each` support and default `Item` members (DISPID 0), as in the original type library.

## License

[MIT](LICENSE)
