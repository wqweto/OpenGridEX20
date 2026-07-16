# Open GridEX 2000 Control

Open reimplementation of Janus GridEX 2000b Control (DAO 3.6 & ADO 2.x)

## Status

Work in progress. The complete public API of the original `GridEX20.ocx` -- both controls, 24 dependent classes, 44 enums, 58 events and all methods/properties incl. help strings -- is stubbed out from the original type library. No functionality is implemented yet.

## Repository Layout

| Folder | Contents |
| ------ | -------- |
| `src` | VB6 ActiveX control project (`OpenGridEX20.vbp`) with `GridEX` and `GEXPreview` user-controls and the `JSColumn(s)`, `JSFormatStyle(s)`, `JSRowData`, etc. dependent classes |
| `doc` | Reference IDL dumps of the original control's type libraries (`GridEX20.idl`, `JSIOSafe.idl`) |
| `typelib` | `OpenGEXHelper` helper type library with the `IObjectSafety` interface used by the controls |

## Building

Compile the helper type library first (requires `mktyplib.exe`), then the control (requires VB6):

    cd typelib
    make.bat
    cd ..\src
    make.bat

`src\make.bat` also normalizes line endings back to CRLF in case git checkout mangled them, so the sources load correctly in the VB6 IDE.

## Notes

- Both `GridEX` and `GEXPreview` implement `IObjectSafety` via the helper type library and report safe for scripting/initialization, matching the original control.
- Collection classes expose hidden `NewEnum` (DISPID -4) for `For Each` support and default `Item` members (DISPID 0), as in the original type library.

## License

[MIT](LICENSE)
