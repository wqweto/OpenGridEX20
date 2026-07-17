@echo off
setlocal
set "VB6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
pushd "%~dp0"

"%VB6%" /make Snapshot.vbp /out errors.log
if errorlevel 1 (
    type errors.log
    exit /b 1
)
if exist GridEX.json del GridEX.json
if exist GEXPreview.json del GEXPreview.json
start /wait Snapshot.exe
if not exist GridEX.json (
    echo FAILED: GridEX.json not produced
    exit /b 1
)
if not exist GEXPreview.json (
    echo FAILED: GEXPreview.json not produced
    exit /b 1
)
echo Snapshot smoke test PASSED
