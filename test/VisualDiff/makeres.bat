@echo off
setlocal
rem --- regenerates VisualDiff.res (DPI-aware manifest) from VisualDiff.rc
rem --- the built .res is committed so building the harness needs no SDK
pushd "%~dp0"

set "RC="
for %%D in (
    "%ProgramFiles(x86)%\Windows Kits\10\bin\10.0.22621.0\x86\rc.exe"
    "%ProgramFiles(x86)%\Microsoft SDKs\Windows\v7.1A\Bin\RC.Exe"
) do if exist %%D set "RC=%%~D"
if not defined RC (
    echo FAILED: rc.exe not found, keeping the committed VisualDiff.res
    exit /b 1
)

"%RC%" /nologo /fo VisualDiff.res VisualDiff.rc
if errorlevel 1 (
    echo FAILED: resource compile
    exit /b 1
)
echo Successfully built VisualDiff.res
