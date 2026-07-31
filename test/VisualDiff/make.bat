@echo off
setlocal
set "VB6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
rem --- optional scenario mask, e.g. make.bat gridlines-* -- recording and
rem --- verifying the whole corpus is the slow part of the loop
set "MASK=%~1"
if "%MASK%"=="" set "MASK=*"
pushd "%~dp0"

powershell -NoProfile -Command "Get-ChildItem *.vbp,*.frm,*.bas | ForEach-Object { $t = [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::Default); $f = $t -replace '(?<!\r)\n', ([string][char]13 + [char]10); if ($f -ne $t) { [IO.File]::WriteAllText($_.FullName, $f, [Text.Encoding]::Default); Write-Host ('Fixed ' + $_.Name) } }"

"%VB6%" /make VisualDiff.vbp /out errors.log
if errorlevel 1 (
    type errors.log
    exit /b 1
)

rem --- the embedded manifest makes the harness DPI aware, so it renders at
rem --- the real system DPI; __COMPAT_LAYER=DPIUNAWARE forces the 96dpi
rem --- (OS-virtualized) mode out of the same executable. Both are verified.
call :verify "DPIUNAWARE" "96dpi"
if errorlevel 1 exit /b 1
call :verify "" "system dpi"
if errorlevel 1 exit /b 1
echo VisualDiff verify PASSED at both DPIs ^(%MASK%^)
exit /b 0

:verify
set "__COMPAT_LAYER=%~1"
if exist VisualDiff.out.txt del VisualDiff.out.txt
start /wait VisualDiff.exe verify %MASK%
if not exist VisualDiff.out.txt (
    echo FAILED: VisualDiff.out.txt not produced ^(%~2^)
    exit /b 1
)
findstr /C:"RESULT: PASSED" VisualDiff.out.txt >nul || (
    echo --- FAILED at %~2 ---
    type VisualDiff.out.txt
    exit /b 1
)
exit /b 0
