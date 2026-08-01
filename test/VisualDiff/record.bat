@echo off
setlocal
rem --- records the golden corpus from the original control at BOTH DPIs and
rem --- self-tests the harness (a fresh original capture must diff to zero).
rem --- Goldens land in golden\<dpi>\, so recording on a 125% machine fills
rem --- golden\96 (forced DPI-unaware) and golden\120 (manifest DPI aware).
pushd "%~dp0"
rem --- optional scenario mask, e.g. record.bat hscrolled
set "MASK=%~1"
if "%MASK%"=="" set "MASK=*"

if not exist VisualDiff.exe (
    echo FAILED: build first with make.bat
    exit /b 1
)
if not exist golden mkdir golden

call :record "DPIUNAWARE" "96dpi"
if errorlevel 1 exit /b 1
call :record "" "system dpi"
if errorlevel 1 exit /b 1
echo Golden corpus recorded and self-test PASSED at both DPIs ^(%MASK%^)
exit /b 0

:record
set "LAYER=%~1"
echo === recording %~2 ===
if exist VisualDiff.out.txt del VisualDiff.out.txt
rem --- watchdog: a wedged run must not hang the loop
powershell -NoProfile -Command "$env:__COMPAT_LAYER = '%LAYER%'; $p = Start-Process '.\VisualDiff.exe' -ArgumentList 'record','%MASK%' -PassThru; if (-not $p.WaitForExit(300000)) { $p.Kill(); Write-Host 'TIMEOUT: killed after 300s'; exit 1 }"
findstr /C:"RESULT: PASSED" VisualDiff.out.txt >nul || (
    type VisualDiff.out.txt
    exit /b 1
)
type VisualDiff.out.txt
if exist VisualDiff.out.txt del VisualDiff.out.txt
rem --- watchdog: a wedged run must not hang the loop
powershell -NoProfile -Command "$env:__COMPAT_LAYER = '%LAYER%'; $p = Start-Process '.\VisualDiff.exe' -ArgumentList 'selftest','%MASK%' -PassThru; if (-not $p.WaitForExit(300000)) { $p.Kill(); Write-Host 'TIMEOUT: killed after 300s'; exit 1 }"
findstr /C:"RESULT: PASSED" VisualDiff.out.txt >nul || (
    type VisualDiff.out.txt
    exit /b 1
)
type VisualDiff.out.txt
exit /b 0
