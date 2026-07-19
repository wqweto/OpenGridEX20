@echo off
setlocal
rem --- records the golden corpus from the original control, then
rem --- self-tests the harness (fresh original capture vs golden = zero diff)
pushd "%~dp0"

if not exist VisualDiff.exe (
    echo FAILED: build first with make.bat
    exit /b 1
)
if not exist golden mkdir golden
if exist VisualDiff.out.txt del VisualDiff.out.txt
start /wait VisualDiff.exe record
findstr /C:"RESULT: PASSED" VisualDiff.out.txt >nul || (
    type VisualDiff.out.txt
    exit /b 1
)
type VisualDiff.out.txt
if exist VisualDiff.out.txt del VisualDiff.out.txt
start /wait VisualDiff.exe selftest
findstr /C:"RESULT: PASSED" VisualDiff.out.txt >nul || (
    type VisualDiff.out.txt
    exit /b 1
)
type VisualDiff.out.txt
echo Golden corpus recorded and self-test PASSED
