@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

rem Double-click launcher for PoE1 Traditional Chinese patch (PoeChinese3).
rem Usage: run_hanhua.cmd [-DryRun] [-ListTargets] [-SkipDownload] [-SetLanguageFr] [-GamePath "..."]

set "PSHOST=powershell.exe"
where /q powershell.exe
if errorlevel 1 set "PSHOST=pwsh.exe"

set "ARGS=%*"
"%PSHOST%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Poe1_TW.ps1" %*
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
  echo [DONE] Chinese patch finished. Launch PoE and pick the 2nd ^(French^) flag.
) else (
  echo [ERROR] Exit code %RC%. Please read the messages above.
)
echo.
echo Re-run this file after every game update.
echo.
echo %ARGS% | findstr /i /c:"-NoPause" >nul
if errorlevel 1 (pause >nul) else (echo [NoPause] skipped.)
exit /b %RC%
