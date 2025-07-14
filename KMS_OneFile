@echo off
setlocal EnableDelayedExpansion

:: 0) Check for elevation, relaunch as Admin if needed
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo [!] Requesting elevation...
  powershell -Command "Start-Process '%~f0' -Verb runAs"
  exit /b
)

:: 1/4) Configure & activate Windows via KMS
echo [1/4] Configuring Windows KMS host...
cscript //nologo "%windir%\system32\slmgr.vbs" /skms kms8.msguides.com:1688 >nul
echo [1/4] Activating Windows...
cscript //nologo "%windir%\system32\slmgr.vbs" /ato >nul
echo.

:: 2/4) Find Office’s ospp.vbs
set "OSPP=%ProgramFiles%\Microsoft Office\Office16\ospp.vbs"
if not exist "%OSPP%" set "OSPP=%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs"
if not exist "%OSPP%" (
  echo [!] Cannot find Office16\ospp.vbs – adjust path and re-run.
  pause & exit /b
)

:: 3/4) Install generic Office ProPlus KMS client key
echo [2/4] Installing generic Office ProPlus KMS client key...
cscript //nologo "%OSPP%" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul
echo.

:: 4/4) Point Click-to-Run at KMS host & activate Office
echo [3/4] Pointing Office Click-to-Run to KMS host...
reg add "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v KMSHostName /t REG_SZ /d kms8.msguides.com /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v KMSHostPort /t REG_DWORD /d 1688 /f >nul

echo [4/4] Restarting ClickToRunSvc...
net stop ClickToRunSvc >nul 2>&1
net start ClickToRunSvc >nul 2>&1

echo [4/4] Activating Office...
cscript //nologo "%OSPP%" /act
echo.

echo ===== All done! =====
pause
endlocal
