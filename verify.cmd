@echo off
setlocal
cd /d "%~dp0"

rem Headless regression check for the diff-apply core logic.
set "CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
set "BIN=%CD%\ExcelMerge.GUI\bin\Debug"

if not exist "%BIN%\ExcelMerge.dll" (
  echo [ERROR] Run build.cmd first.
  exit /b 1
)

echo [1/2] Compile verifier ...
"%CSC%" /nologo /out:%BIN%\apply_verify.exe /reference:%BIN%\ExcelMerge.dll /reference:%BIN%\NPOI.dll /reference:%BIN%\NPOI.OOXML.dll /reference:%BIN%\NPOI.OpenXml4Net.dll /reference:%BIN%\NPOI.OpenXmlFormats.dll %CD%\.smoketest\ApplyVerify.cs
if errorlevel 1 exit /b 1

echo [2/2] Run verifier ...
"%BIN%\apply_verify.exe" "%CD%\.smoketest"
exit /b %errorlevel%