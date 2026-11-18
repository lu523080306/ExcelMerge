@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  ExcelMerge build  (no .NET Targeting Pack fallback setup)
echo ============================================================

set CONFIG=Debug
if not "%1"=="" set CONFIG=%1

rem ---- locate MSBuild ----
set "MSBUILD=C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
if not exist "%MSBUILD%" set "MSBUILD=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
if not exist "%MSBUILD%" (
  echo [ERROR] MSBuild not found. Edit MSBUILD path at top of this script.
  exit /b 1
)

set "NU=%CD%\.nuget\nuget.exe"

echo [1/4] Install .NET Framework reference-assembly packages (net452/net45/net40) ...
"%NU%" install Microsoft.NETFramework.ReferenceAssemblies.net452 -Version 1.0.3 -OutputDirectory packages >nul
"%NU%" install Microsoft.NETFramework.ReferenceAssemblies.net45  -Version 1.0.3 -OutputDirectory packages >nul
"%NU%" install Microsoft.NETFramework.ReferenceAssemblies.net40  -Version 1.0.3 -OutputDirectory packages >nul
if errorlevel 1 exit /b 1

echo [2/4] Restore NuGet packages (packages.config) ...
"%NU%" restore ExcelMerge.sln -ConfigFile .nuget\NuGet.Config >nul || exit /b 1

echo [3/4] Build ExcelMerge.GUI %CONFIG% ...
"%MSBUILD%" ExcelMerge.GUI\ExcelMerge.GUI.csproj /t:Build /p:Configuration=%CONFIG% /m /v:m
if errorlevel 1 (
  echo [ERROR] Build failed.
  exit /b 1
)

echo [4/4] Done. Output: bin\%CONFIG%\ExcelMerge.GUI.exe
echo    Usage: ExcelMerge.GUI.exe diff -s LEFT -d RIGHT
exit /b 0