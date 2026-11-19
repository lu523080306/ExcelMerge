@echo off
setlocal
cd /d "%~dp0"

set "CONFIG=Debug"
set "BUILD_TARGET=Build"
if /I "%~1"=="rebuild" set "BUILD_TARGET=Rebuild"

set "MSBUILD=C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
if not exist "%MSBUILD%" set "MSBUILD=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"

echo ============================================================
echo  ExcelMerge 构建
echo  模式: %BUILD_TARGET%    配置: %CONFIG%
echo ============================================================

if not exist "%MSBUILD%" (
  echo [错误] 找不到 MSBuild，请修改 build.cmd 顶部的 MSBUILD 路径。
  echo.
  pause
  exit /b 1
)

rem ---- 正在运行的程序会锁住 bin\Debug 下的 exe/dll ----
tasklist /FI "IMAGENAME eq ExcelMerge.GUI.exe" 2>nul | findstr /I /C:"ExcelMerge.GUI.exe" >nul
if not errorlevel 1 (
  echo.
  echo [错误] ExcelMerge.GUI.exe 正在运行，构建输出文件被占用。
  echo        请先关闭程序，再重新运行 build.cmd。
  echo.
  pause
  exit /b 1
)

set "NU=%CD%\.nuget\nuget.exe"

echo [1/4] 准备 .NET Framework 引用程序集 ...
"%NU%" install Microsoft.NETFramework.ReferenceAssemblies.net452 -Version 1.0.3 -OutputDirectory packages >nul
"%NU%" install Microsoft.NETFramework.ReferenceAssemblies.net45  -Version 1.0.3 -OutputDirectory packages >nul
"%NU%" install Microsoft.NETFramework.ReferenceAssemblies.net40  -Version 1.0.3 -OutputDirectory packages >nul
if errorlevel 1 (
  echo [错误] 引用程序集准备失败。
  pause
  exit /b 1
)

echo [2/4] 还原 NuGet 包 ...
"%NU%" restore ExcelMerge.sln -ConfigFile .nuget\NuGet.Config >nul || (
  echo [错误] NuGet restore 失败。
  pause
  exit /b 1
)

set "LOG=%CD%\build.log"
echo [3/4] 编译 ExcelMerge.GUI %CONFIG% ...
echo       日志: %LOG%
"%MSBUILD%" ExcelMerge.GUI\ExcelMerge.GUI.csproj /t:%BUILD_TARGET% /p:Configuration=%CONFIG% /m /v:m > "%LOG%" 2>&1
set "MSBUILD_EXIT=%ERRORLEVEL%"

if not "%MSBUILD_EXIT%"=="0" (
  echo.
  echo [错误] 构建失败，退出码 %MSBUILD_EXIT%。
  echo        正在打开日志: %LOG%
  echo.
  start "" notepad "%LOG%"
  pause
  exit /b %MSBUILD_EXIT%
)

echo [4/4] 构建完成
for %%F in ("ExcelMerge.GUI\bin\%CONFIG%\ExcelMerge.GUI.exe") do echo   exe : %%~fF
for %%F in ("ExcelMerge.GUI\bin\%CONFIG%\ExcelMerge.GUI.exe") do echo   时间: %%~tF
echo   日志: %LOG%

if /I "%BUILD_TARGET%"=="Build" (
  echo.
  echo   提示: 当前是增量构建。源码没有修改时 exe 时间戳不会变化，
  echo         看到时间没变不代表构建失败。
  echo         需要强制重新编译请运行: build.cmd rebuild
)

echo.
echo 构建流程结束，按任意键关闭窗口...
pause >nul
exit /b 0
