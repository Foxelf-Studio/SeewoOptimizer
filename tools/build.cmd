@echo off
REM ============================================================
REM  SeewoOptimizer build script (net48 WinForms, no Visual Studio)
REM  Usage: build.cmd [Debug|Release]
REM
REM  Step 0 mirrors the repo source into the build workbench so the
REM  two copies can never drift apart. See sync.cmd for the details
REM  (backup + /MIR + pruning).
REM
REM  net48 requires reference assemblies at G:\SeewoBuild\build\.NETFramework\v4.8\
REM  v4.7.2 is kept there too, so reverting is just a csproj change.
REM ============================================================

setlocal

set ROOT=G:\SeewoBuild
set PROJ=%ROOT%\proj\WindowsFormsApp1

if "%1"=="Debug" (
  set CFG=Debug
) else (
  set CFG=Release
)

echo ============================================================
echo  SeewoOptimizer Build
echo  Configuration: %CFG%
echo ============================================================
echo.

REM --- Step 0: sync sources from the repository into the workbench ---
call "%~dp0sync.cmd"
if errorlevel 1 (
  echo.
  echo [ERROR] Source sync failed - aborting build.
  echo         Fix the sync problem first; building now would use
  echo         stale sources and produce a misleading result.
  goto :fail
)

cd /d "%PROJ%"
if errorlevel 1 (
  echo [ERROR] Cannot enter project directory: %PROJ%
  goto :fail
)

dotnet msbuild WindowsFormsApp1.csproj -t:Rebuild -p:Configuration=%CFG% -p:TargetFrameworkRootPath=%ROOT%\build -v:minimal -nologo

if errorlevel 1 goto :fail

echo.
echo [OK] Build succeeded.
echo Output: %PROJ%\bin\%CFG%\SeewoOpt.exe
dir "%PROJ%\bin\%CFG%\SeewoOpt.exe"
endlocal
exit /b 0

:fail
echo.
echo [FAILED] Build failed with error code %ERRORLEVEL%
endlocal
exit /b 1
