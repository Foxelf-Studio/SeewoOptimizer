@echo off
REM ============================================================
REM  SeewoOpt publish script (sync + build + sign + show upload list)
REM
REM  TWO files must be uploaded together to the GitHub Release:
REM    SeewoOpt.exe        main program
REM    SeewoOpt.exe.sig    RSA-3072 signature (raw binary, not base64)
REM
REM  Private key: G:\SeewoBuild\signing\private_key.pem
REM  NEVER commit the private key or upload it to any online service.
REM
REM  Usage:
REM    publish.cmd          build Release and sign
REM    publish.cmd Debug    build Debug and sign (do not publish)
REM ============================================================

setlocal

set ROOT=G:\SeewoBuild
set KEYDIR=%ROOT%\signing
set PRIKEY=%KEYDIR%\private_key.pem
set PUBKEY=%KEYDIR%\public_key.pem
set PROJ=%ROOT%\proj\WindowsFormsApp1

if "%1"=="Debug" (set CFG=Debug) else (set CFG=Release)

echo ============================================================
echo  SeewoOptimizer Publish
echo  Configuration: %CFG%
echo ============================================================
echo.

if not exist "%PRIKEY%" (
  echo [ERROR] Private key not found: %PRIKEY%
  echo         Generate one with:
  echo         openssl genrsa -out private_key.pem 3072
  echo         openssl rsa -in private_key.pem -pubout -out public_key.pem
  exit /b 1
)

echo [1/4] Syncing sources from the repository...
call "%~dp0sync.cmd"
if errorlevel 1 (
  echo [FAILED] Source sync failed - refusing to sign a stale build
  exit /b 1
)

cd /d "%PROJ%"
if errorlevel 1 (
  echo [ERROR] Cannot enter project directory
  exit /b 1
)

echo [2/4] Building...
dotnet msbuild WindowsFormsApp1.csproj -t:Rebuild -p:Configuration=%CFG% -p:TargetFrameworkRootPath=%ROOT%\build -v:minimal -nologo
if errorlevel 1 (
  echo [FAILED] Build failed
  exit /b 1
)

set OUTDIR=%PROJ%\bin\%CFG%
set EXE=%OUTDIR%\SeewoOpt.exe
set SIG=%OUTDIR%\SeewoOpt.exe.sig

if not exist "%EXE%" (
  echo [ERROR] Output not found: %EXE%
  exit /b 1
)

echo.
echo [3/4] Signing with private key...
openssl dgst -sha256 -sign "%PRIKEY%" -out "%SIG%" "%EXE%"
if errorlevel 1 (
  echo [ERROR] Signing failed
  exit /b 1
)

echo [4/4] Self-check with public key...
openssl dgst -sha256 -verify "%PUBKEY%" -signature "%SIG%" "%EXE%"
if errorlevel 1 (
  echo [ERROR] Signature self-check FAILED - do not publish
  exit /b 1
)

echo.
echo ============================================================
echo  [OK] Build and sign succeeded
echo ============================================================
echo.
dir /-c "%EXE%" "%SIG%"
echo.
echo Upload BOTH files to the GitHub Release:
echo    SeewoOpt.exe
echo    SeewoOpt.exe.sig
echo.
echo SHA256 of exe:
certutil -hashfile "%EXE%" SHA256
echo.
echo NOTE: The .sig file is REQUIRED. Without it the client
echo       refuses to update.
echo.
echo Also: bump version in Properties\AssemblyInfo.cs, then tag.

endlocal
exit /b 0
