@echo off
REM ============================================================
REM  SeewoOptimizer source sync:  repo  ->  G:\SeewoBuild\proj
REM
REM  WHY THIS EXISTS
REM
REM  Two independent copies of the source tree are kept:
REM
REM    _repo\WindowsFormsApp1               <- the real source, under git
REM    G:\SeewoBuild\proj\WindowsFormsApp1  <- the build workbench
REM
REM  Only the workbench is compilable: it sits beside the NuGet
REM  packages\ folder and the self-supplied targeting pack. But
REM  only the repo is authoritative and version controlled.
REM
REM  Nothing kept them in step, so they drifted. A release was once
REM  produced from a workbench copy that was missing the latest
REM  edits: the exe compiled fine and looked right, but it did not
REM  match the committed source. Silent, and expensive to notice -
REM  the binary and the repo just disagreed.
REM
REM  So every build and publish now mirrors the repo first.
REM
REM  MIRROR, NOT MERGE
REM
REM  /MIR makes the workbench an exact copy, deleting files the repo
REM  no longer has. A stale file left behind could still influence
REM  the build. Exactness is worth more than keeping strays.
REM
REM  Checked before enabling /MIR: 26 files on each side and zero
REM  unique to either, so there is nothing to lose.
REM
REM  bin\ and obj\ are excluded - build output, not source.
REM  Mirroring them would force a full rebuild every time.
REM
REM  Callers: build.cmd, publish.cmd
REM ============================================================

setlocal enabledelayedexpansion

REM %~dp0 is the folder holding this script, so the repo is found
REM relative to it and the Chinese directory name never appears in
REM this file's body. Two copies of this script exist:
REM
REM   <root>\sync.cmd               repo is in .\_repo
REM   <root>\_repo\tools\sync.cmd   repo is in ..\..
REM
REM The REPO line below is the ONLY difference between the two
REM copies; everything else is byte identical.
set REPO=%~dp0_repo
set WB=G:\SeewoBuild\proj

if not exist "%REPO%\WindowsFormsApp1" (
  echo [ERROR] Repo source not found: %REPO%\WindowsFormsApp1
  exit /b 1
)

if not exist "%WB%\WindowsFormsApp1" (
  echo [ERROR] Workbench not found: %WB%\WindowsFormsApp1
  exit /b 1
)

echo [sync] Mirroring repo source into the build workbench
echo        from: %REPO%
echo        to:   %WB%
echo.

REM ------------------------------------------------------------
REM  Backup before overwriting.
REM
REM  /MIR deletes as well as copies, so a mistake costs real work.
REM  The tree is small (~26 files, ~710 KB), so copying it is cheap
REM  insurance.
REM
REM  WHY BACK UP EVERY RUN INSTEAD OF ONLY WHEN SOMETHING CHANGED:
REM  the exit code cannot tell "changed" from "unchanged". robocopy
REM  sets the same bit (value 2, "extras") for the excluded obj\ and
REM  bin\ directories as it does for real differences - a raw run
REM  with nothing to do still returns 2. Telling them apart would
REM  mean parsing the summary table, which is brittle and locale
REM  dependent. Given the size, just always back up.
REM
REM  Numbered folders rather than a timestamp: %DATE% depends on the
REM  system locale (yyyy/MM/dd in zh-CN, MM/dd/yyyy in en-US), so
REM  slicing it would produce wrong names on other machines.
REM
REM  Taking the MAX of the existing numbers, not the last one by
REM  name: "10" sorts before "9" lexically, so reading the last
REM  entry would pick the wrong value and overwrite an old backup.
REM ------------------------------------------------------------
set NEXT=1
if exist "%WB%\_backup" (
  for /f %%n in ('dir /b /ad "%WB%\_backup" 2^>nul') do (
    echo %%n| findstr /r "^[0-9][0-9]*$" >nul && (
      if %%n GEQ !NEXT! set NEXT=%%n
    )
  )
  set /a NEXT=!NEXT!+1
)

set BACKUPDIR=%WB%\_backup\%NEXT%

echo [sync] Backing up current workbench to _backup\%NEXT%
robocopy "%WB%\WindowsFormsApp1" "%BACKUPDIR%\WindowsFormsApp1" /E /NFL /NDL /NJH /NJS /NP /XD obj bin >nul
robocopy "%WB%" "%BACKUPDIR%" WindowsFormsApp1.sln /NFL /NDL /NJH /NJS /NP >nul

REM ------------------------------------------------------------
REM  Prune old backups, keeping the newest %KEEP%.
REM
REM  Without this the folder grows without bound - every run makes
REM  one, so that happens fast during a working session.
REM
REM  CUTOFF = NEXT - KEEP + 1, not NEXT - KEEP. %NEXT% is the folder
REM  just created, so the run keeps NEXT-KEEP+1 .. NEXT - that is
REM  KEEP folders. Using NEXT-KEEP would keep KEEP+1 and grow by one
REM  every run.
REM ------------------------------------------------------------
set KEEP=10
set CUTOFF=0
if !NEXT! GTR %KEEP% set /a CUTOFF=!NEXT!-%KEEP%+1

if !CUTOFF! GTR 0 (
  for /f %%n in ('dir /b /ad "%WB%\_backup" 2^>nul') do (
    echo %%n| findstr /r "^[0-9][0-9]*$" >nul && (
      if %%n LSS !CUTOFF! (
        echo [sync] Pruning old backup _backup\%%n
        rd /s /q "%WB%\_backup\%%n" 2>nul
      )
    )
  )
)

REM ------------------------------------------------------------
REM  The mirror itself.
REM
REM  robocopy exit codes are NOT ordinary: 0 means "nothing to do"
REM  and 1-7 all mean success (bits flag what was copied/skipped).
REM  Only 8 and above are real failures. A naive "if errorlevel 1"
REM  would abort on every successful copy.
REM ------------------------------------------------------------
robocopy "%REPO%\WindowsFormsApp1" "%WB%\WindowsFormsApp1" /MIR /NFL /NDL /NJH /NJS /NP /XF *.user /XD obj bin
set RC=%ERRORLEVEL%

if %RC% GEQ 8 (
  echo [ERROR] robocopy failed with code %RC%
  exit /b 1
)

REM Code 0-7 all mean success. The value is a bit field (1 = files
REM copied, 2 = extras, 4 = mismatches), so it says what categories
REM were SEEN, not whether anything actually happened - an excluded
REM obj\ and bin\ alone raise bit 2. So do not claim "updated" here;
REM just confirm the mirror ran clean.
echo [sync] Mirror complete ^(robocopy code %RC%^)

robocopy "%REPO%" "%WB%" WindowsFormsApp1.sln /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 (
  echo [ERROR] Failed to copy the solution file
  exit /b 1
)

echo.
endlocal
exit /b 0
