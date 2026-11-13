@echo off
setlocal EnableExtensions
rem AcadMgdHost XH_STAR: draw a five-point star, save a DWG on the Desktop.
rem Build first: dotnet build xUnitHost.slnx -v q
rem Optional env: ACADMGDHOST_STAR_OUT (full target path)
rem Rerun: the command NEVER overwrites an existing DWG. Move the old one aside,
rem        or set ACADMGDHOST_STAR_OUT to a fresh path before running again.
rem   ACADMGDHOST_STAR_REPORT (UTF-8 key=value report, read by the Fact)
rem /i is mandatory; /s only accepts an ABSOLUTE path.
rem _.quit _y means DISCARD in accoreconsole (opposite of full acad.exe).
set "ACAD_DIR=C:\Program Files\Autodesk\AutoCAD 2020"
set "EXE=%ACAD_DIR%\accoreconsole.exe"
for %%I in ("%~dp0..") do set "ROOT=%%~fI"
set "DLL=%ROOT%\bin\Debug\netstandard2.0\AcadMgdHost.dll"
set "LOG=%TEMP%\AcadMgdHost-star.log"
set "ACADMGDHOST_LOG=%LOG%"
if exist "%LOG%" del /q "%LOG%"
if not defined ACADMGDHOST_STAR_REPORT set "ACADMGDHOST_STAR_REPORT=%TEMP%\AcadMgdHost-star-report.txt"
if exist "%ACADMGDHOST_STAR_REPORT%" del /q "%ACADMGDHOST_STAR_REPORT%"
set "SEED=%TEMP%\AcadMgdHost-star-seed.dwg"
copy /y "%ACAD_DIR%\Express\brkline.dwg" "%SEED%" >nul
attrib -r "%SEED%" >nul
set "SCR=%TEMP%\AcadMgdHost-star.scr"
set "DLLPOSIX=%DLL:\=/%"
set "ACADMGDHOST_AUTORUN=0"

if not exist "%DLL%" echo [AcadMgdHost] build first: dotnet build xUnitHost.slnx & exit /b 2
if not exist "%EXE%" echo [AcadMgdHost] accoreconsole.exe not found: %EXE% & exit /b 2
if not exist "%SEED%" echo [AcadMgdHost] seed DWG not found: %SEED% & exit /b 2

>  "%SCR%" echo (setvar "secureload" 0)
>> "%SCR%" echo (setvar "filedia" 0)
>> "%SCR%" echo (command "_.netload" "%DLLPOSIX%")
>> "%SCR%" echo XH_STAR
>> "%SCR%" echo _.quit _y

echo [AcadMgdHost] script = %SCR%
type "%SCR%"
echo.
"%EXE%" /i "%SEED%" /s "%SCR%"
set "RC=%ERRORLEVEL%"
echo [AcadMgdHost] accoreconsole exit = %RC%
echo [AcadMgdHost] log = %LOG%
if exist "%LOG%" type "%LOG%"
echo [AcadMgdHost] report (UTF-8) = %ACADMGDHOST_STAR_REPORT%
exit /b %RC%
