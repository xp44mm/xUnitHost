@echo off
setlocal EnableExtensions
rem AcadMgdHost smoke test: build first, then run this file.
rem Usage: run.cmd [path-to-a-dwg]
rem   /i is mandatory (accoreconsole 2023 fails with error 53 without it)
rem   /s only accepts an ABSOLUTE path
rem   _.quit _y means DISCARD in accoreconsole (opposite of full acad.exe)
rem   ACADMGDHOST_AUTORUN=1 runs the probe from Initialize (no command registration needed)
set "ACAD_DIR=C:\Program Files\Autodesk\AutoCAD 2020"
set "EXE=%ACAD_DIR%\accoreconsole.exe"
rem Normalize the parent path so NETLOAD gets a clean absolute path.
for %%I in ("%~dp0..") do set "ROOT=%%~fI"
set "DLL=%ROOT%\bin\Debug\netstandard2.0\AcadMgdHost.dll"
rem Console.Out goes to accoreconsole's own accc*.tmp, so keep a separate readable log.
set "LOG=%TEMP%\AcadMgdHost.log"
set "ACADMGDHOST_LOG=%LOG%"
if exist "%LOG%" del /q "%LOG%"
rem Default seed lives under Program Files (read-only) -> copy to TEMP, never touch the original.
set "SEED=%~1"
if not defined SEED (
  set "SEED=%TEMP%\AcadMgdHost-seed.dwg"
  copy /y "%ACAD_DIR%\Express\brkline.dwg" "%TEMP%\AcadMgdHost-seed.dwg" >nul
)
if not defined ACADMGDHOST_AUTORUN set "ACADMGDHOST_AUTORUN=1"
set "SCR=%TEMP%\AcadMgdHost-hello.scr"
set "DLLPOSIX=%DLL:\=/%"

if not exist "%DLL%" echo [AcadMgdHost] build first: dotnet build xUnitHost.slnx & exit /b 2
if not exist "%EXE%" echo [AcadMgdHost] accoreconsole.exe not found: %EXE% & exit /b 2
if not exist "%SEED%" echo [AcadMgdHost] seed DWG not found: %SEED% & exit /b 2

>  "%SCR%" echo (setvar "secureload" 0)
>> "%SCR%" echo (setvar "filedia" 0)
>> "%SCR%" echo (command "_.netload" "%DLLPOSIX%")
>> "%SCR%" echo XH_HELLO
>> "%SCR%" echo XH_DB
>> "%SCR%" echo _.quit _y

echo [AcadMgdHost] script = %SCR%
type "%SCR%"
echo.
"%EXE%" /i "%SEED%" /s "%SCR%"
set "RC=%ERRORLEVEL%"
echo [AcadMgdHost] accoreconsole exit = %RC%
echo [AcadMgdHost] log = %LOG%
if exist "%LOG%" type "%LOG%"
exit /b %RC%
