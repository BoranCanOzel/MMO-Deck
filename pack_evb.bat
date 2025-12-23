@echo off
setlocal

rem Locate EnigmaVBConsole.exe
set "EVB_EXE=EnigmaVBConsole.exe"
if exist "%~dp0EnigmaVBConsole.exe" set "EVB_EXE=%~dp0EnigmaVBConsole.exe"
if exist "%ProgramFiles%\\Enigma Virtual Box\\EnigmaVBConsole.exe" set "EVB_EXE=%ProgramFiles%\\Enigma Virtual Box\\EnigmaVBConsole.exe"
if exist "%ProgramFiles(x86)%\\Enigma Virtual Box\\EnigmaVBConsole.exe" set "EVB_EXE=%ProgramFiles(x86)%\\Enigma Virtual Box\\EnigmaVBConsole.exe"

pushd "%~dp0dist"

if not exist "%EVB_EXE%" (
    echo EnigmaVBConsole.exe not found. Place it in this folder or install Enigma Virtual Box.
    popd
    endlocal
    exit /b 1
)

"%EVB_EXE%" "vb.evb"
set "ERR=%ERRORLEVEL%"
if not "%ERR%"=="0" (
    echo Packaging failed with error %ERR%.
    popd
    endlocal
    exit /b %ERR%
)

popd
endlocal
echo Packaging complete.
