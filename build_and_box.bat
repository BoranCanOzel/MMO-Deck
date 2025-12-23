@echo off
setlocal
set "ROOT=%~dp0"
pushd "%ROOT%"

echo Running PyInstaller...
python -m PyInstaller --onefile --windowed --icon "icon.ico" --name "MMO Deck" --add-data "icon.ico;." main.py
if errorlevel 1 (
    echo PyInstaller build failed.
    popd
    endlocal
    exit /b 1
)

echo Running Enigma Virtual Box...
call "%ROOT%pack_evb.bat"
set "ERR=%ERRORLEVEL%"
if not "%ERR%"=="0" (
    echo Boxing failed (code %ERR%).
    popd
    endlocal
    exit /b %ERR%
)

echo Build and boxing complete.
popd
endlocal
