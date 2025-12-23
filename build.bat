@echo off
setlocal
pushd "%~dp0"

python -m PyInstaller --onefile --windowed --icon "icon.ico" --name "MMO Deck" --add-data "icon.ico;." main.py

popd
endlocal
