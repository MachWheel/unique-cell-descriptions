@echo off

REM Use the python from the venv to run pip
.\venv\Scripts\python.exe -m pip install pyinstaller

REM Directly use pyinstaller from the venv (after pip has installed it)
.\venv\Scripts\pyinstaller.exe -w --onefile main.py --name unique-cell-descriptions --icon icon.ico

REM Delete the build folder
if exist build rmdir /s /q build

REM Delete the unique-cell-descriptions.spec file
if exist unique-cell-descriptions.spec del unique-cell-descriptions.spec

REM Rename the dist folder to _release
if exist dist ren dist _release

REM Open the _release folder in Windows Explorer
start _release

echo Script executed successfully!
