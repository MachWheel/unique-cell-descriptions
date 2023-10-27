@echo off

REM Use the python from the venv to run pip
.\venv\Scripts\python.exe -m pip install pyinstaller

REM Directly use pyinstaller from the venv (after pip has installed it)
.\venv\Scripts\pyinstaller.exe -w --onefile main.py --name unique-cell-descriptions --icon icon.ico

REM Check if an .exe exists in the dist folder
if not exist dist\*.exe (
    echo Script did not compile due to pyinstaller issues.
    exit /b
)

REM Delete the build folder
if exist build rmdir /s /q build

REM Delete the unique-cell-descriptions.spec file
if exist unique-cell-descriptions.spec del unique-cell-descriptions.spec

REM Check if _release folder exists
if exist _release (
    REM If _release exists, delete it and create it again
    rmdir /s /q _release
    mkdir _release
    REM Move all files from dist to _release
    move dist\*.* _release
    REM Delete the dist folder
    rmdir /s /q dist
) else (
    REM If _release doesn't exist, rename dist to _release
    ren dist _release
)

REM Open the _release folder in Windows Explorer
start _release

echo Script executed successfully!
