@echo off
setlocal

REM Build Windows EXE with PyInstaller
py -3 -m pip install --upgrade pip
py -3 -m pip install -r requirements.txt
py -3 -m pip install pyinstaller

REM Clean old artifacts
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist app.spec del /q app.spec

REM Build one-file exe, windowed mode
py -3 -m PyInstaller --noconfirm --clean --onefile --windowed --name dwg2img app.py

echo.
echo Build complete. Output: dist\dwg2img.exe
endlocal
