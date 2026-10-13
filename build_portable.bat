@echo off
echo Building portable single-file executable...
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name "VoltageDropCalc" --clean main.py
echo.
echo Done! File is in dist\VoltageDropCalc.exe
echo You can copy this .exe to any Windows PC and run without installing Python.
pause
