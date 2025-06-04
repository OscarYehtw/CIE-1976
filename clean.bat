@echo off
REM Build the stn4_analyzer.py into an executable using PyInstaller

REM Ensure pyinstaller is installed
pip install pyinstaller

REM Remove previous build/dist folders
rmdir /s /q build
rmdir /s /q dist
rmdir /s /q __pycache__
del /q cie1976_analyzer.spec
del /q *.html
del /q *.exe
