@echo off
rem Builds dist\PDFMaker\PDFMaker.exe (folder build: opens in ~1.5 s; --onefile took ~7 s per launch)
pyinstaller --noconfirm --onedir --windowed --name PDFMaker app.py
