@echo off
rem Builds dist\PDFMaker\PDFMaker.exe. Needs: pip install -r requirements.txt pyinstaller
rem Folder build (not --onefile): opens in ~1.5 s instead of ~7 s.
pyinstaller --noconfirm --onedir --windowed --name PDFMaker ^
  --exclude-module numpy --exclude-module PIL._avif --exclude-module PIL.AvifImagePlugin ^
  app.py || exit /b 1

rem Drop parts the app never uses (~80 MB): software OpenGL, Qt translations,
rem PyMuPDF's AI layout models and its C++ developer files.
set I=dist\PDFMaker\_internal
del /q "%I%\PySide6\opengl32sw.dll"
rmdir /s /q "%I%\PySide6\translations" "%I%\pymupdf\layout" "%I%\pymupdf\mupdf-devel"
