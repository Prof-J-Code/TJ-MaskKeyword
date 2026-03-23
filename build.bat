@echo off
REM Build script for TJ-MaskKeyword (run inside Wine)

REM Install Python dependencies
python -m pip install python-docx pywin32 pyinstaller

REM Build with PyInstaller
python -m PyInstaller --onefile --windowed --name=TJ-MaskKeyword ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=docx ^
    main.py

echo Build complete. Exe is in dist/main.exe