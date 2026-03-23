import PyInstaller.__main__

PyInstaller.__main__.run([
    'main.py',
    '--onefile',
    '--windowed',
    '--name=TJ-MaskKeyword',
    '--add-binary=pythoncom37.dll;.',
    '--add-binary=pythonw37.dll;.',
    '--hidden-import=win32com.client',
    '--hidden-import=pythoncom',
    '--hidden-import=tkinter',
    '--hidden-import=docx',
])