# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec: two executables sharing one folder.
#   tacitus-lite.exe     console, for the Start menu shortcuts and the terminal
#   tacitus-lite-bg.exe  windowless, for the scheduled task (logs to mail/dump.log)
# Build on Windows:  pyinstaller installer/tacitus-lite.spec

from pathlib import Path

root = Path(SPECPATH).parent
entry = str(root / "tacitus_lite.py")
hidden = ["win32com.client", "pythoncom", "pywintypes", "win32timezone", "outlook_dump", "make_draft"]

console = Analysis([entry], pathex=[str(root)], hiddenimports=hidden, noarchive=False)
windowed = Analysis([entry], pathex=[str(root)], hiddenimports=hidden, noarchive=False)

exe_console = EXE(PYZ(console.pure), console.scripts, [], exclude_binaries=True,
                  name="tacitus-lite", console=True, upx=False)
exe_windowed = EXE(PYZ(windowed.pure), windowed.scripts, [], exclude_binaries=True,
                   name="tacitus-lite-bg", console=False, upx=False)

COLLECT(exe_console, console.binaries, console.datas,
        exe_windowed, windowed.binaries, windowed.datas,
        name="TacitusLite", upx=False)
