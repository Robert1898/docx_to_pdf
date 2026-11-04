# -*- mode: python ; coding: utf-8 -*-
# Doua executabile (interfata + linie de comanda) intr-un singur folder (onedir):
# porneste mai repede decat onefile si da mai rar alarme false la antivirus.
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_data_files
from PyInstaller.utils.win32.versioninfo import (
    FixedFileInfo,
    StringFileInfo,
    StringStruct,
    StringTable,
    VarFileInfo,
    VarStruct,
    VSVersionInfo,
)

ROOT = Path(SPECPATH)
sys.path.insert(0, str(ROOT))
import convertor  # noqa: E402 - numele si versiunea au o singura sursa

VERSION = convertor.__version__
NUMBERS = tuple(int(p) for p in VERSION.split(".")) + (0,) * (4 - len(VERSION.split(".")))
GUI_NAME = Path(convertor.GUI_EXE).stem
CLI_NAME = Path(convertor.CLI_EXE).stem


def version_info(name, description):
    return VSVersionInfo(
        ffi=FixedFileInfo(filevers=NUMBERS, prodvers=NUMBERS),
        kids=[
            StringFileInfo([StringTable("041804B0", [
                StringStruct("CompanyName", convertor.PUBLISHER),
                StringStruct("FileDescription", description),
                StringStruct("FileVersion", VERSION),
                StringStruct("InternalName", name),
                StringStruct("LegalCopyright", f"© 2026 {convertor.PUBLISHER} · Licență MIT"),
                StringStruct("OriginalFilename", f"{name}.exe"),
                StringStruct("ProductName", convertor.APP_TITLE),
                StringStruct("ProductVersion", VERSION),
            ])]),
            VarFileInfo([VarStruct("Translation", [0x0418, 1200])]),
        ],
    )


dnd_datas, dnd_binaries, dnd_hidden = collect_all("tkinterdnd2")
common = dict(
    pathex=[str(ROOT)],
    binaries=dnd_binaries,
    datas=[("icon.ico", "."), *dnd_datas, *collect_data_files("ttkbootstrap")],
    hiddenimports=[*dnd_hidden, "win32com.client", "win32com.shell", "pythoncom", "pywintypes"],
    # AVIF nu e folosit nicaieri si ar adauga ~7 MB
    excludes=["pytest", "ruff", "PIL._avif", "PIL.AvifImagePlugin"],
)
a_gui = Analysis(["run_gui.py"], **common)
a_cli = Analysis(["run_cli.py"], **common)

exe_gui = EXE(
    PYZ(a_gui.pure),
    a_gui.scripts,
    [],
    exclude_binaries=True,
    name=GUI_NAME,
    console=False,
    icon="icon.ico",
    version=version_info(GUI_NAME, convertor.APP_TITLE),
    upx=False,
)
exe_cli = EXE(
    PYZ(a_cli.pure),
    a_cli.scripts,
    [],
    exclude_binaries=True,
    name=CLI_NAME,
    console=True,
    icon="icon.ico",
    version=version_info(CLI_NAME, f"{convertor.APP_TITLE} (linie de comandă)"),
    upx=False,
)
COLLECT(
    exe_gui, a_gui.binaries, a_gui.datas,
    exe_cli, a_cli.binaries, a_cli.datas,
    name=convertor.APP_NAME,
    upx=False,
)
