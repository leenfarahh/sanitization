"""Builds the Sanitize installer: python build_installer.py

It writes "Sanitize Setup.exe" one level up, next to this folder (Source) rather than inside it, so opening the
project folder shows the installer and Source, and the installer is never committed to the repository.

The installer holds everything the app needs, Python included, so the computer it goes on needs nothing installed
first. It installs for the current user without admin rights, into %LOCALAPPDATA%\\Programs\\Sanitize; adds Sanitize
to the Start menu and, if left ticked, the desktop; and is removed from Settings > Apps. Running a newer installer
updates an installed copy in place.

Building needs:
- Inno Setup 6.7 or later: winget install JRSoftware.InnoSetup
- 64-bit Python 3.9 or later from python.org, with tcl/tk, to run this script. That Python is copied into the
  installer (without its site-packages, test suite and IDLE), so the copied programs keep the Python Software
  Foundation's signature and antivirus tools treat them as Python, not as an unknown program.
- Microsoft Edge (on every Windows computer), which renders the installer's images from sanitize.svg.
"""

from __future__ import annotations

import compileall
import os
import py_compile
import shutil
import struct
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

import build_icon

HERE = Path(__file__).resolve().parent
ISS = HERE / "Sanitize.iss"
OUT = HERE.parent  # the folder holding Source, outside the repository
APP_FILES = ("sanitize_desktop.py", "sanitize_pptx.py", "sanitize.ico")  # copied by name, never by pattern
XY = f"{sys.version_info.major}{sys.version_info.minor}"
RUNTIME_FILES = ("pythonw.exe", "python3.dll", f"python{XY}.dll", "vcruntime140.dll", "vcruntime140_1.dll", "LICENSE.txt")
SKIP_LIB = {"site-packages", "test", "idlelib", "ensurepip", "turtledemo", "lib2to3", "pydoc_data"}
TCL_DIRS = ("tcl8", "tcl8.6", "tk8.6")  # Tcl, Tk, and the Tcl modules Tk loads (msgcat for its dialogs)

# Wizard images, in CSS pixels at 100% display scaling; each is also rendered at 200% for high-DPI screens
SIDE = (164, 314)  # left of the Finished page
TOP = (55, 55)  # top right of the other pages
LOGO = build_icon.SVG.as_uri()
SIDE_HTML = (  # sized in pixels, not to the window: headless Edge lays pages out wider than a small window
    "<!doctype html><style>html,body{margin:0;background:transparent}"
    f"body{{width:{SIDE[0]}px;height:{SIDE[1]}px;background:#002528;display:flex;flex-direction:column;"
    "align-items:center;justify-content:center;gap:12px;font:600 20px Figtree,'Segoe UI',sans-serif;color:#f4f6f6}"
    f'</style><img src="{LOGO}" width="88" height="88"><div>Sanitize</div>'
)
TOP_HTML = ("<!doctype html><style>html,body{margin:0;background:transparent}img{display:block}</style>"
            f'<img src="{LOGO}" width="{TOP[0]}" height="{TOP[1]}">')


def find_iscc() -> Path:
    found = shutil.which("ISCC")
    if found:
        return Path(found)
    for base in (Path(os.environ.get("LOCALAPPDATA", "")) / "Programs", os.environ.get("ProgramFiles(x86)"),
                 os.environ.get("ProgramFiles")):
        path = Path(base or "") / "Inno Setup 6" / "ISCC.exe"
        if base and path.exists():
            return path
    sys.exit("Inno Setup 6 was not found. Install it (winget install JRSoftware.InnoSetup), then run this again.")


def check_python() -> Path:
    """The python.org install this script runs on (the base install, when run from a virtual environment)."""
    base = Path(sys.base_prefix)
    if sys.platform != "win32" or struct.calcsize("P") != 8 or sys.version_info < (3, 9):
        sys.exit("Run this with 64-bit Python 3.9 or later for Windows, from python.org.")
    if "WindowsApps" in base.parts:
        sys.exit("This is the Microsoft Store Python, which can't be copied. Run this with Python from python.org.")
    missing = [name for name in RUNTIME_FILES if not (base / name).exists()]
    missing += [f"tcl\\{name}" for name in TCL_DIRS if not (base / "tcl" / name).is_dir()]
    if missing or not (base / "DLLs" / "_tkinter.pyd").exists():
        sys.exit(f"{base} is missing {', '.join(missing) or '_tkinter'}. Install Python from python.org with "
                 '"tcl/tk and IDLE" ticked, then run this again.')
    return base


def stage_python(base: Path, dest: Path) -> None:
    """A private copy of this Python: only what the app can use, with no access to other Pythons' packages."""
    dest.mkdir(parents=True)
    for name in RUNTIME_FILES:
        shutil.copy2(base / name, dest / name)
    shutil.copytree(base / "DLLs", dest / "DLLs", ignore=shutil.ignore_patterns("_test*", "_ctypes_test*"))
    lib = base / "Lib"
    shutil.copytree(lib, dest / "Lib", ignore=lambda folder, names: [
        n for n in names if n == "__pycache__" or (Path(folder) == lib and n in SKIP_LIB)])
    for name in TCL_DIRS:
        shutil.copytree(base / "tcl" / name, dest / "tcl" / name)
    # With this file, Python looks for modules only in these folders (relative to it, ".." being the app's folder)
    # and ignores PYTHONPATH, the registry and user site-packages: another Python on the computer can't interfere.
    (dest / f"python{XY}._pth").write_text("Lib\nDLLs\n..\n", encoding="ascii")


def stage(folder: Path) -> Path:
    app = folder / "app"
    stage_python(check_python(), app / "python")
    for name in APP_FILES:
        shutil.copy2(HERE / name, app / name)
    # Precompiled, so the first start is quick and nothing is written into the install folder later. Hash-checked
    # bytecode stays valid whatever timestamps the installer gives the files.
    if not compileall.compile_dir(app, quiet=1, workers=0,
                                  invalidation_mode=py_compile.PycInvalidationMode.CHECKED_HASH):
        sys.exit("Some files did not compile; see the errors above.")
    return app


def wizard_images(folder: Path) -> dict[str, str]:
    """Inno Setup's two wizard images, each as a 100% and a 200% PNG, comma-separated as Inno Setup takes them."""
    edge = build_icon.find_edge()
    folder.mkdir()
    files = {}
    for key, html, (width, height) in (("WizardSide", SIDE_HTML, SIDE), ("WizardTop", TOP_HTML, TOP)):
        paths = []
        for scale in (1, 2):
            path = folder / f"{key}@{scale}x.png"
            path.write_bytes(build_icon.screenshot(edge, html, width, height, folder, scale))
            paths.append(str(path))
        files[key] = ",".join(paths)
    return files


def main() -> None:
    iscc = find_iscc()
    today = date.today()
    version = f"{today.year}.{today.month}.{today.day}"
    with tempfile.TemporaryDirectory(prefix="sanitize-build-") as tmp:
        folder = Path(tmp)
        print("Copying Python and the app...")
        app = stage(folder)
        print("Rendering the installer's images...")
        images = wizard_images(folder / "images")
        print("Building the installer...")
        defines = {"StageDir": str(app), "AppVersion": version, **images}
        done = subprocess.run([str(iscc), "/Q", f"/O{OUT}", *(f"/D{k}={v}" for k, v in defines.items()), str(ISS)])
        if done.returncode:
            sys.exit("Inno Setup could not build the installer; see the errors above.")
    setup = OUT / "Sanitize Setup.exe"
    print(f"{setup}: version {version}, {setup.stat().st_size / 1e6:.1f} MB")


if __name__ == "__main__":
    main()
