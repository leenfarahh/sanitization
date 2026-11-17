"""Builds sanitize.ico, the desktop app's icon, from sanitize.svg: python build_icon.py

Microsoft Edge (on every Windows computer) renders the SVG at each size, so small icons stay sharp; each size is
stored in the .ico as a PNG. Run it after changing the SVG, then run Setup Sanitize.bat so the shortcut shows the
new icon. The .ico is kept in the repository, so setting the app up on a computer doesn't need this.
"""

from __future__ import annotations

import os
import struct
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SVG, ICO = HERE / "sanitize.svg", HERE / "sanitize.ico"
SIZES = (16, 20, 24, 32, 40, 48, 64, 256)  # title bar and taskbar at 100% to 200% display scaling, large icon views


def find_edge() -> Path:
    for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles")):
        path = Path(base or "") / "Microsoft" / "Edge" / "Application" / "msedge.exe"
        if base and path.exists():
            return path
    sys.exit("Microsoft Edge was not found; it renders the SVG.")


def render(edge: Path, size: int, folder: Path) -> bytes:
    """The SVG as a size x size PNG with a transparent background."""
    page, out = folder / f"{size}.html", folder / f"{size}.png"
    page.write_text("<!doctype html><style>html,body{margin:0;background:transparent}img{display:block}</style>"
                    f'<img src="{SVG.as_uri()}" width="{size}" height="{size}">', encoding="utf-8")
    subprocess.run([str(edge), "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                    "--default-background-color=00000000", f"--user-data-dir={folder / 'profile'}",
                    f"--window-size={size},{size}", f"--screenshot={out}", page.as_uri()],
                   capture_output=True, timeout=120)
    png = out.read_bytes() if out.exists() else b""
    if not png.startswith(b"\x89PNG\r\n\x1a\n") or struct.unpack(">II", png[16:24]) != (size, size):
        sys.exit(f"Edge did not render the {size} px icon.")
    return png


def ico(pngs: dict[int, bytes]) -> bytes:
    """An .ico holding one PNG per size (Windows reads PNG icons, transparency included)."""
    head = struct.pack("<HHH", 0, 1, len(pngs))
    offset, body = 6 + 16 * len(pngs), b""
    for size, png in pngs.items():
        head += struct.pack("<BBBBHHII", size % 256, size % 256, 0, 0, 1, 32, len(png), offset + len(body))  # 0 = 256
        body += png
    return head + body


def main() -> None:
    edge = find_edge()
    with tempfile.TemporaryDirectory() as folder:
        pngs = {size: render(edge, size, Path(folder)) for size in SIZES}
    ICO.write_bytes(ico(pngs))
    print(f"{ICO.name}: {', '.join(map(str, SIZES))} px, {ICO.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
