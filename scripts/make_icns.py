from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: python scripts/make_icns.py icon.ico icon.icns", file=sys.stderr)
        return 2
    src = Path(sys.argv[1]).resolve()
    dst = Path(sys.argv[2]).resolve()
    if not src.exists():
        raise FileNotFoundError(src)
    iconutil = shutil.which("iconutil")
    if not iconutil:
        raise RuntimeError("iconutil is required on macOS to create .icns files")

    iconset = dst.with_suffix(".iconset")
    if iconset.exists():
        shutil.rmtree(iconset)
    iconset.mkdir(parents=True)

    image = Image.open(src).convert("RGBA")
    variants = [
        (16, "icon_16x16.png"),
        (32, "icon_16x16@2x.png"),
        (32, "icon_32x32.png"),
        (64, "icon_32x32@2x.png"),
        (128, "icon_128x128.png"),
        (256, "icon_128x128@2x.png"),
        (256, "icon_256x256.png"),
        (512, "icon_256x256@2x.png"),
        (512, "icon_512x512.png"),
        (1024, "icon_512x512@2x.png"),
    ]
    for size, name in variants:
        resized = image.resize((size, size), Image.Resampling.LANCZOS)
        resized.save(iconset / name)

    subprocess.run([iconutil, "-c", "icns", str(iconset), "-o", str(dst)], check=True)
    shutil.rmtree(iconset)
    print(dst)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
