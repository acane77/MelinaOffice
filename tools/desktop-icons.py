#!/usr/bin/env python3
"""Regenerate the desktop app's icons (needs librsvg2-bin and ImageMagick): each app's document icon for Windows
from VO.appIcon, and the suite mark (icon.png, icon.ico, icon.icns) from public/icons/icon.svg.
Called by tools/rebuild_icons.sh."""
from pathlib import Path
import shutil
import struct
import subprocess

ROOT = Path(__file__).resolve().parents[1]


def render(svg, size):
    """svg (text) as a size × size PNG. librsvg keeps the 16-unit SVGs' strokes at the right scale; ImageMagick's
    internal SVG renderer can lose the white letter/screen strokes when enlarging these icons."""
    return subprocess.run(["rsvg-convert", "-w", str(size), "-h", str(size)], input=svg.encode(),
                          capture_output=True, check=True).stdout


def imagemagick():
    if shutil.which("magick"):
        return ["magick"]
    if shutil.which("convert"):
        return ["convert"]
    raise SystemExit("Needs ImageMagick (magick or convert)")


def png_to_ico(png, dest):
    subprocess.run(imagemagick() + ["png:-", "-define", "icon:auto-resize=256,128,64,48,32,24,16", str(dest)],
                   input=png, check=True)


def app_svg(app):
    """the folded-paper app icon from common/suite.js VO.appIcon, as a 256×256 SVG"""
    js = (ROOT / "public/common/suite.js").read_text()
    start = js.find("VO.appIcon = function")
    if start < 0:
        raise SystemExit("Cannot find VO.appIcon")
    depth, i = 0, js.find("{", start)
    while i < len(js):
        if js[i] == "{":
            depth += 1
        elif js[i] == "}":
            depth -= 1
            if depth == 0:
                body = js[start:i + 1]
                break
        i += 1
    else:
        raise SystemExit("Cannot find the end of VO.appIcon")
    svg = subprocess.run(
        ["node", "-e", "const VO = {}; let appIconN = 0; " + body + f"; process.stdout.write(VO.appIcon({app!r}, 256))"],
        capture_output=True, text=True, check=True,
    ).stdout
    return svg.replace("<svg ", '<svg xmlns="http://www.w3.org/2000/svg" ', 1)


for app in ("ecrivaine", "matheuse", "oratrice"):
    png_to_ico(render(app_svg(app), 256), ROOT / "desktop/icons" / f"{app}.ico")
    print(f"desktop/icons/{app}.ico")

web = (ROOT / "public/icons/icon.svg").read_text()
(ROOT / "desktop/icons/icon.png").write_bytes(render(web, 128))
print("desktop/icons/icon.png")
png_to_ico(render(web, 256), ROOT / "desktop/icons/icon.ico")
print("desktop/icons/icon.ico")

# macOS: the icon on Apple's grid, its rounded square 824 of 1024 units wide and centred (the web icon's square is 464
# of 512), each size drawn from the SVG. ICNS is a list of (type, PNG) entries: 16 to 512 points at 1x and 2x, the
# 1024 px one (ic10) being what the App Store asks for.
head, body = web.split("</defs>", 1)
body = body.rsplit("</svg>", 1)[0]
mac = f'{head}</defs>\n<g transform="translate(256 256) scale({412 / 464:.4f}) translate(-256 -256)">{body}</g>\n</svg>\n'
ICNS = [("icp4", 16), ("ic11", 32), ("icp5", 32), ("ic12", 64), ("ic07", 128), ("ic13", 256), ("ic08", 256),
        ("ic14", 512), ("ic09", 512), ("ic10", 1024)]
pngs = {size: render(mac, size) for size in {s for _, s in ICNS}}
entries = b"".join(kind.encode() + struct.pack(">I", 8 + len(pngs[size])) + pngs[size] for kind, size in ICNS)
(ROOT / "desktop/icons/icon.icns").write_bytes(b"icns" + struct.pack(">I", 8 + len(entries)) + entries)
print("desktop/icons/icon.icns")
