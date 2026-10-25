#!/bin/sh
# After editing public/icons/icon.svg or icon-maskable.svg (the suite mark) or VO.appIcon in
# common/suite.js (the three apps), rewrite every PNG, ICO and ICNS that install, Add to Home Screen
# and the desktop app use:
#
#   tools/rebuild_icons.sh
#
# Needs rsvg-convert (librsvg2-bin), ImageMagick (magick or convert) and node (for the app icons).
# Does not start Chromium.
set -eu
ROOT=$(CDPATH= cd -- "$(dirname "$0")/.." && pwd)
cd "$ROOT"

command -v rsvg-convert >/dev/null || { echo "Needs rsvg-convert (librsvg2-bin)" >&2; exit 1; }

png() {
  rsvg-convert -w "$2" -h "$2" "$1" > "$3"
  echo "$3"
}

png public/icons/icon.svg 512 public/icons/icon-512.png
png public/icons/icon.svg 192 public/icons/icon-192.png
png public/icons/icon-maskable.svg 512 public/icons/icon-maskable-512.png
png public/icons/icon-maskable.svg 180 public/icons/apple-touch-icon.png

python3 tools/desktop-icons.py
