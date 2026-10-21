#!/bin/sh
# Build the gitignored Quire VB6 engine.
#   make
#   tools/vb6/build-engine.sh
# Writes public/quire/vendor/vb6/vb6-engine.js (global VB6Engine).
set -eu
cd "$(dirname "$0")/../.."
entry=third_party/VB6/src/runtime/engine-entry.js
if [ ! -f "$entry" ]; then
  echo "VB6 engine entry is missing: $entry" >&2
  echo "Initialize the submodule: git submodule update --init third_party/VB6" >&2
  exit 1
fi
command -v node >/dev/null || { echo "node is required" >&2; exit 1; }
node tools/vb6/build-engine.mjs
