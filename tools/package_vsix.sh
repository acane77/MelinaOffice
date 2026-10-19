#!/bin/sh
# Write public/version.json, then pack the VS Code extension.
#
#   tools/package_vsix.sh [out.vsix]
#
# version is package.json's "version". commit_id is git rev-parse --short HEAD.
# branch is git branch --show-current. build_time is today's date, YYYY-MM-DD.
# The VSIX is written by tools/vscode_vsix.py (tmp/vibeoffice-v<version>-<commit>.vsix
# when no path is given).
set -eu
cd "$(dirname "$0")/.."

version=$(python3 -c 'import json; print(json.load(open("package.json", encoding="utf-8"))["version"])')
commit=$(git rev-parse --short HEAD)
branch=$(git branch --show-current)
build_time=$(date "+%Y-%m-%d")

python3 -c '
import json, sys
from pathlib import Path
version, commit, branch, build_time = sys.argv[1:5]
Path("public/version.json").write_text(json.dumps({
    "version": version,
    "commit_id": commit,
    "branch": branch,
    "build_time": build_time,
}, indent=2) + "\n", encoding="utf-8")
' "$version" "$commit" "$branch" "$build_time"

echo "public/version.json: version=$version commit_id=$commit branch=$branch build_time=$build_time"
exec python3 tools/vscode_vsix.py "$@"
