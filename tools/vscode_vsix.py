#!/usr/bin/env python3
"""Pack a VSIX of this repository so code-server can install it.

    python3 tools/vscode_vsix.py [out.vsix]

The archive holds package.json, vscode/ and public/ (the pages the webview loads).
Install:  code-server --install-extension vibeoffice-1.0.1.vsix
"""
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKIP = {'.git', 'node_modules', '__pycache__', 'target', 'gen'}


def main():
    pkg = json.loads((ROOT / 'package.json').read_text(encoding='utf-8'))
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'tmp' / f"vibeoffice-{pkg['version']}.vsix"
    out.parent.mkdir(parents=True, exist_ok=True)
    ident = pkg['publisher'], pkg['name'], pkg['version']
    manifest = f'''<?xml version="1.0" encoding="utf-8"?>
<PackageManifest Version="2.0.0" xmlns="http://schemas.microsoft.com/developer/vsx-schema/2011">
  <Metadata>
    <Identity Language="en-US" Id="{ident[1]}" Version="{ident[2]}" Publisher="{ident[0]}"/>
    <DisplayName>{pkg["displayName"]}</DisplayName>
    <Description xml:space="preserve">{pkg["description"]}</Description>
  </Metadata>
  <Installation>
    <InstallationTarget Id="Microsoft.VisualStudio.Code"/>
  </Installation>
  <Dependencies/>
  <Assets>
    <Asset Type="Microsoft.VisualStudio.Code.Manifest" Path="extension/package.json" Addressable="true"/>
  </Assets>
</PackageManifest>
'''
    types = '''<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="json" ContentType="application/json"/>
  <Default Extension="js" ContentType="application/javascript"/>
  <Default Extension="html" ContentType="text/html"/>
  <Default Extension="css" ContentType="text/css"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="woff2" ContentType="font/woff2"/>
  <Default Extension="txt" ContentType="text/plain"/>
  <Default Extension="md" ContentType="text/markdown"/>
  <Default Extension="vsixmanifest" ContentType="text/xml"/>
  <Default Extension="xml" ContentType="text/xml"/>
  <Default Extension="words" ContentType="application/octet-stream"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>
'''
    n = 0
    with zipfile.ZipFile(out, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr('extension.vsixmanifest', manifest)
        z.writestr('[Content_Types].xml', types)
        for rel in ('package.json', 'README.md', 'LICENSE'):
            p = ROOT / rel
            if p.is_file():
                z.write(p, 'extension/' + rel)
                n += 1
        for folder in ('vscode', 'public'):
            for p in (ROOT / folder).rglob('*'):
                if not p.is_file() or SKIP.intersection(p.parts):
                    continue
                z.write(p, 'extension/' + p.relative_to(ROOT).as_posix())
                n += 1
    print(f'{out} ({n} files)')


if __name__ == '__main__':
    main()
