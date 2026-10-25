#!/usr/bin/env python3
"""Assemble public/matheuse/index.html: a full document with Matheuse's CSS inlined (after the shared
public/common/ui.css) and the scripts in load order.

    python3 tools/matheuse/build/make.py
"""
import os
D = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(D, '..', '..', '..', 'public', 'matheuse')
# load order; 'common/x' is public/common/x.js, shared by the suite, anything else public/matheuse/js/x.js
SCRIPTS = ['common/core', 'common/i18n', 'common/locales', 'common/xml', 'common/opc', 'common/opc-order', 'xml', 'common/zip', 'common/sha', 'common/crypto', 'common/icons',
           'common/geometry', 'common/metafile', 'common/ui', 'common/macros', 'common/numfmt', 'formula', 'model', 'calc',
           'fn-core', 'fn-lookup', 'fn-stat', 'fn-fin', 'fn-eng', 'common/dml', 'common/charts', 'xchart',
           'pivots', 'threads', 'objects', 'slicers', 'extensions', 'tables', 'preserve', 'xlsx-read', 'xlsx-write', 'csv', 'xmlss', 'styles', 'cf', 'layout', 'render', 'ops', 'clipboard',
           'fninfo', 'grid', 'editor', 'common/spell', 'spell', 'filter', 'drawing', 'commands', 'dialogs', 'common/host-audit',
           'dialogs2', 'dialogs3', 'print', 'panes', 'common/vb-dom-ref', 'vb-dom-ref', 'common/scout/spritepet', 'common/scout/assistant', 'common/search', 'common/ribbon', 'ribbon', 'app',
           'dom', 'dom-ir', 'common/vb-forms', 'vb-host', 'macros', 'common/vendor/xterm/xterm', 'common/vendor/codemirror/codemirror', 'common/vendor/codemirror/vb', 'common/vb-complete', 'common/vb-ide', 'dom-console']
tpl = open(os.path.join(D, 'template.html')).read()
css = open(os.path.join(D, 'matheuse.css')).read()   # after public/common/ui.css, linked by the template
def src(s):
    rel = '../' + s + '.js' if s.startswith('common/') else 'js/' + s + '.js'
    if not os.path.exists(os.path.join(APP, rel)):
        raise SystemExit(f'missing script: {os.path.relpath(os.path.join(APP, rel))}')
    return rel
tags = '\n'.join(f'<script src="{src(s)}"></script>' for s in SCRIPTS)
body = tpl.replace('/*CSS*/', css).replace('<!--SCRIPTS-->', tags)
out = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
       '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n' + body + '\n</html>\n')
open(os.path.join(APP, 'index.html'), 'w').write(out)
print('public/matheuse/index.html', len(out), 'bytes;', tags.count('<script'), 'scripts')
