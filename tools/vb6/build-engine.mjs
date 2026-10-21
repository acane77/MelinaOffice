/** Rebuild the gitignored Quire VB6 engine from third_party/VB6.
 *  make
 *  Writes public/quire/vendor/vb6/vb6-engine.js (global VB6Engine). */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { bundle } from '../../third_party/VB6/tools/bundle.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

export function buildEngine() {
  const entry = path.join(root, 'third_party/VB6/src/runtime/engine-entry.js');
  const outDir = path.join(root, 'public/quire/vendor/vb6');
  const outFile = path.join(outDir, 'vb6-engine.js');
  fs.mkdirSync(outDir, { recursive: true });
  const code = '/* Quire headless VB6 engine. Built by tools/vb6/build-engine.mjs. Not ApplicationHost. */\n'
    + bundle(entry, 'VB6Engine');
  const tmp = outFile + '.' + process.pid + '.tmp';
  fs.writeFileSync(tmp, code);
  fs.renameSync(tmp, outFile);
  return outFile;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const out = buildEngine();
  console.log(out + ' ' + fs.statSync(out).size);
}
