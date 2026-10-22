/** Rebuild the gitignored VB6 engine from third_party/VB6.
 *  make
 *  Writes public/quire/vendor/vb6/vb6-engine.js and the same file under public/lectern/vendor/vb6/ and public/ledger/vendor/vb6/ (global VB6Engine). */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { bundle } from '../../third_party/VB6/tools/bundle.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

function writeBundle(dir, code) {
  fs.mkdirSync(dir, { recursive: true });
  const outFile = path.join(dir, 'vb6-engine.js');
  const tmp = outFile + '.' + process.pid + '.tmp';
  fs.writeFileSync(tmp, code);
  fs.renameSync(tmp, outFile);
  return outFile;
}

export function buildEngine() {
  const entry = path.join(root, 'third_party/VB6/src/runtime/engine-entry.js');
  const code = '/* Headless VB6 engine. Built by tools/vb6/build-engine.mjs. Not ApplicationHost. */\n'
    + bundle(entry, 'VB6Engine');
  writeBundle(path.join(root, 'public/lectern/vendor/vb6'), code);
  writeBundle(path.join(root, 'public/ledger/vendor/vb6'), code);
  return writeBundle(path.join(root, 'public/quire/vendor/vb6'), code);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const out = buildEngine();
  console.log(out + ' ' + fs.statSync(out).size);
}
