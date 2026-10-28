// Đóng gói bản Windows: electron-builder (thư mục chương trình) → NSIS (bộ cài, chạy makensis gốc, không cần Wine) + bản portable .zip
import { execFileSync } from 'node:child_process';
import { readdirSync, statSync, writeFileSync, readFileSync, rmSync, existsSync } from 'node:fs';
import { join, relative } from 'node:path';

const ROOT = new URL('..', import.meta.url).pathname;
const { version } = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'));
const run = (cmd, args, opts = {}) => execFileSync(cmd, args, { cwd: ROOT, stdio: 'inherit', ...opts });

rmSync(join(ROOT, 'release'), { recursive: true, force: true });
run('npx', ['electron-builder', '--win', 'dir', '--x64']);
const SRC = join(ROOT, 'release/win-unpacked');

// danh sách tệp để gỡ cài đặt (tệp trước, thư mục sâu nhất trước)
const files = [], dirs = [];
(function walk(d) {
  for (const n of readdirSync(d)) {
    const p = join(d, n);
    if (statSync(p).isDirectory()) { walk(p); dirs.push(p); } else files.push(p);
  }
})(SRC);
const win = (p) => relative(SRC, p).split('/').join('\\');
writeFileSync(join(ROOT, 'build/uninst-files.nsh'), [
  ...files.map((f) => `  Delete "$INSTDIR\\${win(f)}"`),
  ...dirs.sort((a, b) => b.length - a.length).map((d) => `  RMDir "$INSTDIR\\${win(d)}"`),
].join('\n') + '\n');

const out = join(ROOT, `release/ExcelLike-Setup-${version}.exe`);
run('makensis', ['-V2', `-DVERSION=${version}`, `-DSRC=${SRC}`, `-DOUT=${out}`, join(ROOT, 'build/installer.nsi')]);

const zip = join(ROOT, `release/ExcelLike-Portable-${version}.zip`);
if (existsSync(zip)) rmSync(zip);
run('bash', ['-c', `cd release && rm -rf ExcelLike && cp -r win-unpacked ExcelLike && zip -qr -9 "${zip}" ExcelLike && rm -rf ExcelLike`]);
for (const f of [out, zip]) console.log(`  • ${relative(ROOT, f)}  ${(statSync(f).size / 1048576).toFixed(1)} MB`);
