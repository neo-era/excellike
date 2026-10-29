// Đóng gói app Quản lý Key cho Windows: electron-builder (thư mục) → NSIS
import { execFileSync } from 'node:child_process';
import { copyFileSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
const ROOT = new URL('../keymgr/', import.meta.url).pathname;
const { version } = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'));
const run = (c, a) => execFileSync(c, a, { cwd: ROOT, stdio: 'inherit' });
// lõi key dùng chung với ExcelLike (một nguồn duy nhất)
copyFileSync(join(ROOT, '../electron/license.cjs'), join(ROOT, 'electron/license.cjs'));
rmSync(join(ROOT, 'release'), { recursive: true, force: true });
run('npx', ['electron-builder', '--win', 'dir', '--x64', '--projectDir', ROOT]);
const SRC = join(ROOT, 'release/win-unpacked');
const files = [], dirs = [];
(function walk(d) { for (const n of readdirSync(d)) { const p = join(d, n); if (statSync(p).isDirectory()) { walk(p); dirs.push(p); } else files.push(p); } })(SRC);
const win = (p) => relative(SRC, p).split('/').join('\\');
writeFileSync(join(ROOT, 'build/uninst-files.nsh'), [...files.map((f) => `  Delete "$INSTDIR\\${win(f)}"`), ...dirs.sort((a, b) => b.length - a.length).map((d) => `  RMDir "$INSTDIR\\${win(d)}"`)].join('\n') + '\n');
const out = join(ROOT, `release/ExcelLike-KeyManager-Setup-${version}.exe`);
run('makensis', ['-V2', `-DVERSION=${version}`, `-DSRC=${SRC}`, `-DOUT=${out}`, join(ROOT, 'build/installer.nsi')]);
console.log(`  • ${relative(ROOT, out)}  ${(statSync(out).size / 1048576).toFixed(1)} MB`);
