// Biên dịch build/lic/lic.exe (công cụ kiểm key cho bộ cài) với khóa công khai và danh sách thu hồi hiện tại
import { execFileSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join } from 'node:path';
const ROOT = new URL('..', import.meta.url).pathname;
const require = createRequire(import.meta.url);
const { PUBLIC_KEY, REVOKED } = require(join(ROOT, 'electron/license-config.cjs'));
const D = join(ROOT, 'build/lic');
const pk = Buffer.from(PUBLIC_KEY, 'base64');
if (pk.length !== 32) throw new Error('PUBLIC_KEY phải là 32 byte (base64)');
writeFileSync(join(D, 'pubkey.h'), `static const unsigned char PUBKEY[32] = { ${[...pk].join(', ')} };\n`);
const rv = (REVOKED || []).filter((n) => Number.isInteger(n) && n > 0);
writeFileSync(join(D, 'revoked.h'), `static const unsigned int REVOKED[] = { ${rv.length ? rv.join(', ') : '0'} };\nstatic const unsigned REVOKED_N = ${rv.length};\n`);
execFileSync('x86_64-w64-mingw32-gcc', ['-O2', '-s', '-municode', '-o', join(D, 'lic.exe'), join(D, 'lic.c'), join(D, 'tweetnacl.c'), '-ladvapi32'], { stdio: 'inherit' });
console.log('  • build/lic/lic.exe');
