// Sinh key thật cho kiểm thử bản máy tính (cần tệp khóa ký — không có trong repo)
import { createRequire } from 'node:module';
import { existsSync, readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
const require = createRequire(import.meta.url);
export const L = require('../electron/license.cjs');
const F = process.env.XL_SIGNING_FILE || [join(homedir(), '.excellike-signing/signing.json'), '/home/claude/.excellike-signing/signing.json'].find(existsSync) || join(homedir(), '.excellike-signing/signing.json');
export const signing = existsSync(F) ? JSON.parse(readFileSync(F, 'utf8')) : null;
export const TEST_GUID = '11111111-2222-3333-4444-555555555555';
export const testKey = (o = {}) => {
  if (!signing) throw new Error('Thiếu tệp khóa ký ' + F + ' — không sinh được key kiểm thử');
  return L.makeKey({ serial: 9000, type: 'shared', ...o }, signing.priv);
};
