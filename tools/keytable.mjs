// Sinh bảng ánh xạ phím tắt (markdown) từ src/ui/keymap.ts — chạy: node --experimental-strip-types tools/keytable.mjs
const { KEYMAP } = await import('../src/ui/keymap.ts');
const st = { ok: '✅', later: '◐ sau', no: '— không làm' };
let grp = '', out = '';
for (const k of KEYMAP) {
  if (k.group !== grp) { grp = k.group; out += `\n**${grp}**\n\n| Phím | Hành động | Trạng thái |\n|---|---|---|\n`; }
  out += `| ${k.label.replace(/\|/g, '\\|')} | ${k.desc} | ${st[k.status]}${k.note ? ' (' + k.note + ')' : ''} |\n`;
}
const ok = KEYMAP.filter((k) => k.status === 'ok').length, later = KEYMAP.filter((k) => k.status === 'later').length, no = KEYMAP.filter((k) => k.status === 'no').length;
console.log(`Tổng ${KEYMAP.length} mục: ✅ ${ok} · ◐ ${later} · — ${no}\n` + out);
