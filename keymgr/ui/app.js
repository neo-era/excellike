// Giao diện Key Manager (JS thuần, không thư viện). Mọi dữ liệu người dùng đều qua esc()/textContent.
'use strict';
const $ = (s, r = document) => r.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const main = $('#main');
let S = null;              // trạng thái từ tiến trình chính
let page = 'issue';
let lastIssued = null;
let listFilter = 'all', listQuery = '';

const fmt = (iso) => esc(iso ? String(iso).split('-').reverse().join('/') : '—');
const daysTo = (iso) => { if (!iso) return Infinity; const [y, m, d] = iso.split('-').map(Number); const t = new Date(); return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(t.getFullYear(), t.getMonth(), t.getDate())) / 86400000); };
const typeText = (k) => (k.type === 'machine' ? 'Gắn 1 máy' : 'Dùng mọi máy');
const stateTag = (k) => {
  if (k.state === 'revoked') return '<span class="tag revoked">Đã thu hồi</span>';
  if (k.state === 'expired') return '<span class="tag expired">Hết hạn</span>';
  const d = daysTo(k.expires);
  if (d <= 30) return `<span class="tag soon">Còn ${d} ngày</span>`;
  return '<span class="tag active">Hiệu lực</span>';
};

function toast(msg, bad = false) {
  const t = $('#toast');
  t.textContent = msg; t.className = 'toast on' + (bad ? ' bad' : '');
  clearTimeout(toast.h); toast.h = setTimeout(() => { t.className = 'toast' + (bad ? ' bad' : ''); }, 2600);
}
async function refresh() {
  S = await window.km.state();
  $('#nav-count').textContent = S.keys.length || '';
  const ss = $('#sign-state');
  ss.className = 'pill ' + (S.signing ? 'ok' : 'no');
  ss.textContent = S.signing ? '● Đã có khóa ký' : '● Chưa có khóa ký';
  $('#ver').textContent = 'Phiên bản ' + S.version;
}
function go(p) {
  page = p;
  document.querySelectorAll('.side nav button').forEach((b) => b.setAttribute('aria-current', b.dataset.page === p ? 'page' : 'false'));
  render();
  main.focus({ preventScroll: true });
}
document.querySelectorAll('.side nav button').forEach((b) => b.addEventListener('click', () => go(b.dataset.page)));

/** tin nhắn gửi khách (Zalo/email) kèm hướng dẫn nhập key */
function message(k) {
  return `Key bản quyền ExcelLike (số ${k.serial}) cho ${k.customer}${k.org ? ' — ' + k.org : ''}:\n\n${k.key}\n\n`
    + `Loại: ${typeText(k)}${k.machine ? ' (Mã máy ' + k.machine + ')' : ''}. Hạn dùng: ${k.expires ? 'đến hết ' + fmt(k.expires) : 'không thời hạn'}.\n`
    + 'Cách dùng: chạy bộ cài ExcelLike, dán key vào bước "Kích hoạt bản quyền". Nếu đã cài: mở ExcelLike ▸ Tệp ▸ Tài khoản ▸ Đổi key bản quyền.';
}

function render() {
  if (!S) return;
  if (S.dbError && page !== 'settings') { main.innerHTML = '<h1>Sổ key bị hỏng</h1><p class="sub err" id="dberr"></p><button class="btn primary" id="to-set">Mở Cài đặt để khôi phục</button>'; $('#dberr').textContent = S.dbError; $('#to-set').onclick = () => go('settings'); return; }
  if (!S.signing && page !== 'settings') return renderOnboard();
  ({ issue: renderIssue, list: renderList, verify: renderVerify, settings: renderSettings })[page]();
}

// ---------- lần đầu: nhập khóa ký ----------
function renderOnboard() {
  main.innerHTML = `<div class="onboard">
    <div class="big"><svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="8" cy="15" r="4"/><path d="M10.8 12.2 20 3M16 7l3 3M14 9l2 2"/></svg></div>
    <h1>Nhập khóa ký để bắt đầu</h1>
    <p class="sub">Key Manager cần <b>tệp khóa ký ExcelLike</b> (do người phát triển giao riêng) để tạo key mà ExcelLike chấp nhận.</p>
    <div class="drop card"><button class="btn primary" id="imp">Chọn tệp khóa ký…</button><p class="err" id="imp-err"></p></div>
    <ol><li>Khóa ký được mã hóa và chỉ giải mã được trên tài khoản Windows này.</li><li>Ai có tệp khóa ký đều cấp được key — không gửi tệp qua mạng xã hội, xóa bản tải về sau khi nhập.</li><li>Có thể sao lưu lại khóa ký trong Cài đặt.</li></ol>
  </div>`;
  $('#imp').onclick = async () => {
    const r = await window.km.importSigning();
    if (r.canceled) return;
    if (!r.ok) { $('#imp-err').textContent = r.error; return; }
    toast('Đã nhập khóa ký.');
    await refresh(); go('issue');
  };
}

// ---------- cấp key ----------
function renderIssue() {
  main.innerHTML = `<h1>Cấp key mới</h1><p class="sub">Key số <b>${S.nextSerial}</b> sẽ được tạo. Key ký số nên không làm giả được, ExcelLike kiểm ngay trên máy (không cần mạng).</p>
  <div class="grid2">
    <form class="card" id="f" novalidate>
      <label class="fld"><span>Khách hàng / người dùng <i>*</i></span><input type="text" name="customer" id="customer" autocomplete="off" maxlength="120" placeholder="VD: Nguyễn Văn A"></label>
      <div class="row">
        <label class="fld"><span>Đơn vị</span><input type="text" name="org" autocomplete="off" maxlength="160" placeholder="VD: Xí nghiệp Chiếu sáng 1"></label>
        <label class="fld"><span>Liên hệ</span><input type="text" name="contact" autocomplete="off" maxlength="160" placeholder="SĐT hoặc email"></label>
      </div>
      <div class="fld"><span>Loại key</span><div class="seg" role="group" aria-label="Loại key"><button type="button" data-type="machine" aria-pressed="true">Gắn 1 máy</button><button type="button" data-type="shared" aria-pressed="false">Dùng mọi máy</button></div>
        <small id="type-help">Chỉ dùng được trên máy có Mã máy dưới đây (máy hiện Mã máy ở bước cài đặt và trong màn Kích hoạt).</small></div>
      <label class="fld" id="mc-fld"><span>Mã máy <i>*</i></span><input type="text" name="machine" id="machine" class="mono" autocomplete="off" maxlength="24" placeholder="XXXX-XXXX-XXXXX"></label>
      <div class="row">
        <label class="fld"><span>Thời hạn</span><select name="term" id="term"><option value="0">Vĩnh viễn</option><option value="30">30 ngày (dùng thử)</option><option value="90">3 tháng</option><option value="182">6 tháng</option><option value="365">1 năm</option><option value="730">2 năm</option><option value="date">Đến ngày…</option></select></label>
        <label class="fld" id="date-fld" hidden><span>Hết hạn ngày</span><input type="date" name="expires" id="expires"></label>
      </div>
      <label class="fld"><span>Ghi chú</span><textarea name="note" rows="2" maxlength="500" placeholder="VD: hợp đồng số…, máy phòng kỹ thuật"></textarea></label>
      <p class="err" id="err" role="alert"></p>
      <button class="btn primary" type="submit" id="go">Cấp key</button>
    </form>
    <section class="card result" id="res" aria-live="polite"></section>
  </div>`;
  let type = 'machine';
  const setType = (t) => {
    type = t;
    main.querySelectorAll('.seg button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.type === t)));
    $('#mc-fld').hidden = t !== 'machine';
    $('#type-help').textContent = t === 'machine' ? 'Chỉ dùng được trên máy có Mã máy dưới đây (máy hiện Mã máy ở bước cài đặt và trong màn Kích hoạt).' : 'Nhập được trên máy bất kỳ — nên đặt thời hạn để hạn chế chia sẻ key.';
  };
  main.querySelectorAll('.seg button').forEach((b) => b.addEventListener('click', () => setType(b.dataset.type)));
  $('#term').onchange = () => { $('#date-fld').hidden = $('#term').value !== 'date'; };
  $('#machine').addEventListener('input', (e) => { const v = e.target.value.toUpperCase().replace(/[^0-9A-Z]/g, '').slice(0, 13); e.target.value = [v.slice(0, 4), v.slice(4, 8), v.slice(8)].filter(Boolean).join('-'); });
  showResult(lastIssued);
  $('#f').onsubmit = async (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    let expires = null;
    const term = fd.get('term');
    if (term === 'date') expires = fd.get('expires') || '';
    else if (+term > 0) { const d = new Date(); d.setDate(d.getDate() + +term - 1); /* tính cả hôm nay: 30 ngày = hôm nay + 29 */ expires = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
    if (term === 'date' && !expires) { $('#err').textContent = 'Hãy chọn ngày hết hạn.'; return; }
    $('#go').disabled = true; $('#err').textContent = '';
    const r = await window.km.create({ customer: fd.get('customer'), org: fd.get('org'), contact: fd.get('contact'), type, machine: fd.get('machine'), expires, note: fd.get('note') });
    $('#go').disabled = false;
    if (!r.ok) { $('#err').textContent = r.error; return; }
    lastIssued = r.rec;
    await refresh();
    $('#f').reset(); setType(type); $('#date-fld').hidden = true;
    main.querySelector('.sub b').textContent = S.nextSerial;
    showResult(r.rec);
    toast(`Đã cấp key số ${r.rec.serial}.`);
  };
}
function showResult(k) {
  const res = $('#res');
  if (!k) { res.innerHTML = '<div class="empty"><svg viewBox="0 0 24 24" width="40" height="40" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="8" cy="15" r="4"/><path d="M10.8 12.2 20 3M16 7l3 3M14 9l2 2"/></svg><p>Key vừa cấp sẽ hiện ở đây để chép hoặc gửi cho khách.</p></div>'; return; }
  res.innerHTML = `<h2>Key số ${k.serial} — ${esc(k.customer)}</h2><div class="keybox" id="keybox"></div>
    <dl class="meta"><dt>Loại</dt><dd>${typeText(k)}${k.machine ? ' — <code>' + esc(k.machine) + '</code>' : ''}</dd><dt>Ngày cấp</dt><dd>${fmt(k.issued)}</dd><dt>Hết hạn</dt><dd>${k.expires ? fmt(k.expires) : 'Không thời hạn'}</dd>${k.org ? `<dt>Đơn vị</dt><dd>${esc(k.org)}</dd>` : ''}</dl>
    <div class="btns"><button class="btn primary" id="c-key">Chép key</button><button class="btn" id="c-msg">Chép tin nhắn gửi khách</button><button class="btn" id="s-key">Lưu tệp .txt</button></div>`;
  $('#keybox').textContent = k.key;
  $('#c-key').onclick = async () => { await window.km.copy(k.key); toast('Đã chép key.'); };
  $('#c-msg').onclick = async () => { await window.km.copy(message(k)); toast('Đã chép tin nhắn (dán vào Zalo/email).'); };
  $('#s-key').onclick = async () => { const r = await window.km.saveText(`ExcelLike-key-${k.serial}.txt`, message(k)); if (r.ok) toast('Đã lưu ' + r.path); };
}

// ---------- sổ key ----------
function renderList() {
  const ks = S.keys;
  const c = { all: ks.length, active: ks.filter((k) => k.state === 'active').length, soon: ks.filter((k) => k.state === 'active' && daysTo(k.expires) <= 30).length, expired: ks.filter((k) => k.state === 'expired').length, revoked: ks.filter((k) => k.state === 'revoked').length };
  const lab = { all: 'Tổng số key', active: 'Đang hiệu lực', soon: 'Sắp hết hạn (≤ 30 ngày)', expired: 'Hết hạn', revoked: 'Đã thu hồi' };
  main.innerHTML = `<h1>Sổ key</h1><p class="sub">Mọi key đã cấp, lưu trên máy này (tự sao lưu hằng ngày).</p>
    <div class="stats">${Object.keys(lab).map((f) => `<button class="stat" data-f="${f}" aria-pressed="${listFilter === f}"><b>${c[f]}</b><span>${lab[f]}</span></button>`).join('')}</div>
    <div class="toolbar"><input type="search" id="q" placeholder="Tìm theo tên, đơn vị, mã máy, số key…" aria-label="Tìm key"><span class="sp"></span><button class="btn" id="csv">Xuất CSV</button><button class="btn" id="rv">Danh sách thu hồi</button></div>
    <div class="tbl-wrap"><table><thead><tr><th>Số</th><th>Khách hàng</th><th>Loại</th><th>Mã máy</th><th>Ngày cấp</th><th>Hết hạn</th><th>Trạng thái</th></tr></thead><tbody id="tb"></tbody></table></div>`;
  const q = $('#q'); q.value = listQuery;
  const draw = () => {
    const t = listQuery.trim().toLowerCase();
    const rows = ks.filter((k) => (listFilter === 'all' || (listFilter === 'soon' ? k.state === 'active' && daysTo(k.expires) <= 30 : k.state === listFilter))
      && (!t || [k.serial, k.customer, k.org, k.contact, k.machine, k.note].join(' ').toLowerCase().includes(t))).slice().reverse();
    $('#tb').innerHTML = rows.length ? rows.map((k) => `<tr data-s="${k.serial}" tabindex="0"><td><b>${k.serial}</b></td><td class="cust">${esc(k.customer)}${k.org ? `<small>${esc(k.org)}</small>` : ''}</td><td>${typeText(k)}</td><td><code>${esc(k.machine || '—')}</code></td><td>${fmt(k.issued)}</td><td>${k.expires ? fmt(k.expires) : 'Vĩnh viễn'}</td><td>${stateTag(k)}</td></tr>`).join('')
      : `<tr><td colspan="7" class="empty">${ks.length ? 'Không có key phù hợp.' : 'Chưa cấp key nào. Vào “Cấp key” để tạo key đầu tiên.'}</td></tr>`;
  };
  draw();
  q.oninput = () => { listQuery = q.value; draw(); };
  main.querySelectorAll('.stat').forEach((b) => b.onclick = () => { listFilter = b.dataset.f; renderList(); });
  const open = (e) => { const tr = e.target.closest('tr[data-s]'); if (tr) detail(+tr.dataset.s); };
  $('#tb').onclick = open;
  $('#tb').onkeydown = (e) => { if (e.key === 'Enter') open(e); };
  $('#csv').onclick = async () => {
    const head = ['Số', 'Khách hàng', 'Đơn vị', 'Liên hệ', 'Loại', 'Mã máy', 'Ngày cấp', 'Hết hạn', 'Trạng thái', 'Ghi chú', 'Key'];
    const cell = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const st = { active: 'Hiệu lực', expired: 'Hết hạn', revoked: 'Đã thu hồi' };
    const csv = '﻿' + [head, ...ks.map((k) => [k.serial, k.customer, k.org, k.contact, typeText(k), k.machine, fmt(k.issued), k.expires ? fmt(k.expires) : 'Vĩnh viễn', st[k.state], k.note, k.key])].map((r) => r.map(cell).join(';')).join('\r\n');
    const r = await window.km.saveText(`So-key-ExcelLike-${new Date().toISOString().slice(0, 10)}.csv`, csv);
    if (r.ok) toast('Đã xuất ' + r.path);
  };
  $('#rv').onclick = async () => {
    const list = ks.filter((k) => k.status === 'revoked').map((k) => k.serial);
    await window.km.copy(`REVOKED: [${list.join(', ')}],`);
    toast(list.length ? `Đã chép ${list.length} số key thu hồi — gửi cho người phát triển để đưa vào bản cài kế tiếp.` : 'Chưa có key nào bị thu hồi.');
  };
}
function detail(serial) {
  const k = S.keys.find((x) => x.serial === serial);
  if (!k) return;
  const back = document.createElement('div');
  back.className = 'drawer-back';
  back.innerHTML = `<div class="drawer" role="dialog" aria-modal="true" aria-label="Chi tiết key"><header><h2>Key số ${k.serial}</h2>${stateTag(k)}<button class="x" aria-label="Đóng">✕</button></header>
    <div class="keybox" id="d-key"></div>
    <div class="btns"><button class="btn primary" id="d-ck">Chép key</button><button class="btn" id="d-cm">Chép tin nhắn</button></div>
    <dl class="meta"><dt>Loại</dt><dd>${typeText(k)}${k.machine ? ' — <code>' + esc(k.machine) + '</code>' : ''}</dd><dt>Ngày cấp</dt><dd>${fmt(k.issued)}</dd><dt>Hết hạn</dt><dd>${k.expires ? fmt(k.expires) : 'Không thời hạn'}</dd>${k.revokedAt ? `<dt>Thu hồi lúc</dt><dd>${esc(new Date(k.revokedAt).toLocaleString('vi-VN'))}</dd>` : ''}</dl>
    <label class="fld"><span>Khách hàng</span><input type="text" id="d-customer" maxlength="120"></label>
    <label class="fld"><span>Đơn vị</span><input type="text" id="d-org" maxlength="160"></label>
    <label class="fld"><span>Liên hệ</span><input type="text" id="d-contact" maxlength="160"></label>
    <label class="fld"><span>Ghi chú</span><textarea id="d-note" rows="3" maxlength="500"></textarea></label>
    <div class="btns"><button class="btn primary" id="d-save">Lưu thay đổi</button><button class="btn danger" id="d-rv">${k.status === 'revoked' ? 'Bỏ thu hồi' : 'Thu hồi key'}</button></div>
    <p class="warnbox">Thu hồi: key đã nhập trên máy khách vẫn chạy tới khi họ cài bản ExcelLike mới có danh sách thu hồi (key được kiểm ngay trên máy, không cần mạng).</p>
  </div>`;
  document.body.appendChild(back);
  $('#d-key', back).textContent = k.key;
  for (const f of ['customer', 'org', 'contact', 'note']) $('#d-' + f, back).value = k[f] || '';
  const close = () => { back.remove(); document.removeEventListener('keydown', onKey); };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', onKey);
  back.addEventListener('click', (e) => { if (e.target === back) close(); });
  $('.x', back).onclick = close;
  $('#d-ck', back).onclick = async () => { await window.km.copy(k.key); toast('Đã chép key.'); };
  $('#d-cm', back).onclick = async () => { await window.km.copy(message(k)); toast('Đã chép tin nhắn.'); };
  $('#d-save', back).onclick = async () => {
    const patch = {}; for (const f of ['customer', 'org', 'contact', 'note']) patch[f] = $('#d-' + f, back).value;
    if (!patch.customer.trim()) { toast('Tên khách hàng không được trống.', true); return; }
    await window.km.update(k.serial, patch); await refresh(); close(); renderList(); toast('Đã lưu.');
  };
  $('#d-rv', back).onclick = async () => {
    if (k.status !== 'revoked' && !confirm(`Thu hồi key số ${k.serial} của ${k.customer}?`)) return;
    await window.km.update(k.serial, { status: k.status === 'revoked' ? 'active' : 'revoked' }); await refresh(); close(); renderList();
    toast(k.status === 'revoked' ? 'Đã bỏ thu hồi.' : 'Đã thu hồi key.');
  };
  $('#d-customer', back).focus();
}

// ---------- kiểm tra key ----------
function renderVerify() {
  main.innerHTML = `<h1>Kiểm tra key</h1><p class="sub">Dán key khách gửi lại để xem thông tin, hạn dùng và key có do Key Manager này cấp không.</p>
    <div class="grid2"><div class="card">
      <label class="fld"><span>Key</span><textarea id="v-key" class="mono" rows="4" placeholder="XXXXX-XXXXX-…"></textarea></label>
      <label class="fld"><span>Mã máy (không bắt buộc)</span><input type="text" id="v-mc" class="mono" maxlength="24" placeholder="Kiểm key có dùng được trên máy này không"></label>
      <button class="btn primary" id="v-go">Kiểm tra</button></div>
      <section class="card result" id="v-res" aria-live="polite"><div class="empty"><p>Kết quả hiện ở đây.</p></div></section></div>`;
  $('#v-go').onclick = async () => {
    const r = await window.km.verify($('#v-key').value, $('#v-mc').value.trim());
    const i = r.info;
    $('#v-res').innerHTML = `<h2>${r.ok ? '<span class="tag active">Hợp lệ</span>' : '<span class="tag revoked">Không dùng được</span>'}</h2><p class="v-msg"></p>
      ${i ? `<dl class="meta"><dt>Số key</dt><dd>${i.serial}</dd><dt>Loại</dt><dd>${i.type === 'machine' ? 'Gắn 1 máy — <code>' + esc(i.machine) + '</code>' : 'Dùng mọi máy'}</dd><dt>Ngày cấp</dt><dd>${esc(i.issued)}</dd><dt>Hết hạn</dt><dd>${i.expires ? esc(i.expires) : 'Không thời hạn'}</dd>${r.rec ? `<dt>Khách hàng</dt><dd>${esc(r.rec.customer)}${r.rec.org ? ' — ' + esc(r.rec.org) : ''}</dd><dt>Trong sổ</dt><dd>${stateTag(r.rec)}</dd>` : '<dt>Trong sổ</dt><dd>Không có trong sổ key của máy này</dd>'}</dl>` : ''}`;
    $('.v-msg', $('#v-res')).textContent = r.rec?.state === 'revoked' ? 'Key đã bị thu hồi trong sổ.' : r.message;
  };
}

// ---------- cài đặt ----------
function renderSettings() {
  const fp = S.pub ? S.pub.slice(0, 12) + '…' + S.pub.slice(-6) : '';
  main.innerHTML = `<h1>Cài đặt</h1><p class="sub">Khóa ký, sao lưu và dữ liệu.</p>
    <div class="card">
      <div class="set-row"><div><b>Khóa ký</b><small>${S.signing ? 'Đã nhập. Dấu vân tay: <span class="fp">' + esc(fp) + '</span>' : 'Chưa nhập — chưa cấp được key.'}</small></div><button class="btn${S.signing ? '' : ' primary'}" id="s-imp">${S.signing ? 'Nhập lại tệp khóa ký…' : 'Nhập tệp khóa ký…'}</button>${S.signing ? '<button class="btn" id="s-exp">Sao lưu khóa ký…</button>' : ''}</div>
      <div class="set-row"><div><b>Sao lưu sổ key</b><small>Lưu toàn bộ sổ key ra tệp .json (cất ở ổ khác hoặc USB).</small></div><button class="btn" id="s-bak">Sao lưu…</button></div>
      <div class="set-row"><div><b>Khôi phục sổ key</b><small>Gộp key từ tệp sao lưu vào sổ hiện tại (không ghi đè key đang có).</small></div><button class="btn" id="s-res">Khôi phục…</button></div>
      <div class="set-row"><div><b>Thư mục dữ liệu</b><small class="fp" id="s-dir"></small></div><button class="btn" id="s-open">Mở thư mục</button></div>
    </div>
    <p class="warnbox">Khóa ký là bí mật: ai có nó đều cấp được key. Mất khóa ký thì không cấp thêm key được cho các bản ExcelLike đã phát hành — hãy sao lưu ra USB.</p>`;
  $('#s-dir').textContent = S.dataDir;
  $('#s-imp').onclick = async () => { const r = await window.km.importSigning(); if (r.canceled) return; if (!r.ok) return toast(r.error, true); toast('Đã nhập khóa ký.'); await refresh(); renderSettings(); };
  $('#s-exp') && ($('#s-exp').onclick = async () => { const r = await window.km.exportSigning(); if (r.ok) toast('Đã sao lưu khóa ký: ' + r.path); });
  $('#s-bak').onclick = async () => { const r = await window.km.saveText(`So-key-ExcelLike-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify({ version: 1, nextSerial: S.nextSerial, keys: S.keys.map(({ state, ...k }) => k) }, null, 1)); if (r.ok) toast('Đã sao lưu: ' + r.path); };
  $('#s-res').onclick = async () => { const r = await window.km.restore(); if (r.canceled) return; if (!r.ok) return toast(r.error, true); await refresh(); toast(`Đã khôi phục ${r.added} key.` + (r.bad ? ` Bỏ qua ${r.bad} bản ghi lỗi.` : '') + (r.conflict ? ` ${r.conflict} key trùng số khác nội dung (giữ bản đang có).` : ''), !!(r.bad || r.conflict)); };
  $('#s-open').onclick = () => window.km.openData();
}

refresh().then(render);
