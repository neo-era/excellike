// Giao diện Key Manager — bản iPhone/web (cùng giao diện bản Windows, thêm khóa bằng mật khẩu, giao diện điện thoại).
// Mọi dữ liệu người dùng đều qua esc()/textContent.
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
  ss.textContent = S.signing ? '● Đã mở khóa' : S.sealed ? '● Đang khóa' : '● Chưa có khóa ký';
  $('#ver').textContent = 'Bản iPhone ' + S.version;
  $('#lock-btn').hidden = !S.signing; $('#lock-btn2').hidden = !S.signing;
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
  if (!S.sealed) return renderOnboard();
  if (!S.signing) return renderLock();
  ({ issue: renderIssue, list: renderList, verify: renderVerify, settings: renderSettings })[page]();
}

// ---------- lần đầu: nhập khóa ký + đặt mật khẩu ----------
function renderOnboard() {
  main.innerHTML = `<div class="onboard">
    <div class="big"><svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="8" cy="15" r="4"/><path d="M10.8 12.2 20 3M16 7l3 3M14 9l2 2"/></svg></div>
    <h1>Nhập khóa ký để bắt đầu</h1>
    <p class="sub">Chọn tệp <b>ExcelLike-khoa-ky.json</b> (trong Tệp ▸ OneDrive ▸ Kho_Software_Driver ▸ ExcelLike ▸ Key) rồi đặt mật khẩu mở app trên iPhone này.</p>
    <form class="card" id="ob" novalidate>
      <div class="fld"><span>Tệp khóa ký <i>*</i></span><div class="btns"><button type="button" class="btn" id="ob-file">Chọn tệp…</button><span id="ob-name" class="muted"></span></div>
        <details><summary>Hoặc dán nội dung tệp</summary><textarea id="ob-text" class="mono" rows="4" placeholder='{"product":"ExcelLike", … "priv": "…"}'></textarea></details></div>
      <label class="fld"><span>Mật khẩu mở app <i>*</i></span><input type="password" id="ob-p1" autocomplete="new-password" minlength="6" placeholder="Tối thiểu 6 ký tự"></label>
      <label class="fld"><span>Nhập lại mật khẩu <i>*</i></span><input type="password" id="ob-p2" autocomplete="new-password"></label>
      <p class="err" id="imp-err" role="alert"></p>
      <button class="btn primary" type="submit" id="ob-go">Nhập khóa ký</button>
    </form>
    <ol><li>Khóa ký được mã hóa bằng mật khẩu, chỉ nằm trên iPhone này, không gửi đi đâu.</li><li>App tự khóa sau 10 phút không dùng hoặc khi rời app hơn 2 phút.</li><li>Quên mật khẩu: nhập lại tệp khóa ký, sổ key vẫn giữ nguyên.</li></ol>
  </div>`;
  let fileText = null;
  $('#ob-file').onclick = async () => { const t = await window.km.pickSigningFile(); if (t !== null) { fileText = t; $('#ob-name').textContent = '✓ Đã chọn tệp'; } };
  $('#ob').onsubmit = async (e) => {
    e.preventDefault();
    const text = fileText ?? $('#ob-text').value.trim();
    if (!text) { $('#imp-err').textContent = 'Hãy chọn tệp khóa ký (hoặc dán nội dung tệp).'; return; }
    if ($('#ob-p1').value !== $('#ob-p2').value) { $('#imp-err').textContent = 'Hai mật khẩu không khớp.'; return; }
    $('#ob-go').disabled = true; $('#imp-err').textContent = 'Đang mã hóa…';
    const r = await window.km.importSigning(text, $('#ob-p1').value);
    $('#ob-go').disabled = false;
    if (!r.ok) { $('#imp-err').textContent = r.error; return; }
    $('#imp-err').textContent = '';
    toast('Đã nhập khóa ký.');
    await refresh(); go('issue');
  };
}
// ---------- màn khóa ----------
function renderLock() {
  main.innerHTML = `<div class="onboard lock">
    <div class="big"><svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg></div>
    <h1>Key Manager đang khóa</h1><p class="sub">Nhập mật khẩu để cấp key và xem sổ key.</p>
    <form class="card" id="lk" novalidate>
      <label class="fld"><span>Mật khẩu</span><input type="password" id="lk-p" autocomplete="current-password" autofocus></label>
      <p class="err" id="lk-err" role="alert"></p>
      <button class="btn primary" type="submit" id="lk-go">Mở khóa</button>
    </form>
    <p class="muted center"><button class="linkbtn" id="lk-re">Quên mật khẩu? Nhập lại tệp khóa ký</button></p>
  </div>`;
  $('#lk').onsubmit = async (e) => {
    e.preventDefault();
    $('#lk-go').disabled = true; $('#lk-err').textContent = 'Đang mở…';
    const r = await window.km.unlock($('#lk-p').value);
    $('#lk-go').disabled = false;
    if (!r.ok) { $('#lk-err').textContent = r.error; $('#lk-p').select(); return; }
    await refresh(); render();
  };
  $('#lk-re').onclick = () => renderReimport();
  setTimeout(() => $('#lk-p')?.focus(), 50);
}
function renderReimport() {
  renderOnboard();
  $('.onboard h1').textContent = 'Nhập lại tệp khóa ký';
  $('.onboard .sub').textContent = 'Sổ key trên iPhone được giữ nguyên. Đặt mật khẩu mới cho app.';
}

// ---------- cấp key ----------
function renderIssue() {
  main.innerHTML = `<h1>Cấp key mới</h1><p class="sub">Key số <b>${S.nextSerial}</b> sẽ được tạo (key cấp trên iPhone đánh số từ ${window.km.RANGE.toLocaleString('vi-VN')} để không trùng với bản Windows).</p>
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
    if (matchMedia('(max-width: 760px), (max-height: 520px)').matches) $('#res').scrollIntoView({ behavior: 'smooth', block: 'start' });
    const lb = S.lastBackup ? (Date.now() - Date.parse(S.lastBackup)) / 86400000 : Infinity;
    toast(lb > 7 ? `Đã cấp key số ${r.rec.serial}. Nhớ sao lưu sổ key (Cài đặt).` : `Đã cấp key số ${r.rec.serial}.`);
  };
}
function showResult(k) {
  const res = $('#res');
  if (!k) { res.innerHTML = '<div class="empty"><svg viewBox="0 0 24 24" width="40" height="40" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="8" cy="15" r="4"/><path d="M10.8 12.2 20 3M16 7l3 3M14 9l2 2"/></svg><p>Key vừa cấp sẽ hiện ở đây để chép hoặc gửi cho khách.</p></div>'; return; }
  res.innerHTML = `<h2>Key số ${k.serial} — ${esc(k.customer)}</h2><div class="keybox" id="keybox"></div>
    <dl class="meta"><dt>Loại</dt><dd>${typeText(k)}${k.machine ? ' — <code>' + esc(k.machine) + '</code>' : ''}</dd><dt>Ngày cấp</dt><dd>${fmt(k.issued)}</dd><dt>Hết hạn</dt><dd>${k.expires ? fmt(k.expires) : 'Không thời hạn'}</dd>${k.org ? `<dt>Đơn vị</dt><dd>${esc(k.org)}</dd>` : ''}</dl>
    <div class="btns">${navigator.share ? '<button class="btn primary" id="sh-msg">Gửi cho khách…</button>' : ''}<button class="btn${navigator.share ? '' : ' primary'}" id="c-key">Chép key</button><button class="btn" id="c-msg">Chép tin nhắn gửi khách</button><button class="btn" id="s-key">Lưu tệp .txt</button></div>`;
  $('#keybox').textContent = k.key;
  $('#c-key').onclick = async () => { await window.km.copy(k.key); toast('Đã chép key.'); };
  $('#sh-msg') && ($('#sh-msg').onclick = () => navigator.share({ text: message(k) }).catch(() => {}));
  $('#c-msg').onclick = async () => { await window.km.copy(message(k)); toast('Đã chép tin nhắn (dán vào Zalo/email).'); };
  $('#s-key').onclick = async () => { const r = await window.km.saveText(`ExcelLike-key-${k.serial}.txt`, message(k)); if (r.ok) toast('Đã lưu ' + r.path); };
}

// ---------- sổ key ----------
function renderList() {
  const ks = S.keys;
  const c = { all: ks.length, active: ks.filter((k) => k.state === 'active').length, soon: ks.filter((k) => k.state === 'active' && daysTo(k.expires) <= 30).length, expired: ks.filter((k) => k.state === 'expired').length, revoked: ks.filter((k) => k.state === 'revoked').length };
  const lab = { all: 'Tổng số key', active: 'Đang hiệu lực', soon: 'Sắp hết hạn (≤ 30 ngày)', expired: 'Hết hạn', revoked: 'Đã thu hồi' };
  main.innerHTML = `<h1>Sổ key</h1><p class="sub">Key đã cấp trên iPhone này và key nhập từ bản Windows (Cài đặt ▸ Khôi phục sổ key).</p>
    <div class="stats">${Object.keys(lab).map((f) => `<button class="stat" data-f="${f}" aria-pressed="${listFilter === f}"><b>${c[f]}</b><span>${lab[f]}</span></button>`).join('')}</div>
    <div class="toolbar"><input type="search" id="q" placeholder="Tìm theo tên, đơn vị, mã máy, số key…" aria-label="Tìm key"><span class="sp"></span><button class="btn" id="csv">Xuất CSV</button><button class="btn" id="rv">Danh sách thu hồi</button></div>
    <div class="tbl-wrap"><table><thead><tr><th>Số</th><th>Khách hàng</th><th>Loại</th><th>Mã máy</th><th>Ngày cấp</th><th>Hết hạn</th><th>Trạng thái</th></tr></thead><tbody id="tb"></tbody></table></div><div class="cards" id="cards"></div>`;
  const q = $('#q'); q.value = listQuery;
  const draw = () => {
    const t = listQuery.trim().toLowerCase();
    const rows = ks.filter((k) => (listFilter === 'all' || (listFilter === 'soon' ? k.state === 'active' && daysTo(k.expires) <= 30 : k.state === listFilter))
      && (!t || [k.serial, k.customer, k.org, k.contact, k.machine, k.note].join(' ').toLowerCase().includes(t))).slice().reverse();
    $('#tb').innerHTML = rows.length ? rows.map((k) => `<tr data-s="${k.serial}" tabindex="0"><td><b>${k.serial}</b></td><td class="cust">${esc(k.customer)}${k.org ? `<small>${esc(k.org)}</small>` : ''}</td><td>${typeText(k)}</td><td><code>${esc(k.machine || '—')}</code></td><td>${fmt(k.issued)}</td><td>${k.expires ? fmt(k.expires) : 'Vĩnh viễn'}</td><td>${stateTag(k)}</td></tr>`).join('')
      : `<tr><td colspan="7" class="empty">${ks.length ? 'Không có key phù hợp.' : 'Chưa cấp key nào. Vào “Cấp key” để tạo key đầu tiên.'}</td></tr>`;
    $('#cards').innerHTML = rows.length ? rows.map((k) => `<button class="kc" data-s="${k.serial}"><b>Số ${k.serial} · ${esc(k.customer)}</b>${stateTag(k)}<small>${typeText(k)}${k.machine ? ' · ' + esc(k.machine) : ''}</small><small>Cấp ${fmt(k.issued)} · ${k.expires ? 'hết hạn ' + fmt(k.expires) : 'vĩnh viễn'}${k.org ? ' · ' + esc(k.org) : ''}</small></button>`).join('')
      : `<p class="empty">${ks.length ? 'Không có key phù hợp.' : 'Chưa cấp key nào. Vào “Cấp key” để tạo key đầu tiên.'}</p>`;
  };
  draw();
  q.oninput = () => { listQuery = q.value; draw(); };
  main.querySelectorAll('.stat').forEach((b) => b.onclick = () => { listFilter = b.dataset.f; renderList(); });
  const open = (e) => { const tr = e.target.closest('tr[data-s]'); if (tr) detail(+tr.dataset.s); };
  $('#tb').onclick = open;
  $('#cards').onclick = (e) => { const b = e.target.closest('[data-s]'); if (b) detail(+b.dataset.s); };
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
    <div class="btns"><button class="btn primary" id="d-ck">Chép key</button><button class="btn" id="d-cm">Chép tin nhắn</button>${navigator.share ? '<button class="btn" id="d-sh">Gửi…</button>' : ''}</div>
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
  $('#d-sh', back) && ($('#d-sh', back).onclick = () => navigator.share({ text: message(k) }).catch(() => {}));
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
  const lb = S.lastBackup ? new Date(S.lastBackup).toLocaleString('vi-VN') : 'chưa sao lưu lần nào';
  main.innerHTML = `<h1>Cài đặt</h1><p class="sub">Khóa ký, mật khẩu, sao lưu và đồng bộ với bản Windows.</p>
    <div class="card">
      <div class="set-row"><div><b>Khóa ký</b><small>Đã nhập. Dấu vân tay: <span class="fp">${esc(fp)}</span></small></div><button class="btn" id="s-exp">Sao lưu khóa ký…</button><button class="btn" id="s-imp">Nhập lại tệp khóa ký…</button></div>
      <div class="set-row"><div><b>Mật khẩu mở app</b><small>App tự khóa sau 10 phút không dùng.</small></div><button class="btn" id="s-pw">Đổi mật khẩu…</button><button class="btn" id="s-lock">Khóa ngay</button></div>
      <div class="set-row"><div><b>Sao lưu sổ key</b><small>Lần gần nhất: ${esc(lb)}. Lưu vào Tệp ▸ OneDrive (thư mục ExcelLike\\Key) để mở bằng Key Manager trên Windows.</small></div><button class="btn primary" id="s-bak">Sao lưu…</button></div>
      <div class="set-row"><div><b>Khôi phục / đồng bộ sổ key</b><small>Gộp key từ tệp sao lưu (của iPhone hoặc bản Windows) vào sổ này, không ghi đè key đang có.</small></div><button class="btn" id="s-res">Chọn tệp…</button></div>
      <div class="set-row"><div><b>Đánh số key</b><small>iPhone cấp từ số ${window.km.RANGE.toLocaleString('vi-VN')} trở lên, bản Windows cấp dưới số đó — hai nơi cấp cùng lúc không bao giờ trùng số.</small></div></div>
      <div class="set-row"><div><b>Xóa khỏi iPhone này</b><small>Xóa khóa ký và sổ key trên máy (sao lưu trước!).</small></div><button class="btn danger" id="s-del">Xóa…</button></div>
    </div>
    <p class="warnbox">Khóa ký là bí mật: ai có nó đều cấp được key. Đừng gửi tệp khóa ký qua Zalo/email; xóa bản tải về sau khi nhập.</p>`;
  $('#s-exp').onclick = async () => { const r = await window.km.exportSigning(); if (r.ok) toast('Đã xuất khóa ký.'); else if (r.error) toast(r.error, true); };
  $('#s-imp').onclick = () => renderReimport();
  $('#s-lock').onclick = () => window.km.lock();
  $('#s-pw').onclick = () => {
    const back = document.createElement('div'); back.className = 'drawer-back';
    back.innerHTML = `<form class="drawer" role="dialog" aria-modal="true" aria-label="Đổi mật khẩu"><header><h2>Đổi mật khẩu</h2><button type="button" class="x" aria-label="Đóng">✕</button></header>
      <label class="fld"><span>Mật khẩu hiện tại</span><input type="password" id="p0" autocomplete="current-password"></label>
      <label class="fld"><span>Mật khẩu mới</span><input type="password" id="p1" autocomplete="new-password"></label>
      <label class="fld"><span>Nhập lại mật khẩu mới</span><input type="password" id="p2" autocomplete="new-password"></label>
      <p class="err" id="p-err"></p><div class="btns"><button class="btn primary" type="submit">Đổi mật khẩu</button></div></form>`;
    document.body.appendChild(back);
    const close = () => back.remove();
    $('.x', back).onclick = close; back.addEventListener('click', (e) => { if (e.target === back) close(); });
    $('form', back).onsubmit = async (e) => {
      e.preventDefault();
      if ($('#p1', back).value !== $('#p2', back).value) { $('#p-err', back).textContent = 'Hai mật khẩu mới không khớp.'; return; }
      $('#p-err', back).textContent = 'Đang mã hóa…';
      const r = await window.km.changePass($('#p0', back).value, $('#p1', back).value);
      if (!r.ok) { $('#p-err', back).textContent = r.error; return; }
      close(); toast('Đã đổi mật khẩu.');
    };
    $('#p0', back).focus();
  };
  $('#s-bak').onclick = async () => { const r = await window.km.saveText(`So-key-ExcelLike-iPhone-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify({ version: 1, keys: S.keys.map(({ state, ...k }) => k) }, null, 1)); if (r.ok) { toast('Đã sao lưu sổ key.'); await refresh(); renderSettings(); } };
  $('#s-res').onclick = async () => { const r = await window.km.restore(); if (r.canceled) return; if (!r.ok) return toast(r.error, true); await refresh(); toast(`Đã gộp ${r.added} key.` + (r.bad ? ` Bỏ qua ${r.bad} bản ghi lỗi.` : '') + (r.conflict ? ` ${r.conflict} key trùng số khác nội dung (giữ bản đang có).` : ''), !!(r.bad || r.conflict)); };
  $('#s-del').onclick = async () => {
    if (!confirm('Xóa khóa ký và toàn bộ sổ key khỏi iPhone này? Hãy chắc chắn đã sao lưu sổ key.')) return;
    await window.km.forget(); await refresh(); render(); toast('Đã xóa khỏi iPhone này.');
  };
}

window.addEventListener('km-locked', async () => { await refresh(); render(); toast('Đã khóa Key Manager.'); });
$('#lock-btn').onclick = () => window.km.lock();
$('#lock-btn2').onclick = () => window.km.lock();
// gợi ý cài lên màn hình chính (Safari trên iPhone không có nút cài tự động)
(() => {
  const standalone = window.navigator.standalone || matchMedia('(display-mode: standalone)').matches;
  const ios = /iPhone|iPad|iPod/.test(navigator.userAgent);
  let seen = false; try { seen = localStorage.getItem('km-a2hs') === '1'; } catch { /* bỏ qua */ }
  if (ios && !standalone && !seen) {
    const b = $('#a2hs'); b.hidden = false;
    $('#a2hs .x').onclick = () => { b.hidden = true; try { localStorage.setItem('km-a2hs', '1'); } catch { /* bỏ qua */ } };
  }
})();
if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js', { scope: './' }).catch(() => {});

refresh().then(render);
