// Lớp dữ liệu Key Manager bản iPhone/web: cùng giao diện window.km như bản Windows (preload.cjs),
// nhưng chạy hoàn toàn trên máy: sổ key trong IndexedDB, khóa ký mã hóa AES-GCM bằng mật khẩu (PBKDF2).
'use strict';
(function () {
  const L = window.LIC;
  const VERSION = '1.0.0';
  /** Số key cấp trên iPhone bắt đầu từ đây, để không trùng số với Key Manager trên Windows (1–99.999) */
  const RANGE = 100001;
  const PBKDF2_ITER = 310000;
  const IDLE_MS = 10 * 60 * 1000;

  // ---------- IndexedDB (một kho khóa/giá trị) ----------
  let dbp = null;
  const idb = () => (dbp ??= new Promise((res, rej) => { const r = indexedDB.open('excellike-keymgr', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }));
  async function get(k) { const d = await idb(); return new Promise((res, rej) => { const r = d.transaction('kv').objectStore('kv').get(k); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }
  async function put(k, v) { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); }
  async function del(k) { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction('kv', 'readwrite'); t.objectStore('kv').delete(k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); }

  // ---------- khóa ký: mã hóa bằng mật khẩu ----------
  const enc = new TextEncoder(), dec = new TextDecoder();
  async function deriveKey(pass, salt, iter) {
    const base = await crypto.subtle.importKey('raw', enc.encode(pass), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  }
  async function seal(priv, pass) {
    const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
    const k = await deriveKey(pass, salt, PBKDF2_ITER);
    const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, k, enc.encode(priv)));
    return { v: 1, iter: PBKDF2_ITER, salt: L.b64(salt), iv: L.b64(iv), ct: L.b64(ct) };
  }
  async function unseal(box, pass) {
    const k = await deriveKey(pass, L.unb64(box.salt), box.iter);
    try { return dec.decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: L.unb64(box.iv) }, k, L.unb64(box.ct))); }
    catch { throw new Error('Mật khẩu không đúng.'); }
  }
  let signing = null; // { key: CryptoKey không xuất được, pub, priv } — chỉ trong bộ nhớ khi đã mở khóa
  let idleT = 0;
  function lock() { signing = null; clearTimeout(idleT); window.dispatchEvent(new Event('km-locked')); }
  function touch() { clearTimeout(idleT); if (signing) idleT = setTimeout(lock, IDLE_MS); }
  ['pointerdown', 'keydown'].forEach((e) => window.addEventListener(e, touch, { passive: true }));
  let hiddenAt = 0;
  document.addEventListener('visibilitychange', () => { if (document.hidden) hiddenAt = Date.now(); else if (signing && hiddenAt && Date.now() - hiddenAt > 2 * 60 * 1000) lock(); });

  function parseSigningFile(text) {
    let o;
    try { o = JSON.parse(text); } catch { throw new Error('Không phải tệp khóa ký ExcelLike (JSON lỗi).'); }
    if (!o || typeof o.priv !== 'string') throw new Error('Không phải tệp khóa ký ExcelLike.');
    return o;
  }

  // ---------- sổ key (cùng cấu trúc với bản Windows) ----------
  const clean = (s, n = 200) => String(s ?? '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n);
  const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const parseDate = (s) => { const m = /^(\d{4})-(\d\d)-(\d\d)$/.exec(String(s || '')); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
  const nextOf = (keys) => Math.max(RANGE, ...keys.filter((k) => k.serial >= RANGE).map((k) => k.serial + 1));
  async function readDb() {
    const d = (await get('db')) ?? (await get('db.bak')) ?? { version: 1, keys: [] };
    if (!Array.isArray(d.keys)) throw new Error('Sổ key bị hỏng. Vào Cài đặt ▸ Khôi phục sổ key từ tệp sao lưu.');
    d.nextSerial = nextOf(d.keys);
    return d;
  }
  async function writeDb(d) { const cur = await get('db'); if (cur) await put('db.bak', cur); await put('db', d); }
  function status(k) {
    if (k.status === 'revoked') return 'revoked';
    if (k.expires && L.dayNum(parseDate(k.expires)) < L.dayNum(new Date())) return 'expired';
    return 'active';
  }
  function validRec(k) {
    if (!k || !Number.isInteger(k.serial) || k.serial < 1 || k.serial > 0xffffffff || typeof k.key !== 'string') return null;
    const d = L.decodeKey(k.key);
    if (!d || d.serial !== k.serial) return null;
    return {
      serial: k.serial, key: k.key, type: d.type, machine: d.machine,
      issued: iso(L.dayDate(d.issued)), expires: d.expires ? iso(L.dayDate(d.expires)) : null,
      customer: clean(k.customer, 120) || '(không tên)', org: clean(k.org, 160), contact: clean(k.contact, 160), note: clean(k.note, 500),
      status: k.status === 'revoked' ? 'revoked' : 'active', revokedAt: typeof k.revokedAt === 'string' ? clean(k.revokedAt, 40) : undefined,
      created: typeof k.created === 'string' ? clean(k.created, 40) : new Date().toISOString(),
    };
  }

  // ---------- tệp: chọn / lưu (iPhone: bảng Chia sẻ → Lưu vào Tệp, Zalo…) ----------
  function pickFile(accept) {
    return new Promise((res) => {
      const i = document.createElement('input');
      i.type = 'file'; i.accept = accept; i.style.display = 'none';
      i.onchange = async () => { const f = i.files?.[0]; i.remove(); res(f ? await f.text() : null); };
      i.addEventListener('cancel', () => { i.remove(); res(null); });
      document.body.appendChild(i); i.click();
    });
  }
  async function saveText(name, text) {
    const type = name.endsWith('.json') ? 'application/json' : name.endsWith('.csv') ? 'text/csv' : 'text/plain';
    const file = new File([text], name, { type });
    if (navigator.canShare?.({ files: [file] })) {
      try { await navigator.share({ files: [file], title: name }); return { ok: true, path: name }; }
      catch (e) { if (e?.name === 'AbortError') return { ok: false, canceled: true }; }
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(file); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    return { ok: true, path: name };
  }
  async function copy(t) {
    try { await navigator.clipboard.writeText(String(t)); return true; }
    catch { const ta = document.createElement('textarea'); ta.value = String(t); ta.setAttribute('readonly', ''); ta.style.cssText = 'position:fixed;top:0;opacity:0'; document.body.appendChild(ta); ta.select(); ta.setSelectionRange(0, ta.value.length); const ok = document.execCommand('copy'); ta.remove(); return ok; }
  }

  window.km = {
    VERSION, RANGE,
    async state() {
      const box = await get('signing');
      const base = { web: true, signing: !!signing, sealed: !!box, pub: signing?.pub ?? box?.pub ?? null, version: VERSION, lastBackup: (await get('lastBackup')) ?? null };
      try { const db = await readDb(); return { ...base, keys: db.keys.map((k) => ({ ...k, state: status(k) })), nextSerial: db.nextSerial }; }
      catch (e) { return { ...base, keys: [], nextSerial: 0, dbError: e.message }; }
    },
    pickSigningFile: () => pickFile('.json,application/json,text/plain'),
    /** nhập khóa ký (nội dung tệp .json) và đặt mật khẩu bảo vệ trên máy này */
    async importSigning(text, pass) {
      try {
        if (!(await L.supported())) throw new Error('Trình duyệt này chưa hỗ trợ chữ ký Ed25519. Cần iPhone iOS 17 trở lên (hoặc Chrome/Edge mới).');
        const o = parseSigningFile(text);
        const k = await L.importPriv(o.priv);
        if (o.pub && o.pub !== k.pub) throw new Error('Tệp khóa ký hỏng (khóa công khai không khớp).');
        if (!pass || pass.length < 6) throw new Error('Mật khẩu tối thiểu 6 ký tự.');
        await put('signing', { ...(await seal(o.priv, pass)), pub: k.pub });
        signing = { key: k.key, pub: k.pub, priv: o.priv }; touch();
        return { ok: true, pub: k.pub };
      } catch (e) { return { ok: false, error: e.message }; }
    },
    async unlock(pass) {
      try {
        const box = await get('signing');
        if (!box) throw new Error('Chưa có khóa ký trên máy này.');
        const priv = await unseal(box, pass);
        const k = await L.importPriv(priv);
        signing = { key: k.key, pub: k.pub, priv }; touch();
        return { ok: true };
      } catch (e) { return { ok: false, error: e.message }; }
    },
    lock,
    async changePass(oldPass, newPass) {
      try {
        const box = await get('signing');
        const priv = await unseal(box, oldPass);
        if (!newPass || newPass.length < 6) throw new Error('Mật khẩu mới tối thiểu 6 ký tự.');
        await put('signing', { ...(await seal(priv, newPass)), pub: box.pub });
        return { ok: true };
      } catch (e) { return { ok: false, error: e.message }; }
    },
    async exportSigning() {
      if (!signing) return { ok: false, error: 'Hãy mở khóa trước.' };
      return saveText('ExcelLike-khoa-ky.json', JSON.stringify({ product: 'ExcelLike', note: 'KHÓA BÍ MẬT — ai có tệp này đều cấp được key. Cất giữ cẩn thận.', pub: signing.pub, priv: signing.priv }, null, 2));
    },
    async forget() { lock(); await del('signing'); await del('db'); await del('db.bak'); await del('lastBackup'); return { ok: true }; },
    async create(o) {
      if (!signing) return { ok: false, error: 'Khóa ký đang khóa. Hãy mở khóa.' };
      try {
        const db = await readDb();
        const type = o.type === 'machine' ? 'machine' : 'shared';
        const issued = new Date();
        const expires = o.expires ? parseDate(o.expires) : null;
        if (o.expires && !expires) throw new Error('Ngày hết hạn không hợp lệ');
        if (expires && L.dayNum(expires) < L.dayNum(issued)) throw new Error('Ngày hết hạn phải từ hôm nay trở đi');
        const customer = clean(o.customer, 120);
        if (!customer) throw new Error('Hãy nhập tên khách hàng / người dùng');
        let machine = null;
        if (type === 'machine') { const h = L.parseMachineCode(o.machine); if (!h) throw new Error('Mã máy không hợp lệ (13 ký tự, dạng XXXX-XXXX-XXXXX)'); machine = L.machineCode(h); }
        const serial = db.nextSerial;
        const key = await L.makeKey({ serial, type, machine, issued, expires }, signing.key);
        const rec = { serial, key, type, machine, customer, org: clean(o.org, 160), contact: clean(o.contact, 160), issued: iso(issued), expires: expires ? iso(expires) : null, note: clean(o.note, 500), status: 'active', created: new Date().toISOString(), device: 'iPhone' };
        db.keys.push(rec); delete db.nextSerial;
        await writeDb(db);
        return { ok: true, rec: { ...rec, state: status(rec) } };
      } catch (e) { return { ok: false, error: e.message }; }
    },
    async update(serial, patch) {
      const db = await readDb();
      const k = db.keys.find((x) => x.serial === serial);
      if (!k) return { ok: false, error: 'Không tìm thấy key' };
      for (const f of ['customer', 'org', 'contact', 'note']) if (f in patch) k[f] = clean(patch[f], f === 'note' ? 500 : 160);
      if (patch.status === 'revoked' || patch.status === 'active') { k.status = patch.status; k.revokedAt = patch.status === 'revoked' ? new Date().toISOString() : undefined; }
      delete db.nextSerial; await writeDb(db);
      return { ok: true };
    },
    async verify(text, machineCode) {
      const d = L.decodeKey(text);
      if (!d) return { ok: false, code: 'format', message: L.MESSAGES.format };
      const box = await get('signing');
      const pub = signing?.pub ?? box?.pub;
      if (!pub) return { ok: false, code: 'nosign', message: 'Chưa nhập khóa ký nên không kiểm được chữ ký.' };
      const m = machineCode ? L.parseMachineCode(machineCode) : null;
      const r = await L.checkKey(text, { pub, machine: m || undefined });
      let keys = []; try { keys = (await readDb()).keys; } catch { /* sổ hỏng */ }
      const rec = keys.find((k) => k.serial === d.serial && L.normalize(k.key) === L.normalize(text));
      return { ...r, message: L.MESSAGES[r.code], rec: rec ? { ...rec, state: status(rec) } : null };
    },
    copy,
    async saveText(name, text) { const r = await saveText(name, text); if (r.ok && /^So-key/.test(name) && name.endsWith('.json')) await put('lastBackup', new Date().toISOString()); return r; },
    async restore() {
      const text = await pickFile('.json,application/json');
      if (text === null) return { ok: false, canceled: true };
      try {
        const d = JSON.parse(text);
        if (!d || !Array.isArray(d.keys)) throw new Error('Không phải tệp sổ key');
        let db; try { db = await readDb(); } catch { db = { version: 1, keys: [] }; }
        const have = new Map(db.keys.map((k) => [k.serial, k]));
        let added = 0, bad = 0, conflict = 0;
        for (const raw of d.keys) {
          const k = validRec(raw);
          if (!k) { bad++; continue; }
          const cur = have.get(k.serial);
          if (cur) { if (L.normalize(cur.key) !== L.normalize(k.key)) conflict++; continue; }
          db.keys.push(k); have.set(k.serial, k); added++;
        }
        db.keys.sort((a, b) => a.serial - b.serial); delete db.nextSerial;
        await writeDb(db);
        return { ok: true, added, bad, conflict };
      } catch (e) { return { ok: false, error: e.message }; }
    },
  };
})();
