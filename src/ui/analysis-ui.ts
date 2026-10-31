// Phân tích (0.9): Tìm Mục tiêu, Bảng Dữ liệu, Trình quản lý Kịch bản, Trang tính Dự báo, Lọc Nâng cao, Hợp nhất
import { openDialog, closeDialog, h, field, select, input } from './dialogs';
import { parseSheetRef } from './charts';
import { toA1, rangeToA1, type Range } from '../core/address';
import { quoteSheet } from '../core/formula';
import type { Scenario } from '../core/workbook';
import type { App } from './app';

const refOf = (app: App, g: Range, s = app.grid.sheet) => `${quoteSheet(app.wb.sheetName(s))}!${rangeToA1(g).split(':').map((p) => p.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2')).join(':')}`;
const cellIn = (app: App, el: HTMLInputElement, label: string) => {
  const r = parseSheetRef(app.wb, el.value.trim().replace(/^=/, ''), app.grid.sheet);
  if (!r || r.range.r1 !== r.range.r2 || r.range.c1 !== r.range.c2) throw new Error(`${label}: cần tham chiếu tới một ô (vd B5).`);
  if (r.sheet !== app.grid.sheet) throw new Error(`${label}: ô phải ở trang tính hiện hành.`);
  return { r: r.range.r1, c: r.range.c1 };
};
const numIn = (t: string) => { const n = Number(t.trim().replace(/\./g, '').replace(',', '.')); return isFinite(n) && t.trim() !== '' ? n : null; };
const done = (app: App) => { app.grid.layout(); app.grid.invalidate(); app.syncSelection(); app.scheduleSave(); };
const fmtN = (x: number) => (Math.round(x * 1e9) / 1e9).toLocaleString('vi-VN', { maximumFractionDigits: 9 });

/** Tìm Mục tiêu (Goal Seek) */
export function goalSeekDialog(app: App) {
  const { ar, ac } = app.grid.sel;
  const setI = input(toA1(ar, ac)), toI = input(''), byI = input('');
  setI.classList.add('mono'); byI.classList.add('mono');
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  openDialog('Tìm Mục tiêu', h('div', { class: 'an-form' }, field('Đặt ô', setI), field('Thành giá trị', toI), field('Bằng cách thay đổi ô', byI), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      try {
        const set = cellIn(app, setI, 'Đặt ô'), by = cellIn(app, byI, 'Ô thay đổi');
        const target = numIn(toI.value);
        if (target === null) throw new Error('Giá trị đích phải là số.');
        const s = app.grid.sheet, wb = app.wb;
        const old = wb.rawText(s, by.r, by.c);
        const res = wb.goalSeek(s, set, target, by);
        if (res.ok) wb.setInput(s, by.r, by.c, String(Math.round(res.x * 1e12) / 1e12).replace('.', ','));
        done(app);
        // kết quả như Excel: OK giữ, Hủy trả lại
        const msg = h('div', { class: 'an-form' },
          h('p', {}, res.ok ? `Tìm Mục tiêu với ô ${toA1(set.r, set.c)} đã tìm thấy một lời giải.` : `Tìm Mục tiêu với ô ${toA1(set.r, set.c)} có thể chưa tìm được lời giải.`),
          h('p', {}, `Giá trị đích: ${fmtN(target)}`), h('p', {}, `Giá trị hiện tại: ${fmtN(res.y)}`));
        setTimeout(() => openDialog('Trạng thái Tìm Mục tiêu', msg, [
          { label: 'Hủy', act: () => { if (res.ok) { app.wb.undo(); done(app); } } },
          { label: 'OK', kind: 'primary', act: () => {} },
        ]), 0);
        void old;
      } catch (e) { err.textContent = (e as Error).message; return false; }
    } },
  ]);
  setTimeout(() => toI.focus(), 40);
}

/** Bảng Dữ liệu */
export function dataTableDialog(app: App) {
  const g = app.grid.range;
  if (g.r1 === g.r2 && g.c1 === g.c2) { app.toast('Hãy chọn cả vùng bảng dữ liệu (gồm công thức và các giá trị thử).', 'err'); return; }
  const rowI = input(''), colI = input('');
  rowI.classList.add('mono'); colI.classList.add('mono');
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  openDialog('Bảng Dữ liệu', h('div', { class: 'an-form' }, field('Ô nhập hàng', rowI), field('Ô nhập cột', colI), h('p', { class: 'muted' }, 'Kết quả ghi dạng giá trị; đổi dữ liệu thì chạy lại Bảng Dữ liệu để cập nhật.'), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      try {
        const r = rowI.value.trim() ? cellIn(app, rowI, 'Ô nhập hàng') : null, c = colI.value.trim() ? cellIn(app, colI, 'Ô nhập cột') : null;
        app.wb.dataTable(app.grid.sheet, g, r, c);
        done(app);
      } catch (e) { err.textContent = (e as Error).message; return false; }
    } },
  ]);
  setTimeout(() => rowI.focus(), 40);
}

/** Trình quản lý Kịch bản */
export function scenarioManager(app: App) {
  const s = app.grid.sheet, wb = app.wb;
  let list: Scenario[] = wb.scenarios(s).map((x) => ({ ...x, cells: [...x.cells], values: [...x.values] }));
  const lb = h('select', { size: '6', class: 'nm-list', 'aria-label': 'Kịch bản' }) as HTMLSelectElement;
  const info = h('div', { class: 'an-info' });
  const draw = () => {
    const i = lb.selectedIndex;
    lb.replaceChildren(...list.map((x, j) => h('option', { value: String(j) }, x.name)));
    if (!list.length) lb.append(h('option', { disabled: '' }, 'Chưa có kịch bản. Bấm Thêm…'));
    lb.selectedIndex = Math.min(Math.max(0, i), list.length - 1);
    const sc = list[lb.selectedIndex];
    info.textContent = sc ? `Ô thay đổi: ${sc.cells.map(([r, c]) => toA1(r, c)).join('; ')}${sc.comment ? ' — ' + sc.comment : ''}` : '';
  };
  lb.onchange = draw;
  const save = () => wb.setScenarios(s, list);
  const edit = (sc?: Scenario) => {
    const g = app.grid.range;
    const nm = input(sc?.name ?? ''), cells = input(sc ? sc.cells.map(([r, c]) => toA1(r, c)).join(';') : rangeToA1(g)), cm = input(sc?.comment ?? '');
    cells.classList.add('mono');
    const err = h('p', { class: 'dlg-err', role: 'alert' });
    openDialog(sc ? 'Sửa Kịch bản' : 'Thêm Kịch bản', h('div', { class: 'an-form' }, field('Tên kịch bản', nm), field('Ô thay đổi (cách nhau bởi ;)', cells), field('Chú thích', cm), err), [
      { label: 'Hủy', act: () => { setTimeout(() => scenarioManagerOpen(), 0); } },
      { label: 'OK', kind: 'primary', act: () => {
        const name = nm.value.trim();
        if (!name) { err.textContent = 'Hãy nhập tên kịch bản.'; return false; }
        if (list.some((x) => x !== sc && x.name.toLowerCase() === name.toLowerCase())) { err.textContent = 'Tên kịch bản đã có.'; return false; }
        const pts: [number, number][] = [];
        for (const part of cells.value.split(/[;,]/).map((x) => x.trim()).filter(Boolean)) {
          const r = parseSheetRef(wb, part, s);
          if (!r || r.sheet !== s) { err.textContent = `Ô “${part}” không hợp lệ.`; return false; }
          for (let rr = r.range.r1; rr <= r.range.r2; rr++) for (let cc = r.range.c1; cc <= r.range.c2; cc++) pts.push([rr, cc]);
        }
        if (!pts.length || pts.length > 32) { err.textContent = 'Cần từ 1 đến 32 ô thay đổi.'; return false; }
        // bước 2: giá trị cho từng ô (mặc định giá trị hiện tại)
        const ins = pts.map(([r, c], i) => { const x = input(sc && sc.cells[i] && sc.cells[i][0] === r && sc.cells[i][1] === c ? sc.values[i] : wb.rawText(s, r, c)); return { x, el: field(toA1(r, c), x) }; });
        setTimeout(() => openDialog('Giá trị Kịch bản', h('div', { class: 'an-form an-vals' }, h('p', {}, 'Nhập giá trị cho từng ô thay đổi:'), ...ins.map((x) => x.el)), [
          { label: 'Hủy', act: () => { setTimeout(() => scenarioManagerOpen(), 0); } },
          { label: 'OK', kind: 'primary', act: () => {
            const ns: Scenario = { name, cells: pts, values: ins.map((x) => x.x.value), comment: cm.value.trim() || undefined };
            if (sc) list = list.map((x) => (x === sc ? ns : x)); else list.push(ns);
            save();
            setTimeout(() => scenarioManagerOpen(), 0);
          } },
        ]), 0);
      } },
    ]);
  };
  const bAdd = h('button', { class: 'btn' }, 'Thêm…'), bDel = h('button', { class: 'btn' }, 'Xóa'), bEdit = h('button', { class: 'btn' }, 'Sửa…'), bSum = h('button', { class: 'btn' }, 'Tóm tắt…');
  bAdd.onclick = () => { closeDialog(); edit(); };
  bEdit.onclick = () => { const sc = list[lb.selectedIndex]; if (sc) { closeDialog(); edit(sc); } };
  bDel.onclick = () => { if (!list[lb.selectedIndex]) return; list.splice(lb.selectedIndex, 1); save(); draw(); };
  bSum.onclick = () => {
    if (!list.length) return;
    const resI = input(toA1(app.grid.sel.ar, app.grid.sel.ac)); resI.classList.add('mono');
    const err = h('p', { class: 'dlg-err', role: 'alert' });
    closeDialog();
    openDialog('Tóm tắt Kịch bản', h('div', { class: 'an-form' }, field('Ô kết quả (cách nhau bởi ;)', resI), err), [
      { label: 'Hủy', act: () => {} },
      { label: 'OK', kind: 'primary', act: () => {
        const pts: { r: number; c: number }[] = [];
        for (const part of resI.value.split(/[;,]/).map((x) => x.trim()).filter(Boolean)) { const r = parseSheetRef(wb, part, s); if (!r || r.sheet !== s) { err.textContent = `Ô “${part}” không hợp lệ.`; return false; } pts.push({ r: r.range.r1, c: r.range.c1 }); }
        try { const id = wb.scenarioSummary(s, pts); app.grid.setSheet(id); app.renderSheets(); done(app); } catch (e) { err.textContent = (e as Error).message; return false; }
      } },
    ]);
  };
  const scenarioManagerOpen = () => scenarioManager(app);
  draw();
  openDialog('Trình quản lý Kịch bản', h('div', { class: 'an-sm' }, h('div', {}, lb, info), h('div', { class: 'an-side' }, bAdd, bDel, bEdit, bSum)), [
    { label: 'Đóng', act: () => {} },
    { label: 'Hiện', kind: 'primary', act: () => { const sc = list[lb.selectedIndex]; if (!sc) return false; app.safe(() => wb.showScenario(s, sc)); done(app); return false; } },
  ]);
}

/** Trang tính Dự báo: dự báo tuyến tính (FORECAST.LINEAR) + khoảng tin cậy, kèm biểu đồ đường */
export function forecastDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb;
  let g = app.grid.range;
  if (app.grid.isSingle()) g = wb.currentRegion(s, app.grid.sel.ar, app.grid.sel.ac);
  if (g.c2 - g.c1 < 1) { app.toast('Chọn 2 cột: cột thời gian (ngày hoặc số) và cột giá trị.', 'err'); return; }
  const hasHead = typeof wb.value(s, g.r1, g.c1 + 1) === 'string';
  const d1 = g.r1 + (hasHead ? 1 : 0);
  const n = g.r2 - d1 + 1;
  if (n < 3) { app.toast('Cần ít nhất 3 điểm dữ liệu để dự báo.', 'err'); return; }
  const tl = [...Array(n).keys()].map((i) => wb.value(s, d1 + i, g.c1));
  if (tl.some((v) => typeof v !== 'number')) { app.toast('Cột thời gian phải là ngày hoặc số.', 'err'); return; }
  const steps = input(String(Math.max(1, Math.round(n / 3))));
  const conf = input('95');
  openDialog('Tạo Trang tính Dự báo', h('div', { class: 'an-form' }, field('Số kỳ dự báo', steps), field('Khoảng tin cậy (%)', conf), h('p', { class: 'muted' }, 'Dự báo theo xu hướng tuyến tính (FORECAST.LINEAR). Excel dùng làm trơn hàm mũ (ETS) nên số có thể khác.')), [
    { label: 'Hủy', act: () => {} },
    { label: 'Tạo', kind: 'primary', act: () => {
      const k = Math.min(500, Math.max(1, parseInt(steps.value) || 1)), cf = Math.min(99.9, Math.max(1, numIn(conf.value) ?? 95));
      const t = tl as number[], step = (t[n - 1] - t[0]) / (n - 1);
      const src = quoteSheet(wb.sheetName(s));
      const X = `${src}!$${colL(g.c1)}$${d1 + 1}:$${colL(g.c1)}$${d1 + n}`, Y = `${src}!$${colL(g.c1 + 1)}$${d1 + 1}:$${colL(g.c1 + 1)}$${d1 + n}`;
      const fmt = wb.style(s, d1, g.c1).fmt;
      app.safe(() => wb.group(() => {
        const id = wb.addSheet(wb.uniqueSheetName('Dự báo'));
        const head = hasHead ? wb.display(s, g.r1, g.c1 + 1).text : 'Giá trị';
        const rows: [number, number, string | number | null][] = [[0, 0, "'" + (hasHead ? wb.display(s, g.r1, g.c1).text : 'Thời gian')], [0, 1, "'" + head], [0, 2, `'Dự báo(${head})`], [0, 3, "'Cận dưới tin cậy"], [0, 4, "'Cận trên tin cậy"]];
        for (let i = 0; i < n; i++) { rows.push([1 + i, 0, `=${src}!${colL(g.c1)}${d1 + i + 1}`]); rows.push([1 + i, 1, `=${src}!${colL(g.c1 + 1)}${d1 + i + 1}`]); }
        // điểm cuối thực tế nối với đường dự báo như Excel
        rows.push([n, 2, `=B${n + 1}`]); rows.push([n, 3, `=B${n + 1}`]); rows.push([n, 4, `=B${n + 1}`]);
        const ci = (r: number) => `T.INV.2T(${String(1 - cf / 100).replace('.', ',')};${n}-2)*STEYX(${Y};${X})*SQRT(1+1/${n}+(A${r}-AVERAGE(${X}))^2/DEVSQ(${X}))`;
        for (let j = 1; j <= k; j++) {
          const r = n + j, rr = r + 1;
          rows.push([r, 0, String(Math.round((t[n - 1] + step * j) * 1e9) / 1e9).replace('.', ',')]);
          rows.push([r, 2, `=FORECAST.LINEAR(A${rr};${Y};${X})`]);
          rows.push([r, 3, `=C${rr}-${ci(rr)}`]);
          rows.push([r, 4, `=C${rr}+${ci(rr)}`]);
        }
        wb.writeCells(id, rows);
        if (fmt) wb.setStyle(id, { r1: 1, c1: 0, r2: n + k, c2: 0 }, { fmt });
        wb.setStyle(id, { r1: 0, c1: 0, r2: 0, c2: 4 }, { b: true });
        wb.setStyle(id, { r1: 1, c1: 1, r2: n + k, c2: 4 }, { fmt: '#,##0.00' });
        for (let c = 0; c <= 4; c++) wb.setColWidth(id, c, c === 0 ? 90 : 120);
        const R = (c: number, r1: number, r2: number) => ({ sheet: id, range: { r1, c1: c, r2, c2: c } });
        wb.addChart(id, { type: 'line', title: 'Dự báo ' + head, series: [1, 2, 3, 4].map((c) => ({ name: R(c, 0, 0), cats: R(0, 1, n + k), vals: R(c, 1, n + k) })), anchor: { r: 1, c: 6, dx: 0, dy: 0 }, w: 520, h: 300, legend: 'b' });
        app.grid.setSheet(id);
      }));
      app.renderSheets(); done(app);
    } },
  ]);
}
const colL = (c: number) => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };

/** Lọc Nâng cao */
export function advancedFilterDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb;
  const g0 = app.grid.isSingle() ? wb.currentRegion(s, app.grid.sel.ar, app.grid.sel.ac) : app.grid.range;
  const listI = input(refOf(app, g0)), critI = input(''), toI = input('');
  for (const x of [listI, critI, toI]) x.classList.add('mono');
  const inPlace = h('input', { type: 'radio', name: 'af-mode' }) as HTMLInputElement, copy = h('input', { type: 'radio', name: 'af-mode' }) as HTMLInputElement;
  inPlace.checked = true;
  const uniq = h('input', { type: 'checkbox' }) as HTMLInputElement;
  const toF = field('Sao chép tới', toI);
  const sync = () => { toI.disabled = !copy.checked; };
  inPlace.onchange = copy.onchange = sync; sync();
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  openDialog('Lọc Nâng cao', h('div', { class: 'an-form' },
    h('label', { class: 'chk' }, inPlace, ' Lọc danh sách tại chỗ'), h('label', { class: 'chk' }, copy, ' Sao chép tới vị trí khác'),
    field('Vùng danh sách', listI), field('Vùng điều kiện', critI), toF, h('label', { class: 'chk' }, uniq, ' Chỉ bản ghi duy nhất'), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      try {
        const L = parseSheetRef(wb, listI.value.replace(/^=/, ''), s);
        if (!L || L.sheet !== s || L.range.r2 <= L.range.r1) throw new Error('Vùng danh sách không hợp lệ (cần có hàng tiêu đề).');
        const C = critI.value.trim() ? parseSheetRef(wb, critI.value.replace(/^=/, ''), s) : null;
        if (critI.value.trim() && (!C || C.sheet !== s)) throw new Error('Vùng điều kiện không hợp lệ.');
        let to: { s: number; r: number; c: number } | undefined;
        if (copy.checked) { const T = parseSheetRef(wb, toI.value.replace(/^=/, ''), s); if (!T) throw new Error('Ô “Sao chép tới” không hợp lệ.'); to = { s: T.sheet, r: T.range.r1, c: T.range.c1 }; }
        const n = wb.advancedFilter(s, L.range, C?.range ?? null, { to, unique: uniq.checked });
        done(app);
        app.toast(`Tìm thấy ${n} bản ghi.`, 'ok');
      } catch (e) { err.textContent = (e as Error).message; return false; }
    } },
  ]);
  setTimeout(() => critI.focus(), 40);
}

/** Hợp nhất */
export function consolidateDialog(app: App) {
  const s = app.grid.sheet, wb = app.wb;
  const fn = select([['sum', 'Tổng'], ['count', 'Đếm'], ['average', 'Trung bình'], ['max', 'Lớn nhất'], ['min', 'Nhỏ nhất'], ['product', 'Tích']], 'sum');
  const refI = input(refOf(app, app.grid.range)); refI.classList.add('mono');
  const lb = h('select', { size: '5', class: 'nm-list', 'aria-label': 'Tất cả tham chiếu' }) as HTMLSelectElement;
  const refs: string[] = [];
  const bAdd = h('button', { class: 'btn' }, 'Thêm'), bDel = h('button', { class: 'btn' }, 'Xóa');
  const draw = () => lb.replaceChildren(...refs.map((r) => h('option', {}, r)));
  bAdd.onclick = () => { const t = refI.value.trim().replace(/^=/, ''); if (t && !refs.includes(t)) { refs.push(t); draw(); } };
  bDel.onclick = () => { if (lb.selectedIndex >= 0) { refs.splice(lb.selectedIndex, 1); draw(); } };
  const top = h('input', { type: 'checkbox' }) as HTMLInputElement, left = h('input', { type: 'checkbox' }) as HTMLInputElement;
  top.checked = left.checked = true;
  const err = h('p', { class: 'dlg-err', role: 'alert' });
  openDialog('Hợp nhất', h('div', { class: 'an-form' }, field('Hàm', fn), h('div', { class: 'an-row' }, field('Tham chiếu', refI), bAdd), field('Tất cả tham chiếu', lb), bDel,
    h('div', { class: 'an-row' }, h('label', { class: 'chk' }, top, ' Hàng trên cùng'), h('label', { class: 'chk' }, left, ' Cột bên trái')), h('p', { class: 'muted' }, `Kết quả ghi từ ô ${toA1(app.grid.sel.ar, app.grid.sel.ac)}.`), err), [
    { label: 'Hủy', act: () => {} },
    { label: 'OK', kind: 'primary', act: () => {
      try {
        if (!refs.length && refI.value.trim()) refs.push(refI.value.trim().replace(/^=/, ''));
        const rs = refs.map((t) => { const r = parseSheetRef(wb, t, s); if (!r) throw new Error(`Tham chiếu “${t}” không hợp lệ.`); return r; });
        const out = wb.consolidate({ s, r: app.grid.sel.ar, c: app.grid.sel.ac }, rs, fn.value as 'sum', top.checked, left.checked);
        app.grid.selectRange(out); done(app);
      } catch (e) { err.textContent = (e as Error).message; return false; }
    } },
  ]);
}
