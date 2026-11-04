// Mở / lưu tệp trên máy: dùng File System Access (Chrome/Edge máy tính) nếu có, nếu không thì chọn tệp / tải xuống
const XLSX_TYPE = { description: 'Sổ tính Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } };
const CSV_TYPE = { description: 'CSV', accept: { 'text/csv': ['.csv'] } };
const XLSM_MIME = 'application/vnd.ms-excel.sheet.macroEnabled.12';
const XLSM_TYPE = { description: 'Sổ làm việc Excel có Macro', accept: { [XLSM_MIME]: ['.xlsm'] } };

/** Cầu nối bản máy tính (Electron) — có khi chạy trong ExcelLike.exe */
export interface DeskFile { path: string; name: string; data: Uint8Array }
export const desk = (): any => (window as any).xlDesktop;
export const deskToPicked = (f: DeskFile) => ({ file: new File([f.data as BlobPart], f.name), handle: { path: f.path, name: f.name } });

export const canPick = () => typeof (window as any).showOpenFilePicker === 'function';

export async function pickFile(): Promise<{ file: File; handle?: any } | null> {
  if (desk()) { const f = await desk().openDialog(); return f ? deskToPicked(f) : null; }
  if (canPick()) {
    try {
      const [handle] = await (window as any).showOpenFilePicker({ types: [{ description: 'Excel hoặc CSV', accept: { ...XLSX_TYPE.accept, ...XLSM_TYPE.accept, ...CSV_TYPE.accept } }], excludeAcceptAllOption: false });
      return { file: await handle.getFile(), handle };
    } catch (e) {
      if ((e as Error).name === 'AbortError') return null;
    }
  }
  return new Promise((res) => {
    const inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = '.xlsx,.xlsm,.csv,.txt,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12,text/csv';
    inp.onchange = () => res(inp.files?.[0] ? { file: inp.files[0] } : null);
    inp.addEventListener('cancel', () => res(null));
    inp.click();
  });
}

/**
 * Ghi vào tệp đã mở (nếu còn quyền), không thì hỏi nơi lưu / tải xuống. Trả về handle mới (nếu có).
 * kind 'xlsm': sổ có macro — hộp lưu cho chọn .xlsm hoặc .xlsx; chọn .xlsx thì gọi plain() để lấy bản không macro.
 */
export async function saveBytes(data: Uint8Array | string, name: string, kind: 'xlsx' | 'xlsm' | 'csv', handle?: any, plain?: () => Promise<Uint8Array | null>): Promise<{ handle?: any; how: 'file' | 'download'; plain?: boolean } | null> {
  if (desk()) {
    const r = await desk().save({ path: handle?.path, name, kind, data, saveAs: !handle?.path });
    if (r?.plain) {
      const d2 = plain ? await plain() : null;
      if (!d2) return null;
      const r2 = await desk().save({ path: r.path, name: r.name, kind: 'xlsx', data: d2, saveAs: false });
      return r2 ? { handle: { path: r2.path, name: r2.name }, how: 'file', plain: true } : null;
    }
    return r ? { handle: { path: r.path, name: r.name }, how: 'file' } : null;
  }
  const mime = kind === 'xlsx' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : kind === 'xlsm' ? XLSM_MIME : 'text/csv;charset=utf-8';
  const blob = new Blob([data as BlobPart], { type: mime });
  if (handle) {
    try {
      const ok = (await handle.queryPermission?.({ mode: 'readwrite' })) === 'granted' || (await handle.requestPermission?.({ mode: 'readwrite' })) === 'granted';
      if (ok) { const w = await handle.createWritable(); await w.write(blob); await w.close(); return { handle, how: 'file' }; }
    } catch { /* thử cách khác */ }
  }
  if (typeof (window as any).showSaveFilePicker === 'function') {
    try {
      const h = await (window as any).showSaveFilePicker({ suggestedName: name, types: kind === 'xlsx' ? [XLSX_TYPE] : kind === 'xlsm' ? [XLSM_TYPE, XLSX_TYPE] : [CSV_TYPE] });
      let out: Blob = blob, isPlain = false;
      if (kind === 'xlsm' && /\.xlsx$/i.test(h.name ?? '')) {
        const d2 = plain ? await plain() : null;
        if (!d2) return null;
        out = new Blob([d2 as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); isPlain = true;
      }
      const w = await h.createWritable(); await w.write(out); await w.close();
      return { handle: h, how: 'file', plain: isPlain };
    } catch (e) {
      if ((e as Error).name === 'AbortError') return null;
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  return { how: 'download' };
}

export const baseName = (n: string) => n.replace(/\.(xlsx|xlsm|csv|txt)$/i, '');
export const isBookFile = (n: string) => /\.(xlsx|xlsm)$/i.test(n);
