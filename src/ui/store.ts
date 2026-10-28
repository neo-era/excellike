// Lưu sổ tính trong IndexedDB của trình duyệt (tự lưu + danh sách sổ đã lưu)
export interface BookRec { id: string; name: string; updated: number; data: unknown; handle?: unknown; fileName?: string; fileDirty?: boolean }

const DB = 'excellike', STORE = 'books';

function open(): Promise<IDBDatabase> {
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE, { keyPath: 'id' });
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function tx<T>(mode: IDBTransactionMode, f: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await open();
  return new Promise((res, rej) => {
    // chỉ báo xong khi giao dịch đã ghi hẳn xuống đĩa (oncomplete), không phải lúc yêu cầu thành công:
    // nạp lại trang ngay sau đó sẽ huỷ giao dịch chưa xong → mất sổ vừa mở (tệp lớn ghi lâu)
    const t = db.transaction(STORE, mode, { durability: 'strict' } as any);
    const q = f(t.objectStore(STORE));
    t.oncomplete = () => { db.close(); res(q.result); };
    t.onerror = () => { db.close(); rej(t.error ?? q.error); };
    t.onabort = () => { db.close(); rej(t.error ?? new Error('Ghi dữ liệu bị huỷ (có thể đầy bộ nhớ).')); };
  });
}
export const saveBook = (b: BookRec) => tx('readwrite', (s) => s.put(b));
export const loadBook = (id: string) => tx<BookRec | undefined>('readonly', (s) => s.get(id));
export const deleteBook = (id: string) => tx('readwrite', (s) => s.delete(id));
export async function listBooks(): Promise<BookRec[]> {
  const all = await tx<BookRec[]>('readonly', (s) => s.getAll());
  return all.sort((a, b) => b.updated - a.updated);
}
export function lsGet(k: string) { try { return localStorage.getItem(k); } catch { return null; } }
export function lsSet(k: string, v: string) { try { localStorage.setItem(k, v); } catch { /* bỏ qua */ } }
