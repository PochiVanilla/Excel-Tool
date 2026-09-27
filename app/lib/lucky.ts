// ==========================================
// LỚP KẾT NỐI LUCKYSHEET (thư viện nạp từ CDN, chạy ngoài chu kỳ render của React)
// ==========================================
import { LuckyValue, SheetData, toBlocks } from './sheetText';

type AnyObj = Record<string, LuckyValue>;
type LuckyWindow = Window & { luckysheet?: LuckyValue; LuckyExcel?: LuckyValue };

export const ls = (): LuckyValue => (typeof window !== 'undefined' ? (window as LuckyWindow).luckysheet : undefined);
export const luckyExcel = (): LuckyValue => (typeof window !== 'undefined' ? (window as LuckyWindow).LuckyExcel : undefined);
export const isLuckyReady = () => {
  const l = ls();
  return !!l && typeof l.getSheetData === 'function';
};

export const getFiles = (): AnyObj[] => (isLuckyReady() ? ls().getluckysheetfile() || [] : []);

export const getActiveFile = (): AnyObj | undefined => {
  const files = getFiles();
  return files.find((s) => Number(s.status) === 1) || files[0];
};

export const getActiveOrder = () => {
  const files = getFiles();
  const i = files.findIndex((s) => Number(s.status) === 1);
  return i >= 0 ? i : 0;
};

export const getSheetData = (): SheetData => {
  try {
    return ls().getSheetData() || [];
  } catch {
    return [];
  }
};

export type SelRange = { row: [number, number]; column: [number, number] };

export const getSelection = (): SelRange[] => {
  try {
    return ls().getRange() || [];
  } catch {
    return [];
  }
};

// ==========================================
// LỚP PHỦ XEM TRƯỚC (vẽ trực tiếp lên canvas qua hook, không sửa dữ liệu)
// ==========================================
export const overlay = {
  deleteRows: new Set<number>(),
  keepRows: new Set<number>(),
  focusRows: new Set<number>(),
};

// Trong lúc tự vẽ lại, Luckysheet có thể bắn hook "updated" -> đánh dấu để phía React bỏ qua (tránh vòng lặp)
let selfRefreshing = false;
export const isSelfRefreshing = () => selfRefreshing;

export const refreshSheet = () => {
  selfRefreshing = true;
  try {
    ls().refresh();
  } catch {
    /* Luckysheet chưa sẵn sàng */
  } finally {
    window.setTimeout(() => {
      selfRefreshing = false;
    }, 0);
  }
};

const sameSet = (a: Set<number>, b: Set<number>) => a.size === b.size && [...a].every((x) => b.has(x));

export const setOverlay = (next: Partial<{ deleteRows: Iterable<number>; keepRows: Iterable<number>; focusRows: Iterable<number> }>) => {
  let changed = false;
  (['deleteRows', 'keepRows', 'focusRows'] as const).forEach((k) => {
    const v = next[k];
    if (!v) return;
    const set = new Set(v);
    if (!sameSet(set, overlay[k])) {
      overlay[k] = set;
      changed = true;
    }
  });
  // Chỉ vẽ lại khi lớp phủ thực sự thay đổi
  if (changed) refreshSheet();
};

export const clearOverlay = () => setOverlay({ deleteRows: [], keepRows: [], focusRows: [] });

// Hook vẽ sau mỗi ô: tô đỏ hàng sắp bị xoá, tô xanh vùng được giữ, tô vàng hàng đang soi (Pivot)
export const overlayHooks = {
  cellRenderAfter: (_cell: LuckyValue, pos: AnyObj, _sheet: LuckyValue, ctx: CanvasRenderingContext2D) => {
    const r = pos.r;
    let color = '';
    if (overlay.deleteRows.has(r)) color = 'rgba(220, 38, 38, 0.16)';
    else if (overlay.focusRows.has(r)) color = 'rgba(250, 204, 21, 0.28)';
    else if (overlay.keepRows.has(r)) color = 'rgba(16, 124, 65, 0.08)';
    if (!color) return;
    ctx.save();
    ctx.fillStyle = color;
    ctx.fillRect(pos.start_c, pos.start_r, pos.end_c - pos.start_c, pos.end_r - pos.start_r);
    ctx.restore();
  },
  rowTitleCellRenderAfter: (_rowNum: number, pos: AnyObj, ctx: CanvasRenderingContext2D) => {
    const r = pos.r;
    let color = '';
    if (overlay.deleteRows.has(r)) color = 'rgba(220, 38, 38, 0.55)';
    else if (overlay.focusRows.has(r)) color = 'rgba(234, 179, 8, 0.6)';
    else if (overlay.keepRows.has(r)) color = 'rgba(16, 124, 65, 0.35)';
    if (!color) return;
    ctx.save();
    ctx.fillStyle = color;
    ctx.fillRect(0, pos.top, 4, pos.height);
    ctx.restore();
  },
};

// ==========================================
// ĐIỀU HƯỚNG / CHỌN VÙNG
// ==========================================
export const selectRange = (ranges: SelRange[] | SelRange, scroll = true) => {
  const list = Array.isArray(ranges) ? ranges : [ranges];
  if (list.length === 0) return;
  try {
    ls().setRangeShow(list.map((r) => ({ row: r.row, column: r.column })), { show: true });
    if (scroll) ls().scroll({ targetRow: Math.max(0, list[0].row[0] - 2) });
  } catch (err) {
    console.warn('Không chọn được vùng', err);
  }
};

export const selectRows = (rows: number[], scroll = true) => {
  const data = getSheetData();
  const lastCol = Math.max(0, (data[0]?.length || 1) - 1);
  const blocks = toBlocks(rows).slice(0, 50);
  selectRange(
    blocks.map((b) => ({ row: [b.start, b.end] as [number, number], column: [0, lastCol] as [number, number] })),
    scroll,
  );
};

// Dòng có viền trong phạm vi cột c1..c2 (dùng cho luật "dọn dòng đệm không viền" của cắt gộp)
export const buildRowBorderChecker = (c1: number, c2: number) => {
  const rows = new Set<number>();
  const file = getActiveFile();
  for (const item of file?.config?.borderInfo || []) {
    if (!item) continue;
    if (item.rangeType === 'cell') {
      const v = item.value;
      if (v && v.col_index >= c1 && v.col_index <= c2 && (v.l || v.r || v.t || v.b)) rows.add(v.row_index);
    } else if (item.rangeType === 'range' && item.borderType !== 'border-none') {
      for (const rg of item.range || (item.value && item.value.range) || []) {
        if (Math.max(c1, rg.column[0]) > Math.min(c2, rg.column[1])) continue;
        for (let r = rg.row[0]; r <= rg.row[1]; r++) rows.add(r);
      }
    }
  }
  return (r: number) => rows.has(r);
};

// ==========================================
// XOÁ NHIỀU HÀNG + DỊCH CHUYỂN ẢNH (Luckysheet không tự dời ảnh khi xoá hàng)
// ==========================================
// Trả về true nếu lớp ảnh thay đổi -> phía gọi cần nạp lại workbook (reloadWorkbook) để Luckysheet vẽ lại ảnh
export const deleteRowsWithImages = (rows: number[]): boolean => {
  if (rows.length === 0) return false;
  const file = getActiveFile();
  const rowlen: AnyObj = { ...(file?.config?.rowlen || {}) };
  const defaultRow = Number(file?.defaultRowHeight) || 19;
  const deleted = new Set(rows);
  const heightOf = (r: number) => (rowlen[r] !== undefined ? Number(rowlen[r]) : defaultRow) + 1;

  // Tính lại vị trí ảnh trước khi xoá: ảnh neo ở hàng bị xoá thì bỏ, ảnh bên dưới dời lên
  const images: AnyObj = file?.images ? JSON.parse(JSON.stringify(file.images)) : null;
  let imagesChanged = false;
  if (images) {
    const maxRow = Math.max(...rows) + 1;
    const tops: number[] = [0];
    for (let r = 0; r <= maxRow; r++) tops.push(tops[r] + heightOf(r));
    const rowAt = (y: number) => {
      for (let r = 0; r < tops.length - 1; r++) if (y < tops[r + 1]) return r;
      return Number.MAX_SAFE_INTEGER;
    };
    for (const [id, img] of Object.entries<AnyObj>(images)) {
      const box = img.default || {};
      const anchor = rowAt(Number(box.top) || 0);
      if (deleted.has(anchor)) {
        delete images[id];
        imagesChanged = true;
        continue;
      }
      let shift = 0;
      for (const r of rows) if (r < anchor) shift += heightOf(r);
      if (shift > 0) {
        box.top = Math.max(0, Number(box.top) - shift);
        imagesChanged = true;
      }
    }
  }

  for (const block of toBlocks(rows).reverse()) ls().deleteRow(block.start, block.end);

  if (imagesChanged) {
    const after = getActiveFile();
    if (after) after.images = images;
  }
  return imagesChanged;
};

// Nạp lại toàn bộ workbook từ trạng thái hiện tại (Luckysheet chỉ vẽ lại lớp ảnh khi khởi tạo / đổi sheet)
export const workbookForReload = () => {
  const order = getActiveOrder();
  return snapshotWorkbook().map((s, i) => ({ ...s, status: i === order ? 1 : 0 }));
};

// ==========================================
// CHỤP / KHÔI PHỤC TOÀN BỘ WORKBOOK (cho nút Hoàn tác thao tác cắt gộp)
// ==========================================
export const snapshotWorkbook = (): AnyObj[] => {
  const sheets: AnyObj[] = ls().getAllSheets();
  return sheets.map((s) => {
    const copy = { ...s };
    delete copy.data;
    return copy;
  });
};

// Tự giãn chiều cao hàng theo nội dung (double-click ranh giới hàng như Excel)
export const autoFitRowHeight = (rowIndex: number) => {
  const data = getSheetData();
  const row = data[rowIndex];
  const file = getActiveFile();
  if (!row || !file) return;
  const columnlen: AnyObj = file.config?.columnlen || {};
  const defaultCol = Number(file.defaultColWidth) || 73;
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  if (!ctx) return;
  let maxH = 19;
  for (let c = 0; c < row.length; c++) {
    const cell = row[c];
    if (!cell || (cell.mc && (cell.mc.r !== rowIndex || (cell.mc.rs ?? 1) > 1))) continue;
    const runs: { v: string; fs: number; bl?: number; ff?: string }[] =
      cell.ct?.t === 'inlineStr' && cell.ct.s
        ? cell.ct.s.map((s: AnyObj) => ({ v: String(s.v ?? ''), fs: Number(s.fs || cell.fs || 10), bl: s.bl, ff: s.ff || cell.ff }))
        : [{ v: String(cell.m ?? cell.v ?? ''), fs: Number(cell.fs || 10), bl: cell.bl, ff: cell.ff }];
    const text = runs.map((r) => r.v).join('');
    if (!text.trim()) continue;
    let width = 0;
    const cs = cell.mc?.cs ?? 1;
    for (let k = 0; k < cs; k++) width += columnlen[c + k] !== undefined ? Number(columnlen[c + k]) : defaultCol;
    width = Math.max(10, width - 6);
    const fs = Math.max(...runs.map((r) => r.fs));
    const lineH = Math.ceil(fs * 1.33 * 1.2);
    const wrap = Number(cell.tb) === 2;
    let lines = 0;
    for (const para of text.split(/\r?\n/)) {
      ctx.font = `${runs[0]?.bl ? 'bold ' : ''}${Math.round(fs * 1.33)}px ${runs[0]?.ff || 'Arial'}`;
      const w = ctx.measureText(para).width;
      lines += wrap ? Math.max(1, Math.ceil(w / width)) : 1;
    }
    maxH = Math.max(maxH, lines * lineH + 6);
  }
  try {
    ls().setRowHeight({ [rowIndex]: Math.min(maxH, 600) });
  } catch (err) {
    console.warn('Không giãn được chiều cao hàng', err);
  }
};
