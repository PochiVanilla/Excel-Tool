// ==========================================
// XUẤT FILE EXCEL "THẤY GÌ TẢI NẤY" TỪ TRẠNG THÁI HIỆN TẠI CỦA LUCKYSHEET
// Mọi thao tác trên web (sửa ô, xoá/chèn hàng, định dạng, sheet Pivot...) đều có trong file tải về.
// Giữ: giá trị, công thức, rich-text, font, màu, căn lề, xuống dòng, định dạng số, viền, ô gộp,
//      độ rộng cột, chiều cao hàng, hàng/cột ẩn, cố định dòng, bộ lọc, ảnh, link, cài đặt in của file gốc
// ==========================================
import ExcelJS from 'exceljs';
import type { LuckyValue } from './sheetText';

type AnyObj = Record<string, LuckyValue>;

const BORDER_STYLES: Record<number, ExcelJS.BorderStyle> = {
  1: 'thin',
  2: 'hair',
  3: 'dotted',
  4: 'dashed',
  5: 'dashDot',
  6: 'dashDotDot',
  7: 'double',
  8: 'medium',
  9: 'mediumDashed',
  10: 'mediumDashDot',
  11: 'mediumDashDotDot',
  12: 'slantDashDot',
  13: 'thick',
};

// Danh sách font mặc định của Luckysheet khi ff là chỉ số
const LUCKY_FONTS = ['Times New Roman', 'Arial', 'Tahoma', 'Verdana', 'Microsoft YaHei', 'SimSun', 'SimHei', 'KaiTi', 'FangSong', 'NSimSun'];

// Quy đổi ngược với LuckyExcel: px = round((width - 0.83) * 8 + 5), px = pt / 0.75
const pxToColWidth = (px: number) => Math.max(0.5, Math.round(((px - 5) / 8 + 0.83) * 100) / 100);
const pxToPt = (px: number) => Math.round(px * 0.75 * 100) / 100;

export const colorToArgb = (color?: string): string | undefined => {
  if (!color || typeof color !== 'string') return undefined;
  const c = color.trim();
  let m = c.match(/^#?([0-9a-f]{6})$/i);
  if (m) return 'FF' + m[1].toUpperCase();
  m = c.match(/^#?([0-9a-f])([0-9a-f])([0-9a-f])$/i);
  if (m) return ('FF' + m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  m = c.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (m) return 'FF' + [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('').toUpperCase();
  return undefined;
};

const fontName = (ff: LuckyValue): string | undefined => {
  if (ff === undefined || ff === null || ff === '') return undefined;
  if (typeof ff === 'number' || /^\d+$/.test(String(ff))) return LUCKY_FONTS[Number(ff)] ?? undefined;
  return String(ff);
};

const buildFont = (src: AnyObj, fallback: AnyObj = {}): Partial<ExcelJS.Font> | undefined => {
  const pick = (k: string) => (src[k] !== undefined ? src[k] : fallback[k]);
  const font: Partial<ExcelJS.Font> = {};
  const name = fontName(pick('ff'));
  if (name) font.name = name;
  const fs = Number(pick('fs'));
  if (fs > 0) font.size = fs;
  if (Number(pick('bl')) === 1) font.bold = true;
  if (Number(pick('it')) === 1) font.italic = true;
  if (Number(pick('cl')) === 1) font.strike = true;
  const un = Number(pick('un'));
  if (un === 1) font.underline = true;
  else if (un === 2) font.underline = 'double';
  else if (un === 3) font.underline = 'singleAccounting';
  else if (un === 4) font.underline = 'doubleAccounting';
  const color = colorToArgb(pick('fc'));
  if (color && color !== 'FF000000') font.color = { argb: color };
  const va = Number(src.va);
  if (va === 1) font.vertAlign = 'subscript';
  else if (va === 2) font.vertAlign = 'superscript';
  return Object.keys(font).length ? font : undefined;
};

const buildAlignment = (cell: AnyObj): Partial<ExcelJS.Alignment> | undefined => {
  const a: Partial<ExcelJS.Alignment> = {};
  const ht = cell.ht === undefined ? undefined : Number(cell.ht);
  if (ht === 0) a.horizontal = 'center';
  else if (ht === 1) a.horizontal = 'left';
  else if (ht === 2) a.horizontal = 'right';
  const vt = cell.vt === undefined ? undefined : Number(cell.vt);
  if (vt === 0) a.vertical = 'middle';
  else if (vt === 1) a.vertical = 'top';
  else if (vt === 2) a.vertical = 'bottom';
  if (Number(cell.tb) === 2) a.wrapText = true;
  const tr = cell.tr === undefined ? undefined : Number(cell.tr);
  if (tr === 1) a.textRotation = 45;
  else if (tr === 2) a.textRotation = -45;
  else if (tr === 3) a.textRotation = 'vertical';
  else if (tr === 4) a.textRotation = 90;
  else if (tr === 5) a.textRotation = -90;
  if (cell.rt !== undefined && !isNaN(Number(cell.rt)) && Number(cell.rt) !== 0) a.textRotation = Number(cell.rt);
  return Object.keys(a).length ? a : undefined;
};

type Side = { style: ExcelJS.BorderStyle; color?: { argb: string } } | null;
type CellBorder = { top?: Side; left?: Side; bottom?: Side; right?: Side };

const toSide = (s: AnyObj | undefined | null): Side => {
  if (!s) return null;
  const style = BORDER_STYLES[Number(s.style)];
  if (!style) return null;
  const argb = colorToArgb(s.color);
  return argb ? { style, color: { argb } } : { style };
};

// Dựng bản đồ viền từng ô từ config.borderInfo (cả dạng từng ô lẫn dạng vùng do thanh công cụ tạo)
const buildBorderMap = (borderInfo: AnyObj[] | undefined) => {
  const map = new Map<string, CellBorder>();
  const set = (r: number, c: number, side: keyof CellBorder, value: Side) => {
    const key = `${r}_${c}`;
    const cur = map.get(key) || {};
    cur[side] = value;
    map.set(key, cur);
  };
  for (const item of borderInfo || []) {
    if (!item) continue;
    if (item.rangeType === 'cell' && item.value) {
      const v = item.value;
      const r = Number(v.row_index);
      const c = Number(v.col_index);
      if (v.l !== undefined) set(r, c, 'left', toSide(v.l));
      if (v.r !== undefined) set(r, c, 'right', toSide(v.r));
      if (v.t !== undefined) set(r, c, 'top', toSide(v.t));
      if (v.b !== undefined) set(r, c, 'bottom', toSide(v.b));
      continue;
    }
    if (item.rangeType !== 'range') continue;
    const side = toSide({ style: item.style, color: item.color });
    for (const rg of item.range || []) {
      const [r1, r2] = rg.row;
      const [c1, c2] = rg.column;
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          switch (item.borderType) {
            case 'border-all':
              set(r, c, 'top', side);
              set(r, c, 'bottom', side);
              set(r, c, 'left', side);
              set(r, c, 'right', side);
              break;
            case 'border-outside':
              if (r === r1) set(r, c, 'top', side);
              if (r === r2) set(r, c, 'bottom', side);
              if (c === c1) set(r, c, 'left', side);
              if (c === c2) set(r, c, 'right', side);
              break;
            case 'border-inside':
              if (r < r2) set(r, c, 'bottom', side);
              if (c < c2) set(r, c, 'right', side);
              break;
            case 'border-horizontal':
              if (r < r2) set(r, c, 'bottom', side);
              break;
            case 'border-vertical':
              if (c < c2) set(r, c, 'right', side);
              break;
            case 'border-top':
              if (r === r1) set(r, c, 'top', side);
              break;
            case 'border-bottom':
              if (r === r2) set(r, c, 'bottom', side);
              break;
            case 'border-left':
              if (c === c1) set(r, c, 'left', side);
              break;
            case 'border-right':
              if (c === c2) set(r, c, 'right', side);
              break;
            case 'border-none':
              set(r, c, 'top', null);
              set(r, c, 'bottom', null);
              set(r, c, 'left', null);
              set(r, c, 'right', null);
              break;
          }
        }
      }
    }
  }
  return map;
};

// Lấy dữ liệu 2 chiều của sheet (sheet chưa từng được mở trên Luckysheet chỉ có celldata)
const sheetMatrix = (sheet: AnyObj): AnyObj[][] => {
  if (Array.isArray(sheet.data) && sheet.data.length > 0) return sheet.data;
  const out: AnyObj[][] = [];
  for (const item of sheet.celldata || []) {
    if (!out[item.r]) out[item.r] = [];
    out[item.r][item.c] = item.v;
  }
  return out;
};

const toExcelValue = (cell: AnyObj, link?: AnyObj): ExcelJS.CellValue => {
  if (cell.f) {
    const result = cell.v === undefined || cell.v === null || cell.v === '' ? undefined : cell.v;
    return { formula: String(cell.f).replace(/^=/, ''), result } as ExcelJS.CellFormulaValue;
  }
  let value: ExcelJS.CellValue = null;
  if (cell.ct && cell.ct.t === 'inlineStr' && Array.isArray(cell.ct.s)) {
    const runs = cell.ct.s
      .filter((s: AnyObj) => s && s.v !== undefined && s.v !== null && s.v !== '')
      .map((s: AnyObj) => {
        const font = buildFont(s, cell);
        return font ? { text: String(s.v).replace(/\r\n/g, '\n'), font } : { text: String(s.v).replace(/\r\n/g, '\n') };
      });
    if (runs.length === 0) return null;
    value = runs.length === 1 && !runs[0].font ? runs[0].text : { richText: runs };
  } else if (cell.v !== undefined && cell.v !== null && cell.v !== '') {
    const t = cell.ct?.t;
    if (typeof cell.v === 'number') value = cell.v;
    else if (typeof cell.v === 'boolean') value = cell.v;
    else if ((t === 'n' || t === 'd') && !isNaN(Number(cell.v))) value = Number(cell.v);
    else if (t === 'b') value = String(cell.v).toUpperCase() === 'TRUE';
    else value = String(cell.v).replace(/\r\n/g, '\n');
  }
  if (link && link.linkAddress && value !== null) {
    const text = typeof value === 'object' && value && 'richText' in value ? value.richText.map((r) => r.text).join('') : String(value);
    const hyperlink = link.linkType === 'internal' ? `#${link.linkAddress}` : link.linkAddress;
    return { text, hyperlink, tooltip: link.linkTooltip || undefined } as ExcelJS.CellHyperlinkValue;
  }
  return value;
};

const dataUrlToImage = (src: string): { base64: string; extension: 'png' | 'jpeg' | 'gif' } | null => {
  const m = src.match(/^data:image\/(png|jpe?g|gif);base64,(.+)$/i);
  if (!m) return null;
  const ext = m[1].toLowerCase();
  return { base64: m[2], extension: ext === 'jpg' ? 'jpeg' : (ext as 'png' | 'jpeg' | 'gif') };
};

// Vị trí (px) -> neo ô ExcelJS (chỉ số cột/hàng có phần thập phân)
const pxToAnchor = (px: number, sizes: (i: number) => number) => {
  let pos = 0;
  let i = 0;
  while (i < 100000) {
    const size = sizes(i) + 1; // +1px đường lưới như cách Luckysheet tính toạ độ
    if (pos + size > px) return i + (px - pos) / size;
    pos += size;
    i++;
  }
  return i;
};

export interface ExportOptions {
  // Workbook gốc lúc nhập file (để giữ lại cài đặt trang in)
  original?: ExcelJS.Workbook | null;
}

export const buildWorkbook = (sheets: AnyObj[], opts: ExportOptions = {}) => {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Excel Tool';
  wb.created = new Date();

  const ordered = [...sheets].sort((a, b) => Number(a.order ?? 0) - Number(b.order ?? 0));
  const usedNames = new Set<string>();
  // Ảnh giống nhau (logo / watermark lặp mỗi trang) chỉ lưu 1 lần trong file
  const mediaIds = new Map<string, number>();

  for (const sheet of ordered) {
    let name = String(sheet.name || 'Sheet').replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || 'Sheet';
    while (usedNames.has(name.toLowerCase())) name = `${name.slice(0, 28)} ${usedNames.size}`;
    usedNames.add(name.toLowerCase());

    const config: AnyObj = sheet.config || {};
    const defaultColPx = Number(sheet.defaultColWidth) || 73;
    const defaultRowPx = Number(sheet.defaultRowHeight) || 19;
    const ws = wb.addWorksheet(name, {
      properties: { defaultColWidth: pxToColWidth(defaultColPx), defaultRowHeight: pxToPt(defaultRowPx) },
    });
    if (Number(sheet.hide) === 1) ws.state = 'hidden';
    const tab = colorToArgb(sheet.color);
    if (tab) ws.properties.tabColor = { argb: tab };

    const data = sheetMatrix(sheet);
    const rowCount = data.length;
    let colCount = 0;
    for (const row of data) if (row && row.length > colCount) colCount = row.length;

    // 1. Ô gộp trước (ExcelJS sao chép style ô chính sang ô phụ khi gộp) -> sau đó mới gán style từng ô
    for (const m of Object.values<AnyObj>(config.merge || {})) {
      if (!m || (Number(m.rs) <= 1 && Number(m.cs) <= 1)) continue;
      try {
        ws.mergeCells(m.r + 1, m.c + 1, m.r + Number(m.rs), m.c + Number(m.cs));
      } catch (err) {
        console.warn('Bỏ qua ô gộp không hợp lệ', m, err);
      }
    }

    // 2. Giá trị + định dạng từng ô
    const borders = buildBorderMap(config.borderInfo);
    const links: AnyObj = sheet.hyperlink || {};
    for (let r = 0; r < rowCount; r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (!cell || typeof cell !== 'object') continue;
        const xc = ws.getCell(r + 1, c + 1);
        const isSlave = cell.mc && (cell.mc.r !== r || cell.mc.c !== c);
        if (!isSlave) {
          const value = toExcelValue(cell, links[`${r}_${c}`]);
          if (value !== null) xc.value = value;
        }
        const font = buildFont(cell);
        const isInline = cell.ct && cell.ct.t === 'inlineStr';
        if (font && !isInline) xc.font = font;
        else if (isInline) {
          const f = buildFont({ ff: cell.ff, fs: cell.fs, fc: cell.fc });
          if (f) xc.font = f;
        }
        const alignment = buildAlignment(cell);
        if (alignment) xc.alignment = alignment;
        const bg = colorToArgb(cell.bg);
        if (bg && bg !== 'FFFFFFFF') xc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        const fa = cell.ct?.fa;
        if (fa && fa !== 'General') xc.numFmt = fa;
      }
    }
    borders.forEach((b, key) => {
      const [r, c] = key.split('_').map(Number);
      const border: Partial<ExcelJS.Borders> = {};
      (['top', 'left', 'bottom', 'right'] as const).forEach((side) => {
        const s = b[side];
        if (s) border[side] = s;
      });
      if (Object.keys(border).length) ws.getCell(r + 1, c + 1).border = border;
    });

    // 3. Độ rộng cột, chiều cao hàng, hàng/cột ẩn
    const columnlen: AnyObj = config.columnlen || {};
    const rowlen: AnyObj = config.rowlen || {};
    const colhidden: AnyObj = config.colhidden || {};
    const rowhidden: AnyObj = config.rowhidden || {};
    const maxCol = Math.max(colCount, ...Object.keys(columnlen).map((k) => Number(k) + 1), 0);
    for (let c = 0; c < maxCol; c++) {
      const col = ws.getColumn(c + 1);
      if (columnlen[c] !== undefined) col.width = pxToColWidth(Number(columnlen[c]));
      if (colhidden[c] !== undefined) col.hidden = true;
    }
    for (const [k, px] of Object.entries(rowlen)) {
      const r = Number(k);
      if (r < rowCount + 50) ws.getRow(r + 1).height = pxToPt(Number(px));
    }
    for (const k of Object.keys(rowhidden)) ws.getRow(Number(k) + 1).hidden = true;

    // 4. Cố định dòng/cột, lưới, thu phóng
    const view: Partial<ExcelJS.WorksheetView> & AnyObj = {};
    const fz = sheet.frozen;
    if (fz && fz.type && fz.type !== 'cancel') {
      let xSplit = 0;
      let ySplit = 0;
      const focus = fz.range || {};
      if (fz.type === 'row' || fz.type === 'both') ySplit = 1;
      if (fz.type === 'column' || fz.type === 'both') xSplit = 1;
      if (fz.type === 'rangeRow' || fz.type === 'rangeBoth') ySplit = Number(focus.row_focus ?? 0) + 1;
      if (fz.type === 'rangeColumn' || fz.type === 'rangeBoth') xSplit = Number(focus.column_focus ?? 0) + 1;
      if (xSplit || ySplit) Object.assign(view, { state: 'frozen', xSplit, ySplit });
    }
    if (sheet.showGridLines === 0 || sheet.showGridLines === '0' || sheet.showGridLines === false) view.showGridLines = false;
    const zoom = Number(sheet.zoomRatio);
    if (zoom && zoom !== 1) view.zoomScale = Math.round(zoom * 100);
    if (Object.keys(view).length) ws.views = [view as ExcelJS.WorksheetView];

    // 5. Bộ lọc
    const fs = sheet.filter_select;
    if (fs && Array.isArray(fs.row) && Array.isArray(fs.column)) {
      ws.autoFilter = { from: { row: fs.row[0] + 1, column: fs.column[0] + 1 }, to: { row: fs.row[1] + 1, column: fs.column[1] + 1 } };
    }

    // 6. Ảnh (logo, con dấu, watermark...)
    const colPx = (c: number) => (columnlen[c] !== undefined ? Number(columnlen[c]) : defaultColPx);
    const rowPx = (r: number) => (rowlen[r] !== undefined ? Number(rowlen[r]) : defaultRowPx);
    for (const img of Object.values<AnyObj>(sheet.images || {})) {
      const parsed = img && typeof img.src === 'string' ? dataUrlToImage(img.src) : null;
      if (!parsed) continue;
      const box = img.default || {};
      const left = Number(box.left) || 0;
      const top = Number(box.top) || 0;
      const width = Number(box.width) || Number(img.originWidth) || 50;
      const height = Number(box.height) || Number(img.originHeight) || 50;
      let imageId = mediaIds.get(img.src);
      if (imageId === undefined) {
        imageId = wb.addImage(parsed);
        mediaIds.set(img.src, imageId);
      }
      ws.addImage(imageId, {
        tl: { col: pxToAnchor(left, colPx), row: pxToAnchor(top, rowPx) } as ExcelJS.Anchor,
        ext: { width, height },
      });
    }

    // 7. Cài đặt trang in của file gốc (khổ giấy, hướng, lề...) nếu cùng tên sheet
    const orig = opts.original?.worksheets.find((w) => w.name === sheet.name);
    if (orig) {
      const { printArea, printTitlesRow, printTitlesColumn, ...pageSetup } = orig.pageSetup || ({} as AnyObj);
      void printArea;
      void printTitlesRow;
      void printTitlesColumn;
      ws.pageSetup = { ...ws.pageSetup, ...pageSetup };
      if (orig.headerFooter) ws.headerFooter = { ...orig.headerFooter };
    }
  }

  if (wb.worksheets.length === 0) wb.addWorksheet('Sheet1');
  return wb;
};

export const exportWorkbookBuffer = async (sheets: AnyObj[], opts: ExportOptions = {}) => {
  const wb = buildWorkbook(sheets, opts);
  return wb.xlsx.writeBuffer();
};
