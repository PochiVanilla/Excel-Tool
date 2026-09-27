// ==========================================
// CÁC HÀM HELPER HỖ TRỢ ĐỌC CHỮ / SỐ TRONG Ô LUCKYSHEET
// (Thuần TypeScript, không phụ thuộc DOM -> dùng chung cho Cắt gộp, Pivot, Xuất file)
// ==========================================

export type CompanyMode = 'NORMAL' | 'RPAC';

// Dữ liệu Luckysheet (thư viện JS thuần từ CDN) không có kiểu cố định -> gom về 1 alias duy nhất
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type LuckyValue = any;

export interface LuckyCell {
  v?: LuckyValue;
  m?: LuckyValue;
  f?: string;
  ct?: { t?: string; fa?: string; s?: { v?: string; [k: string]: LuckyValue }[] };
  mc?: { r: number; c: number; rs?: number; cs?: number };
  [key: string]: LuckyValue;
}

export type SheetRow = (LuckyCell | null | undefined)[] | null | undefined;
export type SheetData = SheetRow[];

export const getColLetter = (colIndex: number) => {
  let letter = '';
  while (colIndex >= 0) {
    letter = String.fromCharCode((colIndex % 26) + 65) + letter;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return letter;
};

export const rangeName = (r1: number, c1: number, r2: number, c2: number) =>
  r1 === r2 && c1 === c2
    ? `${getColLetter(c1)}${r1 + 1}`
    : `${getColLetter(c1)}${r1 + 1}:${getColLetter(c2)}${r2 + 1}`;

const HYPERLINK_RE = /HYPERLINK\s*\(\s*["'].*?["']\s*[,;]\s*["'](.*?)["']\s*\)/i;

// Đọc chữ hiển thị của ô (hỗ trợ inlineStr nhiều đoạn định dạng)
// Chế độ R-PAC: bóc nội dung hiển thị khỏi công thức HYPERLINK và loại bỏ URL/domain
export const getCellText = (cell: LuckyCell | null | undefined, mode: CompanyMode = 'NORMAL'): string => {
  if (!cell) return '';
  let text = '';
  if (cell.ct && cell.ct.t === 'inlineStr' && cell.ct.s) {
    text = cell.ct.s.map((s) => s.v ?? '').join('');
  } else if (cell.m !== undefined && cell.m !== null) {
    text = cell.m.toString();
  } else if (cell.v !== undefined && cell.v !== null) {
    text = cell.v.toString();
  }

  if (mode === 'NORMAL') return text;

  const formula = cell.f || '';
  if (formula.toUpperCase().includes('HYPERLINK')) {
    const match = formula.match(HYPERLINK_RE);
    if (match && match[1]) return match[1];
  }
  if (text.toUpperCase().includes('HYPERLINK')) {
    const match = text.match(HYPERLINK_RE);
    if (match && match[1]) text = match[1];
  }
  text = text.replace(/https?:\/\/[^\s]+/gi, '');
  text = text.replace(/www\.[^\s]+/gi, '');
  return text;
};

// Chuẩn hoá để so khớp mẫu: viết thường, bỏ khoảng trắng, số và ký tự đặc biệt
// (giữ nguyên quy tắc so khớp cũ của công cụ cắt gộp)
export const cleanForMatch = (text: string) =>
  text.toLowerCase().replace(/\s+/g, '').replace(/[0-9.,:\/\\#\-]/g, '');

// Bỏ dấu tiếng Việt + viết thường + gộp khoảng trắng -> dùng để nhận diện từ khoá (so luong, don gia...)
export const foldText = (text: string) =>
  text
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/[\s ]+/g, ' ')
    .trim();

// Chuyển chữ dạng số (có thể có phân cách nghìn, đơn vị tiền tệ) thành số.
// Trả về null nếu chuỗi không phải một con số thuần tuý (vd: "6 = 4 x 5", "1/11", "26,260 VND/USD")
export const parseLooseNumber = (raw: string): number | null => {
  let s = raw.replace(/[\s ]/g, '');
  if (s === '') return null;
  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1);
  }
  // Cho phép ký hiệu tiền tệ ở đầu/cuối (USD, VND, đ, $, %)
  s = s.replace(/^(usd|vnd|\$)/i, '').replace(/(usd|vnd|đ|₫|\$)$/i, '');
  if (!/^[-+]?[\d.,]+$/.test(s) || !/\d/.test(s)) return null;
  if (s.startsWith('-')) {
    negative = !negative;
    s = s.slice(1);
  } else if (s.startsWith('+')) {
    s = s.slice(1);
  }

  const lastDot = s.lastIndexOf('.');
  const lastComma = s.lastIndexOf(',');
  if (lastDot !== -1 && lastComma !== -1) {
    // Có cả 2 loại dấu: dấu xuất hiện sau cùng là dấu thập phân
    if (lastDot > lastComma) s = s.replace(/,/g, '');
    else s = s.replace(/\./g, '').replace(',', '.');
  } else if (lastComma !== -1) {
    const parts = s.split(',');
    if (parts.length > 2 || (parts.length === 2 && parts[1].length === 3)) s = s.replace(/,/g, '');
    else s = s.replace(',', '.');
  } else if (lastDot !== -1) {
    const parts = s.split('.');
    if (parts.length > 2) s = s.replace(/\./g, '');
  }

  const n = Number(s);
  if (!isFinite(n)) return null;
  return negative ? -n : n;
};

// Đọc giá trị số của ô (ưu tiên giá trị gốc v, sau đó mới đọc chữ hiển thị)
export const getCellNumber = (cell: LuckyCell | null | undefined): number | null => {
  if (!cell) return null;
  if (typeof cell.v === 'number' && isFinite(cell.v)) return cell.v;
  if (typeof cell.v === 'boolean') return null;
  const text = getCellText(cell).trim();
  if (text === '') return null;
  return parseLooseNumber(text);
};

// Dòng trống thật sự: không có chữ và không nằm trong vùng gộp ô dọc (ô gộp nhiều hàng)
export const isRowBlank = (row: SheetRow, rowIndex: number) => {
  if (!row) return true;
  for (let c = 0; c < row.length; c++) {
    const cell = row[c];
    if (!cell) continue;
    if (cell.mc && cell.mc.r !== rowIndex) return false;
    if (cell.mc && (cell.mc.rs ?? 1) > 1) return false;
    if (getCellText(cell).trim() !== '') return false;
  }
  return true;
};

// Danh sách chữ khác rỗng trên một dòng (kèm chỉ số cột)
export const rowTexts = (row: SheetRow, c1 = 0, c2?: number, mode: CompanyMode = 'NORMAL') => {
  const out: { c: number; text: string }[] = [];
  if (!row) return out;
  const end = c2 === undefined ? row.length - 1 : Math.min(c2, row.length - 1);
  for (let c = c1; c <= end; c++) {
    const t = getCellText(row[c], mode).trim();
    if (t !== '') out.push({ c, text: t });
  }
  return out;
};

export const getRangeText = (row: SheetRow, c1: number, c2: number, mode: CompanyMode = 'NORMAL') => {
  if (!row) return '';
  const textArray: string[] = [];
  for (let i = c1; i <= c2; i++) {
    textArray.push(getCellText(row[i], mode).trim().replace(/\r?\n/g, ' '));
  }
  return textArray.join(' | ');
};

export const getCleanedRangeText = (row: SheetRow, c1: number, c2: number, mode: CompanyMode = 'NORMAL') => {
  if (!row) return '';
  const textArray: string[] = [];
  for (let i = c1; i <= c2; i++) {
    textArray.push(cleanForMatch(getCellText(row[i], mode)));
  }
  return textArray.join('|');
};

export const escapeHtml = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// Gom danh sách chỉ số dòng thành các khối liên tiếp [start, end]
export const toBlocks = (rows: number[]) => {
  const sorted = Array.from(new Set(rows)).sort((a, b) => a - b);
  const blocks: { start: number; end: number }[] = [];
  for (const r of sorted) {
    const last = blocks[blocks.length - 1];
    if (last && r === last.end + 1) last.end = r;
    else blocks.push({ start: r, end: r });
  }
  return blocks;
};

export const formatRowList = (rows: number[]) =>
  toBlocks(rows)
    .map((b) => (b.start === b.end ? `${b.start + 1}` : `${b.start + 1}–${b.end + 1}`))
    .join(', ');
