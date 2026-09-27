// ==========================================
// NHẬN DIỆN CẤU TRÚC BẢNG HOÁ ĐƠN VAT
// (dòng tiêu đề STT / Tên hàng / ĐVT / Số lượng / Đơn giá / Thành tiền,
//  dòng đánh số cột 1 2 3 4 5 6=4x5, dòng số trang, dòng tổng cộng...)
// ==========================================
import { SheetData, SheetRow, foldText, getCellNumber, getCellText, isRowBlank, rowTexts } from './sheetText';

export type ColumnRole = 'stt' | 'name' | 'unit' | 'qty' | 'price' | 'amount';

export type InvoiceColumns = Record<ColumnRole, number>;

export interface InvoiceHeader {
  row: number;
  cols: InvoiceColumns;
  labels: Record<ColumnRole, string>;
  // Dòng đánh số thứ tự cột ngay dưới tiêu đề (1 | 2 | 3 | 4 | 5 | 6 = 4 x 5), -1 nếu không có
  indexRow: number;
}

export type RowKind = 'blank' | 'header' | 'index' | 'page' | 'total' | 'data' | 'group' | 'other';

const ROLE_ORDER: ColumnRole[] = ['stt', 'price', 'unit', 'qty', 'amount', 'name'];

// Nhận diện vai trò của 1 ô tiêu đề (đã bỏ dấu, viết thường)
export const detectRole = (folded: string): ColumnRole | null => {
  if (folded.length === 0 || folded.length > 80) return null;
  for (const role of ROLE_ORDER) {
    switch (role) {
      case 'stt':
        if (/^stt\b/.test(folded) || /^so tt\b/.test(folded) || /^\(?no\.?\)?$/.test(folded)) return 'stt';
        break;
      case 'price':
        if (folded.includes('don gia') || folded.includes('unit price') || /^price\b/.test(folded)) return 'price';
        break;
      case 'unit':
        if (folded.includes('don vi tinh') || /^dvt\b/.test(folded) || /^\(?unit\)?$/.test(folded) || /\(unit\)/.test(folded)) return 'unit';
        break;
      case 'qty':
        if (folded.includes('so luong') || folded.includes('quantity') || /^(qty|sl)\b/.test(folded)) return 'qty';
        break;
      case 'amount':
        if (folded.includes('thanh tien') || folded.includes('tri gia') || /^\(?amount\)?$/.test(folded) || /\(amount\)/.test(folded)) return 'amount';
        break;
      case 'name':
        if (
          folded.includes('ten hang') ||
          folded.includes('hang hoa') ||
          folded.includes('dich vu') ||
          folded.includes('description') ||
          folded.includes('name of goods') ||
          folded.includes('mo ta') ||
          folded.includes('dien giai') ||
          folded.includes('ten san pham')
        )
          return 'name';
        break;
    }
  }
  return null;
};

const emptyCols = (): InvoiceColumns => ({ stt: -1, name: -1, unit: -1, qty: -1, price: -1, amount: -1 });

// Kiểm tra 1 dòng có phải dòng tiêu đề bảng hàng hoá không (>= 3 vai trò, bắt buộc có Số lượng hoặc Thành tiền)
export const parseHeaderRow = (row: SheetRow): { cols: InvoiceColumns; labels: Record<ColumnRole, string> } | null => {
  if (!row) return null;
  const cols = emptyCols();
  const labels: Record<ColumnRole, string> = { stt: '', name: '', unit: '', qty: '', price: '', amount: '' };
  let found = 0;
  for (let c = 0; c < row.length; c++) {
    const text = getCellText(row[c]).trim();
    if (!text) continue;
    const role = detectRole(foldText(text));
    if (role && cols[role] === -1) {
      cols[role] = c;
      labels[role] = text.replace(/\s+/g, ' ').trim();
      found++;
    }
  }
  if (found < 3 || (cols.qty === -1 && cols.amount === -1)) return null;
  return { cols, labels };
};

const INDEX_CELL_RE = /^\(?\d{1,2}\)?$/;
const INDEX_FORMULA_RE = /^\(?\d{1,2}\)?\s*=\s*\(?\d{1,2}\)?\s*[x×*]\s*\(?\d{1,2}\)?$/i;

// Dòng đánh số thứ tự cột: 1 | 2 | 3 | 4 | 5 | 6 = 4 x 5  (hoặc (1) (2) (3)...)
export const isIndexRow = (row: SheetRow) => {
  const texts = rowTexts(row);
  if (texts.length < 2) return false;
  let hasSequence = 0;
  for (const { text } of texts) {
    if (INDEX_FORMULA_RE.test(text)) continue;
    if (!INDEX_CELL_RE.test(text)) return false;
    hasSequence++;
  }
  return hasSequence >= 2;
};

const PAGE_RE = [
  /^(trang|page)\s*\d+\s*(\/|of|trên|tren)\s*\d+$/i,
  /tiếp theo trang trước/i,
  /tiep theo trang truoc/i,
  /^-\s*\d+\s*-$/,
];

// Dòng đánh số trang do chuyển PDF -> Excel sinh ra ("Trang 1/11", "tiếp theo trang trước - trang 2/11")
export const isPageMarkerRow = (row: SheetRow) => {
  const texts = rowTexts(row);
  if (texts.length === 0 || texts.length > 2) return false;
  const joined = texts.map((t) => t.text).join(' ').trim();
  if (joined.length > 80) return false;
  return PAGE_RE.some((re) => re.test(joined));
};

const TOTAL_PREFIXES = [
  'tong cong',
  'tong so luong',
  'total quantity',
  'cong tien hang',
  'thue suat',
  'tien thue',
  'tong tien',
  'so tien viet bang chu',
  'ty gia',
  'total',
  'sub total',
  'subtotal',
  'grand total',
  'vat rate',
  'vat amount',
  'amount in words',
];

// Nhãn tổng (dựa vào ô có chữ đầu tiên trên dòng)
export const getTotalLabel = (row: SheetRow): string | null => {
  const texts = rowTexts(row);
  if (texts.length === 0) return null;
  const label = foldText(texts[0].text);
  return TOTAL_PREFIXES.some((p) => label.startsWith(p)) ? label : null;
};

// Tìm tất cả dòng tiêu đề bảng trong sheet (dòng đầu tiên là tiêu đề gốc, các dòng sau là tiêu đề lặp lại theo trang)
export const findInvoiceHeaders = (sheetData: SheetData, r1 = 0, r2 = sheetData.length - 1): InvoiceHeader[] => {
  const headers: InvoiceHeader[] = [];
  for (let r = Math.max(0, r1); r <= Math.min(r2, sheetData.length - 1); r++) {
    const parsed = parseHeaderRow(sheetData[r]);
    if (!parsed) continue;
    let indexRow = -1;
    // Dòng đánh số cột có thể nằm ngay dưới (bỏ qua tối đa 1 dòng tiêu đề phụ)
    for (let k = r + 1; k <= Math.min(r + 2, sheetData.length - 1); k++) {
      if (isIndexRow(sheetData[k])) {
        indexRow = k;
        break;
      }
    }
    headers.push({ row: r, cols: parsed.cols, labels: parsed.labels, indexRow });
  }
  return headers;
};

// Phân loại 1 dòng theo cột của bảng hoá đơn
export const classifyRow = (sheetData: SheetData, r: number, cols: InvoiceColumns): RowKind => {
  const row = sheetData[r];
  if (isRowBlank(row, r)) return 'blank';
  if (parseHeaderRow(row)) return 'header';
  if (isIndexRow(row)) return 'index';
  if (isPageMarkerRow(row)) return 'page';
  if (getTotalLabel(row)) return 'total';

  const num = (c: number) => (c >= 0 && row ? getCellNumber(row[c]) : null);
  const qty = num(cols.qty);
  const amount = num(cols.amount);
  const price = num(cols.price);
  const nameText = cols.name >= 0 && row ? getCellText(row[cols.name]).trim() : '';
  const sttNum = num(cols.stt);

  if ((qty !== null || amount !== null) && (nameText !== '' || sttNum !== null)) return 'data';
  if (nameText !== '' && qty === null && amount === null && price === null) return 'group';
  return 'other';
};
