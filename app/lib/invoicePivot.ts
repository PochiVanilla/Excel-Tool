// ==========================================
// PIVOT (Polytex - Vải, Chỉ may A&E, Hoá đơn chung, Bảng dữ liệu thủ công)
// Quy trình: nhận diện bảng -> đọc từng dòng hàng -> tách Mã hàng / Mô tả / Màu -> gom nhóm như Excel PivotTable
// Không nhận ra bảng hoá đơn -> Pivot thủ công: dòng tiêu đề của vùng chọn là tên trường, mỗi cột là 1 trường
// ==========================================
import { SheetData, SheetRow, foldText, getCellNumber, getCellText, getColLetter, isRowBlank, rowTexts } from './sheetText';
import {
  ColumnRole,
  InvoiceHeader,
  RowKind,
  classifyRow,
  detectRole,
  findInvoiceHeaders,
  getTotalLabel,
  isIndexRow,
  isPageMarkerRow,
  parseHeaderRow,
} from './invoiceTable';

type KnownFieldId = 'stt' | 'name' | 'item' | 'artNo' | 'description' | 'color' | 'unit' | 'price' | 'quantity' | 'amount';
// Trường của Pivot thủ công: c<chỉ số cột>, vd c2 = cột C
export type PivotFieldId = KnownFieldId | `c${number}`;
export type PivotAgg = 'sum' | 'count' | 'avg' | 'min' | 'max' | 'wavg';
export type InvoiceProfile = 'fabric' | 'thread' | 'generic' | 'table';
// auto: tự nhận diện bảng hoá đơn (không được thì chuyển sang thủ công) · table: luôn Pivot thủ công theo dòng tiêu đề
export type PivotMode = 'auto' | 'table';
export type PivotSort = 'source' | 'asc' | 'desc';

export interface PivotFieldDef {
  id: PivotFieldId;
  label: string;
  kind: 'text' | 'number';
  hint?: string;
}

export interface PivotValueField {
  id: PivotFieldId;
  agg: PivotAgg;
}

export interface PivotLayout {
  rows: PivotFieldId[];
  values: PivotValueField[];
  sort: PivotSort;
}

export interface PivotRecord {
  row: number;
  fields: Partial<Record<PivotFieldId, string | number | null>>;
}

export interface PivotSource {
  profile: InvoiceProfile;
  seller: string;
  header: InvoiceHeader;
  fields: PivotFieldDef[];
  records: PivotRecord[];
  invoiceTotals: { quantity: number | null; amount: number | null };
  formats: Partial<Record<PivotFieldId, string>>;
  skipped: Partial<Record<RowKind, number>>;
  range: { r1: number; r2: number };
}

export interface PivotColumn {
  id: PivotFieldId;
  label: string;
  role: 'row' | 'value';
  agg?: PivotAgg;
  format?: string;
}

export interface PivotResultRow {
  cells: (string | number | null)[];
  sourceRows: number[];
}

export interface PivotResult {
  columns: PivotColumn[];
  rows: PivotResultRow[];
  grandTotal: (string | number | null)[];
  totals: { quantity: number | null; amount: number | null };
  check: { quantity: 'ok' | 'diff' | 'na'; amount: 'ok' | 'diff' | 'na' };
}

export const PROFILE_LABELS: Record<InvoiceProfile, string> = {
  fabric: 'Polytex · Vải (Fabric)',
  thread: 'Chỉ may (Thread · A&E)',
  generic: 'Hoá đơn chung',
  table: 'Bảng dữ liệu (Pivot thủ công)',
};

export const AGG_LABELS: Record<PivotAgg, string> = {
  sum: 'Sum',
  count: 'Count',
  avg: 'Average',
  min: 'Min',
  max: 'Max',
  wavg: 'Bình quân gia quyền',
};

const BASE_FIELDS: Record<KnownFieldId, Omit<PivotFieldDef, 'id'>> = {
  stt: { label: 'STT', kind: 'number' },
  name: { label: 'Tên hàng (đầy đủ)', kind: 'text', hint: 'Toàn bộ nội dung ô tên hàng' },
  item: { label: 'Mã hàng (Item)', kind: 'text' },
  artNo: { label: 'Art No.', kind: 'text' },
  description: { label: 'Mô tả', kind: 'text' },
  color: { label: 'Màu', kind: 'text' },
  unit: { label: 'ĐVT', kind: 'text' },
  price: { label: 'Đơn giá', kind: 'number' },
  quantity: { label: 'Số lượng', kind: 'number' },
  amount: { label: 'Thành tiền', kind: 'number' },
};

// Trường nào có ý nghĩa với từng loại hoá đơn
const PROFILE_FIELDS: Record<Exclude<InvoiceProfile, 'table'>, KnownFieldId[]> = {
  fabric: ['item', 'artNo', 'description', 'color', 'name', 'unit', 'price', 'quantity', 'amount', 'stt'],
  thread: ['name', 'item', 'color', 'unit', 'price', 'quantity', 'amount', 'stt'],
  generic: ['name', 'item', 'description', 'unit', 'price', 'quantity', 'amount', 'stt'],
};

const PROFILE_FIELD_LABELS: Record<Exclude<InvoiceProfile, 'table'>, Partial<Record<KnownFieldId, { label: string; hint?: string }>>> = {
  fabric: {
    item: { label: 'Item', hint: 'Mã nằm trên cùng của ô tên hàng, giữ nguyên, vd: #1002592' },
    description: { label: 'Mô tả', hint: 'Thành phần + khổ + định lượng, vd: 100% POLYESTER ... 195.00 g/m2' },
    color: { label: 'Màu', hint: 'Dòng cuối của ô, vd: 00A BLACK' },
  },
  thread: {
    name: { label: 'Tên hàng (đầy đủ)', hint: 'vd: CHỈ P-CORE ECO100 TEX24 45/2 5000M 145547' },
    item: { label: 'Loại chỉ', hint: 'Tên hàng bỏ mã màu, vd: CHỈ P-CORE ECO100 TEX24 45/2 5000M' },
    color: { label: 'Mã màu', hint: 'Mã cuối tên hàng, vd: 145547, W32109' },
  },
  generic: {
    item: { label: 'Mã / Tên hàng', hint: 'Dòng đầu tiên của ô tên hàng' },
    description: { label: 'Mô tả chi tiết', hint: 'Các dòng còn lại của ô tên hàng' },
  },
};

const baseField = (id: PivotFieldId) => (id in BASE_FIELDS ? BASE_FIELDS[id as KnownFieldId] : undefined);

// Bố cục mặc định theo nghiệp vụ: gom theo mặt hàng + đơn giá, cộng dồn Số lượng & Thành tiền
export const defaultLayout = (profile: InvoiceProfile, fields: PivotFieldDef[] = []): PivotLayout => {
  const values: PivotValueField[] = [
    { id: 'quantity', agg: 'sum' },
    { id: 'amount', agg: 'sum' },
  ];
  if (profile === 'fabric') return { rows: ['item', 'description', 'price'], values, sort: 'source' };
  if (profile === 'table') {
    // Pivot thủ công: 2 cột chữ đầu tiên (+ Đơn giá) vào Rows, cộng dồn Số lượng / Thành tiền (không có thì các cột số khác)
    const has = (id: PivotFieldId) => fields.some((f) => f.id === id);
    const rows: PivotFieldId[] = fields
      .filter((f) => f.kind === 'text')
      .slice(0, 2)
      .map((f) => f.id);
    if (has('price')) rows.push('price');
    let vals = values.filter((v) => has(v.id));
    if (vals.length === 0) {
      vals = fields.filter((f) => f.kind === 'number' && f.id !== 'stt' && f.id !== 'price').map((f) => ({ id: f.id, agg: 'sum' as const }));
    }
    return { rows, values: vals, sort: 'source' };
  }
  return { rows: ['name', 'price'], values, sort: 'source' };
};

export const defaultAggFor = (id: PivotFieldId, kind?: PivotFieldDef['kind']): PivotAgg => {
  if (id === 'price') return 'wavg';
  if ((kind ?? baseField(id)?.kind) === 'number' && id !== 'stt') return 'sum';
  return 'count';
};

export const valueHeader = (label: string, agg: PivotAgg) =>
  agg === 'wavg' ? `${label} (BQ gia quyền)` : `${AGG_LABELS[agg]} of ${label}`;

// ==========================================
// TÁCH THÀNH PHẦN TÊN HÀNG THEO LOẠI HOÁ ĐƠN
// ==========================================
const CODE_LINE_RE = /^#?\s*[A-Z0-9][\w\-./]*$/i;

// Mô tả Polytex: từ con số đứng trước dấu % đầu tiên đến hết "g/m2"
// (tương đương công thức Excel cũ, giữ nguyên nghiệp vụ Mô tả của Pivot Polytex)
export const extractFabricDescription = (text: string) => {
  const percentIndex = text.indexOf('%');
  if (percentIndex === -1) return '';
  let startIndex = percentIndex;
  while (startIndex > 0 && /[\d.]/.test(text[startIndex - 1])) startIndex--;
  const gm2Match = text.slice(percentIndex).match(/g\s*\/\s*m2/i);
  if (!gm2Match || gm2Match.index === undefined) return '';
  const endIndex = percentIndex + gm2Match.index + gm2Match[0].length;
  return text.substring(startIndex, endIndex).replace(/\s+/g, ' ').trim();
};

export interface ParsedName {
  item: string;
  artNo: string;
  description: string;
  color: string;
}

const FABRIC_DESC_START_RE = /\s(?=FEPV\b|Art\s*No|Knitt|Woven|[\w ]*Fabric\b|\d+(?:\.\d+)?\s*%)/i;

export const parseItemName = (raw: string, profile: InvoiceProfile): ParsedName => {
  const text = raw.replace(/\r\n?/g, '\n').trim();
  const lines = text.split('\n').map((l) => l.trim()).filter((l) => l !== '');
  const first = lines[0] || '';
  const oneLine = lines.join(' ');

  if (profile === 'fabric') {
    // Nghiệp vụ Polytex: Item = dòng trên cùng của ô tên hàng, giữ nguyên (vd "#1002592", "#628165-1", "1002592"...)
    // Ô không xuống dòng (mã dính liền mô tả) -> cắt trước phần mô tả "FEPV Art No.: ... Knitting Fabric 100% ..."
    let item = first;
    const cut = first.search(FABRIC_DESC_START_RE);
    if (cut > 0) item = first.slice(0, cut).trim();
    // Dòng trên cùng đã là mô tả (thành phần / khổ vải) -> không có mã hàng
    if (/%|g\s*\/\s*m2|art\s*no/i.test(item)) item = '';
    const art = text.match(/Art\s*No\.?\s*:?\s*([A-Z0-9][A-Z0-9\-]*[A-Z0-9])/i);
    const description = extractFabricDescription(text) || (lines[1] || '');
    const last = lines.length >= 3 ? lines[lines.length - 1] : '';
    const color = last && !/g\s*\/\s*m2|%|art\s*no/i.test(last) ? last : '';
    return { item, artNo: art ? art[1] : '', description, color };
  }

  if (profile === 'thread') {
    const tokens = oneLine.split(/\s+/);
    const lastToken = tokens[tokens.length - 1] || '';
    if (tokens.length >= 2 && /^[A-Z]{0,2}\d{3,}[A-Z]?$/i.test(lastToken)) {
      return { item: tokens.slice(0, -1).join(' '), artNo: '', description: tokens.slice(0, -1).join(' '), color: lastToken };
    }
    return { item: oneLine, artNo: '', description: oneLine, color: '' };
  }

  // Hoá đơn chung: dòng đầu là tên/mã, các dòng sau là mô tả (tên 1 dòng thì Mô tả để trống, không lặp lại Item)
  return { item: first, artNo: '', description: lines.slice(1).join(' '), color: '' };
};

type InvoiceKind = Exclude<InvoiceProfile, 'table'>;

const detectProfile = (names: string[], seller: string): InvoiceKind => {
  if (names.length === 0) return 'generic';
  const foldedSeller = foldText(seller);
  const fabricHits = names.filter((n) => /g\s*\/\s*m2/i.test(n) && n.includes('%')).length;
  // Nghiệp vụ Polytex: tên hàng bắt đầu bằng mã "#1002592" (hoặc có % ... g/m2, hoặc người bán là Polytex)
  const codeHits = names.filter((n) => /^#\s*[\w-]+/.test(n.trim())).length;
  if (fabricHits >= names.length * 0.5 || codeHits >= names.length * 0.5 || (foldedSeller.includes('polytex') && codeHits > 0)) return 'fabric';
  const threadHits = names.filter((n) => /^chi\s/.test(foldText(n))).length;
  if (
    threadHits >= names.length * 0.5 ||
    foldedSeller.includes('chi may') ||
    foldedSeller.includes('american & efird') ||
    foldedSeller.includes('thread')
  )
    return 'thread';
  return 'generic';
};

// Tổng trên hoá đơn để đối chiếu: lấy từ dòng "Tổng cộng" / "Cộng tiền hàng" (trước thuế) của từng hoá đơn
const reconcileTotals = (sheetData: SheetData, totalsRows: number[], qtyCol: number, amountCol: number) => {
  let totalQty: number | null = null;
  let totalAmount: number | null = null;
  let group: { qty: number | null; amountSub: number | null; amountTotal: number | null; lastRow: number } | null = null;
  const flush = () => {
    if (!group) return;
    const amount = group.amountSub ?? group.amountTotal;
    if (group.qty !== null) totalQty = (totalQty ?? 0) + group.qty;
    if (amount !== null) totalAmount = (totalAmount ?? 0) + amount;
    group = null;
  };
  for (const tr of totalsRows) {
    // Nhãn có thể nằm ở bất kỳ ô nào trên dòng (vd: "Tỷ giá ... | Cộng tiền hàng (Total amount): | 2742.78")
    const labels = rowTexts(sheetData[tr]).map((x) => foldText(x.text));
    const isSub = labels.some((l) => l.startsWith('cong tien hang') || l.startsWith('sub'));
    const isTotal = labels.some(
      (l) =>
        (l.startsWith('tong cong') && !l.includes('thanh toan')) ||
        l.startsWith('tong so luong') ||
        l.startsWith('total quantity') ||
        l === 'total' ||
        l.startsWith('total:') ||
        l.startsWith('total ('),
    );
    if (!isSub && !isTotal) continue;
    if (group && tr - group.lastRow > 6) flush();
    if (!group) group = { qty: null, amountSub: null, amountTotal: null, lastRow: tr };
    const row = sheetData[tr];
    const q = qtyCol >= 0 && row ? getCellNumber(row[qtyCol]) : null;
    const a = amountCol >= 0 && row ? getCellNumber(row[amountCol]) : null;
    if (q !== null && group.qty === null) group.qty = q;
    if (a !== null) {
      if (isSub && group.amountSub === null) group.amountSub = a;
      if (isTotal && group.amountTotal === null) group.amountTotal = a;
    }
    group.lastRow = tr;
  }
  flush();
  return { quantity: totalQty as number | null, amount: totalAmount as number | null };
};

// ==========================================
// ĐỌC DỮ LIỆU NGUỒN CHO PIVOT
// ==========================================
type SourceResult = { source: PivotSource } | { error: string };

export const buildPivotSource = (
  sheetData: SheetData,
  bounds?: { r1: number; r2: number },
  opts: { mode?: PivotMode; forceProfile?: InvoiceKind } = {},
): SourceResult => {
  if (opts.mode === 'table') return buildTableSource(sheetData, bounds);
  const invoice = buildInvoiceSource(sheetData, bounds, opts.forceProfile);
  if ('source' in invoice) return invoice;
  // Không nhận ra bảng hoá đơn (vd file đã trích sẵn, tiêu đề cột khác) -> Pivot thủ công theo dòng tiêu đề
  const table = buildTableSource(sheetData, bounds);
  return 'source' in table ? table : invoice;
};

// ---- Pivot theo bảng hoá đơn VAT (nhận diện cột STT / Tên hàng / ĐVT / Số lượng / Đơn giá / Thành tiền) ----
const buildInvoiceSource = (sheetData: SheetData, bounds?: { r1: number; r2: number }, forceProfile?: InvoiceKind): SourceResult => {
  const r1 = bounds ? bounds.r1 : 0;
  const r2 = bounds ? bounds.r2 : sheetData.length - 1;
  let headers = findInvoiceHeaders(sheetData, r1, r2);
  // Vùng chọn không chứa dòng tiêu đề -> tìm tiêu đề ở phía trên vùng chọn
  if (headers.length === 0 && bounds) {
    const above = findInvoiceHeaders(sheetData, 0, r1);
    if (above.length > 0) headers = [above[above.length - 1]];
  }
  if (headers.length === 0) {
    return {
      error:
        'Không tìm thấy dòng tiêu đề bảng hàng hoá (STT, Tên hàng, ĐVT, Số lượng, Đơn giá, Thành tiền). Hãy kiểm tra lại sheet hoặc quét chọn đúng bảng dữ liệu.',
    };
  }
  const header = headers[0];
  const { cols } = header;
  if (cols.name === -1) return { error: 'Không xác định được cột Tên hàng hoá / Mô tả trong dòng tiêu đề.' };

  const seller = (() => {
    for (let r = 0; r < Math.min(sheetData.length, header.row); r++) {
      const texts = rowTexts(sheetData[r]);
      if (texts.length > 0) return texts[0].text.split(/\r?\n/)[0].trim();
    }
    return '';
  })();

  type RawRecord = { row: number; name: string; lineA: string };
  const raws: RawRecord[] = [];
  const skipped: Partial<Record<RowKind, number>> = {};
  const formats: PivotSource['formats'] = {};
  const totalsRows: number[] = [];
  let lineA = '';
  const start = Math.max(header.row + 1, r1);

  for (let r = start; r <= r2 && r < sheetData.length; r++) {
    const kind = classifyRow(sheetData, r, cols);
    if (kind !== 'data') {
      skipped[kind] = (skipped[kind] || 0) + 1;
      if (kind === 'group') lineA = getCellText(sheetData[r]![cols.name]).trim();
      if (kind === 'total') {
        totalsRows.push(r);
        lineA = '';
      }
      continue;
    }
    const row = sheetData[r]!;
    const name = getCellText(row[cols.name]).trim();
    raws.push({ row: r, name, lineA });
    if (!formats.quantity && cols.qty >= 0) formats.quantity = row[cols.qty]?.ct?.fa;
    if (!formats.price && cols.price >= 0) formats.price = row[cols.price]?.ct?.fa;
    if (!formats.amount && cols.amount >= 0) formats.amount = row[cols.amount]?.ct?.fa;
  }

  if (raws.length === 0) return { error: 'Không tìm thấy dòng hàng hoá nào có Số lượng / Thành tiền bên dưới dòng tiêu đề.' };

  const profile = forceProfile ?? detectProfile(raws.map((x) => x.name || x.lineA), seller);

  const records: PivotRecord[] = raws.map(({ row: r, name, lineA: carried }) => {
    const row = sheetData[r]!;
    // Kiểu "Line A": tên hàng nằm ở dòng phía trên, dòng số liệu để trống cột tên
    const ownText = name || carried;
    const parsed = parseItemName(ownText, profile);
    // Dòng Line A chỉ chứa mã hàng (vd "#1002592") -> dùng làm Mã hàng cho các dòng số liệu bên dưới
    if (!parsed.item && carried) {
      const top = profile === 'fabric' ? parseItemName(carried, 'fabric').item : CODE_LINE_RE.test(carried) ? carried.replace(/\s+/g, '') : '';
      if (top && top.length <= 30) parsed.item = top;
    }
    // Polytex không tìm thấy mã # thì để trống Item (không lấy dòng mô tả làm Item)
    if (!parsed.item && profile !== 'fabric') parsed.item = ownText.split(/\r?\n/)[0].trim();

    const num = (c: number) => (c >= 0 ? getCellNumber(row[c]) : null);
    const unitText = cols.unit >= 0 ? getCellText(row[cols.unit]).trim() : '';
    return {
      row: r,
      fields: {
        stt: num(cols.stt),
        name: ownText.replace(/\s*\r?\n\s*/g, ' ').trim(),
        item: parsed.item,
        artNo: parsed.artNo,
        description: parsed.description,
        color: parsed.color,
        unit: unitText,
        price: num(cols.price),
        quantity: num(cols.qty),
        amount: num(cols.amount),
      },
    };
  });

  const invoiceTotals = reconcileTotals(sheetData, totalsRows, cols.qty, cols.amount);

  const fields: PivotFieldDef[] = PROFILE_FIELDS[profile].map((id) => {
    const base = BASE_FIELDS[id];
    const override = PROFILE_FIELD_LABELS[profile][id];
    let label = override?.label ?? base.label;
    // Dùng đúng tiêu đề gốc của hoá đơn cho các cột số liệu
    if (id === 'quantity' && header.labels.qty) label = header.labels.qty;
    if (id === 'price' && header.labels.price) label = header.labels.price;
    if (id === 'amount' && header.labels.amount) label = header.labels.amount;
    if (id === 'unit' && header.labels.unit) label = header.labels.unit;
    return { id, label, kind: base.kind, hint: override?.hint ?? base.hint };
  });

  return {
    source: {
      profile,
      seller,
      header,
      fields,
      records,
      invoiceTotals,
      formats,
      skipped,
      range: { r1: header.row, r2: Math.min(r2, sheetData.length - 1) },
    },
  };
};

// ---- Pivot thủ công (giống Insert PivotTable của Excel): dòng tiêu đề = tên trường, mỗi cột có tiêu đề là 1 trường ----
// Cột có tên kiểu mã (Item, Mã hàng, Art, Lot, PO, Màu...) luôn là trường chữ dù toàn chữ số
const CODE_HEADER_RE = /\b(ma|code|item|art|lot|po|so hd|so hoa don|invoice|mau|color|size|style)\b/;
const ROLE_FIELD: Record<ColumnRole, KnownFieldId> = { stt: 'stt', name: 'name', unit: 'unit', qty: 'quantity', price: 'price', amount: 'amount' };

const lastUsedRow = (sheetData: SheetData) => {
  let r = sheetData.length - 1;
  while (r >= 0 && isRowBlank(sheetData[r], r)) r--;
  return r;
};

const buildTableSource = (sheetData: SheetData, bounds?: { r1: number; r2: number }): SourceResult => {
  const last = lastUsedRow(sheetData);
  const r1 = bounds ? bounds.r1 : 0;
  const r2 = Math.min(bounds ? bounds.r2 : last, last);
  const filled = (row: SheetRow) => rowTexts(row).length;

  // Dòng tiêu đề: vùng chọn -> dòng đầu tiên có >= 2 ô chữ; cả sheet -> ưu tiên dòng tiêu đề hoá đơn (STT | Số lượng | Đơn giá...)
  let hr = -1;
  if (!bounds) {
    for (let r = r1; r <= r2 && hr < 0; r++) if (parseHeaderRow(sheetData[r])) hr = r;
  }
  for (let r = r1; r <= r2 && hr < 0; r++) if (filled(sheetData[r]) >= 2) hr = r;
  if (hr < 0) return { error: 'Không tìm thấy dòng tiêu đề (dòng có tên các cột) trong vùng chọn. Hãy bôi đen bảng dữ liệu từ dòng tiêu đề.' };

  const headRow = sheetData[hr];
  const columns: { c: number; label: string; role: ColumnRole | null; folded: string }[] = [];
  for (let c = 0; headRow && c < headRow.length; c++) {
    const label = getCellText(headRow[c]).replace(/\s+/g, ' ').trim();
    if (!label) continue;
    const folded = foldText(label);
    const dup = columns.some((x) => x.label === label);
    columns.push({ c, label: dup ? `${label} (${getColLetter(c)})` : label, role: detectRole(folded), folded });
  }

  const sig = (row: SheetRow) => columns.map(({ c }) => foldText(getCellText(row?.[c]))).join('|');
  const headSig = sig(headRow);
  const skipped: Partial<Record<RowKind, number>> = {};
  const skip = (k: RowKind) => (skipped[k] = (skipped[k] || 0) + 1);
  const dataRows: number[] = [];
  const totalsRows: number[] = [];
  for (let r = hr + 1; r <= r2; r++) {
    const row = sheetData[r];
    if (isRowBlank(row, r)) {
      // Cả sheet: bảng kết thúc ở hàng trống đầu tiên (giống vùng dữ liệu liền của Excel)
      if (!bounds && dataRows.length > 0) break;
      skip('blank');
      continue;
    }
    if (sig(row) === headSig) skip('header');
    else if (isIndexRow(row)) skip('index');
    else if (isPageMarkerRow(row)) skip('page');
    else if (getTotalLabel(row)) {
      totalsRows.push(r);
      skip('total');
    } else if (columns.every(({ c }) => getCellText(row?.[c]).trim() === '')) skip('other');
    else dataRows.push(r);
  }
  if (dataRows.length === 0) return { error: `Không có dòng dữ liệu nào bên dưới dòng tiêu đề (hàng ${hr + 1}).` };

  // Kiểu trường: cột Số lượng / Đơn giá / Thành tiền / STT là số; cột khác là số nếu >= 80% ô có giá trị là số
  const used = new Set<PivotFieldId>();
  const fieldCols = columns.map((col) => {
    let kind: PivotFieldDef['kind'];
    if (col.role === 'qty' || col.role === 'price' || col.role === 'amount' || col.role === 'stt') kind = 'number';
    else if (col.role === 'name' || col.role === 'unit' || CODE_HEADER_RE.test(col.folded)) kind = 'text';
    else {
      let nonEmpty = 0;
      let nums = 0;
      for (const r of dataRows) {
        const cell = sheetData[r]?.[col.c];
        if (getCellText(cell).trim() === '') continue;
        nonEmpty++;
        if (getCellNumber(cell) !== null) nums++;
      }
      kind = nonEmpty > 0 && nums >= nonEmpty * 0.8 ? 'number' : 'text';
    }
    let id: PivotFieldId = `c${col.c}`;
    if (col.role && !used.has(ROLE_FIELD[col.role])) id = ROLE_FIELD[col.role];
    used.add(id);
    return { ...col, id, kind };
  });

  const formats: PivotSource['formats'] = {};
  const records: PivotRecord[] = dataRows.map((r) => {
    const row = sheetData[r];
    const fields: PivotRecord['fields'] = {};
    for (const f of fieldCols) {
      const cell = row?.[f.c];
      if (f.kind === 'number') {
        fields[f.id] = getCellNumber(cell);
        if (!formats[f.id] && cell?.ct?.fa) formats[f.id] = cell.ct.fa;
      } else fields[f.id] = getCellText(cell).replace(/\s*\r?\n\s*/g, ' ').trim();
    }
    return { row: r, fields };
  });

  const colOf = (id: PivotFieldId) => fieldCols.find((f) => f.id === id)?.c ?? -1;
  const header: InvoiceHeader = {
    row: hr,
    cols: { stt: colOf('stt'), name: colOf('name'), unit: colOf('unit'), qty: colOf('quantity'), price: colOf('price'), amount: colOf('amount') },
    labels: { stt: '', name: '', unit: '', qty: '', price: '', amount: '' },
    indexRow: -1,
  };
  return {
    source: {
      profile: 'table',
      seller: '',
      header,
      fields: fieldCols.map((f) => ({ id: f.id, label: f.label, kind: f.kind, hint: `Cột ${getColLetter(f.c)}` })),
      records,
      invoiceTotals: reconcileTotals(sheetData, totalsRows, header.cols.qty, header.cols.amount),
      formats,
      skipped,
      range: { r1: hr, r2: dataRows[dataRows.length - 1] },
    },
  };
};

// ==========================================
// TÍNH TOÁN PIVOT (giống Excel PivotTable - dạng Tabular)
// ==========================================
const round = (n: number, digits = 6) => {
  const f = Math.pow(10, digits);
  return Math.round(n * f) / f;
};

const aggregate = (records: PivotRecord[], vf: PivotValueField): number | null => {
  if (vf.agg === 'count') return records.filter((r) => r.fields[vf.id] != null && r.fields[vf.id] !== '').length;
  if (vf.agg === 'wavg') {
    let qty = 0;
    let amount = 0;
    for (const r of records) {
      qty += Number(r.fields.quantity) || 0;
      amount += Number(r.fields.amount) || 0;
    }
    if (qty !== 0) return round(amount / qty, 6);
    // Không có số lượng -> quay về trung bình cộng
    return aggregate(records, { id: vf.id, agg: 'avg' });
  }
  const nums = records.map((r) => r.fields[vf.id]).filter((v): v is number => typeof v === 'number');
  if (nums.length === 0) return null;
  switch (vf.agg) {
    case 'sum':
      return round(nums.reduce((a, b) => a + b, 0));
    case 'avg':
      return round(nums.reduce((a, b) => a + b, 0) / nums.length);
    case 'min':
      return Math.min(...nums);
    case 'max':
      return Math.max(...nums);
  }
  return null;
};

const compareCells = (a: (string | number | null)[], b: (string | number | null)[], n: number) => {
  for (let i = 0; i < n; i++) {
    const x = a[i];
    const y = b[i];
    if (x === y) continue;
    if (x === null || x === '') return 1;
    if (y === null || y === '') return -1;
    if (typeof x === 'number' && typeof y === 'number') return x - y;
    const cmp = String(x).localeCompare(String(y), 'vi', { numeric: true, sensitivity: 'base' });
    if (cmp !== 0) return cmp;
  }
  return 0;
};

// Bỏ ký hiệu tiền tệ khỏi mã định dạng của hoá đơn (vd "0.00 [$USD]" -> "0.00"): kết quả Pivot chỉ là số
const plainNumberFormat = (fa?: string) => {
  if (!fa) return undefined;
  const plain = fa
    .split(';')
    .map((part) =>
      part
        .replace(/\[\$[^\]]*\]/g, '')
        .replace(/"[^"]*"/g, '')
        .replace(/\\./g, '')
        .replace(/\b(usd|vnd)\b/gi, '')
        .replace(/[$€₫đ]/g, '')
        .trim(),
    )
    .join(';');
  return /[0#]/.test(plain) ? plain : undefined;
};

export const computePivot = (source: PivotSource, layout: PivotLayout): PivotResult => {
  const fieldDef = (id: PivotFieldId) => source.fields.find((f) => f.id === id);
  const fieldLabel = (id: PivotFieldId) => fieldDef(id)?.label ?? baseField(id)?.label ?? id;
  const formatFor = (id: PivotFieldId, agg?: PivotAgg) => {
    if (agg === 'count') return '#,##0';
    const own = plainNumberFormat(source.formats[id]);
    if (own && own !== 'General') return own;
    if (id === 'quantity') return '#,##0.##';
    if (id === 'price') return '#,##0.0000';
    if (id === 'amount') return '#,##0.00';
    if (id !== 'stt' && fieldDef(id)?.kind === 'number') return '#,##0.##';
    return undefined;
  };

  const columns: PivotColumn[] = [
    ...layout.rows.map((id) => ({ id, label: fieldLabel(id), role: 'row' as const, format: formatFor(id) })),
    ...layout.values.map((vf) => ({
      id: vf.id,
      label: valueHeader(fieldLabel(vf.id), vf.agg),
      role: 'value' as const,
      agg: vf.agg,
      format: formatFor(vf.id, vf.agg),
    })),
  ];

  const groups = new Map<string, { keys: (string | number | null)[]; records: PivotRecord[] }>();
  for (const rec of source.records) {
    const keys = layout.rows.map((id) => {
      const v = rec.fields[id];
      return typeof v === 'number' ? round(v, 8) : v ?? '';
    });
    const k = JSON.stringify(keys);
    let g = groups.get(k);
    if (!g) {
      g = { keys, records: [] };
      groups.set(k, g);
    }
    g.records.push(rec);
  }

  let rows: PivotResultRow[] = Array.from(groups.values()).map((g) => ({
    cells: [...g.keys, ...layout.values.map((vf) => aggregate(g.records, vf))],
    sourceRows: g.records.map((r) => r.row),
  }));

  if (layout.sort !== 'source' && layout.rows.length > 0) {
    rows = [...rows].sort((a, b) => compareCells(a.cells, b.cells, layout.rows.length));
    if (layout.sort === 'desc') rows.reverse();
  }

  const grandTotal: (string | number | null)[] = [
    ...layout.rows.map((_, i) => (i === 0 ? 'Tổng cộng (Grand Total)' : '')),
    ...layout.values.map((vf) => aggregate(source.records, vf)),
  ];
  if (layout.rows.length === 0) grandTotal.unshift('Tổng cộng (Grand Total)');

  const sumOf = (id: 'quantity' | 'amount') => {
    const nums = source.records.map((r) => r.fields[id]).filter((v): v is number => typeof v === 'number');
    return nums.length ? round(nums.reduce((a, b) => a + b, 0)) : null;
  };
  const totals = { quantity: sumOf('quantity'), amount: sumOf('amount') };
  const checkOne = (mine: number | null, invoice: number | null, tol: number) =>
    mine === null || invoice === null ? 'na' : Math.abs(mine - invoice) <= tol ? 'ok' : 'diff';

  return {
    columns: layout.rows.length === 0 ? [{ id: 'name', label: '', role: 'row' }, ...columns] : columns,
    rows: layout.rows.length === 0 ? rows.map((r) => ({ ...r, cells: ['', ...r.cells] })) : rows,
    grandTotal,
    totals,
    check: {
      quantity: checkOne(totals.quantity, source.invoiceTotals.quantity, 0.0001),
      amount: checkOne(totals.amount, source.invoiceTotals.amount, 0.011),
    },
  };
};
