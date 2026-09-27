// ==========================================
// ĐỌC COMMERCIAL INVOICE VẢI (PDF) -> BẢNG "TRÍCH INV VẢI"
// Đầu vào: các mẩu chữ kèm toạ độ do pdf.js trích từ PDF (không phụ thuộc DOM -> test được bằng Node)
// Đầu ra: STT | Item (#) | Thành phần & Định lượng khổ vải | Số lượng | Đơn giá | Trị giá
//         + TỔNG CỘNG đối chiếu TOTAL của invoice + thông tin đóng gói (cuộn, N.W, G.W, CBM)
// ==========================================
import ExcelJS from 'exceljs';
import { parseLooseNumber } from './sheetText';

export interface PdfTextItem {
  str: string;
  x: number;
  y: number;
  w: number;
  page: number;
}

export interface PdfLine {
  page: number;
  y: number;
  items: PdfTextItem[];
  text: string;
}

export interface FabricRow {
  item: string;
  composition: string;
  color: string;
  qty: number;
  price: number;
  amount: number;
  unit: string;
  page: number;
}

export interface FabricInvoice {
  invoiceNo: string;
  date: string;
  seller: string;
  unit: string;
  currency: string;
  rows: FabricRow[];
  total: { qty: number | null; amount: number | null };
  packing: { rolls: number | null; nw: number | null; gw: number | null; cbm: number | null };
  warnings: string[];
  pageCount: number;
}

// ---------- Gom các mẩu chữ thành dòng (cùng trang, cùng toạ độ y) ----------
export const groupLines = (raw: PdfTextItem[]): PdfLine[] => {
  const items = raw.filter((i) => i.str.trim() !== '').sort((a, b) => a.page - b.page || b.y - a.y || a.x - b.x);
  const lines: PdfLine[] = [];
  for (const it of items) {
    const last = lines[lines.length - 1];
    if (last && last.page === it.page && Math.abs(last.y - it.y) <= 2.5) last.items.push(it);
    else lines.push({ page: it.page, y: it.y, items: [it], text: '' });
  }
  for (const line of lines) {
    line.items.sort((a, b) => a.x - b.x);
    line.text = joinItems(line.items);
  }
  return lines;
};

// Nối mẩu chữ: giữ nguyên ký tự; chỉ chèn 1 khoảng trắng khi 2 mẩu cách nhau
const joinItems = (items: PdfTextItem[]) => {
  let out = '';
  let prevEnd: number | null = null;
  for (const it of items) {
    if (prevEnd !== null && it.x - prevEnd > 1 && !out.endsWith(' ') && !it.str.startsWith(' ')) out += ' ';
    out += it.str;
    prevEnd = it.x + it.w;
  }
  return out.replace(/\s+$/, '');
};

// ---------- Nhận diện số lượng / đơn giá / trị giá ở cuối dòng ----------
const UNIT_RE = /(YRDS?|YDS?|MTRS?|MTS|M|KGS|KGM|PCS|ROLLS?)$/i;
const NUMERIC_TOKEN_RE = /^(USD|US\$|\$)?\s*[\d.,]*\d\s*(YRDS?|YDS?|MTRS?|MTS|M|KGS|KGM|PCS)?$/i;

interface Columns {
  descX: number;
  qtyX: number;
  priceX: number;
  amountX: number;
}

interface NumericPart {
  qty: number | null;
  unit: string;
  price: number | null;
  amount: number | null;
  descItems: PdfTextItem[];
  currency: string;
}

const center = (it: PdfTextItem) => it.x + it.w / 2;

const splitNumeric = (line: PdfLine, cols: Columns | null): NumericPart => {
  const items = [...line.items];
  const nums: { it: PdfTextItem; text: string }[] = [];
  let currency = '';
  // Lấy dần từ phải sang trái các mẩu dạng số (kèm "USD" / đơn vị), dừng khi gặp chữ thường
  while (items.length > 0) {
    const it = items[items.length - 1];
    const t = it.str.trim();
    const minX = cols ? cols.descX + 120 : 250;
    if (it.x < minX) break;
    if (/^(USD|US\$|\$)$/i.test(t)) {
      currency = 'USD';
      items.pop();
      continue;
    }
    // Đơn vị bị tách thành ô riêng ("352.00" | "YRD") -> gắn vào số đứng ngay trước
    if (/^(YRDS?|YDS?|MTRS?|KGS|KGM|PCS)$/i.test(t) && items.length >= 2 && /\d$/.test(items[items.length - 2].str.trim())) {
      items.pop();
      const prev = items[items.length - 1];
      items[items.length - 1] = { ...prev, str: prev.str.trim() + t, w: it.x + it.w - prev.x };
      continue;
    }
    if (!NUMERIC_TOKEN_RE.test(t)) break;
    if (/^(USD|US\$)/i.test(t)) currency = 'USD';
    nums.unshift({ it, text: t });
    items.pop();
  }
  const part: NumericPart = { qty: null, unit: '', price: null, amount: null, descItems: items, currency };
  const values = nums.map(({ it, text }) => {
    const unitMatch = text.replace(/^(USD|US\$|\$)\s*/i, '').match(UNIT_RE);
    const clean = text.replace(/^(USD|US\$|\$)\s*/i, '').replace(UNIT_RE, '').trim();
    return { it, value: parseLooseNumber(clean), unit: unitMatch ? unitMatch[1].toUpperCase() : '' };
  });
  if (values.some((v) => v.value === null)) return { ...part, descItems: line.items };

  for (const v of values) {
    if (v.unit) {
      part.qty = v.value;
      part.unit = v.unit;
      continue;
    }
    if (cols) {
      // Gán theo cột tiêu đề gần nhất (số căn phải nên so theo mép phải)
      const right = v.it.x + v.it.w;
      const dQty = Math.abs(right - (cols.qtyX + 20));
      const dPrice = Math.abs(center(v.it) - cols.priceX);
      const dAmount = Math.abs(center(v.it) - cols.amountX);
      const best = Math.min(dQty, dPrice, dAmount);
      if (best === dAmount && part.amount === null) part.amount = v.value;
      else if (best === dPrice && part.price === null) part.price = v.value;
      else if (part.qty === null) part.qty = v.value;
      else if (part.price === null) part.price = v.value;
      else part.amount = v.value;
    } else if (part.price === null && values.length >= 2 + (part.qty === null ? 1 : 0)) part.price = v.value;
    else part.amount = v.value;
  }
  // Không có cột tiêu đề và thiếu đơn vị: 3 số = SL, ĐG, TG
  if (!cols && part.qty === null && values.length === 3) {
    part.qty = values[0].value;
    part.price = values[1].value;
    part.amount = values[2].value;
  }
  return part;
};

// ---------- Thành phần & khổ vải ----------
const WIDTH_RE = /\d+(?:[.,]\d+)?\s*CM\b.*?g\s*\/\s*m\s*2/i;

const LOT_LINE_RE = /^P\d{6,}/i;
const PACKING_LINE_RE = /PACKED\s+IN|NET\s+WEIGHT|GROSS\s+WEIGHT|MEASUREMENT|\bN\.\s*W\b|\bG\.\s*W\b|\bCBM\b|\bROLLS?\b/i;

const extractComposition = (descLines: string[]) => {
  let comp = '';
  let width = '';
  let collecting = false;
  for (const l of descLines) {
    const w = l.match(WIDTH_RE);
    if (!width && w) width = w[0];
    if (!comp) {
      const pct = l.search(/\d+(?:\.\d+)?\s*%/);
      if (pct === -1) continue;
      let part = l.slice(pct);
      const wm = part.match(WIDTH_RE);
      if (wm && wm.index !== undefined) part = part.slice(0, wm.index);
      comp = part.trim();
      // Thành phần bị xuống dòng (vd "100% POLYESTER (MECHANICALLY" + "RECYCLED)") -> nối tiếp tới dòng khổ vải
      collecting = !wm;
      continue;
    }
    if (!collecting) continue;
    if (w || LOT_LINE_RE.test(l)) {
      if (w && w.index) comp = `${comp} ${l.slice(0, w.index).trim()}`.trim();
      collecting = false;
      continue;
    }
    comp = `${comp} ${l.trim()}`;
  }
  return [comp, width].filter((s) => s !== '').join(' ');
};

const lastNumber = (lines: PdfLine[], re: RegExp) => {
  let found: number | null = null;
  for (const l of lines) {
    const m = l.text.match(re);
    if (m) {
      const n = parseLooseNumber(m[1]);
      if (n !== null) found = n;
    }
  }
  return found;
};

// ==========================================
// PHÂN TÍCH INVOICE
// ==========================================
export const parseFabricInvoice = (lines: PdfLine[], pageCount: number): FabricInvoice => {
  const warnings: string[] = [];
  const rows: FabricRow[] = [];
  let cols: Columns | null = null;
  let current: { code: string; desc: string[] } | null = null;
  let compositionOf: string | null = null;
  let total: FabricInvoice['total'] = { qty: null, amount: null };
  let unit = '';
  let currency = '';
  const allText = lines.map((l) => l.text).join('\n');

  const invoiceNo = allText.match(/INVOICE\s*NO\.?\s*[:：]?\s*([A-Z0-9][A-Z0-9\-/]{4,})/i)?.[1] ?? '';
  const date = allText.match(/\bDATE\s*[:：]\s*([0-9A-Za-z][0-9A-Za-z\-/ ,.]{5,20})/i)?.[1]?.trim() ?? '';
  const seller = lines.find((l) => /(CO\.?,?\s*LTD|LIMITED|COMPANY|CÔNG TY)/i.test(l.text))?.text.trim() ?? '';

  // Dòng cuối cùng thuộc bảng hàng (dòng hàng hoặc dòng cộng phụ) -> phần tổng & đóng gói nằm sau dòng này
  let lastTableLine = -1;
  let totalLine = -1;
  for (let li = 0; li < lines.length; li++) {
    const line = lines[li];
    const upper = line.text.toUpperCase();
    // Dòng tiêu đề bảng (lặp lại mỗi trang) -> cập nhật toạ độ cột
    if (/DESCRIPTION/.test(upper) && /(QUANTITY|QTY)/.test(upper) && /AMOUNT/.test(upper)) {
      const find = (re: RegExp) => line.items.find((i) => re.test(i.str.toUpperCase()));
      const d = find(/DESCRIPTION/);
      const q = find(/QUANTITY|QTY/);
      const p = find(/PRICE/);
      const a = find(/AMOUNT/);
      if (d && q && a) cols = { descX: d.x, qtyX: center(q), priceX: p ? center(p) : (center(q) + center(a)) / 2, amountX: center(a) };
      continue;
    }
    if (!cols) continue;

    const num = splitNumeric(line, cols);
    // Bỏ cột SHIPPING MARK bên trái (FEPV:, PO:, ITEM:, G.W:, N.W: ...)
    const descItems = num.descItems.filter((i) => i.x >= cols!.descX - 4);
    const descText = joinItems(descItems).trim();

    // Dòng thông tin đóng gói (PACKED IN, NET/GROSS WEIGHT, MEASUREMENT...) không thuộc bảng hàng
    // (chỉ xét cột mô tả: cột SHIPPING MARK có nhãn G.W:/N.W: nằm cùng dòng với dòng hàng)
    if (PACKING_LINE_RE.test(descText)) continue;
    if (/\bTOTAL\b/i.test(descText) && !/SAY\s+TOTAL|SUB\s*-?\s*TOTAL/i.test(descText)) {
      totalLine = li;
      current = null;
      continue;
    }

    const code = descText.match(/^#\s*[\w\-./]+/);
    if (code) {
      current = { code: code[0].replace(/\s+/g, ''), desc: [] };
      compositionOf = null;
      const rest = descText.slice(code[0].length).trim();
      if (rest) current.desc.push(rest);
      continue;
    }
    if (!current) continue;

    if (num.qty !== null && num.price !== null && num.amount !== null) {
      if (compositionOf === null) {
        compositionOf = extractComposition(current.desc);
        if (!compositionOf) warnings.push(`${current.code}: không tìm thấy thành phần (%) / khổ (CM ... g/m2) trong mô tả.`);
      }
      const diff = Math.abs(num.qty * num.price - num.amount);
      if (diff > Math.max(0.02, Math.abs(num.amount) * 0.0005))
        warnings.push(`${current.code} ${descText}: ${num.qty} x ${num.price} ≠ ${num.amount} (trang ${line.page})`);
      lastTableLine = li;
      unit = unit || num.unit;
      currency = currency || num.currency;
      rows.push({
        item: current.code,
        composition: compositionOf,
        color: descText,
        qty: num.qty,
        price: num.price,
        amount: num.amount,
        unit: num.unit || unit,
        page: line.page,
      });
      continue;
    }
    // Dòng cộng phụ của từng mã (chỉ có số lượng + trị giá) -> bỏ qua
    if (num.qty !== null && num.price === null) {
      lastTableLine = li;
      continue;
    }
    if (descText) current.desc.push(descText);
  }

  // Số tổng: số lượng & trị giá có thể lệch lên/xuống vài pt so với nhãn "TOTAL:" (vd "14750.000YR" / "TOTAL: USD18,118.51" / "D")
  // Không có nhãn TOTAL thì tìm sau dòng hàng cuối, trước "SAY TOTAL"
  const scanTotal = (candidates: PdfLine[]) => {
    for (const l of candidates) {
      const t = l.text.replace(/\s+/g, '');
      const q = t.match(/(?:^|[^\d.,])([\d,]*\d(?:\.\d+)?)(?:YRDS?|YRD?|YDS?|MTRS?)(?!\d)/i) || t.match(/^([\d,]*\d(?:\.\d+)?)(?:YRDS?|YRD?|YDS?|MTRS?)/i);
      const a = t.match(/(?:USD|US\$)([\d,]*\d(?:\.\d+)?)/i);
      if (q && total.qty === null) total = { ...total, qty: parseLooseNumber(q[1]) };
      if (a && total.amount === null) total = { ...total, amount: parseLooseNumber(a[1]) };
    }
  };
  if (totalLine >= 0) {
    const tl = lines[totalLine];
    scanTotal(lines.filter((l) => l.page === tl.page && Math.abs(l.y - tl.y) <= 12 && !PACKING_LINE_RE.test(l.text)));
  } else if (lastTableLine >= 0) {
    const say = lines.findIndex((l, i) => i > lastTableLine && /SAY\s+TOTAL/i.test(l.text));
    const end = say > 0 ? say : Math.min(lines.length, lastTableLine + 8);
    scanTotal(lines.slice(lastTableLine + 1, end).filter((l) => !PACKING_LINE_RE.test(l.text)));
  }

  if (!cols) warnings.push('Không tìm thấy dòng tiêu đề bảng (DESCRIPTION OF GOODS / QUANTITY / AMOUNT).');
  if (rows.length === 0 && cols) warnings.push('Không đọc được dòng hàng nào (dòng có Số lượng + Đơn giá + Trị giá).');
  if (total.qty === null) warnings.push('Không tìm thấy dòng TOTAL trên invoice để đối chiếu.');

  // Thông tin đóng gói nằm sau bảng hàng: có thể có nhãn (N.W:, G.W:) hoặc chỉ có số + đơn vị (ROLLS, KGM, CBM).
  // Không có nhãn thì KGM nhỏ hơn là N.W, lớn hơn là G.W
  const tail = lines.slice(lastTableLine + 1);
  const kgm: number[] = [];
  for (const l of tail) {
    if (/\b(?:N|G)\.\s*W\.?/i.test(l.text)) continue;
    const m = l.text.match(/(\d[\d,]*(?:\.\d+)?)\s*(?:KGM|KGS)\b/i);
    const n = m ? parseLooseNumber(m[1]) : null;
    if (n !== null) kgm.push(n);
  }
  const packing = {
    rolls: lastNumber(tail, /(\d[\d,]*(?:\.\d+)?)\s*(?:ROLLS?|RLS)\b/i) ?? lastNumber(tail, /ROLLS?\s*[:：]\s*(\d[\d,]*(?:\.\d+)?)/i),
    nw:
      lastNumber(tail, /(?:\bN\.\s*W\.?|\bNW\b|NET\s+WEIGHT)\s*[:：]?\s*(\d[\d,]*(?:\.\d+)?)/i) ??
      (kgm.length >= 2 ? Math.min(kgm[0], kgm[1]) : kgm[0] ?? null),
    gw:
      lastNumber(tail, /(?:\bG\.\s*W\.?|\bGW\b|GROSS\s+WEIGHT)\s*[:：]?\s*(\d[\d,]*(?:\.\d+)?)/i) ??
      (kgm.length >= 2 ? Math.max(kgm[0], kgm[1]) : null),
    cbm: lastNumber(tail, /(\d[\d,]*(?:\.\d+)?)\s*CBM\b/i) ?? lastNumber(tail, /CBM\s*[:：]?\s*(\d[\d,]*(?:\.\d+)?)/i),
  };

  return { invoiceNo, date, seller, unit: unit || 'YRD', currency: currency || 'USD', rows, total, packing, warnings, pageCount };
};

export const invoiceSums = (inv: FabricInvoice) => {
  const round = (n: number, d: number) => Math.round(n * 10 ** d) / 10 ** d;
  const qty = round(inv.rows.reduce((a, r) => a + r.qty, 0), 4);
  const amount = round(inv.rows.reduce((a, r) => a + r.amount, 0), 4);
  const ok = (mine: number, theirs: number | null, tol: number) => (theirs === null ? 'na' : Math.abs(mine - theirs) <= tol ? 'ok' : 'diff');
  return { qty, amount, qtyCheck: ok(qty, inv.total.qty, 0.005), amountCheck: ok(amount, inv.total.amount, 0.011) };
};

// ==========================================
// DỰNG FILE EXCEL THEO MẪU "TRÍCH INV VẢI" (sheet Data, có công thức)
// ==========================================
export const buildFabricWorkbook = (inv: FabricInvoice) => {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Excel Tool';
  const thin: Partial<ExcelJS.Borders> = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  };
  const headStyle = (c: ExcelJS.Cell) => {
    c.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border = thin;
  };
  const body = (c: ExcelJS.Cell, numFmt?: string, align: ExcelJS.Alignment['horizontal'] = 'right') => {
    c.font = { name: 'Arial', size: 10 };
    c.border = thin;
    c.alignment = { horizontal: align, vertical: 'middle', wrapText: align === 'left' };
    if (numFmt) c.numFmt = numFmt;
  };
  const totalStyle = (c: ExcelJS.Cell, numFmt?: string) => {
    c.font = { name: 'Arial', size: 11, bold: true };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
    c.border = thin;
    c.alignment = { horizontal: numFmt ? 'right' : 'center', vertical: 'middle' };
    if (numFmt) c.numFmt = numFmt;
  };
  const u = inv.unit || 'YRD';
  const cur = inv.currency || 'USD';

  // ---- Sheet Data ----
  const ws = wb.addWorksheet('Data', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [{ width: 6 }, { width: 14 }, { width: 55 }, { width: 16 }, { width: 14 }, { width: 16 }];
  const head = ['STT', 'Item', 'Thành phần & Định lượng khổ vải', `Số lượng (${u})`, `Đơn giá (${cur})`, `Trị giá (${cur})`];
  ws.addRow(head).eachCell(headStyle);
  inv.rows.forEach((r, i) => {
    const n = i + 2;
    const row = ws.addRow([i + 1, r.item, r.composition, r.qty, r.price, { formula: `D${n}*E${n}`, result: r.amount }]);
    body(row.getCell(1), undefined, 'center');
    body(row.getCell(2), undefined, 'left');
    body(row.getCell(3), undefined, 'left');
    body(row.getCell(4), '#,##0.00');
    body(row.getCell(5), '0.0000');
    body(row.getCell(6), '#,##0.0000');
  });
  const last = inv.rows.length + 1;
  const sums = invoiceSums(inv);
  const totalRow = ws.addRow(['TỔNG CỘNG', '', '', { formula: `SUM(D2:D${last})`, result: sums.qty }, '', { formula: `SUM(F2:F${last})`, result: sums.amount }]);
  ws.mergeCells(totalRow.number, 1, totalRow.number, 3);
  totalRow.eachCell({ includeEmpty: true }, (c, col) => totalStyle(c, col === 4 ? '#,##0.00' : col === 6 ? '#,##0.0000' : undefined));
  ws.addRow([]);
  const packTitle = ws.addRow(['THÔNG TIN ĐÓNG GÓI']);
  ws.mergeCells(packTitle.number, 1, packTitle.number, 3);
  for (let c = 1; c <= 3; c++) {
    const cell = packTitle.getCell(c);
    cell.font = { name: 'Arial', size: 11, bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    cell.border = thin;
    cell.alignment = { horizontal: 'left' };
  }
  const packing: [string, number | null][] = [
    ['Số cuộn (ROLLS):', inv.packing.rolls],
    ['N.W (KGM):', inv.packing.nw],
    ['G.W (KGM):', inv.packing.gw],
    ['CBM:', inv.packing.cbm],
  ];
  // Nhãn gộp A:C (cột A chỉ rộng 6 nên nhãn sẽ bị cắt nếu để riêng), số ở cột D
  for (const [label, value] of packing) {
    const row = ws.addRow([label, '', '', value]);
    ws.mergeCells(row.number, 1, row.number, 3);
    for (let c = 1; c <= 3; c++) body(row.getCell(c), undefined, 'left');
    body(row.getCell(4), '#,##0.00');
  }

  return wb;
};
