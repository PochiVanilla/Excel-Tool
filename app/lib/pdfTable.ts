// ==========================================
// TRÍCH BẢNG HÀNG HOÁ TỪ PDF (hoá đơn VAT, invoice, packing list...) -> EXCEL, giống công cụ PDF -> Excel
// - Tìm dòng tiêu đề bảng (STT / Tên hàng / ĐVT / Số lượng / Đơn giá / Thành tiền) trên từng trang
// - Chia cột theo ĐƯỜNG KẺ DỌC của bảng (chính xác cho bảng có viền); không có viền thì chia theo vị trí tiêu đề cột
// - Mỗi mặt hàng = 1 dòng có số (Số lượng / Thành tiền) + các dòng chữ phía trên / dưới gần nó nhất (tên hàng xuống dòng)
// - Dòng Cộng tiền hàng / Tổng cộng / Thuế suất... giữ lại để Pivot đối chiếu tổng
// Không phụ thuộc DOM -> test được bằng Node
// ==========================================
import ExcelJS from 'exceljs';
import { ColumnRole, detectRole } from './invoiceTable';
import { PdfLine, PdfTextItem, groupLines } from './pdfInvoice';
import { foldText, parseLooseNumber } from './sheetText';

export interface VRule {
  page: number;
  x: number;
  y0: number;
  y1: number;
}

export interface TableRow {
  kind: 'item' | 'group' | 'total';
  cells: (string | number | null)[];
  page: number;
}

export interface ExtractedTable {
  headers: string[];
  roles: Partial<Record<ColumnRole, number>>;
  rows: TableRow[];
  // Dòng chữ đầu trang 1 (tên người bán) và vài dòng thông tin phía trên bảng
  intro: string[];
  pages: number;
}

const NUM_ROLES: ColumnRole[] = ['qty', 'price', 'amount'];
const INDEX_TOKEN_RE = /^\(?\d{1,2}\)?$|^=$|^[x×*]$|^\d{1,2}\s*=\s*\d{1,2}\s*[x×*]\s*\d{1,2}$/i;
const TOTAL_RE =
  /^(cong tien hang|tong cong|tong tien|tong so|thue suat|tien thue|so tien viet bang chu|so tien bang chu|total|sub total|subtotal|grand total|vat rate|vat amount|amount in words|cong\b)/;
// Nhãn tổng rõ ràng -> nhận ra ở bất kỳ đâu trên dòng (vd "Ghi chú: ... Cộng tiền hàng (Total amount): 13.286,26")
const TOTAL_ANY_RE = /(cong tien hang|tong cong|thue suat|tien thue|tong tien thanh toan|so tien (viet )?bang chu|grand total|total amount|amount in words|vat rate)/;
// Nhãn tổng của dòng: từ chữ đầu tiên của nhãn tổng tới trước con số (vd "Cộng tiền hàng (Total amount):")
const totalLabel = (text: string) => {
  const words = text.split(/\s+/).filter(Boolean);
  for (let i = 0; i < words.length; i++) {
    const rest = foldText(words.slice(i).join(' '));
    if (TOTAL_RE.test(rest) || TOTAL_ANY_RE.test(rest.slice(0, 40)) && TOTAL_ANY_RE.exec(rest)?.index === 0) {
      return words
        .slice(i)
        .filter((w) => !isNumberToken(w) || /%$/.test(w))
        .join(' ')
        .trim();
    }
  }
  return text.trim();
};
const isTotalText = (folded: string) => TOTAL_RE.test(folded) || TOTAL_ANY_RE.test(folded);
// Dòng tổng có cột Số lượng ("Tổng cộng: 16.836"); "Cộng tiền hàng" thì số ở cột Số lượng không phải tổng số lượng
const TOTAL_QTY_RE = /(tong cong|tong so luong|total quantity|total qty|^total\b)(?!.*(tien|thanh toan|amount))/;
const FOOTER_RE = /^(trang \d|page \d|tra cuu|nguoi mua hang|nguoi ban hang|ky boi|signature|ma tra cuu|don vi cung cap)/;

const centerX = (it: PdfTextItem) => it.x + Math.max(0, it.w) / 2;
const isNumberToken = (s: string) => parseLooseNumber(s.replace(/\s/g, '')) !== null;
const isIndexLine = (line: PdfLine) => line.items.length >= 3 && line.items.every((it) => INDEX_TOKEN_RE.test(it.str.trim()));

// Tìm dòng tiêu đề bảng trên 1 trang: dòng có STT / No. và trong vùng lân cận có Số lượng + Thành tiền
const findHeaderLine = (lines: PdfLine[]) => {
  for (let i = 0; i < lines.length; i++) {
    const near = lines.filter((l) => Math.abs(l.y - lines[i].y) <= 16);
    const text = foldText(near.map((l) => l.text).join(' '));
    const hasStt = lines[i].items.some((it) => /^(stt|so tt|no\.?|\(no\.?\))$/.test(foldText(it.str.trim())));
    const hasQty = /so luong|quantity|\bqty\b/.test(text);
    const hasAmount = /thanh tien|tri gia|amount/.test(text);
    if (hasStt && hasQty && hasAmount) return i;
  }
  return -1;
};

// Ranh giới cột: các đường kẻ dọc cắt ngang vùng tiêu đề (gộp các đường gần nhau)
const bordersFromRules = (rules: VRule[], yTop: number, yBottom: number) => {
  const xs = rules
    .filter((r) => r.y1 >= yBottom - 2 && r.y0 <= yTop + 2)
    .map((r) => r.x)
    .sort((a, b) => a - b);
  const out: number[] = [];
  for (const x of xs) if (!out.length || x - out[out.length - 1] > 3) out.push(x);
  return out;
};

// Không có đường kẻ: chia cột theo các cụm chữ của tiêu đề (ranh giới = giữa 2 cụm)
const bordersFromHeader = (headerItems: PdfTextItem[]) => {
  const spans = headerItems.map((it) => [it.x, it.x + Math.max(it.w, 4)] as [number, number]).sort((a, b) => a[0] - b[0]);
  const clusters: [number, number][] = [];
  for (const s of spans) {
    const last = clusters[clusters.length - 1];
    if (last && s[0] <= last[1] + 6) last[1] = Math.max(last[1], s[1]);
    else clusters.push([...s]);
  }
  if (clusters.length < 3) return [];
  const borders = [clusters[0][0] - 4];
  for (let i = 1; i < clusters.length; i++) borders.push((clusters[i - 1][1] + clusters[i][0]) / 2);
  borders.push(clusters[clusters.length - 1][1] + 40);
  return borders;
};

const columnOf = (borders: number[], x: number) => {
  if (x < borders[0]) return 0;
  for (let c = 0; c < borders.length - 1; c++) if (x < borders[c + 1]) return c;
  return borders.length - 2;
};

export const extractTable = (items: PdfTextItem[], rules: VRule[], pageCount: number): ExtractedTable | null => {
  const allLines = groupLines(items);
  let headers: string[] | null = null;
  let roles: Partial<Record<ColumnRole, number>> = {};
  const rows: TableRow[] = [];
  const intro: string[] = [];

  for (let page = 1; page <= pageCount; page++) {
    const lines = allLines.filter((l) => l.page === page);
    const h = findHeaderLine(lines);
    if (h < 0) continue;
    // Vùng tiêu đề: các dòng sát dòng STT (tiêu đề 2-3 dòng Việt / Anh), kết thúc ở dòng đánh số cột 1 2 3...
    const hy = lines[h].y;
    // Bảng có đường kẻ: vùng tiêu đề nằm gọn trong các đường kẻ dọc cắt qua dòng STT
    const pageRules = rules.filter((r) => r.page === page);
    const crossing = pageRules.filter((r) => r.y0 - 2 <= hy && hy <= r.y1 + 2 && r.y1 - r.y0 < 80);
    const cellTop = crossing.length >= 3 ? Math.max(...crossing.map((r) => r.y1)) + 1 : hy + 14;
    const block: PdfLine[] = [];
    let bodyStart = h;
    for (let i = 0; i < lines.length; i++) {
      const l = lines[i];
      if (l.y > Math.min(hy + 14, cellTop) || l.y < hy - 30) continue;
      if (i > h && l.items.some((it) => isNumberToken(it.str)) && !isIndexLine(l)) break;
      block.push(l);
      bodyStart = Math.max(bodyStart, i);
      if (isIndexLine(l)) break;
    }
    const headerLines = block.filter((l) => !isIndexLine(l));
    const yTop = Math.max(...block.map((l) => l.y)) + 8;
    const yBottom = Math.min(...block.map((l) => l.y)) - 2;
    let borders = bordersFromRules(pageRules, Math.min(yTop, cellTop), yBottom);
    if (borders.length < 4) borders = bordersFromHeader(headerLines.flatMap((l) => l.items));
    if (borders.length < 4) continue;
    const nCols = borders.length - 1;

    // Tên cột = chữ tiêu đề nằm trong cột (dòng tiếng Việt trước, tiếng Anh sau)
    const labels = Array.from({ length: nCols }, (_, c) =>
      headerLines
        .map((l) =>
          l.items
            .filter((it) => columnOf(borders, centerX(it)) === c)
            .map((it) => it.str.trim())
            .join(' '),
        )
        .filter(Boolean)
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim(),
    );
    if (!headers) {
      headers = labels;
      roles = {};
      labels.forEach((label, c) => {
        const role = detectRole(foldText(label));
        if (role && roles[role] === undefined) roles[role] = c;
      });
      // Dòng chữ phía trên bảng ở trang 1: tên người bán, số hoá đơn... (giữ để Pivot nhận ra người bán)
      intro.push(...lines.slice(0, h).map((l) => l.text).filter((t) => t.trim()).slice(0, 8));
    }
    const sttCol = roles.stt ?? (labels[0] && /stt|no/.test(foldText(labels[0])) ? 0 : -1);
    const numCols = NUM_ROLES.map((r) => roles[r]).filter((c): c is number => c !== undefined);
    if (numCols.length === 0) continue;

    // Thân bảng: từ sau vùng tiêu đề tới dòng tổng / chân trang / hết đường kẻ của bảng
    const tableRules = rules.filter((r) => r.page === page && borders.some((b) => Math.abs(b - r.x) < 3));
    const tableBottom = tableRules.length ? Math.min(...tableRules.map((r) => r.y0)) - 2 : -Infinity;
    type Cells = { line: PdfLine; cells: string[][] };
    const body: Cells[] = [];
    const totals: Cells[] = [];
    for (let i = bodyStart + 1; i < lines.length; i++) {
      const line = lines[i];
      const folded = foldText(line.text);
      if (line.y < tableBottom && !isTotalText(folded)) break;
      if (FOOTER_RE.test(folded)) break;
      if (isIndexLine(line)) continue;
      const cells: string[][] = Array.from({ length: nCols }, () => []);
      let sttTaken = false;
      for (const it of line.items) {
        const s = it.str.trim();
        if (!s) continue;
        let c = columnOf(borders, centerX(it));
        // Cột STT chỉ nhận đúng 1 số thứ tự nhỏ; chữ / số khác lấn sang là của cột kế bên (tên hàng)
        if (c === sttCol && (sttTaken || !/^\d{1,4}$/.test(s))) c = Math.min(nCols - 1, c + 1);
        if (c === sttCol) sttTaken = true;
        cells[c].push(s);
      }
      // Dòng tổng: nhãn ở bất kỳ ô nào (vd "Cộng tiền hàng (Total amount):")
      if (cells.some((cs) => TOTAL_RE.test(foldText(cs.join(' ')))) || TOTAL_ANY_RE.test(folded)) totals.push({ line, cells });
      else body.push({ line, cells });
    }

    const isData = (b: Cells) => numCols.some((c) => b.cells[c].some((s) => isNumberToken(s)));
    const dataIdx = body.map((b, i) => (isData(b) ? i : -1)).filter((i) => i >= 0);
    // Gán mỗi dòng chữ cho dòng số gần nhất theo chiều dọc (trang chỉ có dòng tổng thì không có dòng hàng)
    const owner = body.map((b, i) => {
      if (!dataIdx.length) return -1;
      if (dataIdx.includes(i)) return i;
      let best = dataIdx[0];
      for (const d of dataIdx) if (Math.abs(body[d].line.y - b.line.y) < Math.abs(body[best].line.y - b.line.y)) best = d;
      return best;
    });

    const toCells = (group: Cells[], numbersFrom: Cells | null): (string | number | null)[] =>
      Array.from({ length: nCols }, (_, c) => {
        if (numbersFrom && numCols.includes(c)) {
          const n = numbersFrom.cells[c].map((s) => parseLooseNumber(s.replace(/\s/g, ''))).find((v) => v !== null);
          return n ?? null;
        }
        if (c === sttCol && numbersFrom) {
          const n = numbersFrom.cells[c].map((s) => parseLooseNumber(s)).find((v) => v !== null);
          return n ?? null;
        }
        const text = group
          .map((g) => g.cells[c].join(' '))
          .filter(Boolean)
          .join('\n');
        return text || null;
      });

    for (const d of dataIdx) {
      const group = body.filter((_, i) => owner[i] === d);
      // Hoá đơn vải Polytex: dòng nhóm ("Vải dệt ... Theo hoá đơn thương mại số ...") nằm trên mã "#1012387" -> dòng nhóm riêng
      const codeAt = group.findIndex((g) => g.line.y >= body[d].line.y && g.cells.some((cs) => /^#\s*\w/.test(cs.join(' '))));
      if (codeAt > 0) {
        rows.push({ kind: 'group', cells: toCells(group.slice(0, codeAt), null), page });
        rows.push({ kind: 'item', cells: toCells(group.slice(codeAt), body[d]), page });
      } else rows.push({ kind: 'item', cells: toCells(group, body[d]), page });
    }
    for (const t of totals) {
      const keepQty = TOTAL_QTY_RE.test(foldText(t.line.text));
      // Dòng tổng gọn như Excel: nhãn tổng ở ô đầu tiên (bỏ phần ghi chú lấn sang), số giữ đúng cột Số lượng / Thành tiền
      const cells: (string | number | null)[] = Array.from({ length: nCols }, (_, c) => {
        if (!numCols.includes(c)) return null;
        if (c === roles.qty && !keepQty) return null;
        return t.cells[c].map((x) => parseLooseNumber(x.replace(/\s/g, ''))).find((v) => v !== null) ?? null;
      });
      cells[0] = totalLabel(t.line.text);
      rows.push({ kind: 'total', cells, page });
    }
  }
  if (!headers || rows.every((r) => r.kind !== 'item')) return null;
  return { headers, roles, rows, intro, pages: pageCount };
};

// Dựng file Excel: vài dòng thông tin đầu (người bán...), dòng tiêu đề, các dòng hàng, các dòng tổng
export const buildTableWorkbook = (t: ExtractedTable, sheetName = 'VAT') => {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Excel Tool';
  const ws = wb.addWorksheet(sheetName.slice(0, 31));
  const thin: Partial<ExcelJS.Borders> = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  for (const line of t.intro) ws.addRow([line]).getCell(1).font = { name: 'Arial', size: 10 };
  if (t.intro.length) ws.getRow(1).getCell(1).font = { name: 'Arial', size: 11, bold: true };
  ws.addRow([]);
  const headerRow = ws.addRow(t.headers);
  headerRow.eachCell((c) => {
    c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B3759' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border = thin;
  });
  headerRow.height = 32;
  const numCols = new Set(NUM_ROLES.map((r) => t.roles[r]).filter((c) => c !== undefined));
  // Định dạng số theo số chữ số thập phân thật của cột (tối đa 6) — chỉ dùng mã chuẩn "#,##0.00..." (Luckysheet
  // không đọc được kiểu "0.00##")
  const formats = new Map<number, string>();
  for (const c of numCols) {
    let decimals = c === t.roles.qty ? 0 : 2;
    for (const r of t.rows) {
      const v = r.cells[c as number];
      if (typeof v !== 'number') continue;
      const frac = String(+v.toFixed(6)).split('.')[1] || '';
      decimals = Math.max(decimals, Math.min(6, frac.length));
    }
    formats.set(c as number, decimals ? `#,##0.${'0'.repeat(decimals)}` : '#,##0');
  }
  for (const r of t.rows) {
    const row = ws.addRow(r.cells.map((v) => (v === null ? '' : v)));
    row.eachCell({ includeEmpty: true }, (c, col) => {
      const idx = col - 1;
      c.font = { name: 'Arial', size: 10, bold: r.kind !== 'item' };
      c.border = thin;
      c.alignment = { vertical: 'top', wrapText: typeof c.value === 'string', horizontal: typeof c.value === 'number' ? 'right' : 'left' };
      if (typeof c.value === 'number' && numCols.has(idx)) c.numFmt = formats.get(idx) || '#,##0.00';
    });
    if (r.kind === 'total') row.eachCell({ includeEmpty: true }, (c) => (c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E0EE' } }));
  }
  ws.columns.forEach((col, i) => {
    const role = Object.entries(t.roles).find(([, c]) => c === i)?.[0];
    col.width = role === 'name' ? 48 : role === 'stt' ? 6 : numCols.has(i) ? 15 : Math.min(30, Math.max(10, (t.headers[i] || '').length * 0.6));
  });
  return wb;
};
