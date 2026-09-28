// ==========================================
// TÁCH BỘ CHỨNG TỪ PDF THEO TIÊU ĐỀ TRANG
// (Tờ khai xuất khẩu, Danh sách hàng hoá, Sales Contract, Commercial Invoice, Packing List, VAT...)
// - Mỗi trang: tìm dòng tiêu đề (chữ to / nằm phía trên trang) khớp với từ khoá của từng loại chứng từ
// - Trang không có tiêu đề là trang tiếp theo của chứng từ phía trước
// - Chọn tay loại cho 1 trang -> các trang phía sau theo loại đó cho tới trang chọn tay kế tiếp
//   (hoặc tới trang có tiêu đề nhận ra được) -> chỉ cần chọn trang đầu của mỗi chứng từ
// - Tiêu đề lạ (chưa có trong danh sách, vd "BIÊN BẢN GIAO HÀNG") -> tách thành chứng từ riêng, không gộp vào bộ trước
// - Trang ảnh scan (không có lớp chữ) gom riêng vào scan.pdf, trừ khi OCR đọc được tiêu đề
// Không phụ thuộc DOM -> test được bằng Node
// ==========================================
import { foldText } from './sheetText';

export type DocTypeId =
  | 'to-khai-xuat'
  | 'to-khai-nhap'
  | 'danh-sach'
  | 'phieu-xuat-kho'
  | 'phieu-nhap-kho'
  | 'sales-contract'
  | 'annex'
  | 'commercial-invoice'
  | 'packing-list'
  | 'cdgh'
  | 'vat'
  | 'bill-of-lading'
  | 'co'
  | 'other';

export type PageTarget = DocTypeId | 'scan' | 'skip';

export interface SplitLine {
  text: string;
  // Cỡ chữ (pt) và vị trí mép trên của dòng tính từ đầu trang (0 = trên cùng, 1 = dưới cùng)
  size: number;
  top: number;
}

export interface SplitPage {
  index: number;
  width: number;
  height: number;
  lines: SplitLine[];
  chars: number;
  // Tỉ lệ diện tích trang bị ảnh lớn nhất che phủ (trang scan ~ 1)
  imageCoverage: number;
  isScan: boolean;
  // Kết quả OCR của trang scan (null = đã OCR nhưng không đọc được chữ)
  ocr?: SplitLine[] | null;
}

export interface DocTypeDef {
  id: DocTypeId;
  label: string;
  // Tên file khi xuất (TKX, SC, ANNEX, INV, PKL, VAT...; loại không có tên riêng gom vào Other)
  code: string;
  color: string;
  // Cụm từ tiêu đề (đã bỏ dấu, viết thường)
  phrases: string[];
  // Cụm từ ngắn / chung chung (vd "hop dong"): dòng phải gần như chỉ có đúng cụm này mới tính là tiêu đề
  strict?: string[];
  docNo?: (text: string, lines: SplitLine[]) => string;
}

const firstMatch = (text: string, re: RegExp) => {
  const m = text.match(re);
  return m ? m[1] : '';
};

// Số chứng từ dạng "Số (No.): 3589" (hoá đơn VAT, phiếu xuất / nhập kho)
const soNo = (text: string) => firstMatch(foldText(text), /so\s*\(\s*no\.?\s*\)\s*:?\s*(\d+)/);

// Số tờ khai hải quan: 12 chữ số
const declarationNo = (text: string) => firstMatch(text, /(?:^|\D)(\d{12})(?!\d)/);

// Số invoice nằm cùng dòng với nhãn "INVOICE NO." (vd TPUH26095528)
const INVOICE_NO_RE = /\b([A-Z]{2,6}\d{6,}[A-Z0-9-]*)\b/;
const invoiceNo = (_text: string, lines: SplitLine[]) => {
  for (const l of lines) {
    if (!/invoice\s*no/i.test(l.text)) continue;
    const m = l.text.match(INVOICE_NO_RE);
    if (m) return m[1];
    // Nhãn và số nằm ở 2 khối chữ khác nhau nhưng cùng hàng
    for (const o of lines) {
      if (o === l || Math.abs(o.top - l.top) > 0.012) continue;
      const n = o.text.trim().match(/^([A-Z]{2,6}\d{6,}[A-Z0-9-]*)$/);
      if (n) return n[1];
    }
  }
  return '';
};

const contractNo = (text: string) =>
  firstMatch(text, /S\s*\/?\s*C\s*No\.?\s*:?\s*([A-Z0-9][A-Z0-9-]{3,})/i) || firstMatch(text, /Contract\s*No\.?\s*:?\s*([A-Z0-9][A-Z0-9-/]{3,})/i);

export const OTHER_CODE = 'Other';
export const SCAN_CODE = 'SCAN';

export const DOC_TYPES: DocTypeDef[] = [
  { id: 'to-khai-xuat', label: 'Tờ khai hàng hóa xuất khẩu', code: 'TKX', color: '#2563eb', phrases: ['to khai hang hoa xuat khau'], docNo: declarationNo },
  { id: 'to-khai-nhap', label: 'Tờ khai hàng hóa nhập khẩu', code: OTHER_CODE, color: '#0891b2', phrases: ['to khai hang hoa nhap khau'], docNo: declarationNo },
  {
    id: 'danh-sach',
    label: 'Danh sách hàng hóa',
    code: OTHER_CODE,
    color: '#7c3aed',
    phrases: ['danh sach hang hoa', 'du dieu kien qua khu vuc giam sat'],
    docNo: declarationNo,
  },
  {
    id: 'phieu-xuat-kho',
    label: 'Phiếu xuất kho',
    code: OTHER_CODE,
    color: '#0f766e',
    phrases: ['phieu xuat kho', 'delivery and internal transfer note', 'internal transfer note'],
    docNo: soNo,
  },
  {
    id: 'phieu-nhap-kho',
    label: 'Phiếu nhập kho',
    code: OTHER_CODE,
    color: '#be123c',
    phrases: ['phieu nhap kho', 'goods received note', 'goods receipt note', 'warehouse receipt'],
    docNo: soNo,
  },
  {
    id: 'sales-contract',
    label: 'Hợp đồng (Sales Contract)',
    code: 'SC',
    color: '#ea580c',
    phrases: ['sales contract', 'sale contract', 'hop dong mua ban', 'hop dong ban hang', 'hop dong kinh te', 'hop dong thuong mai'],
    strict: ['hop dong', 'contract'],
    docNo: contractNo,
  },
  {
    // "Phụ lục / Phụ kiện hợp đồng", "Attachment List of Sales Contract/Phụ lục đính kèm" -> cụm dài hơn nên thắng cụm "sales contract"
    id: 'annex',
    label: 'Phụ lục hợp đồng (Annex)',
    code: 'ANNEX',
    color: '#c2410c',
    phrases: ['phu luc hop dong', 'phu kien hop dong', 'attachment list of sales contract', 'phu luc dinh kem', 'contract annex', 'contract appendix'],
    strict: ['phu luc', 'annex', 'appendix'],
    docNo: contractNo,
  },
  { id: 'commercial-invoice', label: 'Commercial Invoice', code: 'INV', color: '#059669', phrases: ['commercial invoice', 'hoa don thuong mai'], docNo: invoiceNo },
  { id: 'packing-list', label: 'Packing List', code: 'PKL', color: '#ca8a04', phrases: ['packing and weight list', 'packing list', 'phieu dong goi'], docNo: invoiceNo },
  {
    id: 'cdgh',
    label: 'Chỉ định giao nhận hàng',
    code: 'CDGH',
    color: '#0369a1',
    phrases: ['chi dinh giao nhan hang', 'chi dinh giao hang', 'chi dinh nhan hang'],
    // Số dạng "29/2026/BSNIN-FEN"
    docNo: (text) => firstMatch(text, /\b(\d{1,4}\/\d{4}\/[A-Z0-9][A-Z0-9-]*)/),
  },
  {
    id: 'vat',
    label: 'Hóa đơn VAT',
    code: 'VAT',
    color: '#db2777',
    phrases: ['hoa don gia tri gia tang', 'vat invoice'],
    docNo: soNo,
  },
  { id: 'bill-of-lading', label: 'Bill of Lading', code: OTHER_CODE, color: '#475569', phrases: ['bill of lading', 'van don'], docNo: (text) => firstMatch(text, /B\s*\/\s*L\s*No\.?\s*:?\s*([A-Z0-9-]{4,})/i) },
  { id: 'co', label: 'C/O', code: OTHER_CODE, color: '#65a30d', phrases: ['certificate of origin', 'giay chung nhan xuat xu'] },
  { id: 'other', label: 'Chứng từ khác', code: OTHER_CODE, color: '#64748b', phrases: [] },
];

export const docTypeOf = (id: DocTypeId) => DOC_TYPES.find((d) => d.id === id)!;
export const SCAN_COLOR = '#b45309';

// ---------- So khớp cụm từ (OCR thì cho phép sai vài ký tự) ----------
const norm = (s: string) =>
  foldText(s)
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();

// Khoảng cách chỉnh sửa nhỏ nhất giữa `pat` và một đoạn con bất kỳ của `text`
const approxSubstringDistance = (text: string, pat: string) => {
  const m = pat.length;
  let prev = new Array(m + 1);
  for (let j = 0; j <= m; j++) prev[j] = j;
  let best = prev[m];
  for (let i = 1; i <= text.length; i++) {
    const cur = new Array(m + 1);
    cur[0] = 0;
    for (let j = 1; j <= m; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (text[i - 1] === pat[j - 1] ? 0 : 1));
    }
    best = Math.min(best, cur[m]);
    prev = cur;
  }
  return best;
};

const matchPhrase = (lineNorm: string, phrase: string, fuzzy: boolean) => {
  if (` ${lineNorm} `.includes(` ${phrase} `)) return true;
  if (!fuzzy) return false;
  // OCR hay dính / tách chữ và sai dấu -> so khớp bỏ khoảng trắng, cho sai ~12% ký tự
  const t = lineNorm.replace(/ /g, '');
  const p = phrase.replace(/ /g, '');
  if (t.includes(p)) return true;
  if (t.length < p.length * 0.6) return false;
  return approxSubstringDistance(t, p) <= Math.floor(p.length * 0.12);
};

export interface PageDetect {
  type: DocTypeId | null;
  title: string;
  docNo: string;
}

// Nhận diện tiêu đề của 1 trang: ưu tiên dòng chữ to, nằm phía trên, ngắn (chỉ có tiêu đề)
export const detectTitle = (lines: SplitLine[], fuzzy = false): PageDetect => {
  if (lines.length === 0) return { type: null, title: '', docNo: '' };
  const sizes = lines.map((l) => l.size).sort((a, b) => a - b);
  const median = sizes[Math.floor(sizes.length / 2)] || 1;
  let best: { score: number; type: DocTypeId; title: string } | null = null;
  for (const line of lines) {
    const big = line.size >= median * 1.25;
    if (line.top > 0.45 && !big) continue;
    const n = norm(line.text);
    if (!n) continue;
    for (const def of DOC_TYPES) {
      const candidates = [...def.phrases.map((p) => ({ p, min: 0.25 })), ...(def.strict || []).map((p) => ({ p, min: 0.7 }))];
      for (const { p: phrase, min } of candidates) {
        // Dòng tiêu đề thường ngắn: bỏ qua câu dài chỉ nhắc tới tên chứng từ (vd "...2 invoice, 2 packing list...")
        if (phrase.length / n.length < min) continue;
        if (!matchPhrase(n, phrase, fuzzy)) continue;
        const score = line.size / median + (1 - line.top) * 1.5 + Math.min(1, phrase.length / n.length);
        if (!best || score > best.score) best = { score, type: def.id, title: line.text.trim() };
      }
    }
  }
  if (!best) return { type: null, title: '', docNo: '' };
  const text = lines.map((l) => l.text).join('\n');
  const def = docTypeOf(best.type);
  return { type: best.type, title: best.title, docNo: def.docNo ? def.docNo(text, lines) : '' };
};

// Tiêu đề lạ: dòng chữ to (>= 1.5 lần cỡ chữ phổ biến) ở 1/3 trên trang, viết hoa, không phải số
export const prominentTitle = (lines: SplitLine[]): string => {
  if (lines.length < 3) return '';
  const sizes = lines.map((l) => l.size).sort((a, b) => a - b);
  const median = sizes[Math.floor(sizes.length / 2)] || 1;
  let best: SplitLine | null = null;
  for (const l of lines) {
    if (l.top > 0.35 || l.size < median * 1.5) continue;
    const text = l.text.trim();
    const letters = text.replace(/[^\p{L}]/gu, '');
    const digits = text.replace(/\D/g, '');
    if (letters.length < 6 || text.length > 90 || digits.length > text.length * 0.2) continue;
    const upper = letters.replace(/[^\p{Lu}]/gu, '').length;
    if (upper < letters.length * 0.7) continue;
    if (!best || l.size > best.size || (l.size === best.size && l.top < best.top)) best = l;
  }
  return best ? best.text.trim() : '';
};

// "PHIẾU GIAO HÀNG" -> "Phiếu giao hàng"
const sentenceCase = (s: string) => {
  const lower = s.toLocaleLowerCase('vi');
  return lower.charAt(0).toLocaleUpperCase('vi') + lower.slice(1);
};

// ---------- Gán trang vào từng bộ chứng từ ----------
export interface PageAssign {
  index: number;
  target: PageTarget;
  segment: number;
  title: string;
  docNo: string;
  how: 'title' | 'continue' | 'ocr' | 'manual' | 'follow' | 'scan' | 'skip';
  // how = 'follow': trang chọn tay phía trước mà trang này đang theo
  from?: number;
}

export interface Segment {
  id: number;
  type: DocTypeId;
  docNo: string;
  pages: number[];
  // Tên chứng từ: theo loại, hoặc theo tiêu đề lạ đọc được trên trang (loại "other")
  label: string;
  titleKey: string;
}

export const assignPages = (pages: SplitPage[], overrides: Record<number, PageTarget> = {}, opts: { ocrToScan?: boolean } = {}) => {
  const assigns: PageAssign[] = [];
  const segments: Segment[] = [];
  let cur: Segment | null = null;
  let prevScanSegment: Segment | null = null;
  // Lựa chọn tay đang "lan" xuống các trang phía sau
  let run: { target: PageTarget; from: number } | null = null;

  const start = (type: DocTypeId, docNo: string, title = '') => {
    const label = type === 'other' && title ? sentenceCase(title) : docTypeOf(type).label;
    cur = { id: segments.length, type, docNo, pages: [], label, titleKey: title ? norm(title) : '' };
    segments.push(cur);
    return cur;
  };

  for (const page of pages) {
    let override = overrides[page.index];
    const readable = !page.isScan;
    // ocrToScan: trang scan luôn vào scan.pdf (kể cả khi OCR đã đọc được tiêu đề), trừ khi chọn tay
    const lines = readable ? page.lines : opts.ocrToScan ? [] : page.ocr || [];
    const det = lines.length ? detectTitle(lines, !readable) : { type: null, title: '', docNo: '' };

    let followed = false;
    if (override) {
      // "Bỏ trang" chỉ áp dụng cho đúng trang đó, không cắt đứt loại đang lan
      if (override !== 'skip') run = { target: override, from: page.index };
    } else if (run) {
      // Gặp trang có tiêu đề nhận ra được -> dừng lan, trang đó theo nhận diện tự động
      if (det.type || (readable && prominentTitle(lines))) run = null;
      else {
        override = run.target;
        followed = true;
      }
    }
    const from = followed && run ? run.from : undefined;

    if (override === 'skip' || override === 'scan') {
      const how = override === 'skip' ? 'skip' : followed ? 'follow' : 'manual';
      assigns.push({ index: page.index, target: override, segment: -1, title: det.title, docNo: det.docNo, how, from });
      if (override === 'scan') prevScanSegment = null;
      continue;
    }

    let type: DocTypeId | null = override ?? det.type;
    let docNo = det.docNo;
    if (override && override !== det.type) {
      const def = docTypeOf(override);
      docNo = def.docNo ? def.docNo(lines.map((l) => l.text).join('\n'), lines) : '';
    }

    if (!readable && !override && !type) {
      // Trang scan không đọc được tiêu đề: đi liền sau trang scan đã nhận ra loại -> cùng bộ, còn lại vào scan.pdf
      const follow = prevScanSegment;
      if (follow && page.ocr !== undefined && assigns.length && assigns[assigns.length - 1].segment === follow.id) {
        follow.pages.push(page.index);
        assigns.push({ index: page.index, target: follow.type, segment: follow.id, title: '', docNo: follow.docNo, how: 'ocr' });
      } else {
        assigns.push({ index: page.index, target: 'scan', segment: -1, title: '', docNo: '', how: 'scan' });
        prevScanSegment = null;
      }
      continue;
    }

    let how: PageAssign['how'] = followed ? 'follow' : override ? 'manual' : readable ? 'title' : 'ocr';
    const c = cur as Segment | null;
    let seg: Segment;
    const strange = !type && readable ? prominentTitle(lines) : '';
    if (strange && (!c || norm(strange) !== c.titleKey)) {
      // Tiêu đề chưa có trong danh sách -> chứng từ riêng (không gộp vào bộ phía trước)
      seg = start('other', '', strange);
      type = 'other';
      docNo = '';
      how = 'title';
    } else if (!type) {
      // Không có tiêu đề -> trang tiếp theo của chứng từ phía trước
      seg = c ?? start('other', '');
      type = seg.type;
      docNo = seg.docNo;
      how = 'continue';
    } else if (!c || c.type !== type || (docNo && c.docNo && docNo !== c.docNo) || (override && !followed)) {
      // Trang chọn tay luôn là trang đầu của 1 chứng từ mới
      seg = start(type, docNo);
    } else {
      seg = c;
      if (!seg.docNo && docNo) seg.docNo = docNo;
    }
    seg.pages.push(page.index);
    assigns.push({ index: page.index, target: type, segment: seg.id, title: det.title || strange, docNo: docNo || seg.docNo, how, from });
    prevScanSegment = readable ? null : seg;
  }
  return { assigns, segments };
};

// PDF chỉ gồm Commercial Invoice (kèm Packing List) -> nên dùng "Đọc PDF" trích Excel thay vì tách
export const isInvoiceOnly = (assigns: PageAssign[]) =>
  assigns.some((a) => a.target === 'commercial-invoice') && assigns.every((a) => a.target === 'commercial-invoice' || a.target === 'packing-list');

// ---------- Danh sách file PDF xuất ra ----------
export type GroupMode = 'type' | 'document';

export interface SplitOutput {
  key: string;
  target: DocTypeId | 'scan';
  code: string;
  label: string;
  fileName: string;
  pages: number[];
  docNos: string[];
  // Các loại chứng từ nằm trong file (file Other có thể gồm nhiều loại)
  kinds: string[];
  color: string;
}

const safeName = (s: string) =>
  s
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

// Tên file = mã chứng từ: TKX.pdf, SC.pdf, ANNEX.pdf, INV.pdf, PKL.pdf, VAT.pdf, Other.pdf, SCAN.pdf
// Nhiều file cùng mã (chế độ tách từng bộ) -> thêm số chứng từ: SC_SCNK267111.pdf (không có số thì SC_2.pdf)
export const buildOutputs = (assigns: PageAssign[], segments: Segment[], mode: GroupMode): SplitOutput[] => {
  const outputs: SplitOutput[] = [];
  const nonEmpty = segments.filter((s) => s.pages.length > 0);
  const make = (key: string, seg: Segment): SplitOutput => {
    const def = docTypeOf(seg.type);
    const other = def.code === OTHER_CODE;
    return { key, target: seg.type, code: def.code, label: other ? 'Chứng từ khác' : def.label, fileName: '', pages: [], docNos: [], kinds: [], color: other ? '#64748b' : def.color };
  };
  const add = (out: SplitOutput, seg: Segment) => {
    out.pages.push(...seg.pages);
    if (seg.docNo && !out.docNos.includes(seg.docNo)) out.docNos.push(seg.docNo);
    if (!out.kinds.includes(seg.label)) out.kinds.push(seg.label);
  };
  if (mode === 'type') {
    // Mỗi mã 1 file (các loại không có mã riêng gom chung vào Other.pdf)
    for (const seg of nonEmpty) {
      const code = docTypeOf(seg.type).code;
      let out = outputs.find((o) => o.key === code);
      if (!out) outputs.push((out = make(code, seg)));
      add(out, seg);
    }
  } else {
    for (const seg of nonEmpty) {
      const out = make(`seg-${seg.id}`, seg);
      if (out.code === OTHER_CODE) out.label = seg.label;
      add(out, seg);
      outputs.push(out);
    }
  }
  outputs.forEach((o) => o.pages.sort((a, b) => a - b));
  const used = new Set<string>();
  const countByCode = new Map<string, number>();
  outputs.forEach((o) => countByCode.set(o.code, (countByCode.get(o.code) || 0) + 1));
  const seen = new Map<string, number>();
  for (const o of outputs) {
    const n = (seen.get(o.code) || 0) + 1;
    seen.set(o.code, n);
    let name = o.code;
    if ((countByCode.get(o.code) || 0) > 1) name = `${o.code}_${o.docNos.length === 1 ? safeName(o.docNos[0]) : n}`;
    // Trùng tên (vd 2 bộ cùng số chứng từ) -> thêm số thứ tự
    let unique = name;
    for (let k = 2; used.has(unique.toLowerCase()); k++) unique = `${name}_${k}`;
    used.add(unique.toLowerCase());
    o.fileName = `${unique}.pdf`;
  }
  const scanPages = assigns.filter((a) => a.target === 'scan').map((a) => a.index);
  if (scanPages.length) {
    outputs.push({ key: 'scan', target: 'scan', code: SCAN_CODE, label: 'Trang scan', fileName: `${SCAN_CODE}.pdf`, pages: scanPages, docNos: [], kinds: ['Trang scan'], color: SCAN_COLOR });
  }
  return outputs;
};

// Liệt kê trang dạng "1–6, 9, 12–14" (đánh số từ 1)
export const pageRanges = (pages: number[]) => {
  const sorted = [...pages].sort((a, b) => a - b);
  const parts: string[] = [];
  for (let i = 0; i < sorted.length; i++) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    parts.push(i === j ? `${sorted[i] + 1}` : `${sorted[i] + 1}–${sorted[j] + 1}`);
    i = j;
  }
  return parts.join(', ');
};
