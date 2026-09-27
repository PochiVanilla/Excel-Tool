// ==========================================
// ENGINE CẮT GỘP (chỉ LẬP KẾ HOẠCH các hàng cần xoá, không đụng vào bảng tính)
// -> Giao diện dùng kế hoạch này để tô màu xem trước, liệt kê và chỉ xoá khi người dùng bấm Áp dụng
// ==========================================
import {
  CompanyMode,
  SheetData,
  SheetRow,
  cleanForMatch,
  getCellText,
  getCleanedRangeText,
  isRowBlank,
  rowTexts,
} from './sheetText';
import { classifyRow, findInvoiceHeaders, isPageMarkerRow } from './invoiceTable';

export type TrimReason =
  | 'repeat-header'
  | 'repeat-footer'
  | 'page-marker'
  | 'blank'
  | 'template'
  | 'spacer'
  | 'rpac-header'
  | 'outside-table'
  | 'index-row';

export const TRIM_REASON_LABELS: Record<TrimReason, string> = {
  'repeat-header': 'Header lặp lại',
  'repeat-footer': 'Footer trùng lặp',
  'page-marker': 'Dòng số trang',
  blank: 'Hàng trống',
  template: 'Khớp mẫu',
  spacer: 'Dòng đệm không viền',
  'rpac-header': 'Header R-pac',
  'outside-table': 'Ngoài bảng hàng hoá',
  'index-row': 'Dòng đánh số cột',
};

export interface TrimBlock {
  start: number;
  end: number;
  reason: TrimReason;
}

export interface TrimPlan {
  rows: number[];
  blocks: TrimBlock[];
  notes: string[];
  // Vùng tiêu đề gốc được giữ lại (để tô xanh khi xem trước)
  keep?: { start: number; end: number };
}

export interface AutoTrimOptions {
  repeatHeaders: boolean;
  repeatFooters: boolean;
  pageMarkers: boolean;
  blankRows: boolean;
  tableOnly: boolean;
}

export interface TemplateSelection {
  pattern: string[];
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

// Gom các lý do theo từng hàng -> danh sách khối liên tiếp cùng lý do
const finalizePlan = (reasons: Map<number, TrimReason>, notes: string[], keep?: TrimPlan['keep']): TrimPlan => {
  const rows = Array.from(reasons.keys()).sort((a, b) => a - b);
  const blocks: TrimBlock[] = [];
  for (const r of rows) {
    const reason = reasons.get(r)!;
    const last = blocks[blocks.length - 1];
    if (last && last.end === r - 1 && last.reason === reason) last.end = r;
    else blocks.push({ start: r, end: r, reason });
  }
  return { rows, blocks, notes, keep };
};

// Hàng cuối cùng có nội dung (Luckysheet luôn đệm thêm hàng trống ở cuối lưới -> không tính là "hàng trống" cần xoá)
const lastContentRow = (sheetData: SheetData) => {
  for (let r = sheetData.length - 1; r >= 0; r--) if (!isRowBlank(sheetData[r], r)) return r;
  return -1;
};

const addRange = (reasons: Map<number, TrimReason>, start: number, end: number, reason: TrimReason) => {
  for (let r = start; r <= end; r++) if (!reasons.has(r)) reasons.set(r, reason);
};

// Chữ ký nội dung của cả dòng (bỏ số, dấu câu, khoảng trắng) để so sánh các dòng header lặp lại
const rowSignature = (row: SheetRow) => cleanForMatch(rowTexts(row).map((t) => t.text).join('|'));
const rowRawSignature = (row: SheetRow) => rowTexts(row).map((t) => t.text.replace(/\s+/g, '').toLowerCase()).join('|');

// Dòng ứng viên có "giống" dòng mẫu không: >= 80% số ô có chữ của mẫu xuất hiện trong dòng ứng viên
const rowSimilar = (candidate: SheetRow, template: SheetRow) => {
  const tCells = rowTexts(template).map((t) => cleanForMatch(t.text)).filter((t) => t !== '');
  if (tCells.length === 0) return false;
  const combined = rowTexts(candidate).map((t) => cleanForMatch(t.text)).join('');
  if (combined === '') return false;
  const hits = tCells.filter((t) => combined.includes(t)).length;
  return hits >= Math.max(1, Math.ceil(tCells.length * 0.8));
};

// ==========================================
// 1. CHẾ ĐỘ TỰ ĐỘNG (khuyên dùng cho hoá đơn VAT chuyển từ PDF)
// Neo theo dòng tiêu đề bảng (STT | Tên hàng | ... | Thành tiền): tiêu đề đầu tiên là gốc,
// mỗi lần tiêu đề xuất hiện lại -> dò ngược lên các dòng giống khối header gốc để xoá cả khối
// ==========================================
export const planAutoTrim = (sheetData: SheetData, opts: AutoTrimOptions): TrimPlan => {
  const reasons = new Map<number, TrimReason>();
  const notes: string[] = [];
  const headers = findInvoiceHeaders(sheetData);
  const lastRow = lastContentRow(sheetData);
  let keep: TrimPlan['keep'];

  if (headers.length === 0) {
    notes.push('Không nhận diện được dòng tiêu đề bảng (STT, Tên hàng, Số lượng...). Chỉ dọn dòng số trang / hàng trống. Với mẫu đặc biệt hãy dùng chế độ "Theo mẫu bôi chọn".');
  } else {
    const master = headers[0];
    const masterEnd = master.indexRow >= 0 ? master.indexRow : master.row;
    let top = 0;
    while (top < master.row && isRowBlank(sheetData[top], top)) top++;
    const templateRows: SheetRow[] = [];
    for (let r = top; r < master.row; r++) if (!isRowBlank(sheetData[r], r)) templateRows.push(sheetData[r]);
    const templateSigs = new Set(templateRows.map(rowSignature).filter((s) => s !== ''));
    // Dòng chỉ có số (vd mã số thuế "3801104574") bị bỏ hết khi chuẩn hoá -> so thêm theo nội dung gốc
    const templateRaw = new Set(templateRows.map(rowRawSignature).filter((s) => s !== ''));
    keep = { start: top, end: masterEnd };

    const isBody = (r: number) => {
      const kind = classifyRow(sheetData, r, master.cols);
      return kind === 'data' || kind === 'group';
    };

    // Vùng chân trang của từng trang: các dòng không phải hàng hoá nằm giữa dòng hàng cuối của trang và khối header trang sau
    const footerRegions: number[][] = [];
    const collectFooter = (from: number, to: number) => {
      const rows: number[] = [];
      for (let r = from; r <= to; r++) {
        const row = sheetData[r];
        if (isRowBlank(row, r) || isPageMarkerRow(row)) continue;
        rows.push(r);
      }
      footerRegions.push(rows);
    };

    let prevEnd = masterEnd;
    for (let k = 1; k < headers.length; k++) {
      const h = headers[k];
      const end = h.indexRow >= 0 ? h.indexRow : h.row;
      // Dò ngược lên: dòng giống khối header gốc thuộc header lặp; dòng trống / số trang xen giữa chỉ "đi qua",
      // chỉ xoá khi người dùng tích tuỳ chọn tương ứng
      const headerRows: number[] = [h.row];
      let blockTop = h.row;
      for (let j = h.row - 1; j > prevEnd; j--) {
        const row = sheetData[j];
        if (isRowBlank(row, j) || isPageMarkerRow(row)) {
          blockTop = j;
          continue;
        }
        if (templateSigs.has(rowSignature(row)) || templateRaw.has(rowRawSignature(row)) || templateRows.some((t) => rowSimilar(row, t))) {
          headerRows.push(j);
          blockTop = j;
          continue;
        }
        break;
      }
      // Dòng trống/số trang phía trên dòng header lặp đầu tiên thuộc về chân trang trước, không thuộc khối header
      const firstHeaderRow = Math.min(...headerRows);
      if (blockTop < firstHeaderRow) blockTop = firstHeaderRow;
      if (opts.repeatHeaders) {
        for (const r of headerRows) addRange(reasons, r, r, 'repeat-header');
        for (let r = h.row + 1; r <= end; r++) addRange(reasons, r, r, 'repeat-header');
      }
      let lastBody = blockTop - 1;
      while (lastBody > prevEnd && !isBody(lastBody)) lastBody--;
      collectFooter(lastBody + 1, blockTop - 1);
      prevEnd = end;
    }
    if (headers.length > 1 && opts.repeatHeaders) notes.push(`Phát hiện ${headers.length - 1} lần lặp lại tiêu đề bảng (giữ tiêu đề gốc ở dòng ${master.row + 1}).`);

    // Chân trang cuối cùng (sau dòng hàng cuối) -> bản được giữ lại khi xoá footer trùng
    let lastBody = lastRow;
    while (lastBody > prevEnd && !isBody(lastBody)) lastBody--;
    collectFooter(lastBody + 1, lastRow);

    if (opts.repeatFooters && footerRegions.length > 1) {
      // Footer mỗi hoá đơn có nội dung khác nhau (số tiền bằng chữ, ngày ký, cách tách ô...) nên so sánh theo CẢ KHỐI:
      // khối chân trang nào có ít nhất 1 dòng giống footer cuối cùng -> xoá cả khối, chỉ giữ lại footer cuối
      const finalRows = footerRegions[footerRegions.length - 1];
      const finalSigs = new Set(finalRows.map((r) => rowSignature(sheetData[r])).filter((s) => s !== ''));
      const looksLikeFinal = (r: number) =>
        finalSigs.has(rowSignature(sheetData[r])) || finalRows.some((f) => rowSimilar(sheetData[r], sheetData[f]) || rowSimilar(sheetData[f], sheetData[r]));
      let removed = 0;
      let blocks = 0;
      footerRegions.slice(0, -1).forEach((rows) => {
        if (rows.length === 0 || !rows.some(looksLikeFinal)) return;
        blocks++;
        rows.forEach((r) => {
          if (!reasons.has(r)) removed++;
          addRange(reasons, r, r, 'repeat-footer');
        });
      });
      if (removed > 0) notes.push(`Phát hiện ${blocks} khối chân trang lặp lại (${removed} dòng) — giữ lại footer cuối cùng.`);
    }

    if (opts.tableOnly) {
      addRange(reasons, 0, master.row - 1, 'outside-table');
      if (master.indexRow >= 0) addRange(reasons, master.indexRow, master.indexRow, 'index-row');
      for (let r = masterEnd + 1; r <= lastRow; r++) {
        if (reasons.has(r)) continue;
        const kind = classifyRow(sheetData, r, master.cols);
        if (kind === 'data' || kind === 'group') continue;
        // Dòng trống / số trang vẫn tuân theo tuỳ chọn riêng
        if (kind === 'blank' || kind === 'page') continue;
        addRange(reasons, r, r, kind === 'index' ? 'index-row' : 'outside-table');
      }
      keep = { start: master.row, end: master.row };
    }
  }

  for (let r = 0; r <= lastRow; r++) {
    if (reasons.has(r)) continue;
    if (opts.pageMarkers && isPageMarkerRow(sheetData[r])) reasons.set(r, 'page-marker');
    else if (opts.blankRows && isRowBlank(sheetData[r], r)) reasons.set(r, 'blank');
  }

  // Bảo vệ tuyệt đối vùng tiêu đề gốc
  if (keep) for (let r = keep.start; r <= keep.end; r++) if (reasons.get(r) !== 'outside-table' && reasons.get(r) !== 'index-row') reasons.delete(r);

  return finalizePlan(reasons, notes, keep);
};

// ==========================================
// 2. CHẾ ĐỘ THEO MẪU BÔI CHỌN (thuật toán chuẩn cũ - Normal Company)
// HEADER: giữ vùng mẫu, xoá các khối lặp lại phía dưới. FOOTER: xoá mọi khối khớp, kể cả vùng mẫu.
// ==========================================
const matchSingleRow = (row: SheetRow, patternRowText: string, c1: number, c2: number, mode: CompanyMode) => {
  if (!row) return false;
  const pCells = patternRowText.split('|').map((s) => s.trim()).filter((s) => s !== '');
  if (pCells.length === 0) return true;
  let rowTextCombined = '';
  for (let c = c1; c <= c2; c++) rowTextCombined += cleanForMatch(getCellText(row[c], mode));
  const matchCount = pCells.filter((p) => rowTextCombined.includes(p)).length;
  return matchCount >= Math.max(1, Math.ceil(pCells.length * 0.8));
};

export const planTemplateTrim = (
  sheetData: SheetData,
  sel: TemplateSelection,
  trimMode: 'HEADER' | 'FOOTER',
  mode: CompanyMode,
  rowHasBorder: (r: number) => boolean,
): TrimPlan => {
  const reasons = new Map<number, TrimReason>();
  const { pattern, c1, c2, r1, r2 } = sel;
  const isHeaderMode = trimMode === 'HEADER';
  const protectedRow = (r: number) => isHeaderMode && r >= r1 && r <= r2;
  const lastRowTemplate = pattern[pattern.length - 1];

  for (let r = 0; r < sheetData.length; r++) {
    if (protectedRow(r)) continue;
    if (!matchSingleRow(sheetData[r], pattern[0], c1, c2, mode)) continue;
    if (isHeaderMode && r <= r2) continue;

    // Tìm dòng cuối của khối (khớp dòng cuối của mẫu) trong phạm vi gần
    let endRow = -1;
    const maxScan = Math.min(sheetData.length - r, pattern.length + 3);
    for (let i = 0; i < maxScan; i++) {
      if (matchSingleRow(sheetData[r + i], lastRowTemplate, c1, c2, mode)) {
        endRow = r + i;
        break;
      }
    }

    if (endRow !== -1) {
      for (let t = r; t <= endRow; t++) if (!protectedRow(t)) addRange(reasons, t, t, 'template');
    } else {
      // Dự phòng: so khớp từng dòng với bất kỳ dòng nào của mẫu
      for (let i = 0; i < pattern.length; i++) {
        const t = r + i;
        if (t >= sheetData.length) break;
        const text = getCleanedRangeText(sheetData[t], c1, c2, mode);
        const matchesAny = text.replace(/\|/g, '') === '' || pattern.some((p) => matchSingleRow(sheetData[t], p, c1, c2, mode));
        if (!matchesAny) break;
        if (!protectedRow(t)) addRange(reasons, t, t, 'template');
        if (matchSingleRow(sheetData[t], lastRowTemplate, c1, c2, mode)) break;
      }
    }

    // Dọn các dòng đệm không có viền xung quanh khối (phân trang)
    const actualEnd = endRow !== -1 ? endRow : r + pattern.length - 1;
    for (let offset = 1; offset <= 2; offset++) {
      const above = r - offset;
      if (above >= 0 && !protectedRow(above) && !rowHasBorder(above)) addRange(reasons, above, above, 'spacer');
      const below = actualEnd + offset;
      if (below < sheetData.length && !protectedRow(below) && !rowHasBorder(below)) addRange(reasons, below, below, 'spacer');
    }
    r = actualEnd;
  }

  const notes: string[] = [];
  if (reasons.size === 0) notes.push(`Không tìm thấy hàng nào khớp với mẫu ${isHeaderMode ? 'Header' : 'Footer'} đã bôi chọn.`);
  return finalizePlan(reasons, notes, isHeaderMode ? { start: r1, end: r2 } : undefined);
};

// ==========================================
// 3. THUẬT TOÁN ĐẶC THÙ CHO CÔNG TY R-PAC (giữ nguyên nghiệp vụ cũ)
// Khối header bắt đầu từ dòng có tên "R-PAC VIỆT NAM" đến dòng tiêu đề STT/Hàng hoá
// ==========================================
export const planRpacTrim = (
  sheetData: SheetData,
  sel: TemplateSelection,
  trimMode: 'HEADER' | 'FOOTER',
  rowHasBorder: (r: number) => boolean,
): TrimPlan => {
  const reasons = new Map<number, TrimReason>();
  const { r1, r2 } = sel;
  const isHeaderMode = trimMode === 'HEADER';

  const cleanedRow = (row: SheetRow) => {
    if (!row) return '';
    let combined = '';
    for (let c = 0; c < row.length; c++) combined += getCellText(row[c], 'RPAC').toLowerCase();
    return combined.replace(/\s+/g, '').replace(/[0-9.,:\/\\#\-]/g, '');
  };
  const isStart = (row: SheetRow) => {
    const s = cleanedRow(row);
    return s.includes('côngtytnhhrpacviệtnam') || s.includes('rpacvietnam') || s.includes('rpacvietnamlimited');
  };
  const isEnd = (row: SheetRow) => {
    const s = cleanedRow(row);
    return (
      (s.includes('stt') || s.includes('no')) &&
      (s.includes('hànghóa') || s.includes('description') || s.includes('diễngiải') || s.includes('tênhàng') || s.includes('đvt') || s.includes('unit'))
    );
  };
  const inMaster = (r: number) => isHeaderMode && r >= r1 && r <= r2;

  for (let r = 0; r < sheetData.length; r++) {
    if (!isStart(sheetData[r])) continue;
    let endRow = -1;
    for (let i = r; i < Math.min(sheetData.length, r + 30); i++) {
      if (isEnd(sheetData[i])) {
        endRow = i;
        break;
      }
    }
    if (endRow === -1) continue;
    const isMasterBlock = (r >= r1 && r <= r2) || r === r1;
    if (!(isHeaderMode && isMasterBlock)) addRange(reasons, r, endRow, 'rpac-header');
    for (let offset = 1; offset <= 2; offset++) {
      const above = r - offset;
      if (above >= 0 && !inMaster(above) && !rowHasBorder(above)) addRange(reasons, above, above, 'spacer');
      const below = endRow + offset;
      if (below < sheetData.length && !inMaster(below) && !rowHasBorder(below)) addRange(reasons, below, below, 'spacer');
    }
    r = endRow;
  }

  const notes: string[] = [];
  if (reasons.size === 0) notes.push('Không tìm thấy khối header R-pac nào (dòng "R-PAC VIỆT NAM" ... dòng STT / Hàng hoá).');
  return finalizePlan(reasons, notes, isHeaderMode ? { start: r1, end: r2 } : undefined);
};

// ==========================================
// 4. XOÁ HÀNG TRỐNG (không xoá hàng nằm trong ô gộp nhiều hàng)
// ==========================================
export const planBlankRows = (sheetData: SheetData, r1 = 0, r2 = sheetData.length - 1): TrimPlan => {
  const reasons = new Map<number, TrimReason>();
  for (let r = Math.max(0, r1); r <= Math.min(r2, lastContentRow(sheetData)); r++) {
    if (isRowBlank(sheetData[r], r)) reasons.set(r, 'blank');
  }
  return finalizePlan(reasons, []);
};
