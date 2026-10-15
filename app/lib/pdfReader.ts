// ==========================================
// ĐỌC LỚP CHỮ CỦA FILE PDF BẰNG pdf.js (nạp từ CDN khi cần, giống cách nạp Luckysheet)
// ==========================================
import type { PdfTextItem } from './pdfInvoice';
import type { SplitLine, SplitPage } from './pdfSplit';
import type { VRule } from './pdfTable';
import type { LuckyValue } from './sheetText';

const PDFJS_VERSION = '3.11.174';
const PDFJS_URL = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.min.js`;
const WORKER_URL = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.min.js`;

let loading: Promise<LuckyValue> | null = null;

const loadPdfJs = (): Promise<LuckyValue> => {
  const w = window as Window & { pdfjsLib?: LuckyValue };
  if (w.pdfjsLib) return Promise.resolve(w.pdfjsLib);
  if (!loading) {
    loading = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = PDFJS_URL;
      s.async = true;
      s.onload = () => {
        if (!w.pdfjsLib) return reject(new Error('Không khởi tạo được pdf.js'));
        w.pdfjsLib.GlobalWorkerOptions.workerSrc = WORKER_URL;
        resolve(w.pdfjsLib);
      };
      s.onerror = () => {
        loading = null;
        reject(new Error('Không tải được thư viện đọc PDF (pdf.js). Kiểm tra kết nối mạng.'));
      };
      document.head.appendChild(s);
    });
  }
  return loading;
};

// ==========================================
// ĐỌC TỪNG TRANG ĐỂ TÁCH BỘ CHỨNG TỪ (chữ + cỡ chữ + vị trí, độ phủ ảnh để nhận trang scan)
// ==========================================
export type PdfDoc = LuckyValue;

export const openPdfDocument = async (bytes: Uint8Array): Promise<PdfDoc> => {
  const pdfjs = await loadPdfJs();
  // pdf.js chuyển buffer sang worker (buffer gốc bị "detach") -> đưa bản sao, giữ bản gốc để cắt trang
  return pdfjs.getDocument({ data: bytes.slice() }).promise;
};

type RawItem = { str?: string; transform: number[]; width: number };

export const readSplitPage = async (doc: PdfDoc, index: number): Promise<SplitPage> => {
  const pdfjs = await loadPdfJs();
  const page = await doc.getPage(index + 1);
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();
  const pieces: { str: string; x: number; y: number; size: number; w: number }[] = [];
  let chars = 0;
  for (const it of content.items as RawItem[]) {
    if (typeof it.str !== 'string' || !it.str.trim()) continue;
    // Đổi sang toạ độ màn hình (gốc trên-trái, đã tính xoay trang)
    const t = pdfjs.Util.transform(viewport.transform, it.transform);
    const size = Math.hypot(t[2], t[3]) || Math.hypot(t[0], t[1]) || 1;
    pieces.push({ str: it.str, x: t[4], y: t[5], size, w: it.width });
    chars += it.str.replace(/\s+/g, '').length;
  }
  pieces.sort((a, b) => a.y - b.y || a.x - b.x);
  const lines: SplitLine[] = [];
  let group: typeof pieces = [];
  const flush = () => {
    if (!group.length) return;
    group.sort((a, b) => a.x - b.x);
    // Cùng hàng nhưng cách xa nhau (vd tiêu đề ở giữa, "SC No.: ..." ở mép phải) -> khối chữ riêng
    let chunk: typeof pieces = [];
    const push = () => {
      if (!chunk.length) return;
      let text = '';
      let prevEnd: number | null = null;
      for (const p of chunk) {
        if (prevEnd !== null && p.x - prevEnd > 1 && !text.endsWith(' ') && !p.str.startsWith(' ')) text += ' ';
        text += p.str;
        prevEnd = p.x + p.w;
      }
      const size = Math.max(...chunk.map((p) => p.size));
      const top = Math.min(...chunk.map((p) => p.y - p.size));
      if (text.trim()) lines.push({ text: text.trim(), size, top: Math.max(0, top / viewport.height) });
      chunk = [];
    };
    for (const p of group) {
      const last = chunk[chunk.length - 1];
      if (last && p.x - (last.x + last.w) > Math.max(24, 3 * Math.max(p.size, last.size))) push();
      chunk.push(p);
    }
    push();
    group = [];
  };
  for (const p of pieces) {
    if (group.length && Math.abs(group[0].y - p.y) > 2.5) flush();
    group.push(p);
  }
  flush();

  // Trang ít chữ -> đo ảnh lớn nhất trên trang (định thức ma trận biến đổi = diện tích ảnh)
  let imageCoverage = 0;
  let hasImage = false;
  if (chars < 150) {
    const ops = await page.getOperatorList();
    const OPS = pdfjs.OPS;
    const IMAGE_OPS = new Set([OPS.paintImageXObject, OPS.paintInlineImageXObject, OPS.paintJpegXObject, OPS.paintImageXObjectRepeat]);
    const [x0, y0, x1, y1] = page.view as number[];
    const pageArea = Math.abs((x1 - x0) * (y1 - y0)) || 1;
    let det = 1;
    const stack: number[] = [];
    for (let i = 0; i < ops.fnArray.length; i++) {
      const fn = ops.fnArray[i];
      const args = ops.argsArray[i];
      if (fn === OPS.save) stack.push(det);
      else if (fn === OPS.restore) det = stack.length ? stack.pop()! : det;
      else if (fn === OPS.transform) det *= args[0] * args[3] - args[1] * args[2];
      else if (fn === OPS.paintFormXObjectBegin) {
        stack.push(det);
        const m = args?.[0] as ArrayLike<number> | null | undefined;
        if (m && m.length >= 4) det *= m[0] * m[3] - m[1] * m[2];
      } else if (fn === OPS.paintFormXObjectEnd) det = stack.length ? stack.pop()! : det;
      else if (IMAGE_OPS.has(fn)) {
        hasImage = true;
        imageCoverage = Math.max(imageCoverage, Math.min(1, Math.abs(det) / pageArea));
      }
    }
  }
  page.cleanup();
  return {
    index,
    width: viewport.width,
    height: viewport.height,
    lines,
    chars,
    imageCoverage,
    isScan: chars < 150 && (imageCoverage >= 0.5 || (chars === 0 && hasImage)),
  };
};

// Vẽ lần lượt từng trang (ảnh thu nhỏ và OCR không vẽ cùng lúc -> tránh bộ nhớ tăng vọt với file scan nhiều trang)
let renderQueue: Promise<unknown> = Promise.resolve();

// Vẽ 1 trang ra canvas (ảnh thu nhỏ / OCR). cropTop: chỉ lấy phần trên của trang (tỉ lệ chiều cao)
export const renderPdfPage = (doc: PdfDoc, index: number, targetWidth: number, cropTop = 1): Promise<HTMLCanvasElement> => {
  const job = renderQueue.then(() => drawPage(doc, index, targetWidth, cropTop));
  renderQueue = job.catch(() => {});
  return job;
};

const drawPage = async (doc: PdfDoc, index: number, targetWidth: number, cropTop: number): Promise<HTMLCanvasElement> => {
  const page = await doc.getPage(index + 1);
  const base = page.getViewport({ scale: 1 });
  const viewport = page.getViewport({ scale: targetWidth / base.width });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height * Math.min(1, Math.max(0.05, cropTop)));
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport }).promise;
  page.cleanup();
  return canvas;
};

export interface PdfTextResult {
  items: PdfTextItem[];
  pageCount: number;
  // Số trang không có lớp chữ (PDF dạng ảnh scan)
  imageOnlyPages: number;
  // Đường kẻ dọc của bảng (chỉ đọc khi cần: withRules)
  rules: VRule[];
}

// Đường kẻ dọc của bảng (toạ độ trang PDF, gốc dưới-trái giống chữ) -> dùng để chia cột khi trích bảng ra Excel
const readVerticalRules = async (pdfjs: LuckyValue, page: LuckyValue, pageNo: number): Promise<VRule[]> => {
  const OPS = pdfjs.OPS;
  const ops = await page.getOperatorList();
  const rules: VRule[] = [];
  let ctm = [1, 0, 0, 1, 0, 0];
  const stack: number[][] = [];
  const pt = (x: number, y: number) => pdfjs.Util.applyTransform([x, y], ctm) as number[];
  const addSeg = (a: number[], b: number[]) => {
    if (Math.abs(a[0] - b[0]) < 1.2 && Math.abs(a[1] - b[1]) > 6) {
      rules.push({ page: pageNo, x: (a[0] + b[0]) / 2, y0: Math.min(a[1], b[1]), y1: Math.max(a[1], b[1]) });
    }
  };
  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i];
    const args = ops.argsArray[i];
    if (fn === OPS.save) stack.push(ctm);
    else if (fn === OPS.restore) ctm = stack.pop() || ctm;
    else if (fn === OPS.transform) ctm = pdfjs.Util.transform(ctm, args);
    else if (fn === OPS.paintFormXObjectBegin) {
      stack.push(ctm);
      const m = args?.[0] as number[] | null;
      if (m && m.length >= 6) ctm = pdfjs.Util.transform(ctm, m);
    } else if (fn === OPS.paintFormXObjectEnd) ctm = stack.pop() || ctm;
    else if (fn === OPS.constructPath) {
      const pathOps = args[0] as number[];
      const coords = args[1] as number[];
      let k = 0;
      let cur: number[] | null = null;
      for (const op of pathOps) {
        if (op === OPS.moveTo) {
          cur = pt(coords[k], coords[k + 1]);
          k += 2;
        } else if (op === OPS.lineTo) {
          const next = pt(coords[k], coords[k + 1]);
          k += 2;
          if (cur) addSeg(cur, next);
          cur = next;
        } else if (op === OPS.curveTo) k += 6;
        else if (op === OPS.curveTo2 || op === OPS.curveTo3) k += 4;
        else if (op === OPS.rectangle) {
          const [x, y, w, h] = coords.slice(k, k + 4);
          k += 4;
          const a = pt(x, y);
          const b = pt(x + w, y + h);
          const width = Math.abs(b[0] - a[0]);
          const height = Math.abs(b[1] - a[1]);
          if (height > 6) {
            // Hình chữ nhật mảnh = 1 đường kẻ; ô có viền = 2 cạnh trái / phải
            if (width < 2.5) rules.push({ page: pageNo, x: (a[0] + b[0]) / 2, y0: Math.min(a[1], b[1]), y1: Math.max(a[1], b[1]) });
            else {
              addSeg([a[0], a[1]], [a[0], b[1]]);
              addSeg([b[0], a[1]], [b[0], b[1]]);
            }
          }
        }
      }
    }
  }
  return rules;
};

export const readPdfText = async (file: File, opts: { withRules?: boolean } = {}): Promise<PdfTextResult> => {
  const pdfjs = await loadPdfJs();
  const data = new Uint8Array(await file.arrayBuffer());
  const doc = await pdfjs.getDocument({ data }).promise;
  const pageCount: number = doc.numPages;
  const items: PdfTextItem[] = [];
  const rules: VRule[] = [];
  let imageOnlyPages = 0;
  for (let p = 1; p <= pageCount; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    let count = 0;
    for (const it of content.items as { str?: string; transform: number[]; width: number }[]) {
      if (typeof it.str !== 'string') continue;
      if (it.str.trim()) count++;
      items.push({ str: it.str, x: it.transform[4], y: it.transform[5], w: it.width, page: p });
    }
    if (count === 0) imageOnlyPages++;
    else if (opts.withRules) rules.push(...(await readVerticalRules(pdfjs, page, p)));
    page.cleanup();
  }
  await doc.destroy();
  return { items, pageCount, imageOnlyPages, rules };
};
