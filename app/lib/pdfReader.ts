// ==========================================
// ĐỌC LỚP CHỮ CỦA FILE PDF BẰNG pdf.js (nạp từ CDN khi cần, giống cách nạp Luckysheet)
// ==========================================
import type { PdfTextItem } from './pdfInvoice';
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

export interface PdfTextResult {
  items: PdfTextItem[];
  pageCount: number;
  // Số trang không có lớp chữ (PDF dạng ảnh scan)
  imageOnlyPages: number;
}

export const readPdfText = async (file: File): Promise<PdfTextResult> => {
  const pdfjs = await loadPdfJs();
  const data = new Uint8Array(await file.arrayBuffer());
  const doc = await pdfjs.getDocument({ data }).promise;
  const pageCount: number = doc.numPages;
  const items: PdfTextItem[] = [];
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
    page.cleanup();
  }
  await doc.destroy();
  return { items, pageCount, imageOnlyPages };
};
