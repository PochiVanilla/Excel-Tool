// ==========================================
// OCR TRANG SCAN BẰNG tesseract.js (nạp từ CDN khi cần, tiếng Việt + tiếng Anh)
// Lần đầu tải ~5MB dữ liệu ngôn ngữ, các lần sau trình duyệt dùng lại bộ nhớ đệm
// ==========================================
import type { SplitLine } from './pdfSplit';
import type { LuckyValue } from './sheetText';

const VERSION = '5.1.1';
const SCRIPT_URL = `https://cdn.jsdelivr.net/npm/tesseract.js@${VERSION}/dist/tesseract.min.js`;
const WORKER_URL = `https://cdn.jsdelivr.net/npm/tesseract.js@${VERSION}/dist/worker.min.js`;
const CORE_URL = `https://cdn.jsdelivr.net/npm/tesseract.js-core@${VERSION}`;

let scriptLoading: Promise<LuckyValue> | null = null;
let workerLoading: Promise<LuckyValue> | null = null;
let onStatus: ((status: string, progress: number) => void) | null = null;

const loadScript = (): Promise<LuckyValue> => {
  const w = window as Window & { Tesseract?: LuckyValue };
  if (w.Tesseract) return Promise.resolve(w.Tesseract);
  if (!scriptLoading) {
    scriptLoading = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = SCRIPT_URL;
      s.async = true;
      s.onload = () => (w.Tesseract ? resolve(w.Tesseract) : reject(new Error('Không khởi tạo được OCR')));
      s.onerror = () => {
        scriptLoading = null;
        reject(new Error('Không tải được thư viện OCR (tesseract.js). Kiểm tra kết nối mạng.'));
      };
      document.head.appendChild(s);
    });
  }
  return scriptLoading;
};

const STATUS_LABELS: Record<string, string> = {
  'loading tesseract core': 'Đang tải bộ OCR',
  'initializing tesseract': 'Đang khởi động OCR',
  'loading language traineddata': 'Đang tải dữ liệu tiếng Việt',
  'initializing api': 'Đang khởi động OCR',
  'recognizing text': 'Đang đọc chữ',
};

const getWorker = (): Promise<LuckyValue> => {
  if (!workerLoading) {
    workerLoading = loadScript()
      .then(async (T) => {
        const worker = await T.createWorker(['vie', 'eng'], 1, {
          workerPath: WORKER_URL,
          corePath: CORE_URL,
          logger: (m: { status: string; progress: number }) => onStatus?.(STATUS_LABELS[m.status] || m.status, m.progress || 0),
        });
        // PSM 11 (chữ rời rạc): vùng tiêu đề có logo, mã vạch, bảng -> chế độ bố cục mặc định hay bỏ sót dòng tiêu đề
        await worker.setParameters({ tessedit_pageseg_mode: '11' });
        return worker;
      })
      .catch((err) => {
        workerLoading = null;
        throw err;
      });
  }
  return workerLoading;
};

type OcrBox = { y0: number; y1: number };

// Đọc chữ trên ảnh -> các dòng kèm cỡ chữ và vị trí (tỉ lệ từ mép trên của cả trang)
// fullHeight: chiều cao cả trang (px) khi canvas chỉ là phần trên của trang
export const ocrCanvas = async (canvas: HTMLCanvasElement, status?: (s: string, p: number) => void, fullHeight = canvas.height): Promise<SplitLine[]> => {
  onStatus = status || null;
  const worker = await getWorker();
  const { data } = await worker.recognize(canvas);
  const lines: SplitLine[] = [];
  for (const l of (data.lines || []) as { text: string; bbox: OcrBox; confidence: number; words?: { bbox: OcrBox }[] }[]) {
    const text = (l.text || '').replace(/\s+/g, ' ').trim();
    if (text.length < 2 || l.confidence < 30) continue;
    // Cỡ chữ = chiều cao chữ phổ biến trong dòng (khung dòng có thể cao gấp đôi khi dính 2 hàng của bảng)
    const heights = (l.words || []).map((w) => w.bbox.y1 - w.bbox.y0).sort((a, b) => a - b);
    const size = heights.length ? heights[Math.floor(heights.length / 2)] : l.bbox.y1 - l.bbox.y0;
    lines.push({ text, size: Math.max(1, size), top: l.bbox.y0 / fullHeight });
  }
  onStatus = null;
  return lines;
};

export const terminateOcr = async () => {
  if (!workerLoading) return;
  const p = workerLoading;
  workerLoading = null;
  try {
    (await p).terminate();
  } catch {
    /* bỏ qua */
  }
};
