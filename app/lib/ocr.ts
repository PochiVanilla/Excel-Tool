// ==========================================
// OCR TRANG SCAN BẰNG tesseract.js (nạp từ CDN khi cần, tiếng Việt + tiếng Anh)
// Lần đầu tải ~5MB dữ liệu ngôn ngữ, các lần sau trình duyệt dùng lại bộ nhớ đệm
// - Nhiều worker chạy song song (mỗi worker dùng 1 nhân CPU) -> file scan nhiều trang đọc nhanh hơn
// - Mỗi worker được làm mới sau vài trang: tesseract.js giữ bộ nhớ WASM tăng dần sau mỗi lần đọc
// ==========================================
import type { SplitLine } from './pdfSplit';
import type { LuckyValue } from './sheetText';

const VERSION = '5.1.1';
const SCRIPT_URL = `https://cdn.jsdelivr.net/npm/tesseract.js@${VERSION}/dist/tesseract.min.js`;
const WORKER_URL = `https://cdn.jsdelivr.net/npm/tesseract.js@${VERSION}/dist/worker.min.js`;
const CORE_URL = `https://cdn.jsdelivr.net/npm/tesseract.js-core@${VERSION}`;

const RECYCLE_AFTER = 8;
const RECOGNIZE_TIMEOUT = 60000;

let scriptLoading: Promise<LuckyValue> | null = null;
let onStatus: ((status: string) => void) | null = null;

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

// Chỉ báo các bước khởi động (tải bộ OCR / dữ liệu); tiến độ đọc từng trang do nơi gọi tự đếm
const STATUS_LABELS: Record<string, string> = {
  'loading tesseract core': 'Đang tải bộ OCR',
  'initializing tesseract': 'Đang khởi động OCR',
  'loading language traineddata': 'Đang tải dữ liệu tiếng Việt',
  'initializing api': 'Đang khởi động OCR',
};

// Số worker chạy song song: chừa nhân CPU cho trình duyệt; máy ít RAM thì dùng ít worker (mỗi worker ~150MB)
export const ocrConcurrency = () => {
  const cores = navigator.hardwareConcurrency || 4;
  const memory = (navigator as Navigator & { deviceMemory?: number }).deviceMemory ?? 8;
  if (cores >= 8 && memory >= 8) return 3;
  if (cores >= 4 && memory >= 4) return 2;
  return 1;
};

type Slot = { worker: Promise<LuckyValue> | null; uses: number; busy: boolean };
const slots: Slot[] = [];
const waiting: (() => void)[] = [];

const createWorker = () =>
  loadScript().then(async (T) => {
    const worker = await T.createWorker(['vie', 'eng'], 1, {
      workerPath: WORKER_URL,
      corePath: CORE_URL,
      logger: (m: { status: string }) => {
        const label = STATUS_LABELS[m.status];
        if (label) onStatus?.(label);
      },
    });
    // PSM 11 (chữ rời rạc): vùng tiêu đề có logo, mã vạch, bảng -> chế độ bố cục mặc định hay bỏ sót dòng tiêu đề
    await worker.setParameters({ tessedit_pageseg_mode: '11' });
    return worker;
  });

const dropWorker = (slot: Slot) => {
  const p = slot.worker;
  slot.worker = null;
  slot.uses = 0;
  if (p) p.then((w) => w.terminate()).catch(() => {});
};

// Lấy 1 worker đang rảnh (chờ nếu tất cả đang bận)
const acquire = async (): Promise<Slot> => {
  while (slots.length < ocrConcurrency()) slots.push({ worker: null, uses: 0, busy: false });
  for (;;) {
    const free = slots.find((s) => !s.busy);
    if (free) {
      free.busy = true;
      return free;
    }
    await new Promise<void>((resolve) => waiting.push(resolve));
  }
};

const release = (slot: Slot) => {
  slot.busy = false;
  waiting.shift()?.();
};

type OcrBox = { x0: number; y0: number; y1: number };

// Đọc chữ trên ảnh -> các dòng kèm cỡ chữ và vị trí (tỉ lệ từ mép trên của cả trang)
// fullHeight: chiều cao cả trang (px) khi canvas chỉ là phần trên của trang
export const ocrCanvas = async (canvas: HTMLCanvasElement, status?: (s: string) => void, fullHeight = canvas.height): Promise<SplitLine[]> => {
  const slot = await acquire();
  let data: LuckyValue;
  let timer = 0;
  try {
    if (slot.uses >= RECYCLE_AFTER) dropWorker(slot);
    if (!slot.worker) {
      onStatus = status || null;
      slot.worker = createWorker();
    }
    const worker = await slot.worker;
    slot.uses++;
    // Trang bị treo (ảnh lỗi / hết bộ nhớ) -> huỷ worker và báo lỗi để trang đó được đọc lại, không làm đứng cả lượt quét
    const timeout = new Promise<never>((_, reject) => {
      timer = window.setTimeout(() => reject(new Error('OCR quá thời gian')), RECOGNIZE_TIMEOUT);
    });
    ({ data } = await Promise.race([worker.recognize(canvas), timeout]));
  } catch (err) {
    dropWorker(slot);
    throw err;
  } finally {
    window.clearTimeout(timer);
    release(slot);
  }
  const lines: SplitLine[] = [];
  for (const l of (data.lines || []) as { text: string; bbox: OcrBox; confidence: number; words?: { bbox: OcrBox }[] }[]) {
    const text = (l.text || '').replace(/\s+/g, ' ').trim();
    if (text.length < 2 || l.confidence < 30) continue;
    // Cỡ chữ = chiều cao chữ phổ biến trong dòng (khung dòng có thể cao gấp đôi khi dính 2 hàng của bảng)
    const heights = (l.words || []).map((w) => w.bbox.y1 - w.bbox.y0).sort((a, b) => a - b);
    const size = heights.length ? heights[Math.floor(heights.length / 2)] : l.bbox.y1 - l.bbox.y0;
    lines.push({ text, size: Math.max(1, size), top: l.bbox.y0 / fullHeight, left: l.bbox.x0 / canvas.width });
  }
  return lines;
};

export const terminateOcr = async () => {
  slots.forEach(dropWorker);
};
