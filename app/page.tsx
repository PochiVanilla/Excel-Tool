'use client';

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

import Icon from './components/Icon';
import StatusBar, { SelectionStats } from './components/StatusBar';
import TrimPane, { TemplateInfo, TrimStrategy } from './components/TrimPane';
import PivotPane, { formatByPattern } from './components/PivotPane';
import ShortcutsDialog from './components/ShortcutsDialog';
import PdfPane from './components/PdfPane';
import SplitBoard from './components/SplitBoard';
import SplitPane from './components/SplitPane';
import { ToastStack, useToasts } from './components/Toasts';

import { exportWorkbookBuffer } from './lib/exportWorkbook';
import { FabricInvoice, buildFabricWorkbook, groupLines, parseFabricInvoice } from './lib/pdfInvoice';
import { readPdfText } from './lib/pdfReader';
import { assignPages, isInvoiceOnly } from './lib/pdfSplit';
import { useSplitPdf } from './lib/useSplitPdf';
import {
  buildPivotSource,
  computePivot,
  defaultLayout,
  InvoiceProfile,
  PivotLayout,
  PivotMode,
  PivotResult,
  PivotSource,
} from './lib/invoicePivot';
import {
  clearOverlay,
  deleteRowsWithImages,
  getActiveFile,
  getActiveOrder,
  getFiles,
  getSelection,
  getSheetData,
  isLuckyReady,
  isSelfRefreshing,
  ls,
  luckyExcel,
  overlayHooks,
  refreshSheet,
  selectRange,
  selectRows,
  setOverlay,
  snapshotWorkbook,
  workbookForReload,
  buildRowBorderChecker,
  autoFitRowHeight,
} from './lib/lucky';
import { ContextAction, autoFitColumns, installSheetInteractions, InteractionHandlers } from './lib/sheetInteractions';
import { LuckyValue, getCellNumber, getCellText, getCleanedRangeText, getRangeText, rangeName } from './lib/sheetText';
import { AutoTrimOptions, TrimBlock, TrimPlan, planAutoTrim, planBlankRows, planRpacTrim, planTemplateTrim } from './lib/trimEngine';

type Pane = 'trim' | 'pivot' | 'pdf' | 'split' | null;
type PdfState = { fileName: string; invoice: FabricInvoice | null; error: string | null; rawLines: string[] };
type UndoEntry = { label: string; sheets: LuckyValue[]; activeOrder: number; title: string };

const PANE_WIDTH_KEY = 'excel-tool.pane-width';
const DEFAULT_PANE_WIDTH = 400;

export default function Home() {
  const { toasts, push: toast, dismiss } = useToasts();

  // ---- Trạng thái chung ----
  const [ready, setReady] = useState(false);
  const [hasFile, setHasFile] = useState(false);
  const [fileName, setFileName] = useState('');
  const [original, setOriginal] = useState<ExcelJS.Workbook | null>(null);
  const [dirty, setDirty] = useState(false);
  const [loading, setLoading] = useState('');
  const [statusMsg, setStatusMsg] = useState('');
  const [stats, setStats] = useState<SelectionStats | null>(null);
  // Vùng đang bôi chọn (hàng) để dùng cho Pivot
  const [selRows, setSelRows] = useState<{ r1: number; r2: number; label: string } | null>(null);
  const [zoom, setZoom] = useState(1);
  const [pane, setPane] = useState<Pane>(null);
  const [paneWidth, setPaneWidth] = useState(() => {
    if (typeof window === 'undefined') return DEFAULT_PANE_WIDTH;
    try {
      const w = Number(localStorage.getItem(PANE_WIDTH_KEY));
      return w >= 320 && w <= 900 ? w : DEFAULT_PANE_WIDTH;
    } catch {
      return DEFAULT_PANE_WIDTH;
    }
  });
  const [showHelp, setShowHelp] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [undoStack, setUndoStack] = useState<UndoEntry[]>([]);
  // Tăng mỗi khi dữ liệu bảng tính thay đổi -> tính lại kế hoạch cắt gộp
  const [sheetVersion, setSheetVersion] = useState(0);

  // ---- Cắt gộp ----
  const [strategy, setStrategy] = useState<TrimStrategy>('auto');
  const [autoOpts, setAutoOpts] = useState<AutoTrimOptions>({ repeatHeaders: true, repeatFooters: true, pageMarkers: true, blankRows: true, tableOnly: false });
  const [trimMode, setTrimMode] = useState<'HEADER' | 'FOOTER'>('HEADER');
  const [template, setTemplate] = useState<TemplateInfo | null>(null);
  const [plan, setPlan] = useState<TrimPlan | null>(null);

  // ---- Pivot ----
  const [pivotScope, setPivotScope] = useState<{ r1: number; r2: number } | null>(null);
  const [pivotSource, setPivotSource] = useState<PivotSource | null>(null);
  const [pivotError, setPivotError] = useState<string | null>(null);
  const [pivotLayout, setPivotLayout] = useState<PivotLayout>(defaultLayout('generic'));
  // auto: tự nhận diện hoá đơn (không được thì Pivot thủ công) · table: Pivot thủ công theo dòng tiêu đề
  const [pivotMode, setPivotMode] = useState<PivotMode>('auto');

  // ---- Trích xuất PDF ----
  const [pdfState, setPdfState] = useState<PdfState | null>(null);
  // ---- Tách bộ chứng từ PDF ----
  const split = useSplitPdf(toast, setLoading);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);
  const openPdfRef = useRef<(file: File) => void>(() => {});
  const splitInputRef = useRef<HTMLInputElement>(null);
  const routePdfRef = useRef<(file: File) => void>(() => {});
  // Nghiệp vụ + danh sách trường của bố cục Pivot hiện tại -> đổi nghiệp vụ / đổi bảng thì dùng bố cục mặc định mới
  const layoutKeyRef = useRef<string | null>(null);
  // Luckysheet gọi hook ngoài chu kỳ render của React -> luôn đọc state mới nhất qua ref
  const stateRef = useRef({ pane, strategy, trimMode, zoom, hasFile, dirty, fileName, pivotMode });
  useLayoutEffect(() => {
    stateRef.current = { pane, strategy, trimMode, zoom, hasFile, dirty, fileName, pivotMode };
  });

  // ==========================================
  // THỐNG KÊ VÙNG CHỌN (đúng quy tắc Excel: Count = ô có dữ liệu, Average = Sum / số ô là số)
  // ==========================================
  const refreshStats = useCallback(() => {
    if (!isLuckyReady()) return;
    const ranges = getSelection();
    if (ranges.length === 0) return setStats(null);
    const first = ranges[0];
    setSelRows(
      first.row[1] > first.row[0]
        ? { r1: first.row[0], r2: first.row[1], label: rangeName(first.row[0], first.column[0], first.row[1], first.column[1]) }
        : null,
    );
    const data = getSheetData();
    const seen = new Set<string>();
    let cells = 0;
    let count = 0;
    let numCount = 0;
    let sum = 0;
    let min: number | null = null;
    let max: number | null = null;
    for (const rg of ranges) {
      cells += (rg.row[1] - rg.row[0] + 1) * (rg.column[1] - rg.column[0] + 1);
      for (let r = rg.row[0]; r <= Math.min(rg.row[1], data.length - 1); r++) {
        const row = data[r];
        if (!row) continue;
        for (let c = rg.column[0]; c <= Math.min(rg.column[1], row.length - 1); c++) {
          const cell = row[c];
          if (!cell || (cell.mc && (cell.mc.r !== r || cell.mc.c !== c))) continue;
          const key = `${r}_${c}`;
          if (seen.has(key)) continue;
          seen.add(key);
          const isNumber = typeof cell.v === 'number' || (cell.ct?.t === 'n' && cell.v !== '' && cell.v !== null && !isNaN(Number(cell.v)));
          if (isNumber) {
            const n = Number(cell.v);
            count++;
            numCount++;
            sum += n;
            min = min === null ? n : Math.min(min, n);
            max = max === null ? n : Math.max(max, n);
          } else if (getCellText(cell).trim() !== '') count++;
        }
      }
    }
    setStats({ cells, count, numCount, sum: Math.round(sum * 1e9) / 1e9, min, max });
  }, []);

  // ==========================================
  // KHỞI TẠO / TẠO LẠI LUCKYSHEET
  // ==========================================
  const handlersRef = useRef({
    onRangeSelect: (range: LuckyValue) => void range,
    onUpdated: () => {},
    onSheetActivate: () => {},
    onCreated: () => {},
  });

  const createSheet = useCallback((sheets?: LuckyValue[], title?: string) => {
    const l = ls();
    if (!l) return;
    setReady(false);
    try {
      l.destroy();
    } catch {
      /* chưa từng tạo */
    }
    l.create({
      container: 'luckysheet-container',
      lang: 'en',
      title: title || 'Excel Tool',
      showinfobar: false,
      showstatisticBar: false,
      // Thanh trạng thái riêng của web đã có Sum/Average/Count + thu phóng -> ẩn bản gốc của Luckysheet
      showstatisticBarConfig: { count: false, view: false, zoom: false },
      enableAddRow: true,
      enableAddBackTop: true,
      data: sheets && sheets.length ? sheets : [{ name: 'Sheet1', status: 1, order: 0, index: 0, celldata: [], config: {} }],
      hook: {
        rangeSelect: (_sheet: LuckyValue, range: LuckyValue) => handlersRef.current.onRangeSelect(range),
        cellRenderAfter: overlayHooks.cellRenderAfter,
        rowTitleCellRenderAfter: overlayHooks.rowTitleCellRenderAfter,
        updated: () => handlersRef.current.onUpdated(),
        sheetActivate: () => handlersRef.current.onSheetActivate(),
        workbookCreateAfter: () => handlersRef.current.onCreated(),
      },
    });
    window.setTimeout(() => handlersRef.current.onCreated(), 300);
  }, []);

  useEffect(() => {
    // Script Luckysheet nạp beforeInteractive nhưng có thể chưa chạy xong khi hydrate -> chờ tối đa 10s
    let tries = 0;
    const timer = window.setInterval(() => {
      tries++;
      if (ls() && typeof ls().create === 'function') {
        window.clearInterval(timer);
        createSheet();
      } else if (tries > 100) {
        window.clearInterval(timer);
        setStatusMsg('Không tải được thư viện bảng tính (Luckysheet). Kiểm tra kết nối mạng rồi tải lại trang.');
      }
    }, 100);
    return () => window.clearInterval(timer);
  }, [createSheet]);

  // Cảnh báo khi rời trang mà chưa tải xuống
  useEffect(() => {
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (!stateRef.current.dirty) return;
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, []);

  // Bảng tính cần tính lại kích thước khi mở/đóng/kéo giãn task pane
  useEffect(() => {
    if (!ready) return;
    const id = requestAnimationFrame(() => {
      try {
        ls().resize();
      } catch {
        /* bỏ qua */
      }
    });
    return () => cancelAnimationFrame(id);
  }, [pane, paneWidth, ready]);

  const markChanged = useCallback(() => {
    setDirty(true);
    setSheetVersion((v) => v + 1);
  }, []);

  // Sau khi xoá hàng làm dời/xoá ảnh: nạp lại workbook để lớp ảnh khớp với dữ liệu, giữ vị trí cuộn gần nhất
  const reloadWorkbook = useCallback(
    (focusRow?: number) => {
      const sheets = workbookForReload();
      createSheet(sheets, stateRef.current.fileName || 'Excel Tool');
      if (focusRow !== undefined) {
        window.setTimeout(() => {
          try {
            ls().scroll({ targetRow: Math.max(0, focusRow - 2) });
          } catch {
            /* bỏ qua */
          }
        }, 400);
      }
    },
    [createSheet],
  );

  // ==========================================
  // HOÀN TÁC (chụp toàn bộ workbook trước các thao tác xoá hàng loạt)
  // ==========================================
  const pushUndo = useCallback((label: string) => {
    try {
      const sheets = snapshotWorkbook();
      const title = ls().getSheet?.()?.name || 'Excel Tool';
      setUndoStack((s) => [...s.slice(-4), { label, sheets, activeOrder: getActiveOrder(), title }]);
    } catch (err) {
      console.error('Không chụp được trạng thái để hoàn tác', err);
    }
  }, []);

  const undoStackRef = useRef(undoStack);
  useLayoutEffect(() => {
    undoStackRef.current = undoStack;
  });

  const undo = useCallback(() => {
    const stack = undoStackRef.current;
    const last = stack[stack.length - 1];
    if (!last) return;
    const sheets = last.sheets.map((s, i) => ({ ...s, status: i === last.activeOrder ? 1 : 0 }));
    undoStackRef.current = stack.slice(0, -1);
    setUndoStack(undoStackRef.current);
    clearOverlay();
    createSheet(sheets, stateRef.current.fileName || last.title);
    setTemplate(null);
    setPivotSource(null);
    setSheetVersion((v) => v + 1);
    setDirty(true);
    toast('info', `Đã hoàn tác: ${last.label}`);
  }, [createSheet, toast]);

  // ==========================================
  // MỞ FILE
  // ==========================================
  const resetWork = () => {
    setTemplate(null);
    setPlan(null);
    setPivotScope(null);
    layoutKeyRef.current = null;
    setPivotSource(null);
    setPivotError(null);
    setUndoStack([]);
    clearOverlay();
  };

  const openFile = useCallback(
    (file: File, opts: { skipConfirm?: boolean; dirty?: boolean; onLoaded?: () => void } = {}) => {
      if (/\.pdf$/i.test(file.name)) {
        routePdfRef.current(file);
        return;
      }
      if (!/\.xlsx$/i.test(file.name)) {
        toast('error', 'Chỉ hỗ trợ file Excel .xlsx hoặc PDF');
        return;
      }
      if (!opts.skipConfirm && stateRef.current.dirty && !window.confirm('Dữ liệu hiện tại chưa được tải xuống. Mở file mới sẽ thay thế toàn bộ. Tiếp tục?')) return;
      const lx = luckyExcel();
      if (!lx) {
        toast('error', 'Thư viện đọc Excel chưa tải xong, vui lòng thử lại sau vài giây.');
        return;
      }
      setLoading(`Đang mở ${file.name}...`);
      // Workbook gốc chỉ dùng để giữ lại cài đặt trang in khi xuất file
      file
        .arrayBuffer()
        .then(async (buf) => {
          const wb = new ExcelJS.Workbook();
          await wb.xlsx.load(buf);
          setOriginal(wb);
        })
        .catch(() => setOriginal(null));

      window.setTimeout(() => {
        try {
          lx.transformExcelToLucky(
            file,
            (exportJson: LuckyValue) => {
              if (!exportJson?.sheets?.length) {
                setLoading('');
                toast('error', 'Không đọc được sheet nào trong file này.');
                return;
              }
              resetWork();
              createSheet(exportJson.sheets, file.name);
              setHasFile(true);
              setFileName(file.name);
              setDirty(!!opts.dirty);
              setZoom(1);
              setLoading('');
              setStatusMsg(`Đã mở ${file.name} · ${exportJson.sheets.length} sheet`);
              opts.onLoaded?.();
            },
            (err: LuckyValue) => {
              setLoading('');
              toast('error', `Lỗi đọc file: ${(err as Error)?.message || err}`);
            },
          );
        } catch (err) {
          setLoading('');
          toast('error', `Lỗi đọc file: ${(err as Error)?.message || err}`);
        }
      }, 50);
    },
    [createSheet, toast],
  );

  // ==========================================
  // ĐỌC PDF COMMERCIAL INVOICE VẢI -> BẢNG "TRÍCH INV VẢI" (mở luôn trên bảng tính)
  // ==========================================
  const openPdf = useCallback(
    async (file: File) => {
      if (stateRef.current.dirty && !window.confirm('Dữ liệu hiện tại chưa được tải xuống. Đọc PDF mới sẽ thay thế toàn bộ. Tiếp tục?')) return;
      setLoading(`Đang đọc ${file.name}...`);
      const fail = (error: string, rawLines: string[] = []) => {
        setLoading('');
        setPdfState({ fileName: file.name, invoice: null, error, rawLines });
        setPane('pdf');
      };
      try {
        const { items, pageCount, imageOnlyPages } = await readPdfText(file);
        if (imageOnlyPages === pageCount) {
          return fail('PDF này là ảnh scan (không có lớp chữ) nên chưa đọc được. Hãy dùng file PDF xuất trực tiếp từ hệ thống (bôi đen được chữ).');
        }
        const lines = groupLines(items);
        const invoice = parseFabricInvoice(lines, pageCount);
        const rawLines = lines.map((l) => `[trang ${l.page}] ${l.text}`);
        if (invoice.rows.length === 0) {
          return fail(
            'Không nhận ra bảng hàng của Commercial Invoice vải (cần có cột DESCRIPTION OF GOODS / QUANTITY / PRICE / AMOUNT và các mã hàng "#..."). Xem chữ đọc được bên dưới.',
            rawLines,
          );
        }
        const buffer = await buildFabricWorkbook(invoice).xlsx.writeBuffer();
        const name = `TRICH_INV_VAI_${invoice.invoiceNo || file.name.replace(/\.pdf$/i, '')}.xlsx`;
        const xlsx = new File([buffer], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        setLoading('');
        openFile(xlsx, {
          skipConfirm: true,
          dirty: true,
          onLoaded: () => {
            setPdfState({ fileName: file.name, invoice, error: null, rawLines });
            setPane('pdf');
            // LuckyExcel không nạp "cố định dòng" -> cố định lại dòng tiêu đề cho sheet Data;
            // Excel tự giãn chiều cao hàng có xuống dòng, Luckysheet thì không -> giãn tay các dòng hàng
            window.setTimeout(() => {
              getFiles().forEach((_, order) => {
                try {
                  ls().setHorizontalFrozen(false, { order });
                } catch {
                  /* bỏ qua */
                }
              });
              for (let r = 0; r <= invoice.rows.length; r++) autoFitRowHeight(r);
            }, 400);
          },
        });
      } catch (err) {
        fail(`Lỗi đọc PDF: ${(err as Error)?.message || err}`);
      }
    },
    [openFile],
  );
  useLayoutEffect(() => {
    openPdfRef.current = openPdf;
  });

  // ==========================================
  // TÁCH BỘ CHỨNG TỪ PDF (Tờ khai, Danh sách hàng hoá, Sales Contract, Invoice, Packing List, VAT...)
  // ==========================================
  const { load: loadSplit } = split;
  const openSplit = useCallback(
    async (file: File) => {
      const st = await loadSplit(file);
      if (!st) return;
      clearOverlay();
      setPane('split');
    },
    [loadSplit],
  );

  // Kéo thả PDF: chỉ là Commercial Invoice (+ Packing List) -> trích Excel; bộ nhiều chứng từ / có trang scan -> Tách PDF
  const routePdf = useCallback(
    async (file: File) => {
      const st = await loadSplit(file);
      if (!st) return;
      if (isInvoiceOnly(assignPages(st.pages).assigns) && !st.pages.some((p) => p.isScan)) {
        openPdf(file);
        return;
      }
      clearOverlay();
      setPane('split');
    },
    [loadSplit, openPdf],
  );
  useLayoutEffect(() => {
    routePdfRef.current = routePdf;
  });

  // ==========================================
  // XUẤT FILE (thấy gì tải nấy)
  // ==========================================
  const exportFile = useCallback(async () => {
    if (!stateRef.current.hasFile) {
      toast('info', 'Hãy mở một file Excel trước.');
      return;
    }
    setLoading('Đang đóng gói file Excel...');
    try {
      try {
        ls().exitEditMode();
      } catch {
        /* không ở chế độ sửa ô */
      }
      const sheets = ls().getAllSheets();
      const buffer = await exportWorkbookBuffer(sheets, { original });
      const base = (stateRef.current.fileName || 'Du_lieu').replace(/\.xlsx$/i, '');
      // File trích từ PDF đã được đặt tên theo số invoice -> giữ nguyên tên; hoá đơn VAT đã cắt gộp / Pivot -> VAT.xlsx
      const outName = /^TRICH_INV/i.test(base) ? `${base}.xlsx` : 'VAT.xlsx';
      saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), outName);
      setDirty(false);
      toast('success', `Đã tải xuống ${outName} (${sheets.length} sheet)`);
    } catch (err) {
      console.error(err);
      toast('error', `Lỗi khi xuất file: ${(err as Error)?.message || err}`);
    } finally {
      setLoading('');
    }
  }, [original, toast]);

  // ==========================================
  // CẮT GỘP: lập kế hoạch trực tiếp (xem trước) mỗi khi đổi tuỳ chọn / vùng mẫu / dữ liệu
  // ==========================================
  const buildTemplateFromSelection = useCallback((range: LuckyValue, forStrategy: TrimStrategy): TemplateInfo | null => {
    const rg = range?.[0];
    if (!rg) return null;
    const [r1, r2] = rg.row;
    const [c1, c2] = rg.column;
    const data = getSheetData();
    const mode = forStrategy === 'rpac' ? 'RPAC' : 'NORMAL';
    const pattern: string[] = [];
    const preview: string[] = [];
    for (let r = r1; r <= Math.min(r2, r1 + 40); r++) {
      pattern.push(getCleanedRangeText(data[r], c1, c2, mode));
      preview.push(getRangeText(data[r], c1, c2, mode).replace(/(\s*\|\s*)+$/g, '').replace(/^(\s*\|\s*)+/g, ''));
    }
    if (pattern.every((t) => t.replace(/\|/g, '').trim() === '')) return null;
    return { r1, r2: Math.min(r2, r1 + 40), c1, c2, rangeName: rangeName(r1, c1, r2, c2), pattern, preview };
  }, []);

  useEffect(() => {
    if (pane !== 'trim' || !ready || !hasFile) {
      if (pane !== 'pivot') clearOverlay();
      return;
    }
    const timer = window.setTimeout(() => {
      const data = getSheetData();
      let next: TrimPlan | null = null;
      if (strategy === 'auto') next = planAutoTrim(data, autoOpts);
      else if (template) {
        const hasBorder = buildRowBorderChecker(template.c1, template.c2);
        next =
          strategy === 'rpac'
            ? planRpacTrim(data, template, trimMode, hasBorder)
            : planTemplateTrim(data, template, trimMode, 'NORMAL', hasBorder);
      }
      setPlan(next);
      const keep: number[] = [];
      if (next?.keep) for (let r = next.keep.start; r <= next.keep.end; r++) keep.push(r);
      setOverlay({ deleteRows: next?.rows || [], keepRows: keep, focusRows: [] });
    }, 120);
    return () => window.clearTimeout(timer);
  }, [pane, strategy, autoOpts, trimMode, template, sheetVersion, ready, hasFile]);

  const applyTrim = useCallback(() => {
    if (!plan || plan.rows.length === 0) return;
    const n = plan.rows.length;
    pushUndo(`Cắt gộp ${n} hàng`);
    clearOverlay();
    setLoading('Đang cắt gộp dữ liệu...');
    window.setTimeout(() => {
      try {
        if (deleteRowsWithImages(plan.rows)) reloadWorkbook(plan.rows[0]);
        setTemplate(null);
        markChanged();
        setStatusMsg(`Đã cắt gộp: xoá và dồn ${n} hàng`);
        toast('success', `Đã xoá và dồn ${n} hàng`, { label: 'Hoàn tác', run: undo });
      } catch (err) {
        toast('error', `Lỗi khi cắt gộp: ${(err as Error)?.message || err}`);
      } finally {
        setLoading('');
      }
    }, 30);
  }, [plan, pushUndo, markChanged, toast, undo, reloadWorkbook]);

  const removeBlankRows = useCallback(
    (bounds?: { r1: number; r2: number }) => {
      if (!stateRef.current.hasFile) return toast('info', 'Hãy mở một file Excel trước.');
      const p = planBlankRows(getSheetData(), bounds?.r1, bounds?.r2);
      if (p.rows.length === 0) return toast('info', bounds ? 'Vùng chọn không có hàng trống nào.' : 'Bảng tính đã sạch, không có hàng trống nào.');
      pushUndo(`Xoá ${p.rows.length} hàng trống`);
      if (deleteRowsWithImages(p.rows)) reloadWorkbook(p.rows[0]);
      markChanged();
      setStatusMsg(`Đã xoá ${p.rows.length} hàng trống`);
      toast('success', `Đã xoá ${p.rows.length} hàng trống`, { label: 'Hoàn tác', run: undo });
    },
    [pushUndo, markChanged, toast, undo, reloadWorkbook],
  );

  // ==========================================
  // PIVOT
  // ==========================================
  const scanPivot = useCallback((scope: { r1: number; r2: number } | null, mode?: PivotMode) => {
    const res = buildPivotSource(getSheetData(), scope ?? undefined, { mode: mode ?? stateRef.current.pivotMode });
    if ('error' in res) {
      setPivotSource(null);
      setPivotError(res.error);
      return;
    }
    setPivotError(null);
    const key = `${res.source.profile}|${res.source.fields.map((f) => `${f.id}:${f.label}`).join(',')}`;
    if (layoutKeyRef.current !== key) {
      layoutKeyRef.current = key;
      setPivotLayout(defaultLayout(res.source.profile, res.source.fields));
    }
    setPivotSource(res.source);
  }, []);

  const openPivot = useCallback(
    (fromSelection: boolean) => {
      if (!stateRef.current.hasFile) return toast('info', 'Hãy mở file hoá đơn Excel trước.');
      const sel = getSelection()[0];
      // Giống Insert PivotTable của Excel: có vùng chọn nhiều hàng thì dùng vùng chọn, không thì tự nhận diện cả sheet
      const scope = sel && (fromSelection || sel.row[1] - sel.row[0] >= 2) ? { r1: sel.row[0], r2: sel.row[1] } : null;
      setPivotScope(scope);

      setPivotSource(null);
      clearOverlay();
      scanPivot(scope);
      setPane('pivot');
    },
    [scanPivot, toast],
  );

  const pivotResult: PivotResult | null = useMemo(
    () => (pivotSource ? computePivot(pivotSource, pivotLayout) : null),
    [pivotSource, pivotLayout],
  );

  const pivotSheetName = (profile: InvoiceProfile) => {
    const base = profile === 'fabric' ? 'Pivot Polytex' : profile === 'thread' ? 'Pivot Chỉ may' : 'Pivot';
    const names = new Set(getFiles().map((f) => String(f.name)));
    if (!names.has(base)) return base;
    let i = 2;
    while (names.has(`${base} (${i})`)) i++;
    return `${base} (${i})`;
  };

  const exportPivot = useCallback(() => {
    if (!pivotResult || !pivotSource) return;
    const { columns, rows, grandTotal } = pivotResult;
    const celldata: LuckyValue[] = [];
    const widths: number[] = columns.map((c) => Math.max(70, Math.min(560, c.label.length * 7 + 20)));
    const put = (r: number, c: number, v: string | number | null, style: Record<string, LuckyValue>) => {
      const col = columns[c];
      const isNum = typeof v === 'number';
      const text = isNum ? formatByPattern(v, col?.format) : v === null || v === undefined ? '' : String(v);
      widths[c] = Math.max(widths[c], Math.min(560, text.length * 7 + 18));
      celldata.push({
        r,
        c,
        v: {
          v: isNum ? v : text,
          m: text,
          // Mã hàng dạng số (vd 1002592) vẫn giữ là chữ để không mất số 0 đầu / không bị cộng nhầm
          ct: isNum ? { fa: col?.format || 'General', t: 'n' } : { fa: '@', t: 's' },
          ff: 'Arial',
          fs: 10,
          vt: 0,
          ht: isNum ? 2 : 1,
          ...style,
        },
      });
    };
    // Tông tím pastel cho sheet Pivot: tiêu đề tím #E0B0FF, dòng xen kẽ tím rất nhạt, dòng tổng tím vừa
    columns.forEach((col, c) => put(0, c, col.label, { bl: 1, bg: '#e0b0ff', fc: '#3f1560', ht: 0, tb: 2 }));
    rows.forEach((row, ri) => row.cells.forEach((v, c) => put(ri + 1, c, v, ri % 2 === 1 ? { bg: '#f7edff' } : {})));
    const totalRow = rows.length + 1;
    grandTotal.forEach((v, c) => put(totalRow, c, v, { bl: 1, bg: '#ecd4ff', fc: '#3f1560' }));
    const lastCol = columns.length - 1;
    const sheetObject = {
      name: pivotSheetName(pivotSource.profile),
      celldata,
      config: {
        columnlen: Object.fromEntries(widths.map((w, i) => [i, w])),
        rowlen: { 0: 36 },
        borderInfo: [
          { rangeType: 'range', borderType: 'border-all', style: '1', color: '#d6a4f5', range: [{ row: [0, totalRow], column: [0, lastCol] }] },
          { rangeType: 'range', borderType: 'border-top', style: '7', color: '#a86ad6', range: [{ row: [totalRow, totalRow], column: [0, lastCol] }] },
        ],
      },
      frozen: { type: 'row' },
    };
    try {
      const order = getFiles().length;
      ls().setSheetAdd({ sheetObject, order });
      window.setTimeout(() => {
        try {
          ls().setSheetActive(order);
        } catch {
          /* sheet đã được kích hoạt */
        }
      }, 50);
      markChanged();
      clearOverlay();
      setPane(null);
      toast('success', `Đã xuất Pivot sang sheet "${sheetObject.name}" · ${rows.length} dòng. Bấm Tải xuống để lưu file.`);
    } catch (err) {
      toast('error', `Lỗi khi tạo sheet mới: ${(err as Error)?.message || err}`);
    }
  }, [pivotResult, pivotSource, markChanged, toast]);

  const copyPivot = useCallback(() => {
    if (!pivotResult) return;
    const line = (cells: (string | number | null)[]) =>
      cells.map((v, i) => (typeof v === 'number' ? formatByPattern(v, pivotResult.columns[i]?.format).replace(/,/g, '') : String(v ?? '').replace(/\t|\n/g, ' '))).join('\t');
    const text = [pivotResult.columns.map((c) => c.label).join('\t'), ...pivotResult.rows.map((r) => line(r.cells)), line(pivotResult.grandTotal)].join('\n');
    navigator.clipboard?.writeText(text).then(
      () => toast('success', 'Đã sao chép bảng Pivot — dán (Ctrl+V) vào Excel là giữ nguyên cột'),
      () => toast('error', 'Trình duyệt không cho phép sao chép'),
    );
  }, [pivotResult, toast]);

  const focusPivotRows = useCallback((rows: number[], jump: boolean) => {
    setOverlay({ focusRows: rows, deleteRows: [], keepRows: [] });
    if (jump) selectRows(rows, true);
    else {
      try {
        ls().scroll({ targetRow: Math.max(0, rows[0] - 2) });
      } catch {
        /* bỏ qua */
      }
    }
  }, []);

  // ==========================================
  // HOOK LUCKYSHEET -> REACT
  // ==========================================
  useLayoutEffect(() => {
    handlersRef.current = {
    onRangeSelect: (range: LuckyValue) => {
      refreshStats();
      const s = stateRef.current;
      if (s.pane === 'trim' && s.strategy !== 'auto') {
        const t = buildTemplateFromSelection(range, s.strategy);
        if (t) setTemplate(t);
      }
    },
    onUpdated: () => {
      if (isSelfRefreshing()) return;
      setDirty(true);
      setSheetVersion((v) => v + 1);
      refreshStats();
    },
    onSheetActivate: () => {
      clearOverlay();
      setTemplate(null);
      setStats(null);
      setSheetVersion((v) => v + 1);
      if (stateRef.current.pane === 'pivot') setPane(null);
      const file = getActiveFile();
      setZoom(Number(file?.zoomRatio) || 1);
    },
    onCreated: () => {
      setReady(true);
      refreshSheet();
    },
    };
  });

  // ==========================================
  // TƯƠNG TÁC CHUỘT / BÀN PHÍM KIỂU EXCEL
  // ==========================================
  const applyZoom = useCallback((z: number) => {
    const next = Math.min(4, Math.max(0.1, Math.round(z * 100) / 100));
    setZoom(next);
    try {
      ls().setSheetZoom(next);
    } catch {
      /* Luckysheet chưa sẵn sàng */
    }
  }, []);

  const runContextAction = useCallback(
    (action: ContextAction) => {
      const sel = getSelection()[0];
      if (!sel) return;
      switch (action) {
        case 'trim-template': {
          const s = stateRef.current.strategy === 'rpac' ? 'rpac' : 'template';
          setStrategy(s);
          setTemplate(buildTemplateFromSelection([sel], s));
          setPane('trim');
          break;
        }
        case 'pivot-selection':
          openPivot(true);
          break;
        case 'blank-rows-selection':
          removeBlankRows({ r1: sel.row[0], r2: sel.row[1] });
          break;
        case 'autofit-cols':
          autoFitColumns(sel.column[0], sel.column[1]);
          markChanged();
          break;
        case 'autofit-rows':
          for (let r = sel.row[0]; r <= Math.min(sel.row[1], sel.row[0] + 2000); r++) autoFitRowHeight(r);
          markChanged();
          break;
        case 'freeze-here':
          ls().setBothFrozen(true, { range: { row_focus: Math.max(0, sel.row[0] - 1), column_focus: Math.max(0, sel.column[0] - 1) } });
          toast('info', `Đã cố định ${sel.row[0]} dòng trên và ${sel.column[0]} cột trái`);
          break;
        case 'unfreeze':
          ls().cancelFrozen();
          break;
        case 'copy-sum': {
          const data = getSheetData();
          let sum = 0;
          for (let r = sel.row[0]; r <= Math.min(sel.row[1], data.length - 1); r++)
            for (let c = sel.column[0]; c <= sel.column[1]; c++) {
              const cell = data[r]?.[c];
              if (cell && typeof cell.v === 'number') sum += cell.v;
              else if (cell && cell.ct?.t === 'n') sum += getCellNumber(cell) || 0;
            }
          const text = String(Math.round(sum * 1e6) / 1e6);
          navigator.clipboard?.writeText(text).then(() => toast('success', `Đã sao chép tổng: ${Number(text).toLocaleString('en-US')}`));
          break;
        }
      }
    },
    [buildTemplateFromSelection, openPivot, removeBlankRows, markChanged, toast],
  );

  const interactionRef = useRef<InteractionHandlers>(null as unknown as InteractionHandlers);
  useLayoutEffect(() => {
    interactionRef.current = {
    getZoom: () => stateRef.current.zoom,
    onZoom: applyZoom,
    onOpenFile: () => fileInputRef.current?.click(),
    onExport: exportFile,
    onHelp: () => setShowHelp(true),
    onEscape: () => {
      if (showHelp) setShowHelp(false);
      else if (stateRef.current.pane) {
        setPane(null);
        clearOverlay();
      }
    },
    onDropFile: openFile,
    onDragState: setDragOver,
    onContextAction: runContextAction,
    onSelectionChange: refreshStats,
    beforeStructureChange: pushUndo,
    onStructureChange: (label, reload) => {
      if (reload) reloadWorkbook();
      markChanged();
      setStatusMsg(label);
    },
    notify: (text) => toast('info', text),
    };
  }, [applyZoom, exportFile, openFile, runContextAction, refreshStats, pushUndo, showHelp, markChanged, toast, reloadWorkbook]);

  useEffect(() => installSheetInteractions(interactionRef), []);

  // Kéo giãn task pane bằng chuột (double-click mép để về mặc định)
  const startPaneResize = (e: React.MouseEvent) => {
    e.preventDefault();
    const startX = e.clientX;
    const startW = paneWidth;
    let latest = startW;
    const move = (ev: MouseEvent) => {
      latest = Math.min(900, Math.max(320, startW + (startX - ev.clientX)));
      setPaneWidth(latest);
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      document.body.classList.remove('is-resizing');
      try {
        localStorage.setItem(PANE_WIDTH_KEY, String(latest));
      } catch {
        /* bỏ qua */
      }
    };
    document.body.classList.add('is-resizing');
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  };

  const closePane = () => {
    setPane(null);
    clearOverlay();
  };

  const lastUndo = undoStack[undoStack.length - 1];

  return (
    <div className="app-shell">
      <input
        ref={fileInputRef}
        type="file"
        id="excel-upload-input"
        accept=".xlsx"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) openFile(file);
          e.target.value = '';
        }}
      />
      <input
        ref={pdfInputRef}
        type="file"
        id="pdf-upload-input"
        accept=".pdf,application/pdf"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) openPdf(file);
          e.target.value = '';
        }}
      />
      <input
        ref={splitInputRef}
        type="file"
        id="split-upload-input"
        accept=".pdf,application/pdf"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) openSplit(file);
          e.target.value = '';
        }}
      />

      {/* 1. Thanh tiêu đề */}
      <header className="titlebar">
        <div className="brand">
          <span className="brand-logo">
            <Icon name="table" size={16} strokeWidth={2} />
          </span>
          <span className="brand-name">Excel Tool</span>
          <span className="brand-tag">Hoá đơn VAT</span>
        </div>
        <div className="titlebar-file" title={fileName}>
          {hasFile ? (
            <>
              <Icon name="file" size={14} />
              <span className="truncate">{fileName}</span>
              {dirty ? <span className="dirty-pill">Chưa tải xuống</span> : <span className="saved-pill">Đã lưu</span>}
            </>
          ) : (
            <span className="opacity-70">Chưa mở file</span>
          )}
        </div>
        {/* Cân đối với khối brand để tên file nằm giữa */}
        <div className="titlebar-actions" aria-hidden="true" />
      </header>

      {/* 2. Ribbon */}
      <nav className="ribbon" aria-label="Công cụ">
        <div className="ribbon-group">
          <div className="ribbon-buttons">
            <button className="ribbon-btn" onClick={() => fileInputRef.current?.click()} title="Mở file Excel (Ctrl+O) — hoặc kéo thả file vào cửa sổ">
              <Icon name="upload" size={26} />
              <span>Mở file</span>
            </button>
            <button className="ribbon-btn is-accent" onClick={exportFile} disabled={!hasFile} title="Tải xuống file đã xử lý (Ctrl+S)">
              <Icon name="download" size={26} />
              <span>Tải xuống</span>
            </button>
          </div>
          <div className="ribbon-caption">Tệp</div>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-buttons">
            <button
              className={`ribbon-btn ${pane === 'pdf' ? 'is-on' : ''}`}
              onClick={() => pdfInputRef.current?.click()}
              title="Đọc Commercial Invoice vải (PDF) và trích ra bảng Excel: STT | Item | Thành phần & khổ | SL | Đơn giá | Trị giá"
            >
              <Icon name="file" size={26} />
              <span>Đọc PDF</span>
            </button>
            <button
              className={`ribbon-btn ${pane === 'split' ? 'is-on' : ''}`}
              onClick={() => {
                if (pane === 'split') closePane();
                else if (split.state) {
                  clearOverlay();
                  setPane('split');
                } else splitInputRef.current?.click();
              }}
              title="Tách bộ chứng từ PDF theo tiêu đề trang — xuất TKX, SC, ANNEX, INV, PKL, VAT, Other; trang scan gom vào SCAN.pdf"
            >
              <Icon name="layers" size={26} />
              <span>Tách PDF</span>
            </button>
          </div>
          <div className="ribbon-caption">PDF</div>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-buttons">
            <button
              className={`ribbon-btn ${pane === 'trim' ? 'is-on' : ''}`}
              onClick={() => (pane === 'trim' ? closePane() : setPane('trim'))}
              disabled={!hasFile}
              title="Xoá header/footer lặp lại khi gộp nhiều trang hoá đơn"
            >
              <Icon name="scissors" size={26} />
              <span>Cắt gộp</span>
            </button>
            <button className="ribbon-btn" onClick={undo} disabled={!lastUndo} title={lastUndo ? `Hoàn tác: ${lastUndo.label}` : 'Chưa có thao tác để hoàn tác'}>
              <Icon name="undo" size={26} />
              <span>Hoàn tác</span>
            </button>
          </div>
          <div className="ribbon-caption">Làm sạch dữ liệu</div>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-buttons">
            <button
              className={`ribbon-btn ${pane === 'pivot' ? 'is-on' : ''}`}
              onClick={() => (pane === 'pivot' ? closePane() : openPivot(false))}
              disabled={!hasFile}
              title="Tổng hợp như PivotTable (tự nhận diện hoá đơn Polytex, chỉ may, hoá đơn chung; không nhận ra thì Pivot thủ công theo dòng tiêu đề)"
            >
              <Icon name="pivot" size={26} />
              <span>Pivot</span>
            </button>
          </div>
          <div className="ribbon-caption">Phân tích</div>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-buttons">
            <button className="ribbon-btn" onClick={() => setShowHelp(true)} title="Thao tác chuột & phím tắt (F1)">
              <Icon name="keyboard" size={26} />
              <span>Phím tắt</span>
            </button>
          </div>
          <div className="ribbon-caption">Trợ giúp</div>
        </div>
      </nav>

      {/* 3. Vùng làm việc: bảng tính + task pane */}
      <main className="workspace">
        <div className={`sheet-area ${pane === 'split' && split.state ? 'is-split' : ''}`}>
          <div id="luckysheet-container" className="sheet-container" />
          {pane === 'split' && split.state && (
            <SplitBoard
              fileName={split.state.fileName}
              pages={split.state.pages}
              assigns={split.assigns}
              autoAssigns={split.autoAssigns}
              outputs={split.outputs}
              overrides={split.state.overrides}
              pageImage={split.pageImage}
              onOverride={split.setOverride}
            />
          )}
          {ready && !hasFile && (
            <div className="welcome">
              <div className="welcome-card">
                <button className="dropzone" onClick={() => fileInputRef.current?.click()}>
                  <span className="dropzone-icon">
                    <Icon name="upload" size={26} />
                  </span>
                  <span className="dropzone-title">Kéo thả file hoá đơn Excel vào đây</span>
                  <span className="dropzone-sub">hoặc bấm để chọn file .xlsx (Ctrl + O) · PDF invoice vải được trích ra Excel · PDF bộ chứng từ được tách theo loại</span>
                </button>
                <div className="feature-grid">
                  <div className="feature">
                    <span className="feature-icon bg-rose">
                      <Icon name="scissors" size={16} />
                    </span>
                    <b>Cắt gộp tự động</b>
                    <span>Xoá header, số trang lặp lại khi chuyển PDF sang Excel — xem trước trước khi xoá.</span>
                  </div>
                  <div className="feature">
                    <span className="feature-icon bg-emerald">
                      <Icon name="pivot" size={16} />
                    </span>
                    <b>Pivot</b>
                    <span>Polytex (vải), chỉ may A&E, hoá đơn chung hoặc bảng bất kỳ (Pivot thủ công) — tự đối chiếu với dòng Tổng cộng.</span>
                  </div>
                  <div className="feature">
                    <span className="feature-icon bg-sky">
                      <Icon name="mouse" size={16} />
                    </span>
                    <b>Thao tác như Excel</b>
                    <span>Chuột phải, kéo fill, AutoFit, Ctrl + lăn chuột, Alt + =, Ctrl + Space...</span>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>

        {pane && (
          <div className="pane-wrap" style={{ width: paneWidth }}>
            <div
              className="pane-resizer"
              onMouseDown={startPaneResize}
              onDoubleClick={() => {
                setPaneWidth(DEFAULT_PANE_WIDTH);
                try {
                  localStorage.setItem(PANE_WIDTH_KEY, String(DEFAULT_PANE_WIDTH));
                } catch {
                  /* bỏ qua */
                }
              }}
              title="Kéo để đổi độ rộng · nhấp đúp để về mặc định"
            />
            {pane === 'trim' ? (
              <TrimPane
                strategy={strategy}
                onStrategy={(s) => {
                  setStrategy(s);
                  if (s === 'auto') return;
                  // Đã bôi chọn sẵn một vùng (nhiều hơn 1 ô) trước khi chuyển chế độ -> dùng luôn làm mẫu
                  const sel = getSelection();
                  const rg = sel[0];
                  const isRange = rg && (rg.row[1] > rg.row[0] || rg.column[1] > rg.column[0]);
                  setTemplate(isRange ? buildTemplateFromSelection(sel, s) : template && buildTemplateFromSelection([{ row: [template.r1, template.r2], column: [template.c1, template.c2] }], s));
                }}
                autoOpts={autoOpts}
                onAutoOpts={setAutoOpts}
                trimMode={trimMode}
                onTrimMode={setTrimMode}
                template={template}
                plan={pane === 'trim' ? plan : null}
                hasFile={hasFile}
                onFocusBlock={(b: TrimBlock) => {
                  const data = getSheetData();
                  selectRange({ row: [b.start, b.end], column: [0, Math.max(0, (data[0]?.length || 1) - 1)] });
                }}
                onApply={applyTrim}
                onClose={closePane}
                undoLabel={lastUndo ? lastUndo.label : null}
                onUndo={undo}
              />
            ) : pane === 'split' ? (
              <SplitPane
                split={split}
                onPickFile={() => splitInputRef.current?.click()}
                onExtractInvoice={async (o) => {
                  const file = await split.outputFile(o);
                  if (file) openPdf(file);
                }}
                onClose={closePane}
              />
            ) : pane === 'pdf' ? (
              <PdfPane
                fileName={pdfState?.fileName ?? ''}
                invoice={pdfState?.invoice ?? null}
                error={pdfState?.error ?? null}
                rawLines={pdfState?.rawLines ?? []}
                onDownload={exportFile}
                onClose={closePane}
                onFocusRow={(i) => {
                  // Dòng hàng thứ i nằm ở hàng i+2 của sheet Data (hàng 1 là tiêu đề)
                  if (getActiveOrder() !== 0) ls().setSheetActive(0);
                  window.setTimeout(() => selectRange({ row: [i + 1, i + 1], column: [0, 5] }), 60);
                }}
              />
            ) : (
              <PivotPane
                source={pivotSource}
                error={pivotError}
                layout={pivotLayout}
                onLayout={setPivotLayout}
                result={pivotResult}
                scopeLabel={pivotScope ? `hàng ${pivotScope.r1 + 1}–${pivotScope.r2 + 1}` : `tự nhận diện cả sheet "${getActiveFile()?.name ?? ''}"`}
                selectionLabel={
                  selRows && selRows.r2 - selRows.r1 >= 1 && !(pivotScope && pivotScope.r1 === selRows.r1 && pivotScope.r2 === selRows.r2) ? selRows.label : null
                }
                onUseSelection={() => {
                  if (!selRows) return;
                  const scope = { r1: selRows.r1, r2: selRows.r2 };
                  setPivotScope(scope);
                  clearOverlay();
                  scanPivot(scope);
                }}
                onUseWholeSheet={
                  pivotScope
                    ? () => {
                        setPivotScope(null);
                        clearOverlay();
                        scanPivot(null);
                      }
                    : null
                }
                onOpenTrim={() => {
                  clearOverlay();
                  setStrategy('auto');
                  setPane('trim');
                }}
                onRescan={() => scanPivot(pivotScope)}
                mode={pivotMode}
                onModeChange={(m) => {
                  setPivotMode(m);
                  clearOverlay();
                  scanPivot(pivotScope, m);
                }}
                onClose={closePane}
                onExport={exportPivot}
                onCopy={copyPivot}
                onFocusRows={focusPivotRows}
              />
            )}
          </div>
        )}
      </main>

      {/* 4. Thanh trạng thái */}
      <StatusBar stats={stats} zoom={zoom} onZoom={applyZoom} message={statusMsg} ready={ready} onCopied={(t) => toast('success', t)} />

      {dragOver && (
        <div className="drop-overlay">
          <div className="drop-overlay-card">
            <Icon name="upload" size={34} />
            <b>Thả file để mở</b>
            <span>Chỉ hỗ trợ .xlsx</span>
          </div>
        </div>
      )}

      {showHelp && <ShortcutsDialog onClose={() => setShowHelp(false)} />}

      {loading && (
        <div className="loading-overlay">
          <div className="spinner" />
          <div className="loading-text">{loading}</div>
        </div>
      )}

      <ToastStack toasts={toasts} dismiss={dismiss} />
    </div>
  );
}
