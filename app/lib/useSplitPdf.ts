'use client';

// ==========================================
// TRẠNG THÁI "TÁCH PDF": đọc trang, nhận diện chứng từ, OCR trang scan, ảnh thu nhỏ, xuất file
// ==========================================
import { useCallback, useMemo, useRef, useState } from 'react';
import { saveAs } from 'file-saver';
import type { PDFDocument } from 'pdf-lib';
import type { ToastTone } from '../components/Toasts';
import { GroupMode, PageTarget, SplitOutput, SplitPage, assignPages, buildOutputs, detectTitle } from './pdfSplit';
import { PdfDoc, openPdfDocument, readSplitPage, renderPdfPage } from './pdfReader';
import { buildZip, extractPages, loadPdfSource, toBlob } from './pdfWriter';
import { ocrCanvas } from './ocr';

export interface SplitState {
  fileName: string;
  pages: SplitPage[];
  overrides: Record<number, PageTarget>;
  mode: GroupMode;
  ocrToScan: boolean;
}

export interface OcrProgress {
  done: number;
  total: number;
  status: string;
  recognized: number;
}

const baseName = (name: string) => name.replace(/\.pdf$/i, '');
// Vùng tiêu đề được OCR: 1/3 trên của trang
const OCR_ZONE = 0.34;

export const useSplitPdf = (notify: (tone: ToastTone, msg: string) => void, setLoading: (msg: string) => void) => {
  const [state, setState] = useState<SplitState | null>(null);
  const [ocr, setOcr] = useState<OcrProgress | null>(null);
  const [busy, setBusy] = useState('');
  const bytesRef = useRef<Uint8Array | null>(null);
  const docRef = useRef<PdfDoc | null>(null);
  const srcRef = useRef<Promise<PDFDocument> | null>(null);
  const thumbsRef = useRef(new Map<string, Promise<string>>());
  // Mỗi lần mở file mới tăng token -> huỷ các tác vụ (OCR, ảnh thu nhỏ) của file cũ
  const tokenRef = useRef(0);

  const release = useCallback(() => {
    tokenRef.current++;
    const doc = docRef.current;
    docRef.current = null;
    bytesRef.current = null;
    srcRef.current = null;
    thumbsRef.current.clear();
    if (doc) doc.destroy().catch(() => {});
  }, []);

  const load = useCallback(
    async (file: File): Promise<SplitState | null> => {
      setLoading(`Đang đọc ${file.name}...`);
      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        release();
        const token = tokenRef.current;
        const doc = await openPdfDocument(bytes);
        docRef.current = doc;
        bytesRef.current = bytes;
        const pages: SplitPage[] = [];
        for (let i = 0; i < doc.numPages; i++) {
          if (token !== tokenRef.current) return null;
          if (i % 5 === 0) setLoading(`Đang nhận diện trang ${i + 1}/${doc.numPages}...`);
          pages.push(await readSplitPage(doc, i));
        }
        const next: SplitState = { fileName: file.name, pages, overrides: {}, mode: 'type', ocrToScan: false };
        setState(next);
        setOcr(null);
        return next;
      } catch (err) {
        notify('error', `Không đọc được PDF: ${(err as Error)?.message || err}`);
        return null;
      } finally {
        setLoading('');
      }
    },
    [notify, setLoading, release],
  );

  const close = useCallback(() => {
    release();
    setState(null);
    setOcr(null);
  }, [release]);

  const { assigns, segments } = useMemo(
    () => (state ? assignPages(state.pages, state.overrides, { ocrToScan: state.ocrToScan }) : { assigns: [], segments: [] }),
    [state],
  );
  const autoAssigns = useMemo(
    () => (state ? assignPages(state.pages, {}, { ocrToScan: state.ocrToScan }).assigns : []),
    [state],
  );
  const outputs = useMemo(() => (state ? buildOutputs(assigns, segments, state.mode) : []), [state, assigns, segments]);

  const patch = useCallback((p: Partial<SplitState>) => setState((s) => (s ? { ...s, ...p } : s)), []);

  const setOverride = useCallback(
    (index: number, target: PageTarget | null) =>
      setState((s) => {
        if (!s) return s;
        const overrides = { ...s.overrides };
        if (target === null) delete overrides[index];
        else overrides[index] = target;
        return { ...s, overrides };
      }),
    [],
  );

  // Ảnh trang (JPEG) theo độ rộng, lưu đệm theo trang + độ rộng
  const pageImage = useCallback((index: number, width: number): Promise<string> => {
    const doc = docRef.current;
    if (!doc) return Promise.reject(new Error('closed'));
    const key = `${index}@${width}`;
    let p = thumbsRef.current.get(key);
    if (!p) {
      p = renderPdfPage(doc, index, width).then((c) => c.toDataURL('image/jpeg', width > 400 ? 0.88 : 0.72));
      p.catch(() => thumbsRef.current.delete(key));
      thumbsRef.current.set(key, p);
    }
    return p;
  }, []);

  // OCR các trang scan chưa đọc: nhận ra tiêu đề thì trang được đưa về đúng loại chứng từ
  const runOcr = useCallback(async () => {
    const doc = docRef.current;
    if (!doc || !state) return;
    const token = tokenRef.current;
    const todo = state.pages.filter((p) => p.isScan && p.ocr === undefined).map((p) => p.index);
    if (todo.length === 0) return;
    let recognized = 0;
    setOcr({ done: 0, total: todo.length, status: 'Đang chuẩn bị OCR', recognized });
    try {
      for (let k = 0; k < todo.length; k++) {
        const index = todo[k];
        const page = state.pages[index];
        // Chỉ đọc 1/3 trên của trang (vùng tiêu đề) ở ~300 dpi: nhanh và bắt tiêu đề chính xác hơn đọc cả trang
        const width = Math.round(Math.min(3000, (page.width / 72) * 300));
        const canvas = await renderPdfPage(doc, index, width, OCR_ZONE);
        if (token !== tokenRef.current) return;
        const fullHeight = (canvas.width / page.width) * page.height;
        const lines = await ocrCanvas(
          canvas,
          (status, p) => {
            if (token === tokenRef.current) setOcr({ done: k, total: todo.length, status: `${status}${p ? ` ${Math.round(p * 100)}%` : ''}`, recognized });
          },
          fullHeight,
        );
        if (token !== tokenRef.current) return;
        if (detectTitle(lines, true).type) recognized++;
        setState((s) => (s ? { ...s, pages: s.pages.map((p) => (p.index === index ? { ...p, ocr: lines.length ? lines : null } : p)) } : s));
        setOcr({ done: k + 1, total: todo.length, status: `Đã đọc trang ${index + 1}`, recognized });
      }
      notify(
        recognized ? 'success' : 'info',
        recognized
          ? `OCR xong ${todo.length} trang scan: nhận ra tiêu đề ${recognized} trang, đã đưa về đúng loại chứng từ.`
          : `OCR xong ${todo.length} trang scan nhưng không nhận ra tiêu đề nào — các trang này vẫn nằm trong SCAN.pdf.`,
      );
    } catch (err) {
      notify('error', `Lỗi OCR: ${(err as Error)?.message || err}`);
      setOcr(null);
    }
  }, [state, notify]);

  const bytesOf = useCallback(async (o: SplitOutput) => {
    const bytes = bytesRef.current;
    if (!bytes) throw new Error('Chưa mở file PDF');
    if (!srcRef.current) srcRef.current = loadPdfSource(bytes);
    return extractPages(await srcRef.current, o.pages, o.fileName.replace(/\.pdf$/i, ''));
  }, []);

  const run = useCallback(
    async <T,>(label: string, job: () => Promise<T>): Promise<T | null> => {
      setBusy(label);
      try {
        return await job();
      } catch (err) {
        notify('error', `Lỗi tạo PDF: ${(err as Error)?.message || err}`);
        return null;
      } finally {
        setBusy('');
      }
    },
    [notify],
  );

  const download = useCallback(
    (o: SplitOutput) =>
      run(`Đang tạo ${o.fileName}...`, async () => {
        saveAs(toBlob(await bytesOf(o)), o.fileName);
      }),
    [run, bytesOf],
  );

  const preview = useCallback(
    (o: SplitOutput) => {
      // Mở tab trước (trong cùng thao tác bấm chuột) để trình duyệt không chặn popup
      const win = window.open('', '_blank');
      run(`Đang tạo ${o.fileName}...`, async () => {
        const url = URL.createObjectURL(toBlob(await bytesOf(o)));
        if (win) win.location.href = url;
        else window.open(url, '_blank');
        window.setTimeout(() => URL.revokeObjectURL(url), 120000);
      }).then((r) => {
        if (r === null) win?.close();
      });
    },
    [run, bytesOf],
  );

  const downloadAll = useCallback(() => {
    if (!state || outputs.length === 0) return;
    const name = `${baseName(state.fileName)}_da_tach.zip`;
    run('Đang đóng gói .zip...', async () => {
      const files: { name: string; bytes: Uint8Array }[] = [];
      for (const o of outputs) files.push({ name: o.fileName, bytes: await bytesOf(o) });
      saveAs(await buildZip(files), name);
      notify('success', `Đã tải ${name} · ${outputs.length} file PDF`);
    });
  }, [state, outputs, run, bytesOf, notify]);

  // File PDF chỉ gồm các trang của 1 nhóm (vd Commercial Invoice -> trích Excel)
  const outputFile = useCallback(
    async (o: SplitOutput) => {
      const r = await run(`Đang tạo ${o.fileName}...`, () => bytesOf(o));
      return r ? new File([r as BlobPart], o.docNos[0] ? `${o.docNos[0]}.pdf` : o.fileName, { type: 'application/pdf' }) : null;
    },
    [run, bytesOf],
  );

  return { state, assigns, autoAssigns, segments, outputs, ocr, busy, load, close, patch, setOverride, pageImage, runOcr, download, preview, downloadAll, outputFile };
};

export type SplitController = ReturnType<typeof useSplitPdf>;
