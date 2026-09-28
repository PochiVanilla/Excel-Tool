'use client';

// Bảng xem trước "Tách PDF": ảnh từng trang xếp theo file PDF sẽ xuất ra
// - Dưới mỗi trang có ô chọn loại chứng từ để sửa tay khi nhận diện sai (hoặc đưa vào SCAN.pdf / bỏ trang)
// - Nhấp vào ảnh để xem to
import { useEffect, useRef, useState } from 'react';
import Icon from './Icon';
import { DOC_TYPES, PageAssign, PageTarget, SCAN_COLOR, SplitOutput, SplitPage, docTypeOf, pageRanges } from '../lib/pdfSplit';

const HOW_LABELS: Record<PageAssign['how'], string> = {
  title: 'Nhận diện theo tiêu đề',
  continue: 'Trang tiếp theo (không có tiêu đề)',
  ocr: 'Trang scan — OCR đọc được',
  manual: 'Đã chọn tay (các trang phía sau theo loại này)',
  follow: 'Theo trang chọn tay phía trước',
  scan: 'Trang scan (ảnh, không có lớp chữ)',
  skip: 'Bỏ ra, không xuất',
};

const targetLabel = (t: PageTarget) => (t === 'scan' ? 'SCAN' : t === 'skip' ? 'Bỏ trang' : `${docTypeOf(t).code} · ${docTypeOf(t).label}`);

function PageThumb({
  page,
  assign,
  autoTarget,
  override,
  pageImage,
  onOverride,
  onZoom,
}: {
  page: SplitPage;
  assign: PageAssign;
  autoTarget: PageTarget;
  override: PageTarget | undefined;
  pageImage: (index: number, width: number) => Promise<string>;
  onOverride: (target: PageTarget | null) => void;
  onZoom: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [src, setSrc] = useState<string | null>(null);

  // Chỉ vẽ ảnh khi trang cuộn tới (file nhiều trang vẫn mở nhanh)
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    let alive = true;
    const io = new IntersectionObserver(
      (entries) => {
        if (!entries.some((e) => e.isIntersecting)) return;
        io.disconnect();
        pageImage(page.index, 220)
          .then((url) => alive && setSrc(url))
          .catch(() => {});
      },
      { rootMargin: '400px' },
    );
    io.observe(el);
    return () => {
      alive = false;
      io.disconnect();
    };
  }, [page.index, pageImage]);

  const badge =
    assign.how === 'scan' ? 'Scan' : assign.how === 'ocr' ? 'Scan · OCR' : assign.how === 'manual' ? 'Chọn tay' : assign.how === 'follow' ? `Theo tr. ${(assign.from ?? 0) + 1}` : assign.how === 'skip' ? 'Bỏ' : null;

  return (
    <div className={`split-thumb ${assign.target === 'skip' ? 'is-skip' : ''}`} ref={ref}>
      <button
        className="split-thumb-img"
        style={{ aspectRatio: `${page.width} / ${page.height}` }}
        onClick={onZoom}
        title={`${HOW_LABELS[assign.how]}${assign.title ? `\nTiêu đề: ${assign.title}` : ''}\nNhấp để xem to`}
      >
        {/* eslint-disable-next-line @next/next/no-img-element -- ảnh dữ liệu (data URL) vẽ từ PDF */}
        {src ? <img src={src} alt={`Trang ${page.index + 1}`} /> : <span className="split-thumb-wait" />}
        {badge && <span className={`split-badge ${assign.how === 'scan' || assign.how === 'ocr' ? 'is-scan' : ''}`}>{badge}</span>}
        {assign.how === 'title' && <span className="split-badge is-title">Tiêu đề</span>}
      </button>
      <div className="split-thumb-foot">
        <span className="split-thumb-no">{page.index + 1}</span>
        <select
          className={`split-select ${override ? 'is-manual' : ''}`}
          value={override ?? ''}
          onChange={(e) => onOverride(e.target.value ? (e.target.value as PageTarget) : null)}
          title="Đổi loại chứng từ của trang này"
        >
          <option value="">
            {assign.how === 'follow' ? `${targetLabel(assign.target)} (theo trang ${(assign.from ?? 0) + 1})` : `${targetLabel(autoTarget)} (tự động)`}
          </option>
          {DOC_TYPES.map((d) => (
            <option key={d.id} value={d.id}>
              {d.code} · {d.label}
            </option>
          ))}
          <option value="scan">SCAN · trang scan</option>
          <option value="skip">Bỏ trang (không xuất)</option>
        </select>
      </div>
    </div>
  );
}

export default function SplitBoard({
  fileName,
  pages,
  assigns,
  autoAssigns,
  outputs,
  overrides,
  pageImage,
  onOverride,
}: {
  fileName: string;
  pages: SplitPage[];
  assigns: PageAssign[];
  // Kết quả nhận diện tự động (chưa tính các trang chọn tay) -> hiện trong ô chọn "Tự động: ..."
  autoAssigns: PageAssign[];
  outputs: SplitOutput[];
  overrides: Record<number, PageTarget>;
  pageImage: (index: number, width: number) => Promise<string>;
  onOverride: (index: number, target: PageTarget | null) => void;
}) {
  const [zoom, setZoom] = useState<{ index: number; src: string | null } | null>(null);
  const skipped = assigns.filter((a) => a.target === 'skip').map((a) => a.index);

  const openZoom = (index: number) => {
    setZoom({ index, src: null });
    pageImage(index, 1100)
      .then((src) => setZoom((z) => (z && z.index === index ? { index, src } : z)))
      .catch(() => {});
  };

  useEffect(() => {
    if (!zoom) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        setZoom(null);
      } else if (e.key === 'ArrowRight' && zoom.index < pages.length - 1) openZoom(zoom.index + 1);
      else if (e.key === 'ArrowLeft' && zoom.index > 0) openZoom(zoom.index - 1);
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  });

  const group = (key: string, title: string, sub: string, color: string, list: number[]) => (
    <section key={key} className="split-group" style={{ ['--split-color' as string]: color }}>
      <header className="split-group-head">
        <span className="split-dot" />
        <b className="truncate">{title}</b>
        <span className="split-group-sub">{sub}</span>
      </header>
      <div className="split-grid">
        {list.map((i) => (
          <PageThumb
            key={i}
            page={pages[i]}
            assign={assigns[i]}
            autoTarget={autoAssigns[i]?.target ?? assigns[i].target}
            override={overrides[i]}
            pageImage={pageImage}
            onOverride={(t) => onOverride(i, t)}
            onZoom={() => openZoom(i)}
          />
        ))}
      </div>
    </section>
  );

  const zoomAssign = zoom ? assigns[zoom.index] : null;

  return (
    <div className="split-board">
      <div className="split-board-head">
        <Icon name="layers" size={16} />
        <b className="truncate">{fileName}</b>
        <span>
          · {pages.length} trang → {outputs.length} file PDF
        </span>
        <span className="ml-auto text-slate-500">Sai loại? Chọn lại ở trang đầu của chứng từ, các trang sau tự theo</span>
      </div>
      {outputs.map((o) => group(o.key, o.fileName, `${o.kinds.join(', ')} · ${o.pages.length} trang · trang ${pageRanges(o.pages)}`, o.color, o.pages))}
      {skipped.length > 0 && group('skip', 'Trang bỏ ra (không xuất)', `${skipped.length} trang`, '#94a3b8', skipped)}

      {zoom && (
        <div className="split-zoom" onClick={() => setZoom(null)}>
          <div className="split-zoom-card" onClick={(e) => e.stopPropagation()}>
            <div className="split-zoom-head">
              <b>Trang {zoom.index + 1}</b>
              {zoomAssign && (
                <span className="chip" style={{ background: zoomAssign.target === 'scan' ? SCAN_COLOR : undefined }}>
                  {targetLabel(zoomAssign.target)}
                </span>
              )}
              <span className="truncate text-slate-500 text-xs">{zoomAssign?.title || (zoomAssign ? HOW_LABELS[zoomAssign.how] : '')}</span>
              <div className="ml-auto flex items-center gap-1">
                <button className="icon-btn" disabled={zoom.index === 0} onClick={() => openZoom(zoom.index - 1)} title="Trang trước (←)">
                  <Icon name="chevronRight" size={16} style={{ transform: 'rotate(180deg)' }} />
                </button>
                <button className="icon-btn" disabled={zoom.index >= pages.length - 1} onClick={() => openZoom(zoom.index + 1)} title="Trang sau (→)">
                  <Icon name="chevronRight" size={16} />
                </button>
                <button className="icon-btn" onClick={() => setZoom(null)} title="Đóng (Esc)">
                  <Icon name="close" size={16} />
                </button>
              </div>
            </div>
            <div className="split-zoom-body">
              {/* eslint-disable-next-line @next/next/no-img-element -- ảnh dữ liệu (data URL) vẽ từ PDF */}
              {zoom.src ? <img src={zoom.src} alt={`Trang ${zoom.index + 1}`} /> : <div className="spinner" />}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
