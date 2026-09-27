'use client';

// Task pane "Tách PDF": danh sách file PDF sẽ xuất (theo loại chứng từ / từng bộ), OCR trang scan, tải từng file hoặc cả gói .zip
import Icon from './Icon';
import { SplitController } from '../lib/useSplitPdf';
import { SplitOutput, pageRanges } from '../lib/pdfSplit';

export default function SplitPane({
  split,
  onPickFile,
  onExtractInvoice,
  onClose,
}: {
  split: SplitController;
  onPickFile: () => void;
  onExtractInvoice: (o: SplitOutput) => void;
  onClose: () => void;
}) {
  const { state, outputs, assigns, ocr, busy } = split;

  if (!state) {
    return (
      <aside className="task-pane">
        <div className="pane-header">
          <div className="pane-title">
            <span className="pane-icon bg-amber">
              <Icon name="layers" size={16} />
            </span>
            <div>
              <div className="pane-title-text">Tách bộ chứng từ PDF</div>
              <div className="pane-subtitle">Chưa chọn file</div>
            </div>
          </div>
          <button className="icon-btn" onClick={onClose} title="Đóng (Esc)">
            <Icon name="close" size={16} />
          </button>
        </div>
        <div className="pane-body">
          <button className="dropzone" onClick={onPickFile}>
            <span className="dropzone-icon">
              <Icon name="upload" size={22} />
            </span>
            <span className="dropzone-title">Chọn file PDF bộ chứng từ</span>
            <span className="dropzone-sub">Tờ khai, Danh sách hàng hoá, Sales Contract, Commercial Invoice, Packing List, VAT...</span>
          </button>
        </div>
      </aside>
    );
  }

  const scanPages = state.pages.filter((p) => p.isScan);
  const ocrPending = scanPages.filter((p) => p.ocr === undefined).length;
  const manual = Object.keys(state.overrides).length;
  const ocrRunning = !!ocr && ocr.done < ocr.total;
  const docCount = new Set(assigns.filter((a) => a.segment >= 0).map((a) => a.segment)).size;

  return (
    <aside className="task-pane">
      <div className="pane-header">
        <div className="pane-title">
          <span className="pane-icon bg-amber">
            <Icon name="layers" size={16} />
          </span>
          <div>
            <div className="pane-title-text">Tách bộ chứng từ PDF</div>
            <div className="pane-subtitle truncate max-w-[240px]" title={state.fileName}>
              {state.fileName}
            </div>
          </div>
        </div>
        <button className="icon-btn" onClick={onClose} title="Đóng (Esc)">
          <Icon name="close" size={16} />
        </button>
      </div>

      <div className="pane-body">
        <div className="card">
          <div className="plan-summary mt-0">
            <div className="plan-number" style={{ color: '#b45309' }}>
              {outputs.length}
            </div>
            <div className="plan-caption">
              file PDF từ {state.pages.length} trang · {docCount} bộ chứng từ
            </div>
          </div>
          <div className="chips mt-2">
            <span className="chip tone-slate">{state.pages.length - scanPages.length} trang có chữ</span>
            {scanPages.length > 0 && <span className="chip tone-amber">{scanPages.length} trang scan</span>}
            {manual > 0 && (
              <button className="chip tone-sky" onClick={() => split.patch({ overrides: {} })} title="Bỏ các lựa chọn tay, về lại nhận diện tự động">
                {manual} trang chọn tay · khôi phục
              </button>
            )}
          </div>
        </div>

        <div className="segmented" style={{ gridTemplateColumns: 'repeat(2, 1fr)' }}>
          {(
            [
              ['type', 'Gộp theo loại', '1 file / loại chứng từ'],
              ['document', 'Tách từng bộ', '1 file / số chứng từ'],
            ] as const
          ).map(([key, label, sub]) => (
            <button key={key} className={`segmented-item ${state.mode === key ? 'is-active' : ''}`} onClick={() => split.patch({ mode: key })}>
              <span>{label}</span>
              <small>{sub}</small>
            </button>
          ))}
        </div>

        {scanPages.length > 0 && (
          <div className="card">
            <div className="field-label">Trang scan</div>
            <p className="card-note">
              <b>{scanPages.length}</b> trang là ảnh scan (không có lớp chữ) — mặc định gom vào <b>SCAN.pdf</b>. Bấm OCR để thử đọc tiêu đề: trang nào đọc được sẽ tự về đúng loại
              chứng từ.
            </p>
            {ocr && (
              <div className="split-progress">
                <div className="split-progress-bar">
                  <span style={{ width: `${Math.round((ocr.done / Math.max(1, ocr.total)) * 100)}%` }} />
                </div>
                <div className="split-progress-text">
                  {ocrRunning ? `${ocr.status} · ${ocr.done}/${ocr.total} trang` : `Đã OCR ${ocr.total} trang · nhận ra ${ocr.recognized} trang`}
                </div>
              </div>
            )}
            <div className="flex items-center gap-2 mt-2">
              <button className="xt-btn xt-btn-ghost" disabled={ocrRunning || ocrPending === 0} onClick={split.runOcr}>
                <Icon name="eye" size={15} /> {ocrPending === 0 ? 'Đã OCR xong' : ocrRunning ? 'Đang OCR...' : `OCR ${ocrPending} trang scan`}
              </button>
            </div>
            <label className="check-row mt-2">
              <input type="checkbox" checked={state.ocrToScan} onChange={(e) => split.patch({ ocrToScan: e.target.checked })} />
              <span className="text-[12px]">Luôn để trang scan trong SCAN.pdf (kể cả khi OCR đọc được)</span>
            </label>
          </div>
        )}

        <div className="card p-0 overflow-hidden">
          <div className="card-head">File PDF sẽ xuất</div>
          <div className="split-outputs">
            {outputs.map((o) => (
              <div key={o.key} className="split-output" style={{ ['--split-color' as string]: o.color }}>
                <div className="split-output-info">
                  <div className="split-output-name" title={o.fileName}>
                    {o.fileName}
                  </div>
                  <div className="split-output-sub">
                    {o.kinds.join(', ')} · {o.pages.length} trang · trang {pageRanges(o.pages)}
                    {o.docNos.length > 1 && ` · ${o.docNos.length} bộ: ${o.docNos.join(', ')}`}
                  </div>
                </div>
                <div className="split-output-actions">
                  {o.target === 'commercial-invoice' && (
                    <button className="icon-btn" onClick={() => onExtractInvoice(o)} title="Trích Commercial Invoice này ra bảng Excel (TRICH_INV_VAI)">
                      <Icon name="table" size={15} />
                    </button>
                  )}
                  <button className="icon-btn" onClick={() => split.preview(o)} title="Xem file PDF này (tab mới)">
                    <Icon name="eye" size={15} />
                  </button>
                  <button className="icon-btn" onClick={() => split.download(o)} title={`Tải ${o.fileName}`}>
                    <Icon name="download" size={15} />
                  </button>
                </div>
              </div>
            ))}
            {outputs.length === 0 && <div className="empty-hint">Không còn trang nào để xuất.</div>}
          </div>
        </div>

        <div className="note">
          <Icon name="info" size={14} className="shrink-0 mt-0.5" />
          <span>
            Mỗi trang được nhận diện theo <b>tiêu đề</b> (Tờ khai, Danh sách hàng hoá, Sales Contract, Commercial Invoice, Packing List, VAT...). Trang không có tiêu đề thuộc chứng từ ở
            trang trước. Nhận sai thì chọn lại loại ngay dưới ảnh trang.
          </span>
        </div>
      </div>

      <div className="pane-footer">
        <button className="xt-btn xt-btn-ghost" onClick={onPickFile} title="Mở file PDF khác">
          <Icon name="upload" size={15} /> File khác
        </button>
        <button className="xt-btn xt-btn-primary ml-auto" disabled={outputs.length === 0 || !!busy} onClick={split.downloadAll}>
          <Icon name="download" size={15} /> {busy ? 'Đang tạo...' : 'Tải tất cả (.zip)'}
        </button>
      </div>
    </aside>
  );
}
