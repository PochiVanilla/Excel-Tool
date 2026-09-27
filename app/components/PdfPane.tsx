'use client';

// Kết quả đọc Commercial Invoice vải từ PDF -> bảng "Trích INV vải" (đã mở sẵn trên bảng tính)
import { useState } from 'react';
import Icon from './Icon';
import { FabricInvoice, invoiceSums } from '../lib/pdfInvoice';

const fmt = (n: number | null, d: number) => (n === null ? '—' : n.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d }));

export default function PdfPane({
  fileName,
  invoice,
  error,
  rawLines,
  onDownload,
  onClose,
  onFocusRow,
}: {
  fileName: string;
  invoice: FabricInvoice | null;
  error: string | null;
  rawLines: string[];
  onDownload: () => void;
  onClose: () => void;
  onFocusRow: (index: number) => void;
}) {
  const [showRaw, setShowRaw] = useState(false);
  const sums = invoice ? invoiceSums(invoice) : null;

  return (
    <aside className="task-pane">
      <div className="pane-header">
        <div className="pane-title">
          <span className="pane-icon bg-sky">
            <Icon name="file" size={16} />
          </span>
          <div>
            <div className="pane-title-text">Trích xuất từ PDF</div>
            <div className="pane-subtitle truncate max-w-[240px]" title={fileName}>
              {fileName}
            </div>
          </div>
        </div>
        <button className="icon-btn" onClick={onClose} title="Đóng (Esc)">
          <Icon name="close" size={16} />
        </button>
      </div>

      <div className="pane-body">
        {error && (
          <div className="note note-warn">
            <Icon name="alert" size={15} className="shrink-0 mt-0.5" />
            <span>{error}</span>
          </div>
        )}

        {invoice && sums && (
          <>
            <div className="card">
              <div className="field-label">Commercial Invoice</div>
              <dl className="pdf-meta">
                <dt>Số invoice</dt>
                <dd>{invoice.invoiceNo || '—'}</dd>
                <dt>Ngày</dt>
                <dd>{invoice.date || '—'}</dd>
                <dt>Người bán</dt>
                <dd title={invoice.seller}>{invoice.seller || '—'}</dd>
                <dt>Số trang</dt>
                <dd>{invoice.pageCount}</dd>
              </dl>
            </div>

            <div className="card">
              <div className="plan-summary mt-0">
                <div className="plan-number" style={{ color: '#0369a1' }}>{invoice.rows.length}</div>
                <div className="plan-caption">dòng màu / lô đã trích</div>
              </div>
              <div className="checks px-0">
                {(
                  [
                    ['Số lượng', sums.qty, invoice.total.qty, sums.qtyCheck, 2, invoice.unit],
                    ['Trị giá', sums.amount, invoice.total.amount, sums.amountCheck, 4, invoice.currency],
                  ] as const
                ).map(([label, mine, theirs, st, d, unit]) =>
                  st === 'na' ? (
                    <div key={label} className="check-badge is-diff">
                      <Icon name="alert" size={13} /> {label}: {fmt(mine, d)} {unit} — không có TOTAL trên invoice để đối chiếu
                    </div>
                  ) : (
                    <div key={label} className={`check-badge ${st === 'ok' ? 'is-ok' : 'is-diff'}`}>
                      <Icon name={st === 'ok' ? 'check' : 'alert'} size={13} />
                      {st === 'ok' ? (
                        <span>
                          {label} khớp TOTAL invoice: <b>{fmt(mine, d)}</b> {unit}
                        </span>
                      ) : (
                        <span>
                          {label} lệch: trích <b>{fmt(mine, d)}</b> ≠ TOTAL <b>{fmt(theirs, d)}</b> {unit}
                        </span>
                      )}
                    </div>
                  ),
                )}
              </div>
            </div>

            <div className="card">
              <div className="field-label">Thông tin đóng gói</div>
              <dl className="pdf-meta">
                <dt>Số cuộn (ROLLS)</dt>
                <dd>{fmt(invoice.packing.rolls, 0)}</dd>
                <dt>N.W (KGM)</dt>
                <dd>{fmt(invoice.packing.nw, 2)}</dd>
                <dt>G.W (KGM)</dt>
                <dd>{fmt(invoice.packing.gw, 2)}</dd>
                <dt>CBM</dt>
                <dd>{fmt(invoice.packing.cbm, 2)}</dd>
              </dl>
            </div>

            {invoice.warnings.length > 0 && (
              <div className="card">
                <div className="field-label">Cần kiểm tra lại</div>
                {invoice.warnings.map((w, i) => (
                  <div key={i} className="note mt-1">
                    <Icon name="info" size={14} className="shrink-0 mt-0.5" />
                    <span>{w}</span>
                  </div>
                ))}
              </div>
            )}

            {invoice.rows.length > 0 && (
              <div className="card p-0 overflow-hidden">
                <div className="card-head flex items-center">
                  <span>Các dòng đã trích</span>
                  <span className="ml-auto text-[10px] font-normal text-slate-400">nhấp để chọn trên bảng tính</span>
                </div>
                <div className="result-wrap mt-0 border-t-0">
                  <table className="result-table">
                    <thead>
                      <tr>
                        <th>#</th>
                        <th>Item</th>
                        <th>Màu / lô</th>
                        <th className="num">SL</th>
                        <th className="num">Đơn giá</th>
                      </tr>
                    </thead>
                    <tbody>
                      {invoice.rows.map((r, i) => (
                        <tr key={i} onClick={() => onFocusRow(i)} title={r.composition}>
                          <td>{i + 1}</td>
                          <td>{r.item}</td>
                          <td>{r.color}</td>
                          <td className="num">{fmt(r.qty, 2)}</td>
                          <td className="num">{fmt(r.price, 4)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        )}

        {rawLines.length > 0 && (
          <div className="card">
            <button className="source-skip mt-0" onClick={() => setShowRaw(!showRaw)}>
              <Icon name={showRaw ? 'chevronDown' : 'chevronRight'} size={12} />
              Xem chữ đọc được từ PDF ({rawLines.length} dòng)
            </button>
            {showRaw && <pre className="pdf-raw">{rawLines.join('\n')}</pre>}
          </div>
        )}
      </div>

      <div className="pane-footer">
        <span className="text-[11px] text-slate-500">Sửa trực tiếp trên bảng tính nếu cần</span>
        <button className="xt-btn xt-btn-primary ml-auto" disabled={!invoice || invoice.rows.length === 0} onClick={onDownload}>
          <Icon name="download" size={15} /> Tải Excel
        </button>
      </div>
    </aside>
  );
}
