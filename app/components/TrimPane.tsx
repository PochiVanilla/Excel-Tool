'use client';

// Bảng điều khiển CẮT GỘP (dạng task pane bên phải như Excel)
// Quy trình mới: chọn chế độ -> hệ thống quét & TÔ ĐỎ trước các hàng sẽ xoá -> kiểm tra -> Áp dụng -> có thể Hoàn tác
import Image from 'next/image';
import Icon from './Icon';
import { TRIM_REASON_LABELS, TrimBlock, TrimPlan, TrimReason, AutoTrimOptions } from '../lib/trimEngine';

export type TrimStrategy = 'auto' | 'template' | 'rpac';

export interface TemplateInfo {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
  rangeName: string;
  pattern: string[];
  preview: string[];
}

const AUTO_OPTIONS: { key: keyof AutoTrimOptions; label: string; hint: string }[] = [
  { key: 'repeatHeaders', label: 'Xoá header lặp lại mỗi trang', hint: 'Giữ header gốc đầu tiên, chỉ xoá khối thông tin công ty + tiêu đề bảng lặp lại' },
  { key: 'repeatFooters', label: 'Xoá footer trùng lặp', hint: 'Chân trang lặp lại ở nhiều trang / nhiều hoá đơn gộp (tổng cộng, chữ ký, cộng chuyển trang...) — giữ lại bản cuối cùng (dòng Tổng cộng của các hoá đơn phía trên cũng bị xoá)' },
  { key: 'pageMarkers', label: 'Xoá dòng số trang', hint: '"Trang 1/11", "tiếp theo trang trước - trang 2/11"...' },
  { key: 'blankRows', label: 'Xoá hàng trống', hint: 'Không xoá hàng nằm trong ô gộp nhiều hàng' },
  { key: 'tableOnly', label: 'Chỉ giữ bảng hàng hoá', hint: 'Bỏ thông tin công ty, dòng 1 2 3 4 5, tổng cộng, chữ ký -> sẵn sàng làm Pivot / nhập liệu' },
];

const REASON_TONE: Record<TrimReason, string> = {
  'repeat-header': 'tone-red',
  'repeat-footer': 'tone-amber',
  'rpac-header': 'tone-red',
  template: 'tone-red',
  'page-marker': 'tone-amber',
  spacer: 'tone-amber',
  blank: 'tone-slate',
  'outside-table': 'tone-violet',
  'index-row': 'tone-violet',
};

export default function TrimPane({
  strategy,
  onStrategy,
  autoOpts,
  onAutoOpts,
  trimMode,
  onTrimMode,
  template,
  plan,
  onFocusBlock,
  onApply,
  onClose,
  undoLabel,
  onUndo,
  hasFile,
}: {
  strategy: TrimStrategy;
  onStrategy: (s: TrimStrategy) => void;
  autoOpts: AutoTrimOptions;
  onAutoOpts: (o: AutoTrimOptions) => void;
  trimMode: 'HEADER' | 'FOOTER';
  onTrimMode: (m: 'HEADER' | 'FOOTER') => void;
  template: TemplateInfo | null;
  plan: TrimPlan | null;
  onFocusBlock: (b: TrimBlock) => void;
  onApply: () => void;
  onClose: () => void;
  undoLabel: string | null;
  onUndo: () => void;
  hasFile: boolean;
}) {
  const counts = new Map<TrimReason, number>();
  plan?.blocks.forEach((b) => counts.set(b.reason, (counts.get(b.reason) || 0) + (b.end - b.start + 1)));
  const total = plan?.rows.length ?? 0;

  return (
    <aside className="task-pane">
      <div className="pane-header">
        <div className="pane-title">
          <span className="pane-icon bg-rose">
            <Icon name="scissors" size={16} />
          </span>
          <div>
            <div className="pane-title-text">Cắt gộp dữ liệu</div>
            <div className="pane-subtitle">Xoá header / footer lặp lại khi gộp nhiều trang</div>
          </div>
        </div>
        <button className="icon-btn" onClick={onClose} title="Đóng (Esc)">
          <Icon name="close" size={16} />
        </button>
      </div>

      <div className="pane-body">
        <div className="segmented">
          {(
            [
              ['auto', 'Tự động', 'Khuyên dùng'],
              ['template', 'Theo mẫu', 'Bôi chọn'],
              ['rpac', 'R-pac', 'Lọc link'],
            ] as const
          ).map(([key, label, sub]) => (
            <button key={key} className={`segmented-item ${strategy === key ? 'is-active' : ''}`} onClick={() => onStrategy(key)}>
              <span>{label}</span>
              <small>{sub}</small>
            </button>
          ))}
        </div>

        {strategy === 'auto' ? (
          <div className="card">
            <p className="card-note">
              Tự nhận diện dòng tiêu đề bảng <b>STT · Tên hàng · Số lượng · Đơn giá · Thành tiền</b>. Lần xuất hiện đầu tiên là header gốc, các lần sau là header lặp do
              chuyển PDF → Excel.
            </p>
            {AUTO_OPTIONS.map((opt) => (
              <label key={opt.key} className="check-row">
                <input type="checkbox" checked={autoOpts[opt.key]} onChange={(e) => onAutoOpts({ ...autoOpts, [opt.key]: e.target.checked })} />
                <span>
                  <span className="check-label">{opt.label}</span>
                  <span className="check-hint">{opt.hint}</span>
                </span>
              </label>
            ))}
          </div>
        ) : (
          <div className="card">
            <div className="field-label">Chức năng xử lý</div>
            <div className="radio-cards">
              <button className={`radio-card ${trimMode === 'HEADER' ? 'is-active' : ''}`} onClick={() => onTrimMode('HEADER')}>
                <b>Xoá Header</b>
                <small>Giữ vùng mẫu gốc</small>
              </button>
              <button className={`radio-card ${trimMode === 'FOOTER' ? 'is-active' : ''}`} onClick={() => onTrimMode('FOOTER')}>
                <b>Xoá Footer</b>
                <small>Xoá cả vùng mẫu</small>
              </button>
            </div>
            <div className="field-label mt-3 flex items-center gap-2">
              Vùng mẫu
              {strategy === 'rpac' && <Image src="/rpac_logo.png" alt="R-pac" width={64} height={16} style={{ height: 16, width: "auto" }} />}
            </div>
            {template ? (
              <div className="template-box">
                <div className="template-head">
                  <span className="kbd">{template.rangeName}</span>
                  <span>{template.r2 - template.r1 + 1} hàng mẫu</span>
                </div>
                <div className="template-rows">
                  {template.preview.map((line, i) => (
                    <div key={i} className="template-row">
                      <span className="template-row-num">{template.r1 + i + 1}</span>
                      <span className="truncate">{line || '(trống)'}</span>
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="empty-hint">
                <Icon name="mouse" size={18} />
                <span>Bôi đen vùng {trimMode === 'HEADER' ? 'tiêu đề' : 'chân trang'} làm mẫu trực tiếp trên bảng tính. Vùng chọn sẽ tự được nhận làm mẫu.</span>
              </div>
            )}
          </div>
        )}

        <div className="card">
          <div className="flex items-baseline justify-between">
            <div className="field-label mb-0">Xem trước</div>
            <div className="legend">
              <span className="legend-dot bg-red-500/70" /> xoá
              <span className="legend-dot bg-emerald-600/60 ml-2" /> giữ
            </div>
          </div>
          {!hasFile ? (
            <div className="empty-hint">
              <Icon name="file" size={18} />
              <span>Hãy mở file Excel trước.</span>
            </div>
          ) : !plan ? (
            <div className="empty-hint">
              <Icon name="mouse" size={18} />
              <span>Chưa có vùng mẫu để quét.</span>
            </div>
          ) : (
            <>
              <div className="plan-summary">
                <div className="plan-number">{total}</div>
                <div className="plan-caption">hàng sẽ bị xoá</div>
              </div>
              {counts.size > 0 && (
                <div className="chips">
                  {Array.from(counts.entries()).map(([reason, n]) => (
                    <span key={reason} className={`chip ${REASON_TONE[reason]}`}>
                      {TRIM_REASON_LABELS[reason]} · {n}
                    </span>
                  ))}
                </div>
              )}
              {plan.notes.map((note, i) => (
                <div key={i} className="note">
                  <Icon name="info" size={14} className="shrink-0 mt-0.5" />
                  <span>{note}</span>
                </div>
              ))}
              {plan.blocks.length > 0 && (
                <div className="block-list" role="list">
                  {plan.blocks.map((b, i) => (
                    <button key={i} className="block-item" onClick={() => onFocusBlock(b)} title="Nhấp để chọn các hàng này trên bảng tính">
                      <span className={`block-bar ${REASON_TONE[b.reason]}`} />
                      <span className="block-rows">{b.start === b.end ? `Hàng ${b.start + 1}` : `Hàng ${b.start + 1}–${b.end + 1}`}</span>
                      <span className="block-reason">{TRIM_REASON_LABELS[b.reason]}</span>
                      <span className="block-count">{b.end - b.start + 1}</span>
                    </button>
                  ))}
                </div>
              )}
            </>
          )}
        </div>
      </div>

      <div className="pane-footer">
        {undoLabel && (
          <button className="xt-btn xt-btn-ghost" onClick={onUndo} title={undoLabel}>
            <Icon name="undo" size={15} /> Hoàn tác
          </button>
        )}
        <button className="xt-btn xt-btn-danger ml-auto" disabled={!plan || total === 0} onClick={onApply}>
          <Icon name="scissors" size={15} /> Áp dụng xoá {total > 0 ? `${total} hàng` : ''}
        </button>
      </div>
    </aside>
  );
}
