'use client';

// Thanh trạng thái kiểu Excel: Average / Count / Numerical Count / Min / Max / Sum + thanh trượt thu phóng
// - Nhấp vào 1 số liệu để sao chép (như Excel 365)
// - Chuột phải lên thanh trạng thái để chọn số liệu hiển thị (Customize Status Bar)
import { useEffect, useRef, useState } from 'react';
import Icon from './Icon';

export interface SelectionStats {
  cells: number;
  count: number;
  numCount: number;
  sum: number;
  min: number | null;
  max: number | null;
}

type StatKey = 'average' | 'count' | 'numCount' | 'min' | 'max' | 'sum';

const STAT_LABELS: Record<StatKey, string> = {
  average: 'Average',
  count: 'Count',
  numCount: 'Numerical Count',
  min: 'Min',
  max: 'Max',
  sum: 'Sum',
};

const DEFAULT_VISIBLE: StatKey[] = ['average', 'count', 'sum'];
const STORAGE_KEY = 'excel-tool.statusbar';

export const formatNumber = (n: number) =>
  n.toLocaleString('en-US', { maximumFractionDigits: Math.abs(n) >= 1000 ? 2 : 4 });

export default function StatusBar({
  stats,
  zoom,
  onZoom,
  message,
  ready,
  onCopied,
}: {
  stats: SelectionStats | null;
  zoom: number;
  onZoom: (z: number) => void;
  message: string;
  ready: boolean;
  onCopied: (text: string) => void;
}) {
  const [visible, setVisible] = useState<StatKey[]>(() => {
    if (typeof window === 'undefined') return DEFAULT_VISIBLE;
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
      if (Array.isArray(saved)) return saved.filter((k): k is StatKey => k in STAT_LABELS);
    } catch {
      /* bỏ qua: trình duyệt chặn localStorage */
    }
    return DEFAULT_VISIBLE;
  });
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!menu) return;
    const close = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) setMenu(null);
    };
    document.addEventListener('mousedown', close);
    return () => document.removeEventListener('mousedown', close);
  }, [menu]);

  const toggle = (key: StatKey) => {
    const next = visible.includes(key) ? visible.filter((k) => k !== key) : [...visible, key];
    const ordered = (Object.keys(STAT_LABELS) as StatKey[]).filter((k) => next.includes(k));
    setVisible(ordered);
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(ordered));
    } catch {
      /* bỏ qua */
    }
  };

  const values: Record<StatKey, number | null> = {
    average: stats && stats.numCount > 0 ? stats.sum / stats.numCount : null,
    count: stats ? stats.count : null,
    numCount: stats ? stats.numCount : null,
    min: stats ? stats.min : null,
    max: stats ? stats.max : null,
    sum: stats && stats.numCount > 0 ? stats.sum : null,
  };

  // Giống Excel: chỉ hiện khi chọn từ 2 ô trở lên; Average/Sum/Min/Max chỉ hiện khi có số
  const showStats = !!stats && stats.cells > 1 && stats.count > 0;
  const zoomPct = Math.round(zoom * 100);

  const copy = (label: string, value: number) => {
    const text = String(Math.round(value * 1e6) / 1e6);
    navigator.clipboard?.writeText(text).then(
      () => onCopied(`Đã sao chép ${label}: ${formatNumber(value)}`),
      () => onCopied('Trình duyệt không cho phép sao chép'),
    );
  };

  return (
    <footer
      className="statusbar"
      onContextMenu={(e) => {
        e.preventDefault();
        setMenu({ x: e.clientX, y: e.clientY });
      }}
    >
      <div className="statusbar-left">
        <span className={`status-dot ${ready ? 'is-ready' : ''}`} />
        <span className="statusbar-message" title={message}>
          {message || (ready ? 'Sẵn sàng' : 'Đang tải bảng tính...')}
        </span>
      </div>

      <div className="statusbar-right">
        {showStats &&
          visible.map((key) => {
            const v = values[key];
            if (v === null) return null;
            return (
              <button key={key} className="stat-chip" title={`Nhấp để sao chép ${STAT_LABELS[key]}`} onClick={() => copy(STAT_LABELS[key], v)}>
                <span className="stat-label">{STAT_LABELS[key]}:</span> <strong>{formatNumber(v)}</strong>
              </button>
            );
          })}

        <div className="zoom-control" title="Ctrl + lăn chuột để thu phóng">
          <button className="zoom-btn" onClick={() => onZoom(Math.max(0.1, Math.round((zoom - 0.1) * 10) / 10))} aria-label="Thu nhỏ">
            <Icon name="minus" size={14} />
          </button>
          <input
            type="range"
            min={10}
            max={400}
            step={5}
            value={zoomPct}
            onChange={(e) => onZoom(Number(e.target.value) / 100)}
            className="zoom-slider"
            aria-label="Thu phóng"
          />
          <button className="zoom-btn" onClick={() => onZoom(Math.min(4, Math.round((zoom + 0.1) * 10) / 10))} aria-label="Phóng to">
            <Icon name="plus" size={14} />
          </button>
          <button className="zoom-value" onClick={() => onZoom(1)} title="Nhấp để về 100%">
            {zoomPct}%
          </button>
        </div>
      </div>

      {menu && (
        <div ref={menuRef} className="popover-menu" style={{ left: Math.min(menu.x, window.innerWidth - 240), bottom: 30 }}>
          <div className="popover-title">Customize Status Bar</div>
          {(Object.keys(STAT_LABELS) as StatKey[]).map((key) => (
            <button key={key} className="popover-item" onClick={() => toggle(key)}>
              <span className="popover-check">{visible.includes(key) && <Icon name="check" size={14} />}</span>
              <span className="flex-1">{STAT_LABELS[key]}</span>
              <span className="popover-hint">{showStats && values[key] !== null ? formatNumber(values[key]!) : ''}</span>
            </button>
          ))}
        </div>
      )}
    </footer>
  );
}
