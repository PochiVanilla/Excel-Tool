'use client';

// Thông báo nổi góc dưới (thay cho alert chặn màn hình), có thể kèm nút hành động (vd: Hoàn tác)
import { useCallback, useRef, useState } from 'react';
import Icon from './Icon';

export type ToastTone = 'success' | 'error' | 'info' | 'warn';

export interface Toast {
  id: number;
  tone: ToastTone;
  text: string;
  action?: { label: string; run: () => void };
}

export const useToasts = () => {
  const [toasts, setToasts] = useState<Toast[]>([]);
  const seq = useRef(0);

  const dismiss = useCallback((id: number) => setToasts((list) => list.filter((t) => t.id !== id)), []);

  const push = useCallback(
    (tone: ToastTone, text: string, action?: Toast['action'], ms = 4500) => {
      const id = ++seq.current;
      setToasts((list) => [...list.slice(-3), { id, tone, text, action }]);
      window.setTimeout(() => dismiss(id), action ? ms + 3500 : ms);
    },
    [dismiss],
  );

  return { toasts, push, dismiss };
};

const TONE_ICON = { success: 'check', error: 'alert', info: 'info', warn: 'alert' } as const;

export function ToastStack({ toasts, dismiss }: { toasts: Toast[]; dismiss: (id: number) => void }) {
  return (
    <div className="toast-stack" aria-live="polite">
      {toasts.map((t) => (
        <div key={t.id} className={`toast toast-${t.tone}`}>
          <Icon name={TONE_ICON[t.tone]} size={16} className="shrink-0 mt-0.5" />
          <span className="flex-1">{t.text}</span>
          {t.action && (
            <button
              className="toast-action"
              onClick={() => {
                t.action!.run();
                dismiss(t.id);
              }}
            >
              {t.action.label}
            </button>
          )}
          <button className="toast-close" onClick={() => dismiss(t.id)} aria-label="Đóng">
            <Icon name="close" size={14} />
          </button>
        </div>
      ))}
    </div>
  );
}
