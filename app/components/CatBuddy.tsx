'use client';

// Mèo đồng hành ở góc dưới trái (sticker do người dùng cung cấp trong public/mascots/):
// - Bình thường đứng yên, rê chuột vào thì cử động; bấm vào để nhận mẹo dùng tool
// - Phản ứng theo kết quả công việc: nghe sự kiện "xt-cat" (bắn từ thông báo / Pivot / Tách PDF) -> đổi sticker + bong bóng câu nói
//   + phản ứng "giữ" (sticky: Pivot xong, Tách PDF xong): mèo giữ nguyên trạng thái tới khi tạo trang Excel mới (catReset)
//   + phản ứng thường (thông báo, mẹo): hiện vài giây rồi quay về trạng thái đang giữ
// - Lúc rảnh thỉnh thoảng tự làm trò; có nút ẩn / hiện (nhớ trong trình duyệt)
import { useEffect, useRef, useState } from 'react';

export type CatMood = 'success' | 'error' | 'please' | 'hmm';

const MOODS: Record<CatMood | 'idle' | 'fun', string[]> = {
  idle: ['/mascots/r-hmm.webp'],
  success: ['/mascots/r-success.webp'],
  error: ['/mascots/r-error.gif'],
  please: ['/mascots/r-please.webp'],
  hmm: ['/mascots/r-hmm.webp'],
  fun: ['/mascots/open.gif', '/mascots/split.gif', '/mascots/trim.gif', '/mascots/pivot.gif', '/mascots/download.gif', '/mascots/undo.gif'],
};

const DEFAULT_TEXT: Record<CatMood, string[]> = {
  success: ['Xong rồi nè! 😎', 'Ngon lành cành đào!', 'Giàu to rồi!'],
  error: ['Ối, lỗi rồi!', 'Cứu mèo với!', 'Hỏng rồi hỏng rồi!'],
  please: ['Kiểm tra lại giúp mèo nha 🥺', 'Hơi lệch xíu nè…'],
  hmm: ['Hừm…', 'Nhớ đọc kỹ nha!'],
};

const TIPS = [
  'Bôi đen bảng rồi bấm Pivot để tổng hợp đúng vùng chọn.',
  'Kéo thả file .xlsx hoặc .pdf thẳng vào cửa sổ là mở luôn.',
  'Tách PDF: chỉ cần chọn loại cho trang đầu, các trang sau tự theo.',
  'Ctrl + lăn chuột để phóng to / thu nhỏ bảng tính.',
  'Nhấp đúp vào viền dưới số dòng để tự giãn chiều cao dòng.',
  'Cắt gộp xong nhớ xem cột "Xem trước" trước khi bấm Áp dụng xoá.',
  'Ctrl + S để tải file đã xử lý.',
  'Ở Tách PDF, bấm nút bảng ở file VAT để trích ra Excel rồi Pivot luôn.',
];

const STORAGE_KEY = 'excel-tool.cat-buddy';
const pick = <T,>(list: T[]) => list[Math.floor(Math.random() * list.length)];

// Gọi từ bất kỳ đâu để mèo phản ứng; sticky = giữ trạng thái tới khi catReset()
export const catReact = (mood: CatMood, text?: string, opts: { sticky?: boolean } = {}) => {
  window.dispatchEvent(new CustomEvent('xt-cat', { detail: { mood, text, sticky: !!opts.sticky } }));
};

// Tạo trang Excel mới (mở / trích file, xuất sheet Pivot) -> mèo về trạng thái ban đầu
export const catReset = () => {
  window.dispatchEvent(new CustomEvent('xt-cat-reset'));
};

export default function CatBuddy() {
  const [visible, setVisible] = useState(() => {
    try {
      return typeof window === 'undefined' || localStorage.getItem(STORAGE_KEY) !== 'off';
    } catch {
      return true;
    }
  });
  // Phản ứng ngắn (vài giây) đè lên trạng thái đang giữ
  const [src, setSrc] = useState<string | null>(null);
  const [say, setSay] = useState<string | null>(null);
  // Trạng thái giữ (Pivot xong / Tách PDF xong) tới khi tạo trang Excel mới
  const [sticky, setSticky] = useState<{ src: string; say: string } | null>(null);
  const [stickySay, setStickySay] = useState(false);
  const [hover, setHover] = useState(false);
  const timer = useRef(0);
  const sayTimer = useRef(0);

  const show = (nextSrc: string | null, text: string | null, ms: number) => {
    window.clearTimeout(timer.current);
    // thêm "#t" để GIF chạy lại từ đầu mỗi lần phản ứng
    setSrc(nextSrc ? `${nextSrc}#${Date.now()}` : null);
    setSay(text);
    timer.current = window.setTimeout(() => {
      setSrc(null);
      setSay(null);
    }, ms);
  };

  // Phản ứng theo sự kiện công việc
  useEffect(() => {
    const onReact = (e: Event) => {
      const { mood, text, sticky: keep } = (e as CustomEvent<{ mood: CatMood; text?: string; sticky?: boolean }>).detail;
      const line = text || pick(DEFAULT_TEXT[mood]);
      if (keep) {
        // Giữ sticker; bong bóng câu nói hiện ~8 giây rồi thu lại (rê chuột vào mèo để xem lại)
        window.clearTimeout(timer.current);
        setSrc(null);
        setSay(null);
        setSticky({ src: `${pick(MOODS[mood])}#${Date.now()}`, say: line });
        setStickySay(true);
        window.clearTimeout(sayTimer.current);
        sayTimer.current = window.setTimeout(() => setStickySay(false), 8000);
        return;
      }
      show(pick(MOODS[mood]), line, mood === 'error' || mood === 'please' ? 6000 : 4500);
    };
    const onReset = () => {
      setSticky(null);
      setStickySay(false);
    };
    window.addEventListener('xt-cat', onReact);
    window.addEventListener('xt-cat-reset', onReset);
    return () => {
      window.removeEventListener('xt-cat', onReact);
      window.removeEventListener('xt-cat-reset', onReset);
    };
  }, []);

  // Lúc rảnh: 40-70 giây tự làm trò 1 lần (không làm trò khi đang giữ trạng thái)
  useEffect(() => {
    if (!visible || sticky) return;
    let t = 0;
    const loop = () => {
      t = window.setTimeout(() => {
        if (!document.hidden) show(pick(MOODS.fun), null, 3200);
        loop();
      }, 40000 + Math.random() * 30000);
    };
    loop();
    return () => window.clearTimeout(t);
  }, [visible, sticky]);

  useEffect(
    () => () => {
      window.clearTimeout(timer.current);
      window.clearTimeout(sayTimer.current);
    },
    [],
  );

  const toggle = (on: boolean) => {
    setVisible(on);
    try {
      localStorage.setItem(STORAGE_KEY, on ? 'on' : 'off');
    } catch {
      /* chế độ riêng tư: không lưu được thì thôi */
    }
  };

  if (!visible) {
    return (
      <button className="cat-buddy-show" onClick={() => toggle(true)} title="Gọi mèo đồng hành ra">
        🐱
      </button>
    );
  }

  const still = '/mascots/r-hmm.png';
  const playing = src || sticky?.src || (hover ? MOODS.idle[0] : null);
  const bubble = say || (sticky && (stickySay || hover) ? sticky.say : null);

  return (
    <div className={`cat-buddy ${sticky ? 'is-sticky' : ''}`} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}>
      {bubble && <div className="cat-buddy-say">{bubble}</div>}
      <button className="cat-buddy-btn" onClick={() => show(MOODS.hmm[0], `Mẹo: ${pick(TIPS)}`, 7000)} title="Bấm để nhận mẹo dùng tool">
        {/* eslint-disable-next-line @next/next/no-img-element -- sticker động (GIF / WebP) giữ nguyên file gốc */}
        <img src={playing || still} alt="Mèo đồng hành" draggable={false} />
      </button>
      <button className="cat-buddy-hide" onClick={() => toggle(false)} title="Ẩn mèo (bấm 🐱 ở góc để gọi lại)" aria-label="Ẩn mèo">
        ×
      </button>
    </div>
  );
}
