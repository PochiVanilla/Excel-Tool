// ==========================================
// HIỆU ỨNG NÚT BẤM
// - Rê chuột vào nút ribbon: sticker mèo (GIF) chạy lại từ đầu (components/CatMascot)
// - Bấm nút: nút nảy, dấu chân mèo / cá / cuộn len bắn ra, bong bóng "Meow!" bật lên
// - Không chặn chuột (pointer-events: none), tự dọn sau khi chạy xong; tắt nếu người dùng chọn giảm chuyển động
// ==========================================

const CLICK_TARGETS = '.ribbon-btn, .xt-btn, .segmented-item, .split-output .icon-btn, .pane-footer button, .dropzone';
const PARTICLES = ['🐾', '🐾', '🐾', '🐟', '🧶', '✨', '💜'];
const PHRASES = ['Meow!', 'Meo meo!', 'Nyan~', 'Purr…', 'Xong meo!', 'Mrrr!', 'Meo ơi!', 'Tuyệt meo!'];

let layer: HTMLDivElement | null = null;

const getLayer = () => {
  if (!layer || !layer.isConnected) {
    layer = document.createElement('div');
    layer.className = 'fx-layer';
    document.body.appendChild(layer);
  }
  return layer;
};

const spawn = (el: HTMLElement, ms: number) => {
  getLayer().appendChild(el);
  window.setTimeout(() => el.remove(), ms);
};

const pick = <T,>(list: T[]) => list[Math.floor(Math.random() * list.length)];

const burst = (x: number, y: number) => {
  const count = 9;
  for (let i = 0; i < count; i++) {
    const p = document.createElement('span');
    p.className = 'fx-particle';
    p.textContent = pick(PARTICLES);
    const angle = (Math.PI * 2 * i) / count + Math.random() * 0.6;
    const dist = 50 + Math.random() * 60;
    p.style.left = `${x}px`;
    p.style.top = `${y}px`;
    p.style.setProperty('--dx', `${Math.cos(angle) * dist}px`);
    p.style.setProperty('--dy', `${Math.sin(angle) * dist - 30}px`);
    p.style.setProperty('--rot', `${(Math.random() - 0.5) * 540}deg`);
    p.style.setProperty('--size', `${14 + Math.random() * 10}px`);
    spawn(p, 1100);
  }
};

const bubble = (target: HTMLElement) => {
  const r = target.getBoundingClientRect();
  const b = document.createElement('span');
  b.className = 'fx-bubble';
  b.textContent = pick(PHRASES);
  b.style.left = `${Math.min(window.innerWidth - 60, Math.max(60, r.left + r.width / 2))}px`;
  // Nút sát mép trên (ribbon) -> bong bóng hiện phía dưới nút để không bị khuất
  b.style.top = `${r.top < 110 ? r.bottom + 52 : r.top - 6}px`;
  b.style.setProperty('--tilt', `${(Math.random() - 0.5) * 16}deg`);
  spawn(b, 1300);
};

const boing = (target: HTMLElement) => {
  target.animate(
    [
      { transform: 'scale(1)' },
      { transform: 'scale(0.88, 0.9)', offset: 0.25 },
      { transform: 'scale(1.07, 1.04) rotate(-1.5deg)', offset: 0.55 },
      { transform: 'scale(0.98) rotate(1deg)', offset: 0.8 },
      { transform: 'scale(1)' },
    ],
    { duration: 500, easing: 'cubic-bezier(0.28, 1.6, 0.5, 1)' },
  );
};

// Gán lại src để GIF chạy lại từ khung hình đầu mỗi lần rê chuột vào nút
const restartGif = (btn: HTMLElement) => {
  const gif = btn.querySelector<HTMLImageElement>('img.m-gif');
  const src = gif?.dataset.src;
  if (gif && src) gif.src = `${src}#${Date.now()}`;
};

export const installFunClicks = () => {
  const reduce = window.matchMedia?.('(prefers-reduced-motion: reduce)');
  const onClick = (e: MouseEvent) => {
    if (reduce?.matches) return;
    const target = (e.target as HTMLElement | null)?.closest<HTMLElement>(CLICK_TARGETS);
    if (!target || (target as HTMLButtonElement).disabled) return;
    boing(target);
    burst(e.clientX, e.clientY);
    bubble(target);
  };
  const onOver = (e: MouseEvent) => {
    const btn = (e.target as HTMLElement | null)?.closest<HTMLElement>('.ribbon-btn');
    if (!btn || btn.contains(e.relatedTarget as Node) || (btn as HTMLButtonElement).disabled) return;
    restartGif(btn);
  };
  document.addEventListener('click', onClick, true);
  document.addEventListener('mouseover', onOver, true);
  return () => {
    document.removeEventListener('click', onClick, true);
    document.removeEventListener('mouseover', onOver, true);
  };
};
