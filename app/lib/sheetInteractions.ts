// ==========================================
// TƯƠNG TÁC NHANH KIỂU EXCEL BỔ SUNG CHO LUCKYSHEET
// (Luckysheet đã có sẵn: kéo chọn, Shift/Ctrl + click, Fill handle, kéo viền để di chuyển,
//  double-click ranh giới cột, Ctrl + C/X/V/Z/Y/B/I/F/H, Ctrl + mũi tên, menu chuột phải cơ bản)
// Phần bổ sung: Ctrl + lăn chuột (zoom), Shift + lăn chuột (cuộn ngang), double-click ranh giới hàng,
//  menu chuột phải "Công cụ nhanh", phím tắt Ctrl+O/S, Ctrl+Space, Shift+Space, Ctrl+-, Ctrl++, Alt+=,
//  Ctrl+PageUp/PageDown, F1, kéo thả file vào cửa sổ
// ==========================================
import { getActiveFile, getActiveOrder, getFiles, getSelection, getSheetData, ls, autoFitRowHeight, deleteRowsWithImages } from './lucky';
import { getCellNumber, getColLetter } from './sheetText';

export type ContextAction =
  | 'trim-template'
  | 'pivot-selection'
  | 'blank-rows-selection'
  | 'autofit-cols'
  | 'autofit-rows'
  | 'freeze-here'
  | 'unfreeze'
  | 'copy-sum';

export interface InteractionHandlers {
  getZoom: () => number;
  onZoom: (z: number) => void;
  onOpenFile: () => void;
  onExport: () => void;
  onHelp: () => void;
  onEscape: () => void;
  onDropFile: (file: File) => void;
  onDragState: (over: boolean) => void;
  onContextAction: (action: ContextAction) => void;
  onSelectionChange: () => void;
  onStructureChange: (label: string, reload?: boolean) => void;
  beforeStructureChange: (label: string) => void;
  notify: (text: string) => void;
}

const CONTAINER_ID = 'luckysheet-container';

// Luckysheet luôn giữ ô nhập ẩn (display:block, rộng vài px); khi đang sửa ô nó chuyển sang display:flex
const isEditingCell = () => {
  const box = document.getElementById('luckysheet-input-box');
  if (!box) return false;
  return window.getComputedStyle(box).display === 'flex';
};

const isOurFormField = (el: EventTarget | null) => {
  if (!(el instanceof HTMLElement)) return false;
  if (el.closest('#' + CONTAINER_ID)) return false;
  return el.isContentEditable || ['INPUT', 'TEXTAREA', 'SELECT'].includes(el.tagName);
};

const sheetSize = () => {
  const data = getSheetData();
  return { rows: data.length, cols: data[0]?.length || 0 };
};

// Vùng chọn đang là cả hàng / cả cột?
const selectionKind = () => {
  const sel = getSelection()[0];
  if (!sel) return null;
  const { rows, cols } = sheetSize();
  const fullRows = sel.column[0] === 0 && sel.column[1] >= cols - 1;
  const fullCols = sel.row[0] === 0 && sel.row[1] >= rows - 1;
  return { sel, fullRows, fullCols };
};

// Tự giãn độ rộng cột theo nội dung của TẤT CẢ các hàng (Luckysheet gốc chỉ đo các hàng đang hiển thị)
export const autoFitColumns = (c1: number, c2: number) => {
  const data = getSheetData();
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  if (!ctx) return;
  const widths: Record<number, number> = {};
  for (let c = c1; c <= c2; c++) {
    let max = 40;
    for (let r = 0; r < data.length; r++) {
      const cell = data[r]?.[c];
      if (!cell || cell.mc) continue;
      const text =
        cell.ct?.t === 'inlineStr' && cell.ct.s ? cell.ct.s.map((s: { v?: string }) => s.v ?? '').join('') : String(cell.m ?? cell.v ?? '');
      if (!text) continue;
      const fs = Number(cell.fs || 10);
      ctx.font = `${Number(cell.bl) === 1 ? 'bold ' : ''}${Math.round(fs * 1.333)}px ${cell.ff || 'Arial'}`;
      for (const line of text.split(/\r?\n/)) max = Math.max(max, ctx.measureText(line).width + 12);
    }
    widths[c] = Math.min(Math.ceil(max), 800);
  }
  ls().setColumnWidth(widths);
};

// Alt + = : chèn =SUM() cho vùng số liền kề phía trên (hoặc ngay dưới vùng chọn dọc)
const autoSum = (notify: (t: string) => void) => {
  const sel = getSelection()[0];
  if (!sel) return;
  const data = getSheetData();
  const c = sel.column[0];
  let target = sel.row[0];
  let top: number;
  let bottom: number;
  if (sel.row[1] > sel.row[0]) {
    top = sel.row[0];
    bottom = sel.row[1];
    target = sel.row[1] + 1;
  } else {
    bottom = target - 1;
    while (bottom >= 0 && getCellNumber(data[bottom]?.[c]) === null) bottom--;
    top = bottom;
    while (top - 1 >= 0 && getCellNumber(data[top - 1]?.[c]) !== null) top--;
  }
  if (bottom < 0 || top > bottom) {
    notify('AutoSum: không tìm thấy vùng số phía trên ô đang chọn');
    return;
  }
  const col = getColLetter(c);
  ls().setCellValue(target, c, `=SUM(${col}${top + 1}:${col}${bottom + 1})`);
  ls().setRangeShow({ row: [target, target], column: [c, c] });
};

export const installSheetInteractions = (handlersRef: { current: InteractionHandlers }) => {
  const h = () => handlersRef.current;
  let sheetFocused = false;
  let zoomRaf = 0;
  let pendingZoom = 0;

  const container = () => document.getElementById(CONTAINER_ID);

  // Theo dõi người dùng đang thao tác trên bảng tính hay trên panel của web
  const onMouseDown = (e: MouseEvent) => {
    const c = container();
    const target = e.target as Node;
    const inSheet = !!c && c.contains(target);
    // Các menu / hộp thoại Luckysheet gắn vào body cũng tính là đang ở bảng tính
    const inLuckyPopup = target instanceof HTMLElement && !!target.closest('[id^="luckysheet"], [class*="luckysheet"]');
    sheetFocused = inSheet || inLuckyPopup;
  };

  // Ctrl + lăn chuột: zoom (chặn zoom cả trang của trình duyệt). Shift + lăn chuột: cuộn ngang
  const onWheel = (e: WheelEvent) => {
    const c = container();
    if (!c || !c.contains(e.target as Node)) return;
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      e.stopPropagation();
      const base = pendingZoom || h().getZoom();
      const step = e.deltaY < 0 ? 0.1 : -0.1;
      pendingZoom = Math.min(4, Math.max(0.1, Math.round((base + step) * 10) / 10));
      if (!zoomRaf) {
        zoomRaf = requestAnimationFrame(() => {
          h().onZoom(pendingZoom);
          pendingZoom = 0;
          zoomRaf = 0;
        });
      }
      return;
    }
    if (e.shiftKey && Math.abs(e.deltaX) < 1 && Math.abs(e.deltaY) > 0) {
      const bar = document.getElementById('luckysheet-scrollbar-x');
      if (bar) {
        e.preventDefault();
        e.stopPropagation();
        bar.scrollLeft += e.deltaY;
      }
    }
  };

  // Double-click ranh giới giữa 2 hàng ở cột tiêu đề hàng -> tự giãn chiều cao hàng phía trên
  const onDblClick = (e: MouseEvent) => {
    const rowsHeader = document.getElementById('luckysheet-rows-h');
    if (!rowsHeader || !rowsHeader.contains(e.target as Node)) return;
    const file = getActiveFile();
    if (!file) return;
    const zoom = Number(file.zoomRatio) || 1;
    const rowlen = file.config?.rowlen || {};
    const rowhidden = file.config?.rowhidden || {};
    const defaultRow = Number(file.defaultRowHeight) || 19;
    const rect = rowsHeader.getBoundingClientRect();
    const scrollBar = document.getElementById('luckysheet-scrollbar-y');
    const y = e.clientY - rect.top + (scrollBar ? scrollBar.scrollTop : rowsHeader.scrollTop);
    const { rows } = sheetSize();
    let bottom = 0;
    for (let r = 0; r < rows; r++) {
      if (rowhidden[r] !== undefined) continue;
      bottom += Math.round((rowlen[r] !== undefined ? Number(rowlen[r]) : defaultRow) * zoom) + 1;
      if (Math.abs(y - bottom) <= 4) {
        const sel = getSelection()[0];
        const targets = sel && sel.row[0] <= r && r <= sel.row[1] && sel.row[1] > sel.row[0] ? range(sel.row[0], sel.row[1]) : [r];
        targets.forEach(autoFitRowHeight);
        e.preventDefault();
        return;
      }
      if (bottom > y + 6) return;
    }
  };

  const range = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i);

  const onKeyDown = (e: KeyboardEvent) => {
    const key = e.key;
    const ctrl = e.ctrlKey || e.metaKey;
    if (isOurFormField(e.target) && !(ctrl && (key === 's' || key === 'o'))) {
      if (key === 'Escape') h().onEscape();
      return;
    }
    const editing = isEditingCell();

    if (ctrl && !e.shiftKey && !e.altKey && (key === 'o' || key === 'O')) {
      e.preventDefault();
      h().onOpenFile();
      return;
    }
    if (ctrl && !e.altKey && (key === 's' || key === 'S')) {
      e.preventDefault();
      if (editing) ls().exitEditMode?.();
      h().onExport();
      return;
    }
    if (key === 'F1' || (ctrl && key === '/')) {
      e.preventDefault();
      h().onHelp();
      return;
    }
    if (editing) return;
    if (key === 'Escape') {
      h().onEscape();
      return;
    }
    if (!sheetFocused) return;

    const stop = () => {
      e.preventDefault();
      e.stopPropagation();
      e.stopImmediatePropagation();
    };

    // Ctrl + Space: chọn cả cột · Shift + Space: chọn cả hàng
    if (key === ' ' && (ctrl || e.shiftKey) && !(ctrl && e.shiftKey)) {
      const sel = getSelection()[0];
      if (!sel) return;
      stop();
      const { rows, cols } = sheetSize();
      if (ctrl) ls().setRangeShow({ row: [0, Math.max(0, rows - 1)], column: sel.column });
      else ls().setRangeShow({ row: sel.row, column: [0, Math.max(0, cols - 1)] });
      h().onSelectionChange();
      return;
    }

    // Ctrl + PageUp / PageDown: chuyển sheet
    if (ctrl && (key === 'PageUp' || key === 'PageDown')) {
      stop();
      const files = getFiles();
      const cur = getActiveOrder();
      const next = key === 'PageUp' ? cur - 1 : cur + 1;
      if (next >= 0 && next < files.length) ls().setSheetActive(next);
      return;
    }

    // Ctrl + - : xoá hàng/cột đang chọn · Ctrl + Shift + = (Ctrl + +): chèn hàng/cột
    if (ctrl && !e.altKey && (key === '-' || key === 'Subtract')) {
      const kind = selectionKind();
      if (!kind) return;
      stop();
      const { sel, fullRows, fullCols } = kind;
      if (fullCols && !fullRows) {
        h().beforeStructureChange(`Xoá cột ${getColLetter(sel.column[0])}:${getColLetter(sel.column[1])}`);
        ls().deleteColumn(sel.column[0], sel.column[1]);
        h().onStructureChange('Đã xoá cột');
      } else {
        const rows = range(sel.row[0], sel.row[1]);
        h().beforeStructureChange(`Xoá ${rows.length} hàng`);
        const reload = deleteRowsWithImages(rows);
        h().onStructureChange(`Đã xoá ${rows.length} hàng (${sel.row[0] + 1}–${sel.row[1] + 1})`, reload);
      }
      return;
    }
    if (ctrl && !e.altKey && (key === '+' || key === 'Add' || (e.shiftKey && key === '='))) {
      const kind = selectionKind();
      if (!kind) return;
      stop();
      const { sel, fullRows, fullCols } = kind;
      if (fullCols && !fullRows) {
        ls().insertColumn(sel.column[0], { number: sel.column[1] - sel.column[0] + 1 });
        h().onStructureChange('Đã chèn cột');
      } else {
        ls().insertRow(sel.row[0], { number: sel.row[1] - sel.row[0] + 1 });
        h().onStructureChange(`Đã chèn ${sel.row[1] - sel.row[0] + 1} hàng`);
      }
      return;
    }

    // Alt + = : AutoSum
    if (e.altKey && !ctrl && key === '=') {
      stop();
      autoSum(h().notify);
      return;
    }
  };

  // Chọn vùng bằng bàn phím (Shift + mũi tên...) -> cập nhật lại thanh trạng thái
  let keyTimer = 0;
  const onKeyUp = () => {
    if (!sheetFocused) return;
    window.clearTimeout(keyTimer);
    keyTimer = window.setTimeout(() => h().onSelectionChange(), 60);
  };

  // Menu chuột phải: chèn nhóm "Công cụ nhanh" lên đầu menu gốc của Luckysheet
  const QUICK: { action: ContextAction; label: string; hint?: string }[] = [
    { action: 'trim-template', label: 'Dùng vùng chọn làm mẫu cắt gộp' },
    { action: 'pivot-selection', label: 'Tạo Pivot từ vùng chọn' },
    { action: 'blank-rows-selection', label: 'Xoá hàng trống trong vùng chọn' },
    { action: 'autofit-cols', label: 'Tự giãn độ rộng cột (AutoFit)' },
    { action: 'autofit-rows', label: 'Tự giãn chiều cao hàng' },
    { action: 'freeze-here', label: 'Cố định dòng/cột tại ô này (Freeze)' },
    { action: 'unfreeze', label: 'Bỏ cố định dòng/cột' },
    { action: 'copy-sum', label: 'Sao chép tổng (Sum) vùng chọn' },
  ];

  const injectQuickMenu = () => {
    const menu = document.getElementById('luckysheet-rightclick-menu');
    if (!menu || window.getComputedStyle(menu).display === 'none') return;
    if (!menu.querySelector('#xt-quick-menu')) {
      const group = document.createElement('div');
      group.id = 'xt-quick-menu';
      group.className = 'luckysheet-mousedown-cancel';
      group.innerHTML =
        '<div class="xt-quick-title luckysheet-mousedown-cancel">Công cụ nhanh</div>' +
        QUICK.map(
          (q) =>
            `<div class="luckysheet-cols-menuitem luckysheet-mousedown-cancel xt-quick-item" data-action="${q.action}"><div class="luckysheet-cols-menuitem-content luckysheet-mousedown-cancel">${q.label}</div></div>`,
        ).join('') +
        '<div class="luckysheet-menuseparator luckysheet-mousedown-cancel xt-quick-sep" role="separator"></div>';
      group.addEventListener('mousedown', (ev) => ev.stopPropagation());
      group.addEventListener('click', (ev) => {
        const item = (ev.target as HTMLElement).closest<HTMLElement>('[data-action]');
        if (!item) return;
        menu.style.display = 'none';
        h().onContextAction(item.dataset.action as ContextAction);
      });
      menu.insertBefore(group, menu.firstChild);
    }
    // Menu dài hơn sau khi chèn -> đẩy lên cho vừa màn hình
    const rect = menu.getBoundingClientRect();
    if (rect.bottom > window.innerHeight - 4) menu.style.top = `${Math.max(4, window.innerHeight - rect.height - 8)}px`;
    if (rect.right > window.innerWidth - 4) menu.style.left = `${Math.max(4, window.innerWidth - rect.width - 8)}px`;
  };
  // Luckysheet hiện menu ngay khi nhả chuột phải (menu nằm dưới con trỏ nên sự kiện contextmenu rơi vào chính menu)
  const onRightMouseUp = (e: MouseEvent) => {
    if (e.button !== 2) return;
    const c = container();
    const target = e.target as Node;
    const menu = document.getElementById('luckysheet-rightclick-menu');
    if (!(c && c.contains(target)) && !(menu && menu.contains(target))) return;
    sheetFocused = true;
    window.setTimeout(injectQuickMenu, 0);
    window.setTimeout(injectQuickMenu, 80);
  };

  // Kéo thả file Excel vào bất kỳ đâu trên cửa sổ
  let dragDepth = 0;
  const hasFiles = (e: DragEvent) => !!e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  const onDragEnter = (e: DragEvent) => {
    if (!hasFiles(e)) return;
    dragDepth++;
    h().onDragState(true);
  };
  const onDragOver = (e: DragEvent) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
  };
  const onDragLeave = (e: DragEvent) => {
    if (!hasFiles(e)) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) h().onDragState(false);
  };
  const onDrop = (e: DragEvent) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    e.stopPropagation();
    dragDepth = 0;
    h().onDragState(false);
    const file = e.dataTransfer?.files?.[0];
    if (file) h().onDropFile(file);
  };

  document.addEventListener('mousedown', onMouseDown, true);
  document.addEventListener('wheel', onWheel, { capture: true, passive: false });
  document.addEventListener('dblclick', onDblClick, true);
  document.addEventListener('keydown', onKeyDown, true);
  document.addEventListener('keyup', onKeyUp, true);
  document.addEventListener('mouseup', onRightMouseUp, true);
  window.addEventListener('dragenter', onDragEnter, true);
  window.addEventListener('dragover', onDragOver, true);
  window.addEventListener('dragleave', onDragLeave, true);
  window.addEventListener('drop', onDrop, true);

  return () => {
    document.removeEventListener('mousedown', onMouseDown, true);
    document.removeEventListener('wheel', onWheel, true);
    document.removeEventListener('dblclick', onDblClick, true);
    document.removeEventListener('keydown', onKeyDown, true);
    document.removeEventListener('keyup', onKeyUp, true);
    document.removeEventListener('mouseup', onRightMouseUp, true);
    window.removeEventListener('dragenter', onDragEnter, true);
    window.removeEventListener('dragover', onDragOver, true);
    window.removeEventListener('dragleave', onDragLeave, true);
    window.removeEventListener('drop', onDrop, true);
    if (zoomRaf) cancelAnimationFrame(zoomRaf);
  };
};
