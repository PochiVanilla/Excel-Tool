'use client';

// Bảng tra cứu thao tác chuột & phím tắt kiểu Excel (F1 hoặc Ctrl + /)
import Icon from './Icon';

const MOUSE: [string, string][] = [
  ['Kéo chuột / Shift + Click', 'Chọn vùng / mở rộng vùng chọn'],
  ['Ctrl + Click / Ctrl + Kéo', 'Chọn nhiều vùng rời nhau'],
  ['Click tiêu đề hàng / cột', 'Chọn cả hàng / cột (kéo để chọn nhiều)'],
  ['Double-click ô', 'Sửa nội dung ô'],
  ['Kéo nút vuông góc vùng chọn', 'Điền nhanh (Fill) / tạo dãy số'],
  ['Double-click nút Fill', 'Điền nhanh xuống hết vùng dữ liệu bên cạnh'],
  ['Kéo viền vùng chọn', 'Di chuyển dữ liệu'],
  ['Double-click ranh giới cột', 'Tự giãn độ rộng cột theo nội dung'],
  ['Double-click ranh giới hàng', 'Tự giãn chiều cao hàng theo nội dung'],
  ['Kéo ranh giới hàng / cột', 'Chỉnh chiều cao / độ rộng'],
  ['Ctrl + Lăn chuột', 'Phóng to / thu nhỏ bảng tính'],
  ['Shift + Lăn chuột', 'Cuộn ngang'],
  ['Chuột phải vào ô', 'Menu nhanh: làm mẫu cắt gộp, tạo Pivot, giãn cột, cố định dòng...'],
  ['Chuột phải thanh trạng thái', 'Chọn số liệu hiển thị (Sum, Average, Count, Min, Max...)'],
  ['Click số liệu ở thanh trạng thái', 'Sao chép số đó'],
  ['Double-click tab sheet', 'Đổi tên sheet'],
  ['Kéo thả file .xlsx / .pdf vào cửa sổ', 'Mở file Excel · PDF invoice vải thì trích Excel · PDF bộ chứng từ thì mở Tách PDF'],
  ['Pivot: click / double-click dòng kết quả', 'Soi / nhảy tới các dòng nguồn trên hoá đơn'],
];

const KEYS: [string, string][] = [
  ['Ctrl + O', 'Mở file Excel'],
  ['Ctrl + S', 'Tải xuống file đã xử lý'],
  ['Ctrl + C / X / V', 'Sao chép / Cắt / Dán'],
  ['Ctrl + Z / Ctrl + Y', 'Hoàn tác / Làm lại'],
  ['Ctrl + B / I / U', 'In đậm / nghiêng / gạch chân'],
  ['Ctrl + F / Ctrl + H', 'Tìm kiếm / Thay thế'],
  ['Ctrl + Mũi tên', 'Nhảy tới cuối vùng dữ liệu'],
  ['Ctrl + Shift + Mũi tên', 'Chọn tới cuối vùng dữ liệu'],
  ['Ctrl + Space', 'Chọn cả cột'],
  ['Shift + Space', 'Chọn cả hàng'],
  ['Ctrl + A', 'Chọn tất cả'],
  ['Ctrl + Shift + =  (Ctrl + +)', 'Chèn hàng / cột'],
  ['Ctrl + -', 'Xoá hàng / cột đang chọn'],
  ['Alt + =', 'AutoSum: tính tổng nhanh'],
  ['Ctrl + PageUp / PageDown', 'Chuyển sheet trước / sau'],
  ['Ctrl + ;', 'Nhập ngày hiện tại'],
  ['F1 hoặc Ctrl + /', 'Mở bảng phím tắt này'],
  ['Esc', 'Đóng bảng điều khiển / menu'],
];

export default function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal" onMouseDown={(e) => e.stopPropagation()} role="dialog" aria-label="Thao tác chuột và phím tắt">
        <div className="pane-header">
          <div className="pane-title">
            <span className="pane-icon bg-slate">
              <Icon name="keyboard" size={16} />
            </span>
            <div>
              <div className="pane-title-text">Thao tác nhanh kiểu Excel</div>
              <div className="pane-subtitle">Chuột & bàn phím được hỗ trợ trên bảng tính</div>
            </div>
          </div>
          <button className="icon-btn" onClick={onClose}>
            <Icon name="close" size={16} />
          </button>
        </div>
        <div className="modal-body">
          <section>
            <h3 className="modal-section">
              <Icon name="mouse" size={14} /> Chuột
            </h3>
            <dl className="shortcut-list">
              {MOUSE.map(([k, v]) => (
                <div key={k} className="shortcut-row">
                  <dt>{k}</dt>
                  <dd>{v}</dd>
                </div>
              ))}
            </dl>
          </section>
          <section>
            <h3 className="modal-section">
              <Icon name="keyboard" size={14} /> Bàn phím
            </h3>
            <dl className="shortcut-list">
              {KEYS.map(([k, v]) => (
                <div key={k} className="shortcut-row">
                  <dt>
                    {k.split(' / ').map((part, i) => (
                      <span key={i}>
                        {i > 0 && <span className="text-slate-400"> / </span>}
                        <kbd className="kbd">{part}</kbd>
                      </span>
                    ))}
                  </dt>
                  <dd>{v}</dd>
                </div>
              ))}
            </dl>
          </section>
        </div>
      </div>
    </div>
  );
}
