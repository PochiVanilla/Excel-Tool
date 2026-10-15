// Sticker mèo trên nút ribbon (ảnh GIF do người dùng cung cấp, đặt trong public/mascots/):
// bình thường hiện ảnh đứng <kind>.png (khung hình đầu), rê chuột vào nút thì chạy <kind>.gif từ đầu
// (lib/funClick.ts gán lại src để GIF tua về khung đầu mỗi lần rê chuột)
export type MascotKind = 'open' | 'pdf' | 'split' | 'trim' | 'undo' | 'pivot' | 'download';

export default function CatMascot({ kind }: { kind: MascotKind }) {
  return (
    <span className={`mascot mascot-gif mascot-${kind}`} aria-hidden="true">
      {/* eslint-disable-next-line @next/next/no-img-element -- GIF động phải giữ nguyên file gốc */}
      <img className="m-still" src={`/mascots/${kind}.png`} alt="" draggable={false} />
      {/* eslint-disable-next-line @next/next/no-img-element -- GIF động phải giữ nguyên file gốc */}
      <img className="m-gif" src={`/mascots/${kind}.gif`} data-src={`/mascots/${kind}.gif`} alt="" draggable={false} />
    </span>
  );
}
