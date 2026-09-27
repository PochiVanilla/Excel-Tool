import type { Metadata } from "next";
import Script from "next/script";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel Tool · Xử lý hoá đơn VAT",
  description: "Cắt gộp header/footer lặp lại, Pivot (hoá đơn Polytex, chỉ may A&E hoặc bảng bất kỳ) và thao tác bảng tính như Excel",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="vi">
      <head>
        {/* CSS của Luckysheet */}
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/css/pluginsCss.css" />
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/plugins.css" />
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/css/luckysheet.css" />
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/assets/iconfont/iconfont.css" />
      </head>
      <body>
        {children}

        {/* Script của Luckysheet (Tải trước khi trang tương tác) */}
        <Script src="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/js/plugin.js" strategy="beforeInteractive" />
        <Script src="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/luckysheet.umd.js" strategy="beforeInteractive" />
        <Script src="https://cdn.jsdelivr.net/npm/luckyexcel/dist/luckyexcel.umd.js" strategy="beforeInteractive" />
      </body>
    </html>
  );
}
