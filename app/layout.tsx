import type { Metadata } from "next";
import Script from "next/script";
import "./globals.css";

export const metadata: Metadata = {
  title: "Giao diện Google Sheets - AI Cắt Gộp & Xuất Excel",
  description: "Công cụ xử lý Excel tốc độ cao",
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