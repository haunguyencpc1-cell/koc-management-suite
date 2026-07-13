import "./globals.css";

export const metadata = {
  title: "KOC Management Suite — SO4 CPC1HN",
  description: "Quản lý Tổng hợp, Xếp hạng, Rời bỏ và Phân tích KOC TikTok Shop",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="vi">
      <body>{children}</body>
    </html>
  );
}
