import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "考勤数据处理平台",
  description: "GUS 考勤排班数据分析全流程处理平台",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="zh-CN">
      <body className="min-h-screen bg-gray-50 text-gray-900 antialiased">{children}</body>
    </html>
  );
}
