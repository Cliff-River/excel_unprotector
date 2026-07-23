import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
    title: "Excel Unprotector - 快速解除 Excel 文件工作表保护",
    description: "专业的 Excel 工作表保护解除工具，支持批量移除保护，操作简单安全",
    icons: {
        icon: "/favicon.ico",
    },
};

export default function RootLayout({
    children,
}: Readonly<{
    children: React.ReactNode;
}>) {
    return (
        <html lang="zh-CN">
            <body>{children}</body>
        </html>
    );
}