import type { Metadata } from "next";
import { headers } from "next/headers";
import "./globals.css";

export async function generateMetadata(): Promise<Metadata> {
  const requestHeaders = await headers();
  const host = requestHeaders.get("x-forwarded-host") ?? requestHeaders.get("host") ?? "localhost:3000";
  const protocol = requestHeaders.get("x-forwarded-proto") ?? "https";

  return {
    metadataBase: new URL(`${protocol}://${host}`),
    title: "万相台报表工作流看板",
    description: "展示万相台月报、半月报与周报从数据投喂到 Excel 交付的完整工作流。",
    openGraph: {
      title: "万相台报表工作流看板",
      description: "万相台报表从数据投喂、周期计算、模板填充到 Excel 交付的工作流。",
      type: "website",
      images: [{ url: "/og.png", width: 1600, height: 1000, alt: "万相台报表工作流看板" }],
    },
    twitter: {
      card: "summary_large_image",
      title: "万相台报表工作流看板",
      description: "万相台报表从数据投喂到 Excel 交付的工作流。",
      images: ["/og.png"],
    },
  };
}

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
