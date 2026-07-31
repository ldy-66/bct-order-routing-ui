import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "报单排序配置 - BCT OMS",
  description: "BCT OMS 报单顺序与客户交易方向配置界面",
  icons: { icon: "/favicon.svg" },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="zh-CN"><body>{children}</body></html>;
}
