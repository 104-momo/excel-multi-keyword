import {
  Filter,
  ShieldCheck,
  Zap,
  FileSpreadsheet,
  Github,
} from "lucide-react";

import { Badge } from "@/components/ui/badge";
import { FilterTool } from "@/components/filter-tool";
import { ThemeToggle } from "@/components/theme-toggle";

export default function Home() {
  return (
    <div className="flex min-h-screen flex-col bg-background">
      <header className="sticky top-0 z-50 w-full border-b bg-background/80 backdrop-blur supports-[backdrop-filter]:bg-background/60">
        <div className="mx-auto flex h-14 max-w-6xl items-center gap-3 px-4">
          <div className="flex items-center gap-2 font-semibold">
            <span className="flex h-7 w-7 items-center justify-center rounded-md bg-emerald-600 text-white shadow-sm">
              <Filter className="h-4 w-4" />
            </span>
            <span className="tracking-tight">表格多关键词筛选</span>
          </div>
          <div className="ml-auto flex items-center gap-1">
            <a
              href="https://github.com/SheetJS/sheetjs"
              target="_blank"
              rel="noreferrer"
              className="hidden text-sm text-muted-foreground transition-colors hover:text-foreground sm:inline-flex"
            >
              基于 SheetJS
            </a>
            <ThemeToggle />
          </div>
        </div>
      </header>

      <main className="mx-auto w-full max-w-6xl flex-1 px-4 py-8 sm:py-12">
        <section className="mb-8 sm:mb-10">
          <h1 className="text-3xl font-bold tracking-tight sm:text-4xl">
            Excel 多关键词筛选工具
          </h1>
          <p className="mt-3 max-w-2xl text-base text-muted-foreground sm:text-lg">
            选定筛选列，输入多个关键词（空格、逗号、换行均可分隔），一键筛出所有包含行。
            替代 WPS「高级筛选」繁琐的多步骤操作，数据全程在浏览器本地处理。
          </p>
          <div className="mt-5 flex flex-wrap gap-2">
            <Badge variant="secondary" className="gap-1.5 py-1.5">
              <ShieldCheck className="h-3.5 w-3.5 text-emerald-600" />
              本地处理 · 不上传服务器
            </Badge>
            <Badge variant="secondary" className="gap-1.5 py-1.5">
              <Zap className="h-3.5 w-3.5 text-amber-500" />
              输入即筛选
            </Badge>
            <Badge variant="secondary" className="gap-1.5 py-1.5">
              <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600" />
              支持 xlsx / xls / csv
            </Badge>
          </div>
        </section>

        <FilterTool />
      </main>

      <footer className="mt-auto border-t bg-muted/30">
        <div className="mx-auto flex max-w-6xl flex-col gap-2 px-4 py-6 text-sm text-muted-foreground sm:flex-row sm:items-center sm:justify-between">
          <span>
            数据仅在浏览器本地处理，关闭页面即清除，不会上传到任何服务器。
          </span>
          <span className="flex items-center gap-1.5">
            <Github className="h-3.5 w-3.5" />
            Powered by Next.js · SheetJS
          </span>
        </div>
      </footer>
    </div>
  );
}
