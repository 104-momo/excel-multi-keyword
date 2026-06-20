"use client";

import * as React from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  FileSpreadsheet,
  RefreshCw,
  Download,
  Search,
  Trash2,
  Layers,
  Table2,
  Sparkles,
  Info,
  CheckCircle2,
  XCircle,
} from "lucide-react";
import { toast } from "sonner";

import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  Tooltip,
  TooltipContent,
  TooltipProvider,
  TooltipTrigger,
} from "@/components/ui/tooltip";

type CellValue = string | number | boolean | Date | null;
type RowData = CellValue[];

type SheetData = {
  headers: string[];
  rows: RowData[];
};

type MatchMode = "any" | "all" | "none";

/* -------------------------------------------------------------------------- */
/*  Helpers                                                                    */
/* -------------------------------------------------------------------------- */

const SEPARATOR_RE = /[ ,，\t\n\r;；|]+/;

function parseKeywords(text: string, trim: boolean): string[] {
  return text
    .split(SEPARATOR_RE)
    .map((s) => (trim ? s.trim() : s))
    .filter((s) => s.length > 0);
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function formatDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  if (d.getHours() || d.getMinutes() || d.getSeconds()) {
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return `${y}-${m}-${day} ${hh}:${mm}`;
  }
  return `${y}-${m}-${day}`;
}

function cellToString(v: CellValue): string {
  if (v == null) return "";
  if (v instanceof Date) return formatDate(v);
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return Number.isFinite(v) ? String(v) : "";
  return String(v);
}

function highlightMatch(
  text: string,
  keywords: string[],
  caseSensitive: boolean,
): React.ReactNode {
  if (!text) return "";
  const kws = keywords.filter(Boolean);
  if (kws.length === 0) return text;
  const pattern = kws.map(escapeRegExp).join("|");
  const re = new RegExp(`(${pattern})`, caseSensitive ? "g" : "gi");
  const parts = text.split(re);
  return parts.map((part, i) =>
    i % 2 === 1 ? (
      <mark
        key={i}
        className="rounded-sm bg-amber-200/80 px-0.5 text-foreground dark:bg-amber-500/40"
      >
        {part}
      </mark>
    ) : (
      <React.Fragment key={i}>{part}</React.Fragment>
    ),
  );
}

function colLetter(index: number): string {
  let n = index;
  let s = "";
  n += 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

/* -------------------------------------------------------------------------- */
/*  Sample data                                                                */
/* -------------------------------------------------------------------------- */

function buildSampleWorkbook(): XLSX.WorkBook {
  const headers = [
    "订单号",
    "客户名称",
    "城市",
    "产品",
    "数量",
    "金额",
    "日期",
  ];
  const cities = [
    "北京",
    "上海",
    "广州",
    "深圳",
    "杭州",
    "成都",
    "武汉",
    "南京",
    "西安",
    "重庆",
  ];
  const products = [
    "笔记本电脑",
    "4K显示器",
    "机械键盘",
    "无线鼠标",
    "蓝牙耳机",
    "USB扩展坞",
    "人体工学椅",
  ];
  const names = [
    "张伟",
    "王芳",
    "李娜",
    "刘洋",
    "陈静",
    "杨磊",
    "赵敏",
    "黄强",
    "周杰",
    "吴婷",
    "徐明",
    "孙丽",
  ];
  const rows: RowData[] = Array.from({ length: 48 }, (_, i) => {
    const d = new Date(2024, 0, 2);
    d.setDate(d.getDate() + i * 2);
    return [
      `DD20240${1000 + i}`,
      names[i % names.length],
      cities[i % cities.length],
      products[i % products.length],
      (i % 5) + 1,
      Math.round((i + 1) * 156.8),
      d,
    ];
  });
  const aoa: unknown[][] = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
  const wb = XLSX.utils.book_new();
  wb.properties = { ...wb.properties, creator: "Z.ai" };
  XLSX.utils.book_append_sheet(wb, ws, "订单明细");
  return wb;
}

/* -------------------------------------------------------------------------- */
/*  Component                                                                  */
/* -------------------------------------------------------------------------- */

export function FilterTool() {
  const [fileName, setFileName] = React.useState("");
  const [fileSize, setFileSize] = React.useState(0);
  const [workbook, setWorkbook] = React.useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);
  const [activeSheet, setActiveSheet] = React.useState("");
  const [hasHeader, setHasHeader] = React.useState(true);

  const [data, setData] = React.useState<SheetData | null>(null);
  const [parsing, setParsing] = React.useState(false);

  const [filterColumn, setFilterColumn] = React.useState("");
  const [keywordsText, setKeywordsText] = React.useState("");
  const [matchMode, setMatchMode] = React.useState<MatchMode>("any");
  const [caseSensitive, setCaseSensitive] = React.useState(false);
  const [trimWhitespace, setTrimWhitespace] = React.useState(true);

  const [previewLimit, setPreviewLimit] = React.useState(200);
  const [dragOver, setDragOver] = React.useState(false);

  const fileInputRef = React.useRef<HTMLInputElement>(null);

  /* ----------------------------- file handling ---------------------------- */

  const handleFile = React.useCallback(async (file: File) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (!ext || !["xlsx", "xls", "csv"].includes(ext)) {
      toast.error("仅支持 xlsx / xls / csv 格式");
      return;
    }
    setParsing(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      if (!wb.SheetNames.length) {
        toast.error("文件中没有可读的工作表");
        setParsing(false);
        return;
      }
      setWorkbook(wb);
      setFileName(file.name);
      setFileSize(file.size);
      setSheetNames(wb.SheetNames);
      setActiveSheet(wb.SheetNames[0]);
      setKeywordsText("");
      setPreviewLimit(200);
      toast.success(`已加载 ${file.name}`);
    } catch (e) {
      console.error(e);
      toast.error("文件解析失败，请检查格式是否正确");
    } finally {
      setParsing(false);
    }
  }, []);

  const loadSample = React.useCallback(() => {
    const wb = buildSampleWorkbook();
    setWorkbook(wb);
    setFileName("示例订单表.xlsx");
    setFileSize(0);
    setSheetNames(wb.SheetNames);
    setActiveSheet(wb.SheetNames[0]);
    setKeywordsText("");
    setPreviewLimit(200);
    toast.success("已加载示例数据，可试试在「城市」列输入 北京 上海");
  }, []);

  const reset = React.useCallback(() => {
    setWorkbook(null);
    setFileName("");
    setFileSize(0);
    setSheetNames([]);
    setActiveSheet("");
    setData(null);
    setFilterColumn("");
    setKeywordsText("");
    setPreviewLimit(200);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }, []);

  /* --------------------------- parse active sheet ------------------------- */

  React.useEffect(() => {
    if (!workbook || !activeSheet) {
      setData(null);
      return;
    }
    const ws = workbook.Sheets[activeSheet];
    if (!ws) {
      setData(null);
      return;
    }
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: false,
    }) as unknown[][];
    if (aoa.length === 0) {
      setData({ headers: [], rows: [] });
      return;
    }
    let headers: string[];
    let body: unknown[][];
    if (hasHeader) {
      const firstRow = (aoa[0] as unknown[]) || [];
      headers = firstRow.map((h, i) =>
        h == null || String(h).trim() === "" ? `列${i + 1}` : String(h),
      );
      body = aoa.slice(1);
    } else {
      const maxLen = aoa.reduce(
        (m, r) => Math.max(m, (r as unknown[])?.length ?? 0),
        0,
      );
      headers = Array.from({ length: maxLen }, (_, i) => `列${i + 1}`);
      body = aoa;
    }
    const width = headers.length;
    const rows: RowData[] = body.map((r) => {
      const arr = ((r as unknown[]) || []).slice(0, width) as RowData;
      while (arr.length < width) arr.push(null);
      return arr;
    });
    setData({ headers, rows });
    setFilterColumn(width > 0 ? "0" : "");
    setPreviewLimit(200);
  }, [workbook, activeSheet, hasHeader]);

  /* ------------------------------ derivation ------------------------------ */

  const keywords = React.useMemo(
    () => parseKeywords(keywordsText, trimWhitespace),
    [keywordsText, trimWhitespace],
  );

  const colIndex = React.useMemo(() => {
    const n = Number(filterColumn);
    return Number.isFinite(n) ? n : -1;
  }, [filterColumn]);

  const filteredRows = React.useMemo(() => {
    if (!data || data.rows.length === 0) return [];
    if (colIndex < 0 || colIndex >= data.headers.length) return data.rows;
    if (keywords.length === 0) return data.rows;
    const kws = caseSensitive
      ? keywords
      : keywords.map((k) => k.toLowerCase());
    return data.rows.filter((row) => {
      const raw = row[colIndex];
      const str = cellToString(raw);
      const hay = caseSensitive ? str : str.toLowerCase();
      if (matchMode === "any") return kws.some((k) => hay.includes(k));
      if (matchMode === "all") return kws.every((k) => hay.includes(k));
      return !kws.some((k) => hay.includes(k));
    });
  }, [data, colIndex, keywords, matchMode, caseSensitive]);

  const totalRows = data?.rows.length ?? 0;
  const matchedCount = filteredRows.length;
  const previewRows = React.useMemo(
    () => filteredRows.slice(0, previewLimit),
    [filteredRows, previewLimit],
  );
  const isFiltering = keywords.length > 0 && colIndex >= 0;

  /* -------------------------------- export -------------------------------- */

  const handleDownload = React.useCallback(() => {
    if (!data || filteredRows.length === 0) {
      toast.error("没有可导出的结果");
      return;
    }
    const aoa: unknown[][] = [data.headers, ...filteredRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    wb.properties = { ...wb.properties, creator: "Z.ai" };
    XLSX.utils.book_append_sheet(wb, ws, "筛选结果");
    const base = fileName
      ? fileName.replace(/\.(xlsx|xls|csv)$/i, "")
      : "表格";
    const outName = `${base}_筛选_${filteredRows.length}行.xlsx`;
    XLSX.writeFile(wb, outName);
    toast.success(`已导出 ${filteredRows.length} 行到 ${outName}`);
  }, [data, filteredRows, fileName]);

  /* -------------------------------- render -------------------------------- */

  return (
    <div className="flex flex-col gap-6">
      {/* Step 1: Upload */}
      <Card className="overflow-hidden">
        <CardHeader className="pb-4">
          <div className="flex items-center gap-2">
            <span className="flex h-6 w-6 items-center justify-center rounded-full bg-emerald-600 text-xs font-semibold text-white">
              1
            </span>
            <CardTitle className="text-base">上传表格</CardTitle>
          </div>
          <CardDescription>
            拖拽或点击上传 Excel / CSV 文件，文件仅在本地浏览器解析。
          </CardDescription>
        </CardHeader>
        <CardContent>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) handleFile(f);
            }}
          />
          {!workbook ? (
            <div
              role="button"
              tabIndex={0}
              onClick={() => fileInputRef.current?.click()}
              onKeyDown={(e) => {
                if (e.key === "Enter" || e.key === " ") {
                  e.preventDefault();
                  fileInputRef.current?.click();
                }
              }}
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={(e) => {
                e.preventDefault();
                setDragOver(false);
              }}
              onDrop={(e) => {
                e.preventDefault();
                setDragOver(false);
                const f = e.dataTransfer.files?.[0];
                if (f) handleFile(f);
              }}
              className={cn(
                "flex flex-col items-center justify-center gap-3 rounded-lg border-2 border-dashed px-6 py-12 text-center transition-colors cursor-pointer",
                "hover:border-emerald-400 hover:bg-emerald-50/50 dark:hover:bg-emerald-950/10",
                dragOver
                  ? "border-emerald-500 bg-emerald-50 dark:bg-emerald-950/20"
                  : "border-muted-foreground/25",
                parsing && "pointer-events-none opacity-60",
              )}
            >
              <div className="flex h-12 w-12 items-center justify-center rounded-full bg-emerald-100 text-emerald-700 dark:bg-emerald-950/40 dark:text-emerald-400">
                {parsing ? (
                  <RefreshCw className="h-5 w-5 animate-spin" />
                ) : (
                  <Upload className="h-5 w-5" />
                )}
              </div>
              <div>
                <p className="text-sm font-medium">
                  {parsing ? "正在解析…" : "点击或拖拽文件到此处上传"}
                </p>
                <p className="mt-1 text-xs text-muted-foreground">
                  支持 .xlsx / .xls / .csv，文件不会上传到服务器
                </p>
              </div>
              <div className="mt-2 flex items-center gap-2">
                <Button
                  type="button"
                  size="sm"
                  variant="outline"
                  onClick={(e) => {
                    e.stopPropagation();
                    loadSample();
                  }}
                >
                  <Sparkles className="h-3.5 w-3.5" />
                  加载示例数据
                </Button>
              </div>
            </div>
          ) : (
            <div className="flex flex-col gap-4">
              <div className="flex flex-wrap items-center gap-3 rounded-lg border bg-muted/40 p-3">
                <div className="flex h-10 w-10 items-center justify-center rounded-md bg-emerald-100 text-emerald-700 dark:bg-emerald-950/40 dark:text-emerald-400">
                  <FileSpreadsheet className="h-5 w-5" />
                </div>
                <div className="min-w-0 flex-1">
                  <p className="truncate text-sm font-medium" title={fileName}>
                    {fileName}
                  </p>
                  <p className="text-xs text-muted-foreground">
                    {fileSize > 0 ? `${formatBytes(fileSize)} · ` : ""}
                    {totalRows} 行 · {data?.headers.length ?? 0} 列
                    {sheetNames.length > 1
                      ? ` · ${sheetNames.length} 个工作表`
                      : ""}
                  </p>
                </div>
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => fileInputRef.current?.click()}
                >
                  <RefreshCw className="h-3.5 w-3.5" />
                  换一个
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  className="text-muted-foreground"
                  onClick={reset}
                >
                  <Trash2 className="h-3.5 w-3.5" />
                  清除
                </Button>
              </div>

              {/* Sheet + header options */}
              <div className="grid gap-3 sm:grid-cols-2">
                {sheetNames.length > 1 && (
                  <div className="flex flex-col gap-1.5">
                    <Label className="text-xs text-muted-foreground">
                      <Layers className="mr-1 inline h-3 w-3" />
                      工作表
                    </Label>
                    <Select
                      value={activeSheet}
                      onValueChange={setActiveSheet}
                    >
                      <SelectTrigger className="w-full">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        {sheetNames.map((name) => (
                          <SelectItem key={name} value={name}>
                            {name}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                )}
                <div className="flex items-end justify-between gap-3 rounded-lg border p-3">
                  <div>
                    <Label className="text-sm">首行为表头</Label>
                    <p className="mt-0.5 text-xs text-muted-foreground">
                      关闭后，所有行均视为数据，列名以 A、B、C 命名
                    </p>
                  </div>
                  <Switch
                    checked={hasHeader}
                    onCheckedChange={setHasHeader}
                  />
                </div>
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Step 2 + 3: only after data loaded */}
      {data && data.headers.length > 0 && (
        <>
          <Card>
            <CardHeader className="pb-4">
              <div className="flex items-center gap-2">
                <span className="flex h-6 w-6 items-center justify-center rounded-full bg-emerald-600 text-xs font-semibold text-white">
                  2
                </span>
                <CardTitle className="text-base">设置筛选条件</CardTitle>
              </div>
              <CardDescription>
                选定要筛选的列，输入多个关键词，结果实时更新。
              </CardDescription>
            </CardHeader>
            <CardContent className="flex flex-col gap-5">
              {/* column select */}
              <div className="flex flex-col gap-1.5">
                <Label className="flex items-center gap-1.5 text-sm">
                  <Table2 className="h-3.5 w-3.5" />
                  筛选列
                </Label>
                <Select value={filterColumn} onValueChange={setFilterColumn}>
                  <SelectTrigger className="w-full">
                    <SelectValue placeholder="请选择要筛选的列" />
                  </SelectTrigger>
                  <SelectContent className="max-h-72">
                    {data.headers.map((h, i) => (
                      <SelectItem key={i} value={String(i)}>
                        <span className="flex items-center gap-2">
                          <Badge
                            variant="outline"
                            className="font-mono text-[10px]"
                          >
                            {colLetter(i)}
                          </Badge>
                          <span className="truncate">{h}</span>
                        </span>
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              {/* keywords */}
              <div className="flex flex-col gap-1.5">
                <div className="flex items-center justify-between">
                  <Label
                    htmlFor="keywords"
                    className="flex items-center gap-1.5 text-sm"
                  >
                    <Search className="h-3.5 w-3.5" />
                    关键词
                  </Label>
                  <span className="text-xs text-muted-foreground">
                    支持空格 / 逗号 / 换行分隔
                  </span>
                </div>
                <Textarea
                  id="keywords"
                  value={keywordsText}
                  onChange={(e) => setKeywordsText(e.target.value)}
                  placeholder={
                    "输入多个关键词，用空格、逗号或换行分隔，例如：\n北京 上海 广州"
                  }
                  className="min-h-[96px] resize-y font-mono text-sm"
                />
                {keywords.length > 0 && (
                  <div className="flex flex-wrap items-center gap-1.5 pt-1">
                    <span className="text-xs text-muted-foreground">
                      已识别 {keywords.length} 个关键词：
                    </span>
                    {keywords.slice(0, 30).map((k, i) => (
                      <Badge
                        key={`${k}-${i}`}
                        variant="secondary"
                        className="gap-1 font-mono"
                      >
                        {k}
                      </Badge>
                    ))}
                    {keywords.length > 30 && (
                      <span className="text-xs text-muted-foreground">
                        …等 {keywords.length - 30} 个
                      </span>
                    )}
                  </div>
                )}
              </div>

              {/* match mode */}
              <div className="flex flex-col gap-1.5">
                <Label className="text-sm">匹配方式</Label>
                <ToggleGroup
                  type="single"
                  value={matchMode}
                  onValueChange={(v) => v && setMatchMode(v as MatchMode)}
                  className="flex w-full overflow-hidden rounded-md border"
                >
                  <ToggleGroupItem
                    value="any"
                    className="flex-1 gap-1.5 rounded-none border-0 data-[state=on]:bg-emerald-600 data-[state=on]:text-white"
                  >
                    包含任一
                  </ToggleGroupItem>
                  <ToggleGroupItem
                    value="all"
                    className="flex-1 gap-1.5 rounded-none border-0 data-[state=on]:bg-emerald-600 data-[state=on]:text-white"
                  >
                    包含全部
                  </ToggleGroupItem>
                  <ToggleGroupItem
                    value="none"
                    className="flex-1 gap-1.5 rounded-none border-0 data-[state=on]:bg-emerald-600 data-[state=on]:text-white"
                  >
                    不包含任一
                  </ToggleGroupItem>
                </ToggleGroup>
                <p className="text-xs text-muted-foreground">
                  {matchMode === "any" &&
                    "只要该列包含任意一个关键词，即保留该行"}
                  {matchMode === "all" &&
                    "该列必须同时包含全部关键词，才保留该行"}
                  {matchMode === "none" &&
                    "该列不包含任何一个关键词，才保留该行（用于排除）"}
                </p>
              </div>

              {/* options */}
              <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex items-center gap-3">
                  <Switch
                    id="case"
                    checked={caseSensitive}
                    onCheckedChange={setCaseSensitive}
                  />
                  <Label htmlFor="case" className="cursor-pointer text-sm">
                    区分大小写
                  </Label>
                </div>
                <Separator
                  orientation="vertical"
                  className="hidden h-4 sm:block"
                />
                <div className="flex items-center gap-3">
                  <Switch
                    id="trim"
                    checked={trimWhitespace}
                    onCheckedChange={setTrimWhitespace}
                  />
                  <Label htmlFor="trim" className="cursor-pointer text-sm">
                    去除关键词首尾空格
                  </Label>
                </div>
              </div>

              {/* live result summary */}
              <div className="flex flex-wrap items-center gap-2 rounded-lg border bg-muted/40 px-3 py-2 text-sm">
                {isFiltering ? (
                  <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                ) : (
                  <Info className="h-4 w-4 text-muted-foreground" />
                )}
                <span>
                  {isFiltering ? (
                    <>
                      已筛出{" "}
                      <span className="font-semibold text-emerald-700 dark:text-emerald-400">
                        {matchedCount}
                      </span>{" "}
                      行 / 共 {totalRows} 行
                    </>
                  ) : (
                    <>输入关键词后将自动筛选（当前为预览全部数据）</>
                  )}
                </span>
              </div>
            </CardContent>
          </Card>

          {/* Step 3: results */}
          <Card>
            <CardHeader className="pb-4">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <span className="flex h-6 w-6 items-center justify-center rounded-full bg-emerald-600 text-xs font-semibold text-white">
                    3
                  </span>
                  <CardTitle className="text-base">筛选结果</CardTitle>
                  <Badge variant="secondary" className="ml-1">
                    {isFiltering
                      ? `${matchedCount} / ${totalRows} 行`
                      : `${totalRows} 行（预览）`}
                  </Badge>
                </div>
                <Button
                  size="sm"
                  onClick={handleDownload}
                  disabled={filteredRows.length === 0}
                >
                  <Download className="h-3.5 w-3.5" />
                  导出为 Excel
                </Button>
              </div>
              <CardDescription>
                {isFiltering
                  ? `已筛出 ${matchedCount} 行，下方预览前 ${Math.min(
                      previewRows.length,
                      previewLimit,
                    )} 行，点击「显示更多」可继续加载。`
                  : "尚未输入关键词，当前展示全部数据。输入关键词后此处将自动筛选。"}
              </CardDescription>
            </CardHeader>
            <CardContent>
              {filteredRows.length === 0 ? (
                <div className="flex flex-col items-center justify-center gap-2 rounded-lg border border-dashed py-14 text-center">
                  <XCircle className="h-8 w-8 text-muted-foreground/50" />
                  <p className="text-sm font-medium">没有匹配的行</p>
                  <p className="text-xs text-muted-foreground">
                    请检查筛选列、关键词与匹配方式是否正确。
                  </p>
                </div>
              ) : (
                <div className="flex flex-col gap-3">
                  <div className="max-h-[28rem] overflow-auto rounded-lg border">
                    <Table>
                      <TableHeader className="sticky top-0 z-10 bg-muted/95 backdrop-blur">
                        <TableRow className="hover:bg-transparent">
                          <TableHead className="w-12 text-center text-xs text-muted-foreground">
                            #
                          </TableHead>
                          {data.headers.map((h, i) => (
                            <TableHead
                              key={i}
                              className={cn(
                                "text-xs",
                                i === colIndex &&
                                  "bg-emerald-100/70 font-semibold text-emerald-800 dark:bg-emerald-950/40 dark:text-emerald-300",
                              )}
                            >
                              <span className="flex items-center gap-1.5">
                                <span className="font-mono text-[10px] text-muted-foreground">
                                  {colLetter(i)}
                                </span>
                                <span className="truncate">{h}</span>
                                {i === colIndex && (
                                  <TooltipProvider>
                                    <Tooltip>
                                      <TooltipTrigger asChild>
                                        <span className="ml-1 inline-flex h-3.5 w-3.5 items-center justify-center rounded-full bg-emerald-600 text-[8px] font-bold text-white">
                                          √
                                        </span>
                                      </TooltipTrigger>
                                      <TooltipContent>
                                        当前筛选列
                                      </TooltipContent>
                                    </Tooltip>
                                  </TooltipProvider>
                                )}
                              </span>
                            </TableHead>
                          ))}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {previewRows.map((row, ri) => (
                          <TableRow key={ri}>
                            <TableCell className="text-center text-xs text-muted-foreground">
                              {ri + 1}
                            </TableCell>
                            {data.headers.map((_, ci) => {
                              const val = row[ci];
                              const str = cellToString(val);
                              const isFilterCol = ci === colIndex;
                              return (
                                <TableCell
                                  key={ci}
                                  className={cn(
                                    "text-sm",
                                    isFilterCol &&
                                      "bg-emerald-50/60 dark:bg-emerald-950/20",
                                  )}
                                >
                                  {isFilterCol && isFiltering
                                    ? highlightMatch(
                                        str,
                                        keywords,
                                        caseSensitive,
                                      )
                                    : str || (
                                        <span className="text-muted-foreground/40">
                                          —
                                        </span>
                                      )}
                                </TableCell>
                              );
                            })}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>

                  <div className="flex items-center justify-between text-xs text-muted-foreground">
                    <span>
                      显示 {previewRows.length} / {matchedCount} 行
                      {isFiltering ? `（共 ${totalRows} 行）` : ""}
                    </span>
                    {previewRows.length < matchedCount && (
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={() =>
                          setPreviewLimit((l) => l + 500)
                        }
                      >
                        显示更多（+500）
                      </Button>
                    )}
                  </div>
                </div>
              )}
            </CardContent>
          </Card>
        </>
      )}
    </div>
  );
}
