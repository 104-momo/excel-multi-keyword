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
  Pencil,
  Undo2,
  Plus,
  Lock,
  RefreshCcw,
  FileDown,
  Highlighter,
  X,
  Snowflake,
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

function deepClone<T>(v: T): T {
  if (typeof structuredClone === "function") return structuredClone(v);
  return JSON.parse(JSON.stringify(v)) as T;
}

/** Preserve numeric type when the original cell was a number. */
function coerceValue(newVal: string, original: CellValue): CellValue {
  if (newVal === "") return null;
  if (typeof original === "number") {
    const n = Number(newVal);
    if (newVal.trim() !== "" && Number.isFinite(n)) return n;
  }
  return newVal;
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
/*  EditableCell — double-click to edit, commit on blur/Enter                 */
/* -------------------------------------------------------------------------- */

type EditableCellProps = {
  rowIndex: number;
  colIndex: number;
  value: CellValue;
  isFilterCol: boolean;
  isFiltering: boolean;
  keywords: string[];
  caseSensitive: boolean;
  isValueHighlighted: boolean;
  isSelectionSource: boolean;
  onCommit: (rowIndex: number, colIndex: number, newVal: string) => void;
  onSelectValue: (rowIndex: number, colIndex: number, value: string) => void;
};

const EditableCell = React.memo(function EditableCell({
  rowIndex,
  colIndex,
  value,
  isFilterCol,
  isFiltering,
  keywords,
  caseSensitive,
  isValueHighlighted,
  isSelectionSource,
  onCommit,
  onSelectValue,
}: EditableCellProps) {
  const [editing, setEditing] = React.useState(false);
  const [draft, setDraft] = React.useState("");
  const inputRef = React.useRef<HTMLInputElement>(null);

  const str = cellToString(value);

  React.useEffect(() => {
    if (editing && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [editing]);

  const startEdit = React.useCallback(() => {
    setDraft(str);
    setEditing(true);
  }, [str]);

  const commit = React.useCallback(() => {
    setEditing(false);
    if (draft !== str) {
      onCommit(rowIndex, colIndex, draft);
    }
  }, [draft, str, onCommit, rowIndex, colIndex]);

  const cancel = React.useCallback(() => {
    setEditing(false);
    setDraft(str);
  }, [str]);

  const handleSelect = React.useCallback(() => {
    onSelectValue(rowIndex, colIndex, str);
  }, [onSelectValue, rowIndex, colIndex, str]);

  if (editing) {
    return (
      <input
        ref={inputRef}
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            commit();
          } else if (e.key === "Escape") {
            e.preventDefault();
            cancel();
          }
        }}
        className="w-full rounded-sm bg-background px-1 py-0.5 text-sm outline-none ring-2 ring-emerald-500"
      />
    );
  }

  return (
    <div
      onClick={handleSelect}
      onDoubleClick={startEdit}
      className={cn(
        "min-h-[1.4rem] cursor-text px-1 py-0.5 transition-colors",
        isValueHighlighted &&
          "bg-rose-100 dark:bg-rose-950/40",
        isSelectionSource &&
          "ring-2 ring-inset ring-rose-500 bg-rose-200/80 dark:bg-rose-900/50",
      )}
      title="单击高亮同列相同值，双击编辑"
    >
      {isFilterCol && isFiltering
        ? highlightMatch(str, keywords, caseSensitive)
        : str || (
            <span className="text-muted-foreground/40">—</span>
          )}
    </div>
  );
});

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

  /** Pristine copy of the parsed data, used for "undo all". */
  const originalDataRef = React.useRef<SheetData | null>(null);
  const [editCount, setEditCount] = React.useState(0);
  /** When set, the displayed row set is frozen (editing mode). */
  const [editSnapshot, setEditSnapshot] = React.useState<number[] | null>(
    null,
  );

  const [filterColumn, setFilterColumn] = React.useState("");
  const [keywordsText, setKeywordsText] = React.useState("");
  const [matchMode, setMatchMode] = React.useState<MatchMode>("any");
  const [caseSensitive, setCaseSensitive] = React.useState(false);
  const [trimWhitespace, setTrimWhitespace] = React.useState(true);

  const [previewLimit, setPreviewLimit] = React.useState(200);
  const [dragOver, setDragOver] = React.useState(false);

  /** Click-to-highlight: the cell the user clicked; matching values in the same column get highlighted. */
  const [highlight, setHighlight] = React.useState<{
    row: number;
    col: number;
    value: string;
  } | null>(null);

  /**
   * Freeze panes: -1 = nothing frozen, 0 = header only (default),
   * N (>0) = header + first N data rows frozen (sticky) while scrolling.
   */
  const [freezePanes, setFreezePanes] = React.useState(0);

  /** Column range to hide: [start, end) */
  const [hiddenStartColumn, setHiddenStartColumn] = React.useState(0);
  const [hiddenEndColumn, setHiddenEndColumn] = React.useState(0);

  /** Row range to hide: [start, end) */
  const [hiddenStartRow, setHiddenStartRow] = React.useState(0);
  const [hiddenEndRow, setHiddenEndRow] = React.useState(0);

  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const headerRef = React.useRef<HTMLTableSectionElement>(null);
  const frozenRowRefs = React.useRef<Map<number, HTMLTableRowElement>>(
    new Map(),
  );

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
    originalDataRef.current = null;
    setEditCount(0);
    setEditSnapshot(null);
    setFilterColumn("");
    setKeywordsText("");
    setPreviewLimit(200);
    setHighlight(null);
    setFreezePanes(0);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }, []);

  /* --------------------------- parse active sheet ------------------------- */

  React.useEffect(() => {
    if (!workbook || !activeSheet) {
      setData(null);
      originalDataRef.current = null;
      return;
    }
    const ws = workbook.Sheets[activeSheet];
    if (!ws) {
      setData(null);
      originalDataRef.current = null;
      return;
    }
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: false,
    }) as unknown[][];
    if (aoa.length === 0) {
      const empty: SheetData = { headers: [], rows: [] };
      setData(empty);
      originalDataRef.current = deepClone(empty);
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
    const next: SheetData = { headers, rows };
    setData(next);
    originalDataRef.current = deepClone(next);
    setEditCount(0);
    setEditSnapshot(null);
    setHighlight(null);
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

  const liveFilteredIndices = React.useMemo(() => {
    if (!data || data.rows.length === 0) return [];
    if (colIndex < 0 || colIndex >= data.headers.length || keywords.length === 0)
      return data.rows.map((_, i) => i);
    const kws = caseSensitive
      ? keywords
      : keywords.map((k) => k.toLowerCase());
    const out: number[] = [];
    data.rows.forEach((row, i) => {
      const raw = row[colIndex];
      const str = cellToString(raw);
      const hay = caseSensitive ? str : str.toLowerCase();
      const hit =
        matchMode === "any"
          ? kws.some((k) => hay.includes(k))
          : matchMode === "all"
            ? kws.every((k) => hay.includes(k))
            : !kws.some((k) => hay.includes(k));
      if (hit) out.push(i);
    });
    return out;
  }, [data, colIndex, keywords, matchMode, caseSensitive]);

  // keep a ref of the latest live indices for use in stable callbacks
  const liveFilteredIndicesRef = React.useRef<number[]>([]);
  React.useEffect(() => {
    liveFilteredIndicesRef.current = liveFilteredIndices;
  }, [liveFilteredIndices]);

  // Clear the editing snapshot whenever filter criteria or data source change
  React.useEffect(() => {
    setEditSnapshot(null);
  }, [colIndex, keywordsText, matchMode, caseSensitive, trimWhitespace, activeSheet, hasHeader]);

  const displayedIndices = editSnapshot ?? liveFilteredIndices;

  const totalRows = data?.rows.length ?? 0;
  const displayedCount = displayedIndices.length;
  const previewIndices = React.useMemo(
    () => displayedIndices.slice(0, previewLimit),
    [displayedIndices, previewLimit],
  );
  const isFiltering = keywords.length > 0 && colIndex >= 0;
  const isLocked = editSnapshot !== null;

  /* --------------------------------- edits -------------------------------- */

  const handleCellCommit = React.useCallback(
    (rowIndex: number, colIdx: number, newVal: string) => {
      setData((prev) => {
        if (!prev) return prev;
        const newRows = prev.rows.map((row, i) => {
          if (i !== rowIndex) return row;
          const newRow = [...row];
          newRow[colIdx] = coerceValue(newVal, row[colIdx]);
          return newRow;
        });
        return { ...prev, rows: newRows };
      });
      setEditCount((c) => c + 1);
      setEditSnapshot((snap) => snap ?? liveFilteredIndicesRef.current);
    },
    [],
  );

  const handleDeleteRow = React.useCallback((rowIndex: number) => {
    setData((prev) => {
      if (!prev) return prev;
      return {
        ...prev,
        rows: prev.rows.filter((_, i) => i !== rowIndex),
      };
    });
    setEditSnapshot((snap) =>
      snap
        ? snap
            .filter((idx) => idx !== rowIndex)
            .map((idx) => (idx > rowIndex ? idx - 1 : idx))
        : null,
    );
    setEditCount((c) => c + 1);
  }, []);

  const handleAddRow = React.useCallback(() => {
    setData((prev) => {
      if (!prev) return prev;
      const blank = new Array(prev.headers.length).fill(null) as RowData;
      const newIdx = prev.rows.length;
      setEditSnapshot((snap) =>
        snap ? [...snap, newIdx] : [...liveFilteredIndicesRef.current, newIdx],
      );
      return { ...prev, rows: [...prev.rows, blank] };
    });
    setEditCount((c) => c + 1);
    setPreviewLimit((l) => Math.max(l, displayedCount + 1));
  }, [displayedCount]);

  const handleUndoAll = React.useCallback(() => {
    if (!originalDataRef.current) return;
    setData(deepClone(originalDataRef.current));
    setEditSnapshot(null);
    setEditCount(0);
    toast.info("已撤销全部编辑，恢复到原始数据");
  }, []);

  const refreshFilter = React.useCallback(() => {
    setEditSnapshot(null);
  }, []);

  /* --------------------------- click-to-highlight ------------------------- */

  const handleSelectValue = React.useCallback(
    (rowIndex: number, colIdx: number, value: string) => {
      setHighlight((prev) =>
        prev && prev.row === rowIndex && prev.col === colIdx
          ? null
          : { row: rowIndex, col: colIdx, value },
      );
    },
    [],
  );

  const clearHighlight = React.useCallback(() => setHighlight(null), []);

  /** Number of visible cells matching the highlighted value in its column. */
  const highlightMatchCount = React.useMemo(() => {
    if (!highlight || !data) return 0;
    let count = 0;
    for (const idx of displayedIndices) {
      if (cellToString(data.rows[idx]?.[highlight.col]) === highlight.value) {
        count += 1;
      }
    }
    return count;
  }, [highlight, data, displayedIndices]);

  const highlightColHeader = React.useMemo(() => {
    if (!highlight || !data) return "";
    return data.headers[highlight.col] ?? "";
  }, [highlight, data]);

  /* ------------------------------ freeze panes ---------------------------- */

  /** Number of data rows currently frozen (clamped to visible rows). */
  const frozenCount = React.useMemo(() => {
    if (freezePanes <= 0) return 0;
    return Math.min(freezePanes, previewIndices.length);
  }, [freezePanes, previewIndices.length]);

  const [layout, setLayout] = React.useState<{
    header: number;
    rows: number[];
  }>({ header: 0, rows: [] });

  React.useLayoutEffect(() => {
    const measure = () => {
      const headerH =
        headerRef.current?.getBoundingClientRect().height ?? 0;
      const rows: number[] = [];
      for (let i = 0; i < frozenCount; i++) {
        const origIdx = previewIndices[i];
        const el = frozenRowRefs.current.get(origIdx);
        rows.push(el?.getBoundingClientRect().height ?? 0);
      }
      setLayout((prev) => {
        const same =
          prev.header === headerH &&
          prev.rows.length === rows.length &&
          prev.rows.every((h, i) => h === rows[i]);
        return same ? prev : { header: headerH, rows };
      });
    };
    measure();
    const ro = new ResizeObserver(measure);
    if (headerRef.current) ro.observe(headerRef.current);
    frozenRowRefs.current.forEach((el) => ro.observe(el));
    return () => ro.disconnect();
  }, [frozenCount, previewIndices, data]);

  /** Cumulative sticky top offset (px) for each frozen data row. */
  const frozenTopOffsets = React.useMemo(() => {
    const offsets: number[] = [];
    let acc = layout.header;
    for (const h of layout.rows) {
      offsets.push(acc);
      acc += h;
    }
    return offsets;
  }, [layout]);

  const headerSticky = freezePanes !== -1;

  /* --------------------------------- export ------------------------------- */

  const exportRows = React.useCallback(
    (rows: RowData[], label: string) => {
      if (!data || rows.length === 0) {
        toast.error("没有可导出的结果");
        return;
      }
      const aoa: unknown[][] = [data.headers, ...rows];
      const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
      const wb = XLSX.utils.book_new();
      wb.properties = { ...wb.properties, creator: "Z.ai" };
      XLSX.utils.book_append_sheet(wb, ws, "筛选结果");
      const base = fileName
        ? fileName.replace(/\.(xlsx|xls|csv)$/i, "")
        : "表格";
      const outName = `${base}_${label}_${rows.length}行.xlsx`;
      XLSX.writeFile(wb, outName);
      toast.success(`已导出 ${rows.length} 行到 ${outName}`);
    },
    [data, fileName],
  );

  const handleExportFiltered = React.useCallback(() => {
    const rows = displayedIndices.map((i) => data!.rows[i]);
    exportRows(rows, "筛选结果");
  }, [displayedIndices, data, exportRows]);

  const handleExportFull = React.useCallback(() => {
    if (!data) return;
    exportRows(data.rows, "完整表格");
  }, [data, exportRows]);

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
                    {editCount > 0 ? ` · 已编辑 ${editCount} 处` : ""}
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
          {/* Step 2: filter settings */}
          <Card>
            <CardHeader className="pb-4">
              <div className="flex items-center gap-2">
                <span className="flex h-6 w-6 items-center justify-center rounded-full bg-emerald-600 text-xs font-semibold text-white">
                  2
                </span>
                <CardTitle className="text-base">设置筛选条件</CardTitle>
              </div>
              <CardDescription>
                设置冻结表头、隐藏行列、筛选列与关键词，结果实时更新。
              </CardDescription>
            </CardHeader>
            <CardContent className="flex flex-col gap-5">
              {/* display settings */}
              <div className="flex flex-wrap items-start gap-4">
                <div className="flex items-center gap-2 text-sm">
                  <Snowflake className="h-4 w-4 text-sky-500" />
                  <span className="text-muted-foreground">冻结表头</span>
                  <div className="flex items-center gap-1">
                    <input
                      type="number"
                      min="1"
                      max="10"
                      value={freezePanes < 1 ? "1" : freezePanes}
                      onChange={(e) => {
                        const val = Math.max(1, Math.min(10, Number(e.target.value) || 1));
                        setFreezePanes(val);
                      }}
                      className="w-16 rounded border px-2 py-1 text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500"
                    />
                    <span className="text-xs text-muted-foreground">行（最多10行）</span>
                  </div>
                  {freezePanes > 0 && (
                    <span className="text-xs text-muted-foreground">
                      （滚动时前 {frozenCount} 行固定）
                    </span>
                  )}
                </div>
                <div className="flex items-center gap-2 text-sm">
                  <span className="text-muted-foreground">隐藏列</span>
                  <span className="text-muted-foreground">从</span>
                  <input
                    type="number"
                    min="1"
                    max={data.headers.length}
                    value={hiddenStartColumn + 1}
                    onChange={(e) => {
                      const val = Math.max(1, Math.min(data.headers.length, Number(e.target.value) || 1));
                      setHiddenStartColumn(val - 1);
                    }}
                    className="w-16 rounded border px-2 py-1 text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <span className="text-muted-foreground">到</span>
                  <input
                    type="number"
                    min="1"
                    max={data.headers.length}
                    value={hiddenEndColumn}
                    onChange={(e) => {
                      const val = Math.max(1, Math.min(data.headers.length, Number(e.target.value) || 1));
                      setHiddenEndColumn(val);
                    }}
                    className="w-16 rounded border px-2 py-1 text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <span className="text-muted-foreground">列</span>
                  {hiddenStartColumn < hiddenEndColumn && (
                    <span className="text-xs text-muted-foreground">
                      （隐藏第 {hiddenStartColumn + 1}-{hiddenEndColumn} 列）
                    </span>
                  )}
                </div>
                <div className="flex items-center gap-2 text-sm">
                  <span className="text-muted-foreground">隐藏行</span>
                  <span className="text-muted-foreground">从</span>
                  <input
                    type="number"
                    min="1"
                    max={totalRows}
                    value={hiddenStartRow + 1}
                    onChange={(e) => {
                      const val = Math.max(1, Math.min(totalRows, Number(e.target.value) || 1));
                      setHiddenStartRow(val - 1);
                    }}
                    className="w-16 rounded border px-2 py-1 text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <span className="text-muted-foreground">到</span>
                  <input
                    type="number"
                    min="1"
                    max={totalRows}
                    value={hiddenEndRow}
                    onChange={(e) => {
                      const val = Math.max(1, Math.min(totalRows, Number(e.target.value) || 1));
                      setHiddenEndRow(val);
                    }}
                    className="w-16 rounded border px-2 py-1 text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <span className="text-muted-foreground">行</span>
                  {hiddenStartRow < hiddenEndRow && (
                    <span className="text-xs text-muted-foreground">
                      （隐藏第 {hiddenStartRow + 1}-{hiddenEndRow} 行）
                    </span>
                  )}
                </div>
              </div>

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
                    isLocked ? (
                      <>
                        结果已锁定（编辑模式）· 当前{" "}
                        <span className="font-semibold text-emerald-700 dark:text-emerald-400">
                          {displayedCount}
                        </span>{" "}
                        行 / 共 {totalRows} 行
                      </>
                    ) : (
                      <>
                        已筛出{" "}
                        <span className="font-semibold text-emerald-700 dark:text-emerald-400">
                          {displayedCount}
                        </span>{" "}
                        行 / 共 {totalRows} 行
                      </>
                    )
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
                <div className="flex flex-wrap items-center gap-2">
                  <span className="flex h-6 w-6 items-center justify-center rounded-full bg-emerald-600 text-xs font-semibold text-white">
                    3
                  </span>
                  <CardTitle className="text-base">筛选与编辑</CardTitle>
                  <Badge variant="secondary" className="ml-1">
                    {isFiltering
                      ? isLocked
                        ? `已锁定 ${displayedCount} / ${totalRows} 行`
                        : `${displayedCount} / ${totalRows} 行`
                      : `${totalRows} 行`}
                  </Badge>
                  {editCount > 0 && (
                    <Badge
                      variant="outline"
                      className="gap-1 border-amber-400 text-amber-700 dark:text-amber-400"
                    >
                      <Pencil className="h-3 w-3" />
                      已编辑 {editCount} 处
                    </Badge>
                  )}
                  {highlight && (
                    <Badge
                      variant="outline"
                      className="gap-1 border-rose-400 bg-rose-50 text-rose-700 dark:bg-rose-950/30 dark:text-rose-300"
                    >
                      <Highlighter className="h-3 w-3" />
                      高亮「{highlightColHeader} = {highlight.value || "(空)"}」
                      <span className="text-rose-500/80">
                        {highlightMatchCount} 处
                      </span>
                      <button
                        type="button"
                        aria-label="清除高亮"
                        onClick={clearHighlight}
                        className="ml-0.5 inline-flex items-center justify-center rounded-sm hover:bg-rose-200/70 dark:hover:bg-rose-900/50"
                      >
                        <X className="h-3 w-3" />
                      </button>
                    </Badge>
                  )}
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  {editCount > 0 && (
                    <Button
                      size="sm"
                      variant="ghost"
                      className="text-muted-foreground"
                      onClick={handleUndoAll}
                    >
                      <Undo2 className="h-3.5 w-3.5" />
                      撤销全部编辑
                    </Button>
                  )}
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={handleExportFull}
                    disabled={totalRows === 0}
                  >
                    <FileDown className="h-3.5 w-3.5" />
                    导出完整表格
                  </Button>
                  <Button
                    size="sm"
                    onClick={handleExportFiltered}
                    disabled={displayedCount === 0}
                  >
                    <Download className="h-3.5 w-3.5" />
                    导出筛选结果
                  </Button>
                </div>
              </div>
              <CardDescription className="flex flex-wrap items-center gap-x-3 gap-y-1">
                <span className="inline-flex items-center gap-1">
                  <Pencil className="h-3.5 w-3.5 text-emerald-600" />
                  双击编辑，回车保存、Esc 取消
                </span>
                <span className="inline-flex items-center gap-1">
                  <Highlighter className="h-3.5 w-3.5 text-rose-500" />
                  单击单元格高亮同列相同值，再次单击取消
                </span>
                {isLocked && (
                  <span className="inline-flex items-center gap-1 text-amber-700 dark:text-amber-400">
                    <Lock className="h-3.5 w-3.5" />
                    编辑期间结果已锁定，
                    <button
                      type="button"
                      className="underline underline-offset-2 hover:opacity-80"
                      onClick={refreshFilter}
                    >
                      点此刷新筛选
                    </button>
                  </span>
                )}
              </CardDescription>
            </CardHeader>
            <CardContent>
              {displayedCount === 0 ? (
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
                    <table className="w-full caption-bottom text-sm">
                      <TableHeader
                        ref={headerRef}
                        className={cn(
                          "bg-muted/95 backdrop-blur",
                          headerSticky && "sticky top-0 z-20",
                        )}
                      >
                        <TableRow className="hover:bg-transparent">
                          <TableHead className="w-12 text-center text-xs text-muted-foreground">
                            #
                          </TableHead>
                          {data.headers.map((h, i) => {
                            if (i >= hiddenStartColumn && i < hiddenEndColumn) return null;
                            return (
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
                            );
                          })}
                          <TableHead className="w-10 text-center text-xs text-muted-foreground">
                            操作
                          </TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {previewIndices.map((origIdx, displayRow) => {
                          if (displayRow >= hiddenStartRow && displayRow < hiddenEndRow) return null;
                          const isFrozen = displayRow < frozenCount;
                          return (
                          <TableRow
                            key={origIdx}
                            ref={
                              isFrozen
                                ? (el) => {
                                    if (el)
                                      frozenRowRefs.current.set(origIdx, el);
                                    else
                                      frozenRowRefs.current.delete(origIdx);
                                  }
                                : undefined
                            }
                            className={cn(
                              "group",
                              isFrozen && "bg-background",
                            )}
                            style={
                              isFrozen
                                ? {
                                    position: "sticky",
                                    top: frozenTopOffsets[displayRow] ?? 0,
                                    zIndex: 10,
                                  }
                                : undefined
                            }
                          >
                            <TableCell className="text-center text-xs text-muted-foreground">
                              {displayRow + 1}
                            </TableCell>
                            {data.headers.map((_, ci) => {
                              if (ci >= hiddenStartColumn && ci < hiddenEndColumn) return null;
                              const val = data.rows[origIdx]?.[ci];
                              const isFilterCol = ci === colIndex;
                              const cellStr = cellToString(val);
                              const isValueHighlighted =
                                highlight !== null &&
                                highlight.col === ci &&
                                highlight.value === cellStr;
                              const isSelectionSource =
                                highlight !== null &&
                                highlight.row === origIdx &&
                                highlight.col === ci;
                              return (
                                <TableCell
                                  key={ci}
                                  className={cn(
                                    "p-0",
                                    isFilterCol &&
                                      "bg-emerald-50/60 dark:bg-emerald-950/20",
                                  )}
                                >
                                  <EditableCell
                                    rowIndex={origIdx}
                                    colIndex={ci}
                                    value={val}
                                    isFilterCol={isFilterCol}
                                    isFiltering={isFiltering}
                                    keywords={keywords}
                                    caseSensitive={caseSensitive}
                                    isValueHighlighted={isValueHighlighted}
                                    isSelectionSource={isSelectionSource}
                                    onCommit={handleCellCommit}
                                    onSelectValue={handleSelectValue}
                                  />
                                </TableCell>
                              );
                            })}
                            <TableCell className="text-center">
                              <button
                                type="button"
                                aria-label="删除该行"
                                title="删除该行"
                                onClick={() => handleDeleteRow(origIdx)}
                                className="inline-flex h-6 w-6 items-center justify-center rounded text-muted-foreground opacity-0 transition-all hover:bg-destructive/10 hover:text-destructive group-hover:opacity-100"
                              >
                                <Trash2 className="h-3.5 w-3.5" />
                              </button>
                            </TableCell>
                          </TableRow>
                          );
                        })}
                      </TableBody>
                    </table>
                  </div>

                  <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-muted-foreground">
                    <span>
                      显示 {previewIndices.length} / {displayedCount} 行
                      {isFiltering ? `（共 ${totalRows} 行）` : ""}
                    </span>
                    <div className="flex items-center gap-2">
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={handleAddRow}
                      >
                        <Plus className="h-3.5 w-3.5" />
                        新增一行
                      </Button>
                      {previewIndices.length < displayedCount && (
                        <Button
                          size="sm"
                          variant="outline"
                          onClick={() => setPreviewLimit((l) => l + 500)}
                        >
                          <RefreshCcw className="h-3.5 w-3.5" />
                          显示更多（+500）
                        </Button>
                      )}
                    </div>
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
