# Worklog — Excel 多关键词筛选工具

---
Task ID: 1
Agent: main (Z.ai Code)
Task: 构建一个 WPS 表格多关键词筛选工具：选定筛选列 + 输入多个关键词（空格/逗号/换行分隔）→ 一键筛出包含行并导出。替代 WPS「高级筛选」繁琐步骤。

Work Log:
- 调用 xlsx skill 了解表格处理能力；确认本工具为纯计算任务（无 AI 需求），采用全客户端方案以保护用户数据隐私（文件不上传服务器）。
- 安装 SheetJS（xlsx@0.18.5）用于浏览器端 Excel 解析与导出。
- 新建 `src/components/theme-provider.tsx` 与 `src/components/theme-toggle.tsx`，接入 next-themes（system + 防闪烁脚本）。
- 更新 `src/app/layout.tsx`：中文元信息、ThemeProvider、sonner Toaster。
- 重写 `src/app/page.tsx`：sticky header、hero（标题+特性徽章）、FilterTool、sticky footer（min-h-screen flex flex-col + mt-auto）。
- 新建 `src/components/filter-tool.tsx`（核心）：
  - 拖拽/点击上传 xlsx/xls/csv，内置「加载示例数据」。
  - 多工作表选择、首行表头开关。
  - 筛选列下拉（带 A/B/C 列字母徽章）。
  - 关键词解析：支持空格/逗号/中文逗号/分号/中文分号/竖线/换行/制表符分隔；实时显示已识别关键词徽章。
  - 三种匹配模式：包含任一(OR) / 包含全部(AND) / 不包含任一(NOT)。
  - 选项：区分大小写、去除关键词首尾空格。
  - 实时筛选 + 筛选列高亮(绿底) + 命中关键词 `<mark>` 高亮(黄)。
  - 结果表 sticky 表头、序号列、最大高度滚动、显示更多(+500)。
  - 一键导出为 xlsx（保留原始单元格类型，日期保留为日期型）。
- 修复示例数据日期显示为序列号的问题：`aoa_to_sheet(aoa, { cellDates: true })`。
- 配置 `next.config.ts` 的 `allowedDevOrigins` 消除预览域跨域告警。
- Agent Browser 端到端验证（全部通过）。
- VLM 视觉评估确认设计专业、无样式问题。

Stage Summary:
- 核心文件：`src/components/filter-tool.tsx`（~620 行）、`src/app/page.tsx`、`src/app/layout.tsx`、`src/components/theme-provider.tsx`、`src/components/theme-toggle.tsx`、`next.config.ts`。
- 验证结果：示例加载✓ / 列选择✓ / 多关键词解析✓ / 实时筛选(15/48)✓ / 三种匹配模式(任一15、全部0、排除33)✓ / 关键词高亮(mark)✓ / 筛选列高亮✓ / 导出 xlsx(11行×7列，日期保留)✓ / 移动端无横向溢出✓ / sticky footer(diff=0)✓ / VLM 视觉评估无问题✓。
- 设计决策：全客户端处理（隐私优先，文件不上传），emerald 绿色为主色调（非蓝/靛），shadcn/ui 组件库。

---
Task ID: 2
Agent: main (Z.ai Code)
Task: 在现有筛选工具上增加编辑功能：筛选后可双击编辑单元格、增删行，并导出编辑后的表格。

Work Log:
- 重写 `src/components/filter-tool.tsx`，核心改动：
  - 新增 `EditableCell`（React.memo 优化）—— 双击进入编辑，回车/失焦提交，Esc 取消；编辑态显示 emerald ring 输入框。
  - 数据模型改为基于原始行索引：`liveFilteredIndices`（useMemo）+ `editSnapshot`（编辑时锁定的索引快照）→ `displayedIndices`。
  - `handleCellCommit`：不可变更新 data，首次编辑时锁定快照，保证编辑中的行不会因重新筛选而消失；数值列保持 number 类型（coerceValue）。
  - `handleDeleteRow`：删除行并同步重映射快照索引。
  - `handleAddRow`：新增空行并追加到快照使其立即可见。
  - `handleUndoAll`：基于 `originalDataRef`（pristine 深拷贝，structuredClone 保留 Date）一键还原。
  - `editCount` 计数 + 「已编辑 N 处」徽章 + 「撤销全部编辑」按钮。
  - 编辑锁定时显示「结果已锁定（编辑模式）· 点此刷新筛选」提示。
  - 表格新增「操作」列，删除按钮 group-hover 显示，降低视觉噪音。
  - 双导出：`handleExportFiltered`（导出筛选结果，含编辑）/ `handleExportFull`（导出完整表格，含全部编辑），均 cellDates:true 保留日期类型。
  - 修改筛选条件/切换工作表/表头开关时自动清除编辑快照（回到实时筛选）。
- 更新 `src/app/page.tsx`：hero 文案与徽章增加「双击编辑 · 增删行」。
- Agent Browser 端到端验证全部通过（见 Stage Summary）。
- VLM 视觉评估确认设计专业、编辑能力可见、无样式问题。

Stage Summary:
- 验证结果：
  - 双击编辑单元格（北京→天津）✓ 编辑持久化 ✓
  - 编辑锁定（行不消失，10行保持）✓ 「已编辑 1 处」徽章 ✓
  - 刷新筛选（天津行移出→9行，编辑保留在完整数据）✓
  - 删除行（9→8行，editCount→2）✓
  - 新增行（8→9行，空行可编辑，填入 NEW-ROW-001，editCount→4）✓
  - 撤销全部（恢复 48 行原始数据，editCount 归零，北京还原）✓
  - 导出筛选结果（10行，仅北京/上海）✓
  - 导出完整表格（48行，含天津编辑 + NEW-ROW-001）✓
  - 数值类型保持（金额 157→9999 为 int）✓ 日期保持 datetime ✓
  - 移动端无横向溢出 ✓ 无控制台错误 ✓ sticky footer ✓
- 关键文件：`src/components/filter-tool.tsx`（~900 行）、`src/app/page.tsx`。
