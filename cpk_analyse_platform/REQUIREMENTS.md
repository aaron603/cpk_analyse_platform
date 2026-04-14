# 产线数据分析AI平台 — 需求说明文档 v2.0

**项目**：Zillnk Efficiency Improvement Group — 产线数据分析AI平台  
**代码目录**：`D:\Gitlab\cpk_analyse_platform`  
**更新日期**：2026-04-14

---

## 一、产品概述

面向Zillnk质量工程师的桌面工具（Python tkinter GUI），覆盖三大数据来源：

| 功能 | 数据来源 | 状态 |
|------|---------|------|
| 功能一：本地测试站数据分析 | 本地站测试数据（xlsx/json/log） | **已实现** |
| 功能二：深科技 MES 数据分析 | 深科技 MES 导出数据 | 待实现（规划中） |
| 功能三：立讯 MES 数据分析 | 立讯 MES 导出数据 | 待实现（规划中） |

顶层窗口为 `ttk.Notebook`，三个 Tab 对应三大功能：

```
[ 本地数据分析 ] [ 深科技MES分析 ] [ 立讯MES分析 ]
```

---

## 二、GUI 布局（功能一）

### 2.1 整体布局（从上到下）

```
┌─────────────────────────────────────────────────────┐
│  Section 1：输入 / 输出配置                          │
├──────────────────────────────┬──────────────────────┤
│  Section 2：测试工站配置      │  Section 3：分析模式  │
│  （含工站合并，默认折叠）     │  （含文件类型选择）   │
├──────────────────────────────┴──────────────────────┤
│  Section 4：操作 + 进度条                            │
└─────────────────────────────────────────────────────┘
运行日志：独立弹出窗口（点击"查看运行日志"按钮打开）
```

> **关键设计**：
> - Section 2 与 Section 3 **左右并列**显示（左侧宽度弹性，右侧固定 270px）
> - 窗口高度**自适应内容**（启动时自动计算所需高度，不留空白）
> - 运行日志为**独立 Toplevel 弹窗**（760×420，深色主题，关闭后隐藏不销毁）

---

### 2.2 Section 1：输入 / 输出配置

| 字段 | 说明 |
|------|------|
| 发货Excel（可选） | 仅"最后一次pass数据"模式必填；第一列=序列号，第二列=发货产品编码 |
| 输出目录 | 所有分析结果的存放根目录 |

---

### 2.3 Section 2：测试工站配置

#### 2.3.1 工站列表

每行：`[工站类型 Entry(宽10)] [测试数据文件夹 Entry(可扩展)] [浏览] [删除]`  
底部：`[＋ 添加工站]` 按钮  
支持 ↑/↓ 键在行间切换焦点（同列对齐）。

#### 2.3.2 工站合并配置（默认折叠）

**需求背景**：同一类产品测试时可能有多台同类工站（FT1_1、FT1_2…），若测试内容完全相同需合并分析。

**入口**：Section 2 底部分隔线下方，点击 `▶ 工站合并配置` 展开。

```
展开后显示：
  说明：将指定工站类型的数据合并到目标工站类型中分析
  ┌──────────┐  → 合并到  ┌──────────┐  [✕]
  │  FT2     │           │  FT1     │
  └──────────┘           └──────────┘
  [＋ 添加合并规则]
```

**逻辑规则**：
- 每条规则：`源工站类型 → 目标工站类型`
- 合并后，源工站的数据在提取和分析阶段归入目标工站类型
- 支持多条规则（如 FT2→FT1、FT3→FT1）
- 默认无规则（功能默认关闭）
- 配置持久化到 `app_config.json`

---

### 2.4 Section 3：分析模式

#### 2.4.1 分析文件类型（单选）

```
分析文件类型：  ● xlsx   ○ json
```

- 决定 CPK 分析时读取 xlsx 还是 json 子文件夹
- 默认选 xlsx，配置持久化

#### 2.4.2 分析模式下拉（5 种）

| 模式显示名 | 内部值 | Excel 要求 | 触发故障分析 |
|-----------|--------|-----------|------------|
| 最后一次pass数据 | `latest_pass` | **必填** | 否 |
| 所选文件夹分析 | `folder_direct` | 否 | 否 |
| 全部成功数据 | `all_pass` | 否 | 否 |
| 所有数据（含失败） | `all_with_fail` | 否 | **是** |
| 仅失败数据 | `fail_only` | 否 | **是** |

鼠标悬停下拉框显示对应模式详细说明 Tooltip。

#### 2.4.3 故障分析方式（仅 all_with_fail / fail_only 时显示）

```
故障分析方式：  [基础版（规则库）  ▼]
```

| 选项 | 说明 |
|------|------|
| 基础版（规则库） | 仅使用 fault_rules 表中的关键词规则匹配 |
| 增强版（规则库+Ollama） | 规则库 + 本地 Ollama LLM 辅助分类（需安装 Ollama） |

---

### 2.5 Section 4：操作

- `[开始分析]` 按钮（运行中变为 `[停止分析]`，可随时中止）
- `[查看运行日志]` 按钮（打开/激活日志弹窗）
- 进度标签（显示当前步骤和百分比）
- 进度条（0–100%）

#### 运行日志弹窗

- 独立 Toplevel 窗口，默认隐藏，点按钮显示
- 深色主题（bg `#1e1e2e`，fg `#a8d8a8`）
- 内置"清空"按钮
- 关闭弹窗 = 隐藏（不销毁），再次打开自动滚到最新行
- 同时写入 `analysis_log_<时间戳>.txt` 文件

#### 异常弹框

- 读取发货 Excel 失败 → `messagebox.showerror` 提示文件占用/格式问题
- `_run_analysis` 主流程任意异常 → 弹框显示异常信息，引导查看运行日志

---

### 2.6 菜单栏

| 菜单 | 子项 | 功能 |
|------|------|------|
| 工具 | 加载故障关系描述文件… | 导入 YAML 格式的故障规则文件到 fault_database.db |
| 帮助 | 使用帮助 | 弹出使用说明文本窗口 |
| 帮助 | 关于 | 版本信息 |

---

## 三、分析模式详细说明

### 模式 1：最后一次pass数据（`latest_pass`）

**触发条件**：需发货Excel（第一列序列号、第二列产品编码）

**处理流程**：
```
1. 读取发货Excel → 获得条码列表
2. 遍历所有配置工站的 TestResult 目录结构
3. 对每个条码，找到所有时间戳目录，筛选最新一次全pass记录
4. 提取 xlsx 和 json 文件（保持原文件名）
5. 按工站类型分类存入输出目录：
     {out_dir}/{station_type}/xlsx/{原文件名.xlsx}
     {out_dir}/{station_type}/json/{原文件名.json}
   （若合并规则有效，源站数据写入目标站目录）
6. 生成 missing_barcodes.xlsx（未找到pass记录的条码清单）
7. 对每类工站下的 xlsx/json 文件夹执行 CPK 分析
8. 生成 HTML 报告
```

---

### 模式 2：所选文件夹分析（`folder_direct`）

**触发条件**：无需 Excel

**处理流程**：
```
1. 跳过目录遍历和提取步骤
2. 直接对每个工站配置的目录下的 xlsx（或 json）文件执行 CPK 分析
3. 生成 HTML 报告
```

**适用场景**：用户已手动整理好数据文件夹，直接分析。

---

### 模式 3：全部成功数据（`all_pass`）

**触发条件**：无需 Excel

**处理流程**：
```
1. 遍历所有配置工站的 TestResult 目录结构
2. 对每个条码的每个时间戳目录，只要该次测试全部通过就提取
   （一个条码可能被提取多次，对应多次成功测试）
3. 提取 xlsx/json 到输出目录（同模式1的分类方式）
4. 对每类工站做 CPK 分析，生成 HTML 报告
```

**与模式1的区别**：模式1每个条码只取最新一次pass，模式3取所有pass记录。

---

### 模式 4：所有数据（含失败）（`all_with_fail`）

**触发条件**：无需 Excel；**自动触发故障分析**

**核心价值**：跨站比对 — 同一模块在不同设备的测试数据关联分析

**处理流程**：
```
1. 遍历所有配置工站目录，收集全部记录（pass + fail）
2. Pass 记录写入故障库（fault_type='测试通过'），供跨站比对参考
3. Fail 记录完整提取：failed_items、equip_errors、first_fail_desc 等
4. 通过 station_machine（EQP_ID）识别跨站记录，分析站级/DUT级问题
5. 生成故障分析库（fault_database.db）
6. 生成包含失败数据的 CPK HTML 报告（蓝pass/红fail直方图）
```

**跨站推断逻辑**：

| 观察到的跨站模式 | 推断 |
|----------------|------|
| BC001 在 FT1_1 失败（COM6错误），在 FT1_2 通过 | **站级问题**：FT1_1 设备故障 |
| BC001 在 FT1_1 失败（RX偏低），在 FT1_2 也失败（同项） | **DUT问题**：产品本身硬件缺陷 |

---

### 模式 5：仅失败数据（`fail_only`）

**触发条件**：无需 Excel；**自动触发故障分析**

**核心价值**：失败规律总结，完善 all_with_fail 模式的故障分析库

**处理流程**：
```
1. 快速预筛选：只处理 fail 记录（跳过所有 pass）
2. 对每条 fail 记录提取结构化故障数据
3. 规则匹配 + 可选 LLM 分析，输出 fault_type
4. 生成故障模式统计（fail_patterns）：
     - 高频失败测试项 TOP20
     - 高频设备/通信错误 TOP10
     - 故障分类分布
     - 未分类故障样本（最多20条，是添加新规则的直接线索）
```

> **设计意图**：先用 `fail_only` 快速总结规律、完善规则库；再用 `all_with_fail` 建立完整的含跨站比对的故障分析库。

---

## 四、测试数据目录结构

### 4.1 生产标准目录层级

```
{station_root}/                      ← 用户在UI配置的工站文件夹
  TestResult/                        ← 固定目录名
    {product_category}/              ← 产品类别（ORBI_B3、ORBI_B40等）
      {station_type}/                ← 工站类型（FT1、FT2、Aging、VSWR等）
        debug*/                      ← 自动跳过（调试目录）
        {product_code}/              ← 产品编号/版本（X11_X11、R1B等）
          {barcode}/                 ← 单条码 或 双条码(BC1_BC2)
            {YYYYMMDDHHMMSS[_sfx]}/  ← 单次测试记录（时间戳目录）
              Test_Result_*_{BC}.xlsx     ← 主测试数据
              *_MEASUREMENT_Zillnk.json  ← B3B40/PA板，含首次失败描述
              ate_test_log.log            ← 主日志
              ate_test_log.html
              Failed_points_*.txt         ← RRU类，失败项摘要
              file_bk/
                env_config.yml            ← 仪器VISA地址、EQP_ID
                env_comp/*.csv            ← 链路插损校准数据（freq,s21）
                test_limits.yml           ← 测试限值配置
                test_cases.yml            ← 测试用例配置
              RU1_Log_{BC}/               ← DUT通信日志
              TM1_Log/                    ← 仪器通信日志
              {TestItemName}/             ← 截图目录（以测试项命名）
              screen_*.png
              DutInfo_Mes=*.json
```

### 4.2 目录遍历策略

| 场景 | 处理方式 |
|------|---------|
| 双条码目录 `BC1_BC2` | 取第一个条码为主条码，`barcode_full` 保存完整名称 |
| 时间戳带后缀 `20251008163411_NT` | `_NT` 等后缀为测试类型标识，解析时忽略后缀 |
| 调试目录 `debug/Debug/DEBUG` | 自动跳过，不计入分析 |
| 无 TestResult 子目录 | 降级为直接扫描条码目录（兼容拷贝数据） |
| 工站合并规则有效 | 遍历时将源站数据归入目标站类型 |

### 4.3 xlsx 与 json 格式说明

`Test_Result_*_{BC}.xlsx`：
- 每个 sheet 页对应一个测试项（sheet名 = 测试项名）
- 列：测试子项名称、测试值、上限(USL)、下限(LSL)、结果(result)、start_time、stop_time、station

`*_MEASUREMENT_Zillnk.json`：
- 与 xlsx 内容基本一致
- `DutInfo` 字段含：`ProductName`、`SiteName`（工站号）、`Result`、`FirstFailCaseDescription`

---

## 五、输出文件

```
{out_dir}/
  {station_type}/xlsx/          ← 提取的测试 xlsx 文件（原文件名）
  {station_type}/json/          ← 对应的测试 json 文件（原文件名）
  missing_barcodes.xlsx         ← 缺失/异常条码汇总
  cpk_report.html               ← 交互式 CPK 分析报告
  fault_database.db             ← 故障分析定位关系库（SQLite）
  analysis_log_<时间戳>.txt     ← 本次运行完整过程日志
```

---

## 六、HTML 报告（规划中重构）

### 6.1 顶部标题栏

```
{ProductName} 分析报告 - Zillnk
生成时间：2026-04-14 14:30:00
工站概况：FT1 × 5台  |  VSWR × 1台
```

**ProductName 获取优先级**：
1. json 文件中 `DutInfo.ProductName`
2. TestResult 下一级目录名（如 ORBI_B3）
3. 兜底：空字符串

### 6.2 工站 Tab → 测试项 Tab（两级结构）

- 第一级 Tab：每个工站类型（FT1、FT2、VSWR…）
- 第二级 Tab：每个工站 Tab 内，按 xlsx sheet 页名称分子 Tab

### 6.3 测试项统计表

| 测试子项 | 样本数 | 均值 | 标准差 | LSL | USL | 通过率 | [Cp] | [Cpl] | [Cpu] | [Cpk] |
|---------|--------|------|--------|-----|-----|--------|------|-------|-------|-------|

- 通过率色标：绿(100%) / 橙(≥95%) / 红(<95%)
- 方括号列默认**隐藏**，搜索框输入 `cpk`（不区分大小写）后显示

### 6.4 搜索栏

- 可筛选测试项名称
- 输入 `cpk` 触发：① 显示 CPK 列 ② 出现"产品数据检索"Tab
- 范围搜索：选择测试子项 + 输入数值范围，查询命中条码

### 6.5 正态分布图

- 横向全宽（min-width: 100%）
- 含均值线（实线）、LSL/USL 虚线
- 含失败数据时：蓝色=pass，红色=fail（堆叠直方图）
- 默认展示第一个测试子项，用户可切换

### 6.6 产品数据检索 Tab（隐藏，输入 cpk 后出现）

三层展开式：

| 层级 | 展示内容 |
|------|---------|
| 第一层：条码列表 | 条码 / 测试开始时间 / 测试结束时间 / pass-fail状态 |
| 第二层：测试项列表 | 测试项名 / start_time / stop_time / pass-fail |
| 第三层：测试子项明细 | 该 sheet 所有列完整数据（测量值、上下限、结果等） |

顶部检索栏：工站类型、工站号（SiteName）、状态、起止时间、条码关键字

---

## 七、CPK 分析

### 7.1 数据来源

- 按分析文件类型选择（xlsx 或 json）读取提取目录下的文件
- 每类工站独立分析（合并规则有效时，合并后视为同一工站）

### 7.2 子项过滤（跳过不做 CPK）

- 结果为字符串型（非数值）
- 所有值都相同的固定值项
- 样本数不足（< 5）

### 7.3 统计量

| 类别 | 字段 |
|------|------|
| 基础（始终显示） | 样本数(n)、均值(μ)、标准差(σ)、LSL、USL、通过率 |
| CPK（默认隐藏） | Cp、Cpl（下单边）、Cpu（上单边）、Cpk |

---

## 八、故障分析模块

### 8.1 数据库表结构（SQLite）

**`fault_rules`**（规则表，支持 CRUD + 用户文件导入）：

| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER | 自增主键 |
| keywords | TEXT | 逗号分隔关键词 |
| fault_type | TEXT | 故障分类标签 |
| suggestion | TEXT | 建议处置措施 |
| created_at / updated_at | TEXT | 时间戳 |

**`fault_records`**（每条测试记录一行）：

| 字段 | 类型 | 说明 |
|------|------|------|
| barcode | TEXT | 主条码 |
| barcode_full | TEXT | 完整条码（含双条码） |
| station | TEXT | 用户配置的工站标签 |
| station_machine | TEXT | 物理设备编号（来自 EQP_ID） |
| product_category | TEXT | 产品类别（ORBI_B3 等） |
| product_code | TEXT | 产品编号（X11_X11 等） |
| test_time | TEXT | 测试时间 |
| status | TEXT | pass / fail / unknown |
| fault_type | TEXT | 故障分类 |
| first_fail_desc | TEXT | ATE首次失败描述（MEASUREMENT JSON） |
| failed_items | JSON | 结构化失败项（含偏差） |
| equip_errors | JSON | 检测到的设备/通信错误 |
| instruments | JSON | 仪器配置（VISA地址映射） |
| log_excerpt | TEXT | 日志关键行摘要 |
| llm_analysis | JSON | Ollama增强分析结果 |

**`fault_stats`**：各类故障计数汇总

### 8.2 分析流程

```
ate_test_log.log
  ↓ _parse_critical_lines()      → failed_items + status
  ↓ _detect_equip_errors()       → equip_errors（6种设备/通信错误模式）

file_bk/env_config.yml
  ↓ _parse_env_config()          → instruments（VISA地址 + EQP_ID）

*_MEASUREMENT_Zillnk.json
  ↓ _read_measurement_json()     → first_fail_desc + result（B3B40专有）

Failed_points_*.txt              → 失败项文本摘要（RRU/Apricot专有）

↓ _match_rules()                 → fault_type（优先级：设备错误 > 失败项名 > 日志关键词）
↓ Ollama LLM（增强版可选）       → 提升未分类故障识别率
```

### 8.3 `CRITICAL - <string>` 日志格式（双产品通用）

```
{timestamp} - CRITICAL - <string> - {TestItemName}, data={value}({unit}), limit=[{lsl},{usl}], result=Pass/Fail
```

示例：
```
2026-04-03 08:29:46 - CRITICAL - <string> - PA CURR CHECK ORBI BandA CH0, data=122.25(mA), limit=[90, 120], result=Fail
2025-10-11 19:08:48 - CRITICAL - <string> - ZBOOT, data=1.1.5, limit=['1.1.4', '1.1.4'], result=Fail
```

### 8.4 故障分类优先级

```
1. 检测到设备/通信错误（equip_errors）→ 优先归为"设备类故障"
2. 失败测试项名称匹配规则             → 归为对应子系统故障
3. 全日志关键词匹配规则               → 兜底分类
4. Ollama LLM分析（增强版）           → 进一步识别未分类故障
```

### 8.5 用户可维护的故障关系描述

**入口**：菜单 `工具 → 加载故障关系描述文件…`

**文件格式**（YAML）：
```yaml
rules:
  - keywords: "PAM接触,接触异常"
    fault_type: "夹具接触不良"
    suggestion: "检查测试夹具金手指，清洁后重测"
  - keywords: "ZBOOT版本不匹配,固件版本"
    fault_type: "DUT固件版本"
    suggestion: "重新烧录匹配版本固件"
```

**导入行为**：
- 新规则追加到 `fault_rules` 表
- 相同 keywords 的规则以最新为准（更新）
- 支持多次导入（历史累积，不覆盖全表）
- 人工可通过 SQLite 工具直接查看/编辑规则表

### 8.6 fail_only 模式输出（规则库优化依据）

`_generate_fail_patterns()` 在 fail_only 模式完成后生成：

```python
{
  'total_fail':           int,            # 处理的失败记录总数
  'top_failed_items':     [(item, n)...], # 高频失败测试项 TOP20
  'top_equip_errors':     [(label, n)...],# 高频设备/通信错误 TOP10
  'top_fault_types':      [(type, n)...], # 故障分类分布
  'unclassified_samples': [str...],       # 未分类故障的 first_fail_desc（最多20条）
}
```

`unclassified_samples` 中的描述直接来自 ATE 的 `FirstFailCaseDescription`，是**添加新规则的最直接线索**。

---

## 九、已知产品类型

| 产品 | 目录特征 | 条码格式 | 关键文件 |
|------|---------|---------|---------|
| B3/B40 PA板 | `TestResult/ORBI_B3(40)/FT1/X11_X11/` | 双条码 WVxxxxxx_WVxxxxxx | `*_MEASUREMENT_Zillnk.json`（含FirstFailDesc） |
| Apricot RRU | 拷贝结构，无TestResult层，直接 `FT1/R1B/` | 单条码 BFxxxxxxx | `Failed_points_*.txt`；截图目录以测试项命名 |

---

## 十、文件结构

```
cpk_analyse_platform/
  main.py                   GUI主窗口（LocalAnalysisTab + PlaceholderTab × 2）
  core/
    data_extractor.py       run_extraction(mode)，5种提取模式；discover_barcodes()
    cpk_calculator.py       CPK计算，values存(barcode,value,is_pass)三元组
    html_report.py          自包含HTML报告（待重构：两级Tab、隐藏CPK列）
    fault_db.py             SQLite持久层（fault_rules/fault_records/fault_stats）
    fault_analyzer.py       故障分析主流程（结构化提取+规则匹配+Ollama增强）
  requirements.txt
  REQUIREMENTS.md           本文件
  app_config.json           用户配置持久化（自动生成）
```

---

## 十一、配置持久化（app_config.json）

```json
{
  "excel_path": "...",
  "out_dir": "...",
  "stations": [
    {"type": "FT1", "folder": "E:/..."}
  ],
  "cpk_mode": "最后一次pass数据",
  "fault_level": "基础版（规则库）",
  "file_type": "xlsx",
  "merge_rules": [
    {"src": "FT2", "dst": "FT1"}
  ]
}
```

---

## 十二、待实现功能

### P1（高价值）
- [ ] HTML 报告重构：工站Tab → 测试项Tab两级结构，CPK列默认隐藏，产品数据检索Tab
- [ ] `env_comp/*.csv` 校准漂移检测：跨时间比较同路径补偿值，检测链路插损变化
- [ ] `test_limits.yml` 跨批次哈希比对：检测限值变更时间点与失败率突变关联

### P2（中优先级）
- [ ] 时间窗口聚类：同一工站1小时内 ≥3 个不同DUT同一项失败 → 仪器健康告警
- [ ] 共失败关联性矩阵：`item_correlations` 表，发现高关联失败项
- [ ] `all_with_fail` 跨站推断自动化（目前输出跨站条码列表，需人工判断）

### P3（后期）
- [ ] 故障分析HTML报告（帕累托图 + 跨站分析图）
- [ ] 截图内容AI分析（需Ollama视觉模型）
- [ ] 增强版Ollama提示词优化（结构化数据传入）

### 功能二/三（规划中）
- [ ] 功能二：深科技 MES 导出数据 CPK 分析（待样本数据）
- [ ] 功能三：立讯 MES 导出数据 CPK 分析（待样本数据）
