# 产线数据分析AI平台 — 需求说明文档

**项目**：Zillnk Efficiency Improvement Group — 产线数据分析AI平台  
**代码目录**：`D:\Gitlab\cpk_analyse_platform`  
**更新日期**：2026-04-12

---

## 一、产品概述

面向Zillnk质量工程师的桌面工具（Python tkinter GUI），用于：
1. 本地测试站数据CPK分析 + HTML报告生成
2. 故障自动分类与定位关系库建立
3. 跨站同一模块测试数据关联分析

---

## 二、测试数据目录结构（生产标准）

每台测试站本地数据根目录固定命名为 `TestResult`：

```
{station_root}/                 ← 用户在UI中配置的"测试数据文件夹"
  TestResult/
    {product_category}/         ← 产品类别，如 ORBI_B3、ORBI_B40
      {station_type}/           ← 工站类型，如 FT1、FT2、Aging、VSWR
        debug*/                 ← 调试测试目录，自动跳过
        {product_code}/         ← 产品编号/夹具版本，如 X11_X11、R1B
          {barcode}/            ← 单条码 或 双条码(BC1_BC2)
            {YYYYMMDDHHMMSS[_suffix]}/   ← 单次测试记录目录
              Test_Result_*_{BC}.xlsx         必有（pass时通常无）
              *_MEASUREMENT_Zillnk.json       B3B40/PA板类，含首次失败描述
              ate_test_log.log                主日志（核心分析源）
              ate_test_log.html               同内容HTML版本
              Failed_points_*.txt             RRU/Apricot类，失败项摘要
              file_bk/
                env_config.yml               仪器VISA地址+设备编号（结构化）
                env_comp/*.csv               射频链路插损校准数据（freq,s21）
                test_limits.yml              测试限值配置
                test_cases.yml               测试用例配置
              RU1_Log_{BC}/                  DUT通信日志（SSH/串口）
              TM1_Log/                       仪器通信日志（GPIB/USB）
              {TestItemName}/                截图目录（以测试项命名）
              screen_*.png                   测试截图
              DutInfo_Mes=*.json             DUT基本信息
```

### 特殊情况处理

| 场景 | 处理方式 |
|------|---------|
| 双条码目录 `BC1_BC2` | 取第一个条码为主条码，`barcode_full` 保存完整名称 |
| 时间戳带后缀 `20251008163411_NT` | `_NT` 等后缀为测试类型标识，解析时忽略后缀 |
| 调试目录 `debug/Debug/DEBUG` | 自动跳过，不计入分析 |
| 无 TestResult 子目录 | 降级为直接扫描条码目录（兼容拷贝数据） |
| 同一模块在多台设备测试 | 通过 `station_machine`（EQP_ID）区分，写入跨站记录表 |

---

## 三、已实现功能

### 功能一：本地测试数据分析（已完成）

**4种分析模式：**

| 模式 | 模式值 | Excel | 说明 |
|------|--------|-------|------|
| 最后一次pass cpk分析 | `latest_pass` | 必填 | 每条码取最新一次成功记录 |
| 分析全部成功数据 | `all_pass` | 可选 | 全部成功记录，不触发故障分析 |
| 分析所有数据(含失败) | `all` | 可选 | 成功+失败全量，**触发故障分析** |
| 只分析失败数据 | `fail_only` | 可选 | 仅失败记录，**触发故障分析** |

**两种故障分析模式的核心差异：**

| | `all`（分析所有数据） | `fail_only`（只分析失败数据） |
|-|----------------------|------------------------------|
| **处理对象** | 全部记录（pass + fail） | 仅fail记录（快速跳过pass） |
| **核心价值** | 跨站比对：同一模块在不同设备的测试数据关联分析 | 失败规律总结：高频失败项/设备错误模式挖掘 |
| **跨站分析** | ★★★★★ 主要场景 — pass记录入库供比对 | ★★☆ 仅fail跨站比对 |
| **规则优化** | ★★★ 为故障库提供完整上下文 | ★★★★★ 主要场景 — 输出未分类故障样本建议新规则 |
| **输出** | cross_station_barcodes 跨站列表 | fail_patterns（高频失败项TOP10、设备错误TOP5、未分类样本） |
| **典型用途** | 换站重测DUT的故障定位（FT1_1失败→FT1_2通过→站级问题） | 批量失败归因、完善故障规则库 |

> **设计意图：**  
> `all` 模式 + `fail_only` 模式**互补使用**：先用 `fail_only` 快速总结规律、完善规则库；  
> 再用 `all` 模式建立完整的含跨站比对的故障分析库。

**输出文件：**
- `cpk_report.html`：交互式CPK报告，含直方图（蓝Pass/红Fail）、通过率列
- `missing_barcodes.xlsx`：缺失条码汇总
- `fault_database.db`：故障定位关系库（SQLite）
- `analysis_log_*.txt`：完整运行日志

---

## 四、故障分析模块

### 4.1 数据库表结构

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
| fault_type | TEXT | 故障分类（规则匹配或LLM） |
| first_fail_desc | TEXT | ATE首次失败描述（MEASUREMENT JSON） |
| failed_items | JSON | 结构化失败项列表（含偏差） |
| equip_errors | JSON | 检测到的设备/通信错误 |
| instruments | JSON | 仪器配置（VISA地址映射） |
| log_excerpt | TEXT | 日志关键行摘要 |
| llm_analysis | JSON | Ollama增强分析结果 |

**`fault_rules`**：关键词规则表，支持CRUD

**`fault_stats`**：各类故障计数汇总

### 4.2 分析流程

```
ate_test_log.log
  ↓ _parse_critical_lines()       → failed_items + status（最高精度）
  ↓ _detect_equip_errors()        → equip_errors（设备/通信问题）
  
file_bk/env_config.yml
  ↓ _parse_env_config()           → instruments（VISA地址 + EQP_ID）

*_MEASUREMENT_Zillnk.json
  ↓ _read_measurement_json()      → first_fail_desc + result（B3B40专有）

Failed_points_*.txt               → 失败项文本摘要（RRU/Apricot专有）

↓ _match_rules()                  → fault_type（优先级：设备错误 > 失败项名 > 日志关键词）
↓ Ollama LLM（可选增强版）        → 提升未分类故障识别率
```

### 4.3 `CRITICAL - <string>` 日志格式（双产品通用）

```
{timestamp} - CRITICAL - <string> - {TestItemName}, data={value}({unit}), limit=[{lsl}, {usl}], result=Pass/Fail
```

示例：
```
2026-04-03 08:29:46 - CRITICAL - <string> - PA CURR CHECK ORBI BandA CH0, data=122.25(mA), limit=[90, 120], result=Fail
2025-10-11 19:08:48 - CRITICAL - <string> - ZBOOT, data=1.1.5, limit=['1.1.4', '1.1.4'], result=Fail
```

### 4.4 故障分类优先级

```
1. 检测到设备/通信错误（equip_errors）→ 优先归为"设备类故障"
2. 失败测试项名称匹配规则             → 归为对应子系统故障
3. 全日志关键词匹配规则               → 兜底分类
4. Ollama LLM分析（增强版）           → 进一步识别未分类故障
```

### 4.5 跨站分析（all 模式重点）

同一条码的模块在多台同类设备上测试时（生产中换站重测场景），系统自动：
- 记录每条记录的 `station_machine`（来自 `env_config.yml` 中的 EQP_ID，如 `FT_1`、`FT_2`）
- `all` 模式下 pass 记录也写入 DB，保留完整跨站上下文
- 通过 `get_cross_station_barcodes()` 查询跨站条码列表
- 运行完毕后在日志中输出跨站条码统计，标注含失败的条码

**跨站推断逻辑（人工/后续自动化）：**

| 观察到的跨站模式 | 推断 |
|----------------|------|
| BC001 在 FT1_1 失败（COM6错误），在 FT1_2 通过 | **站级问题**：FT1_1的COM6开关故障 |
| BC001 在 FT1_1 失败（RX增益偏低），在 FT1_2 也失败（相同项） | **DUT问题**：产品本身硬件缺陷 |
| 多个DUT在同一工站1小时内同一项失败 | **仪器健康告警**（P2待实现） |

### 4.6 fail_only 模式输出（规则库优化依据）

`_generate_fail_patterns()` 在 fail_only 模式完成后生成：

```python
{
  'total_fail':          int,           # 处理的失败记录总数
  'top_failed_items':    [(item, n)...], # 高频失败测试项 TOP20
  'top_equip_errors':    [(label, n)...], # 高频设备/通信错误 TOP10
  'top_fault_types':     [(type, n)...], # 故障分类分布
  'unclassified_samples': [str...],      # 未分类故障的 first_fail_desc（最多20条）
}
```

`unclassified_samples` 中的描述直接来自 ATE 的 `FirstFailCaseDescription`（如"测试板上PAM接触异常，请检查"），是**添加新规则的最直接线索**。

---

## 五、待实现功能

### P1（高价值）
- [ ] `env_comp/*.csv` 校准漂移检测：跨时间比较同路径补偿值，检测链路插损变化
- [ ] `test_limits.yml` 跨批次哈希比对：检测限值变更时间点与失败率突变关联

### P2（中优先级）
- [ ] 时间窗口聚类：同一工站1小时内 ≥3 个不同DUT同一项失败 → 仪器告警
- [ ] 共失败相关性矩阵：`item_correlations` 表，发现高关联失败项
- [ ] 改进LLM提示词：已结构化数据传入，提升分类准确率

### P3（后期）
- [ ] 故障分析HTML报告（帕累托图 + 跨站分析图）
- [ ] 截图内容AI分析（需Ollama视觉模型）
- [ ] `instrument_health` 仪器健康状态表

### 功能二/三（规划中）
- [ ] 功能二：深科技 MES 导出数据 CPK 分析
- [ ] 功能三：立讯 MES 导出数据 CPK 分析

---

## 六、文件结构

```
cpk_analyse_platform/
  main.py                   GUI主窗口，LocalAnalysisTab
  core/
    data_extractor.py       run_extraction(mode)，4种提取模式
    cpk_calculator.py       CPK计算，values存(barcode,value,is_pass)三元组
    html_report.py          自包含HTML报告，直方图蓝pass/红fail
    fault_db.py             SQLite持久层，表结构+CRUD+跨站查询
    fault_analyzer.py       故障分析主流程，新目录结构遍历+结构化提取
  requirements.txt
  REQUIREMENTS.md           本文件
  app_config.json           用户配置持久化（自动生成）
```

---

## 七、已知产品类型

| 产品 | 目录特征 | 条码格式 | 关键文件 |
|------|---------|---------|---------|
| B3/B40 PA板 | `TestResult/ORBI_B3(40)/FT1/X11_X11/` | 双条码 WVxxxxxx_WVxxxxxx | `*_MEASUREMENT_Zillnk.json`（含FirstFailDesc） |
| Apricot RRU | `FT1/R1B/`（拷贝结构无TestResult层） | 单条码 BFxxxxxxx | `Failed_points_*.txt`，截图目录以测试项命名 |
