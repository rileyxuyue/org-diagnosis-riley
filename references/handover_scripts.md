# 交接文档：脚本清单与工作流

本文档面向接手人，说明本skill 内 35 个 Python 文件各自的职责、上下游关系，以及已知的坑。

---

## 一、总览：三条链路

```
链路A（主链路，产出最终报告）
  extract_open_feedback → step1_extract_bg → generate_source_md_with_llm
    → generate_html_report → html2md → refresh_section_13
  由 run_report_pipeline.py 统一串起

链路B（敬满 Word 报告，独立产物）
  【敬满】/report_tool_final/generate_report.py
    → data_loader + module1~5 + report_builder.js（Node）
  产出 docx，不进入链路A

链路C（一次性修补，按需单独跑）
  update_* / patch_* / fix_* 系列
  只改已生成 HTML/MD 的局部片段
```

**交接重点**：日常只需要跑链路A。链路B 是独立的敬满 docx 工具，链路C 是历史遗留的补丁脚本。

---

## 二、链路A：主链路（7 个文件）

### 1. `run_report_pipeline.py`（总控入口）

**作用**：流水线调度器，用 `subprocess` 按顺序调其余脚本。

**参数**：
- `--bgs`：BG 列表，默认 `all`（10 个 BG）
- `--steps`：可选`extract-open-feedback` / `step1-extract` / `generate-source-md` / `build-html` / `convert-md` / `refresh-13`，默认后三个

**常用命令**：
```bash
# 已有源 md，只重建报告（最常用、最稳定）
python3 run_report_pipeline.py --bgs CSIG --steps build-html convert-md refresh-13

# 全链路重建（含 LLM 生成源 md）
python3 run_report_pipeline.py --bgs CSIG --steps extract-open-feedback step1-extract generate-source-md build-html convert-md refresh-13
```

### 2. `【全面反馈】/extract_open_feedback.py`（Step 0：数据拼接）

**作用**：把组织诊断表、全面反馈表、BP观察表按人名join 起来。

- **输入**：`【组织诊断结果】/组织诊断2025全年最终版.xlsx`、`【全面反馈】/全面反馈25H2v2.xlsx`、`{BG}部门负责人-1BP观察.xlsx`
- **输出**：`【全面反馈】/全面反馈开放题{BG}.xlsx`
- **实现特点**：不用 pandas 读 xlsx，而是用 `zipfile` + `ElementTree` 直接解析 XML（规避 openpyxl 兼容问题）
- **调用方式**：`python3 extract_open_feedback.py CSIG TEG`，或不带参数进交互模式

### 3. `【全面反馈】/step1_extract_bg.py`（Step 1：结构化抽取）

**作用**：把上一步的宽表，按 `<闪光点>`/`<更多期待>`/`<bp观察>` 标签解析成每位管理者一条条原文。

- **输入**：`全面反馈开放题{BG}.xlsx`（找不到时**自动回退** `全面反馈开放题_模板.xlsx`）
- **输出**：`【全面反馈】/output/{bg}/step1_output.xlsx`，含 4 个 sheet：`部门概览` / `管理者小计` / `原文清单` / `异常日志`
- **会调 LLM**：仅当 BP 文本格式不规范时，用模型判断正负向极性
- **调用方式**：`python3 step1_extract_bg.py --bg CSIG`

> `step1_extract_csig.py` / `step1_extract_s3.py` 是这个脚本的**前身**（BG 硬编码版），保留仅为追溯历史，新流程不用。

### 4. `【全面反馈】/generate_source_md_with_llm.py`（Step 2：LLM 生成 1.3 源文）

**作用**：把 `step1_output.xlsx` 按部门喂给大模型，产出结构化 Markdown。

- **输入**：`step1_output.xlsx` + `output/报告生成_Prompt2.md`（提示词）+ `output/examples/14_示例部门B.md`（格式样例）
- **输出**：`【全面反馈】/output/{bg}/{dept_idx}_{dept_name}.md`
- **模型配置**：环境变量 `ORG_DIAG_LLM_URL` / `ORG_DIAG_LLM_MODEL` / `ORG_DIAG_LLM_API_KEY`，`temperature=0`
- **调用方式**：`python3 generate_source_md_with_llm.py --bg CSIG`

> **想复刻既有结果就跳过这步**，直接把现成的 `output/{bg}/*.md` 放进去。模型输出无法保证逐字一致。

### 5. `generate_html_report.py`（Step 3：主生成器，304 KB）

**作用**：整个 skill 的核心，读所有 xlsx，渲染出完整 HTML 报告。

- **输出**：`报告/{bg}/html/*.html`
- **调用方式**：`python3 generate_html_report.py --batch CSIG`（或不带参数进交互模式）

**内部结构**（其他脚本会 import 这些类，改动需谨慎）：

| 类/ 函数 | 行号 | 职责 |
|---|---|---|
| `ExcelReader` | 20 | Excel 读取底层封装 |
| `ResignationAnalyzer` | 219 | 离职原因分析，**会调 LLM** |
| `BPObservationLoader` | 669 | BP 观察数据（点赞/提醒关注标签） |
| `JianGangLoader` | 767 | 兼岗数据 |
| `OpenFeedbackLoader` | 827 | 全面反馈干部侧 / 组织侧 |
| `generate_report_data` | 2307 | 汇总单份报告所需全部数据 |
| `JingmanOpenLoader` | 2723 | 敬满开放题 |
| `WordCloudLoader` | 3202 | 词云数据 |
| `build_org_chart_data` | 3339 | 组织架构图 |
| `YidongDataLoader` | 3521 | 异动数据 |
| `JingmanDataLoader` | 3804 | 敬满主数据 |
| `JingmanDetailLoader` | 4041 | 敬满逐题细分 |
| `generate_html_report` | 4509 | HTML 渲染主函数 |
| `batch_generate` / `main` | 6846 / 7003 | 批量与命令行入口 |

**模板文件回退机制**：5 处数据加载都是「先找 `{BG}_xxx.xlsx`，找不到就用 `模板_xxx.xlsx`」，所以只放通用模板也能跑通。

### 6. `html2md.py`（Step 4：格式转换）

**作用**：用 `markdownify` 把 HTML 转 Markdown，转换前先剥掉 `<style>`/`<script>`/`<head>`/工具栏/词云/footer。

- **输入** `报告/{bg}/html/*.html` → **输出** `报告/{bg}/md/*.md`
- **调用方式**：`python3 html2md.py --bgs CSIG`

### 7. `refresh_section_13.py`（Step 5：只刷 1.3）

**作用**：拿 `output/{bg}/*.md` 覆盖报告里的「1.3 开放题总结」，**其余章节一律不动**。

- **同时改** HTML 和 MD 两份产物
- **调用方式**：`python3 refresh_section_13.py --bgs CSIG`

> 这是 `update_csig_section_1_3.py` 的通用化版本。要改 1.3 只用这个，别手改HTML。

---

## 三、链路B：敬满 Word 报告（9 个文件，独立）

入口 `【敬满】/report_tool_final/generate_report.py`，产出 **docx**（不是 HTML），走 Python 算数据 → JSON → Node.js 渲染。

```
generate_report.py（主入口）
├── config.py                路径配置
├── data_loader.py               读 3 个 Excel、定义 47 道子题、格式化函数
├── module1_core_table.py        核心维度表（敬业度/满意度 各4列）
├── module2_risk_table.py        风险区间表（BG末20%/末10%）
├── module3_chart.py             47 题击败率双向柱状图 → PNG
├── module4_bottom_table.py      排名靠后明细（击败率≤10%）
├── module5_subdivision_table.py 细分项大表（4 分组，排名/增幅差值>10）
└── report_builder.js            Node.js 渲染 docx
```

**调用方式**：
```bash
cd 【敬满】/report_tool_final
python3 generate_report.py "BG/部门名"   # 单个部门
python3 generate_report.py --all          # 全部
```

**依赖**：Node.js（渲染 docx）、matplotlib（画图）、中文字体。

### 敬满数据预处理（3 个）

| 脚本 | 作用 | 备注 |
|---|---|---|
| `extract_jingman_open_answers.py` | 从各部门 Excel 汇总成 `{BG}_敬满开放题汇总.xlsx` / `关键词汇总.xlsx` | 需要 `【敬满】/开放题/` 源目录 |
| `analyze_sentiment_v2.py` | 关键词汇总 → 关键词分析，逐词判情感，**调 LLM**，带本地缓存 | 每部门最多取词频前 200 |
| `test_s3_sentiment.py` | 上者的 S3 单BG 测试版 |仅调试用 |

---

## 四、链路C：修补脚本（11 个，按需单跑）

这些都是**历史上为修特定问题写的**，只改已生成产物的局部片段。日常不用，遇到对应问题再翻出来。

### update系列（内容更新）

| 脚本 | 作用 | 依赖 |
|---|---|---|
| `update_wordcloud.py` | 只替换 HTML 里的词云 section | 被下面两个 import |
| `update_jingman_open.py` | 只更新 2.3.3 敬满开放题（通用版，支持 s3/cdg/ieg） | `JingmanOpenLoader` + `update_wordcloud` |
| `update_s3_jingman_open.py` | 同上，S3 专用版 | `JingmanOpenLoader`+`WordCloudLoader`+`update_wordcloud` |
| `update_section_13.py` | 单个文件的 1.3 更新（传html 路径 + md 路径） | `OpenFeedbackLoader` |
| `update_csig_section_1_3.py` | CSIG 专用 1.3 更新（**已被 `refresh_section_13.py` 取代**） | — |

### patch 系列（补内容）

| 脚本 | 作用 |
|---|---|
| `patch_232_subdivision.py` | 给 CDG/S3 补上 2.3.2 细分项表|
| `patch_s3_233.py` | 只更新 S3 的 2.3.3（敬满开放题+词云） |
| `patch_small_dept_wordcloud.py` | ≤20 人小部门词频门槛降到 ≥1，重算词云 |

### fix 系列（修样式/渲染 bug）

| 脚本 | 作用 |
|---|---|
| `fix_cadre_dashed_box.py` | 把跑到虚线框外的 summary 段落移回框内 |
| `fix_jingman_open_low_count.py` | 删掉累计人数 <2 的维度并重排编号 |
| `fix_org_chart_batch.py` | 批量刷新组织架构图（增量替换） |
| `fix_wordcloud_position.py` | 词云从 footer 内移到footer 前 + 自适应高度 |

---

## 五、其他

| 脚本 | 作用 |
|---|---|
| `scripts/prepare_workspace.py` | 把`assets/project_template/` 复制到目标目录，用法`python3 scripts/prepare_workspace.py <目标目录>` |
| `【全面反馈】/generate_bp_observation_template.py` | 从组织诊断表提人名，为每个 BG 生成空的 BP 观察填写模板 |
| `desensitize_report.py` | 报告脱敏工具，**映射表已清空**，用前需自己填 `_person_list` / `ORG_MAP` / `BIZ_MAP` |

---

## 六、已知问题（接手必读）

以下都是**已存在但未修**的问题，按优先级排列：

### 1. 链路C 脚本的文件名与当前模板不匹配（高）

本 skill 把按 BG 拆分的 xlsx 合并成了单个 `模板_*.xlsx`，但链路C 脚本里仍写死了 BG 名：

- `update_wordcloud.py`、`update_jingman_open.py`、`update_s3_jingman_open.py` 的 `BG_CONFIG` 引用 `CSIG_敬满开放题关键词分析.xlsx` 等
- `patch_small_dept_wordcloud.py` 的 `FILE_PAIRS` 同样问题
- `step1_extract_csig.py` 引用 `全面反馈开放题CSIG.xlsx`

**影响**：这些脚本直接跑会报文件不存在。**主链路（链路A）不受影响**，因为 `generate_html_report.py` 有回退机制。

**修法**：要么把模板复制成对应 BG 名，要么给这些脚本也加回退逻辑。

### 2. `report_tool_final/config.py` 路径对不上（高）

`config.py` 里写的是：
```python
DEPT_FILE = "data/数据结构样本.xlsx"
BG_FILE   = "data/BG相关数据样本.xlsx"
```
但 `data/` 下实际是 `全量敬满数据.xlsx` 和 `BG相关数据.xlsx`。**跑链路B 前必须先改 `config.py`**。

### 3. `module3_chart.py` 字体路径是 Linux 的（中）

```python
font_path = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
```
macOS 上会直接抛异常。需改成本机中文字体（如 `/System/Library/Fonts/PingFang.ttc`）。

### 4. `extract_jingman_open_answers.py` 缺源目录（中）

它读 `【敬满】/开放题/`，期望文件名形如 `{序号}_2025_{BG}_{部门}_敬满_整合版.xlsx`。模板工程里没有这个目录，需自建。

### 5. 链路C 脚本大量硬编码日期后缀（低）

如 `update_wordcloud.py` 里的 `_20260409.html`、`_20260327_180000.html`。换一批报告就得改。

---

## 七、环境与依赖

```bash
python3 -m pip install -r requirements.txt
```

**LLM 配置**（链路A 的 step2、离职分析、敬满情感分析都用）：
```bash
export ORG_DIAG_LLM_URL="https://your-endpoint/v1/chat/completions"
export ORG_DIAG_LLM_MODEL="gpt-4.1-mini"
export ORG_DIAG_LLM_API_KEY="sk-xxx"
```
不设则回退到本地 LM Studio 默认地址。

**额外依赖**：链路B 需要 Node.js 和中文字体。

---

## 八、接手建议路径

1. 先只跑链路A 的后三步，确认能出 HTML/MD：
   ```bash
   python3 run_report_pipeline.py --bgs CSIG --steps build-html convert-md refresh-13
   ```
2. 对照 `assets_examples/final_output/示例/` 检查产物结构
3. 需要重建 1.3 时再配 LLM，跑 `step1-extract` + `generate-source-md`
4. 链路B、链路C 按需启用，启用前先看上面「已知问题」
