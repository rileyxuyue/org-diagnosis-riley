# 架构文档：三条链路与 35 个 Python 文件

本文档从**架构视角**说明 skill 内35 个 Python 文件如何组织成三条链路，以及数据在链路中如何流动。

> 需要查单个脚本的详细职责和已知问题，见 `handover_scripts.md`。

---

## 一、整体架构

```
┌──────────────────────────────────────────────────────────────┐
│输入层（xlsx 模板）                      │
│  【组织诊断结果】 【组织架构信息】 【全面反馈】 【异动】 【敬满】      │
└───────────┬──────────────────────────────────┬───────────────┘
            │                                  │
            ▼                                  ▼
  ┌──────────────────────┐          ┌─────────────────────┐
  │  链路A：主链路（9）│          │  链路B：敬满docx（11）│
  │  产出 HTML + MD       │          │  产出 docx           │
  └──────────┬───────────┘          └─────────────────────┘
             │                              （独立，不交汇）
             ▼
  ┌──────────────────────┐
  │  报告/{bg}/html + md  │◄─────┐
  └──────────────────────┘        │
                │ 局部覆写
                        ┌─────────┴────────────┐
                        │  链路C：修补（12）     │
                        └──────────────────────┘

                ┌──────────────────────┐
                      │  独立工具（3）         │
                      └──────────────────────┘
```

**三条链路的关系**：
- **A 和 B 完全独立**，输入有交集（都读【敬满】数据），但产物不同、互不调用
- **C 依赖 A 的产物**，且大量`import generate_html_report`，是 A 的补丁层
- 日常工作**只需要跑 A**

---

## 二、链路A：主链路（9 个文件）

**目标**：xlsx → `报告/{bg}/html/*.html` + `报告/{bg}/md/*.md`

### 数据流

```
【组织诊断结果】/组织诊断2025全年最终版.xlsx
【全面反馈】/全面反馈25H2v2.xlsx                ┐
【全面反馈】/{BG}部门负责人-1BP观察.xlsx              │
                │                │
                    ▼  ① extract_open_feedback.py   │
        全面反馈开放题{BG}.xlsx                       │
                    │                              │
                    ▼  ② step1_extract_bg.py        │  这三步产出1.3 的源文
        output/{bg}/step1_output.xlsx│  （已有 md 时可整段跳过）
        （4 sheet：部门概览/管理者小计/原文清单/异常日志） │
                    │                              │
                    ▼  ③ generate_source_md_with_llm.py  ← 调LLM
        output/{bg}/{idx}_{dept}.md                 ┘
                    │
                    │        ┌── 【组织架构信息】/组织机构信息.xlsx
                    │        ├── 【组织架构信息】/岗位信息表.xlsx
                    │        ├── 【异动】/模板25年离职明细.xlsx
                    │        ├── 【敬满】/模板_敬满开放题分析.xlsx
                    │        └── 【全面反馈】/模板-干部侧问题总结.xlsx
                    │                  │
                    ▼                  ▼
              ④ generate_html_report.py  ← 调 LLM（离职原因分析）
                    │
                    ▼
            报告/{bg}/html/*.html
                    │
                    ▼  ⑤ html2md.py
            报告/{bg}/md/*.md
                    │
                    ▼  ⑥ refresh_section_13.py  ← 回读 output/{bg}/*.md
            只覆写 1.3，其余章节不动
```

### 文件归属

| # | 文件 | 角色 |
|---|---|---|
| 1 | `run_report_pipeline.py` | **总控入口**，subprocess 串起 ①~⑥ |
| 2 | `【全面反馈】/extract_open_feedback.py` | ① 三表 join |
| 3 | `【全面反馈】/step1_extract_bg.py` | ② 标签解析为结构化原文 |
| 4 | `【全面反馈】/generate_source_md_with_llm.py` | ③ LLM 生成 1.3 源文 |
| 5 | `generate_html_report.py` | ④ **核心生成器**（304 KB） |
| 6 | `html2md.py` | ⑤ HTML → MD |
| 7 | `refresh_section_13.py` | ⑥ 只刷 1.3 |
| 8 | `【全面反馈】/step1_extract_csig.py` | ② 的前身，CSIG 硬编码版（保留追溯） |
| 9 | `【全面反馈】/step1_extract_s3.py` | ② 的前身，S3 硬编码版（保留追溯） |

### 两种执行姿势

```bash
# 姿势1：已有源md，只重建报告 —— 最常用、结果最稳定
python3 run_report_pipeline.py --bgs CSIG --steps build-html convert-md refresh-13

# 姿势2：全链路，含 LLM 重新生成 1.3 源文
python3 run_report_pipeline.py --bgs CSIG \
  --steps extract-open-feedback step1-extract generate-source-md build-html convert-md refresh-13
```

> **要复刻既有报告就用姿势1**。姿势2 经过 LLM，即使 `temperature=0` 也无法保证逐字一致。

### 核心生成器的内部分层

`generate_html_report.py` 是链路A 的重心，也是链路C 的依赖源。内部按「加载器 → 组装 → 渲染」分层：

```
加载层（各管一类数据源）
  ExcelReader              Excel 读取底层
  ResignationAnalyzer      异动/离职（调 LLM）
  BPObservationLoader      BP 观察标签
  JianGangLoader           兼岗
  OpenFeedbackLoader       全面反馈干部侧/组织侧
  JingmanOpenLoader        敬满开放题
  WordCloudLoader          词云
  YidongDataLoader         异动
  JingmanDataLoader        敬满主数据
  JingmanDetailLoader      敬满逐题细分
        │
        ▼
组装层
  generate_report_data()汇总单份报告全部数据
  build_org_chart_data()   组织架构图
  build_resignation_data() / build_jingman_data()
        │
        ▼
渲染层
  generate_html_report()   吐出完整 HTML
        │
        ▼
入口层
  batch_generate() / interactive_generate() / main()
```

**模板回退机制**：5 处数据加载都遵循「先找 `{BG}_xxx.xlsx`，找不到退回 `模板_xxx.xlsx`」，所以模板工程只放通用模板也能跑通。这是链路A 不受「已知问题1」影响的原因。

---

## 三、链路B：敬满docx（11 个文件）

**目标**：敬满数据 → `output/报告_{部门}.docx`

**与链路A 无任何调用关系**，是套独立工具。

### 数据流

```
【敬满】/开放题/{序号}_2025_{BG}_{部门}_敬满_整合版.xlsx
                │
                    ▼  extract_jingman_open_answers.py
        {BG}_敬满开放题汇总.xlsx
        {BG}_敬满开放题关键词汇总.xlsx
                    │
                    ▼  analyze_sentiment_v2.py  ← 调 LLM逐词判情感（带本地缓存）
        {BG}_敬满开放题关键词分析.xlsx
                    │
                    │┌── data/全量敬满数据.xlsx
                    │   ├── data/BG相关数据.xlsx
                    │   └── data/题目与标题对照表.xlsx
                    ▼         │
              generate_report.py
                    │
        ┌───────────┼───────────┬───────────┬───────────┐
        ▼           ▼           ▼           ▼           ▼
    module1     module2     module3     module4     module5
    核心维度表   风险区间表   柱状图PNG   排名靠后    细分项大表
        └───────────┴─────┬─────┴───────────┴───────────┘
                          ▼  JSON payload
                   report_builder.js（Node.js）
                          ▼
                  output/报告_{部门}.docx
```

### 文件归属

| # | 文件 | 角色 |
|---|---|---|
| 10 | `【敬满】/report_tool_final/generate_report.py` | **主入口**，调度各 module，转 JSON 交给 Node |
| 11 | `【敬满】/report_tool_final/config.py` | 三个 Excel 的路径配置 |
| 12 | `【敬满】/report_tool_final/data_loader.py` | 读 Excel、定义 47 道子题、格式化函数 |
| 13 | `【敬满】/report_tool_final/module1_core_table.py` | 敬业度/满意度 核心维度表 |
| 14 | `【敬满】/report_tool_final/module2_risk_table.py` | BG 末20% / 末10% 风险区间表 |
| 15 | `【敬满】/report_tool_final/module3_chart.py` | 47 题击败率双向柱状图 → PNG |
| 16 | `【敬满】/report_tool_final/module4_bottom_table.py` | 击败率 ≤10% 明细表 |
| 17 | `【敬满】/report_tool_final/module5_subdivision_table.py` | 细分项大表（4 分组） |
| 18 | `【敬满】/extract_jingman_open_answers.py` | 前置：部门 Excel → BG 汇总 |
| 19 | `【敬满】/analyze_sentiment_v2.py` | 前置：关键词情感分析（调 LLM） |
| 20 | `【敬满】/test_s3_sentiment.py` | 上者的 S3 单 BG 调试版 |

> `report_builder.js` 是 Node 脚本，不计入 35 个 Python 文件。

### 47 道子题的结构

`data_loader.py` 里 `Q47` 是整条链路的计算基准：

```
敬业度 4 道：say_q1 / stay_q1 / stay_q2 / strive_q1
满意度 43 道：
  gb_ 开头9 道（Great Boss：中干 1 + 直接上级 8）
  gj_ 开头  5 道（工作本身）
  gr_ 开头  5 道（绩效/薪酬/福利/晋升）
  gc_ 开头 24 道（团队/协作/文化/客户/沟通/未来/活力/人才/多样性）
```

每道题读 6 个派生列：`_fav`（今年正向比例）、`_fav2024`、`_fav_bg`（击败率）、`_fav2024_bg`、`_fav_diff`（分值变化）、`_fav_growth`（增幅）。

### 调用方式

```bash
cd 【敬满】/report_tool_final
python3 generate_report.py "BG/部门名"   # 单个
python3 generate_report.py --all          # 全部
```

**额外依赖**：Node.js（渲染 docx）、matplotlib + 中文字体（画图）。

---

## 四、链路C：修补脚本（12 个文件）

**目标**：对**已生成**的 HTML/MD 做局部覆写，不重跑主链路。

**特征**：
- 输入输出都是 `报告/{bg}/html|md/` 下的成品
- 大量 `import generate_html_report` 复用加载器
- 都是历史上为修具体问题写的，**日常不跑**

### 三个子系列

```
update系列（5）—— 更新内容片段
  update_wordcloud.py ──────────┐ 提供 load_keyword_data / gen_wc_section / match_dept
                │
  update_jingman_open.py ◄───────┤ 2.3.3 敬满开放题（通用版 s3/cdg/ieg）
  update_s3_jingman_open.py ◄────┘ 2.3.3（S3 专用）
  update_section_13.py             1.3（单文件，传 html + md 路径）
  update_csig_section_1_3.py       1.3（CSIG专用，已被 refresh_section_13 取代）

patch 系列（3）—— 补充缺失内容
  patch_232_subdivision.py         给 CDG/S3 补 2.3.2 细分项表
  patch_s3_233.py                  只更新 S3 的 2.3.3（开放题+词云）
  patch_small_dept_wordcloud.py    ≤20 人部门词频门槛降到 ≥1，重算词云

fix 系列（4）—— 修渲染/样式 bug
  fix_cadre_dashed_box.py          summary 段落移回虚线框内
  fix_jingman_open_low_count.py    删除累计人数 <2 的维度并重排编号
  fix_org_chart_batch.py           批量刷新组织架构图
  fix_wordcloud_position.py        词云移出 footer + 自适应高度
```

### 文件归属

| # | 文件 | 修改目标章节 | 依赖 |
|---|---|---|---|
| 21 | `update_wordcloud.py` | 词云 section | — |
| 22 | `update_jingman_open.py` | 2.3.3 | `JingmanOpenLoader` +21 |
| 23 | `update_s3_jingman_open.py` | 2.3.3 | `JingmanOpenLoader` `WordCloudLoader` + 21 |
| 24 | `update_section_13.py` | 1.3 | `OpenFeedbackLoader` |
| 25 | `update_csig_section_1_3.py` | 1.3 | — |
| 26 | `patch_232_subdivision.py` | 2.3.2 | `JingmanDetailLoader` |
| 27 | `patch_s3_233.py` | 2.3.3 | `JingmanOpenLoader` `WordCloudLoader` |
| 28 | `patch_small_dept_wordcloud.py` | 词云 | 调 LLM |
| 29 | `fix_cadre_dashed_box.py` | 1.3 样式 | — |
| 30 | `fix_jingman_open_low_count.py` | 2.3.3 | — |
| 31 | `fix_org_chart_batch.py` | 组织架构图 | `build_org_chart_data` `ExcelReader` |
| 32 | `fix_wordcloud_position.py` | 词云位置 | — |

> **启用前必读**：链路C 多数脚本写死了 `{BG}_xxx.xlsx` 文件名和日期后缀（如 `_20260409.html`），而模板已统一为 `模板_xxx.xlsx`，直接跑会报文件不存在。详见 `handover_scripts.md` 的「已知问题」。

---

## 五、独立工具（3 个文件）

不属于任何链路，按需单独使用。

| # | 文件 | 用途 |
|---|---|---|
| 33 | `scripts/prepare_workspace.py` | 把 `assets/project_template/` 复制到目标目录，初始化工作区 |
| 34 | `【全面反馈】/generate_bp_observation_template.py` | 从组织诊断表提人名，生成空的 BP 观察填写模板 |
| 35 | `desensitize_report.py` | 报告脱敏（映射表已清空，用前需填 `_person_list` / `ORG_MAP` / `BIZ_MAP`） |

---

## 六、35 个文件分布速查

| 链路 | 数量 | 产物 | 是否日常使用 |
|---|---|---|---|
| A主链路 | 9（含 2 个历史版本） | `报告/{bg}/html` + `md` | ✅ 是|
| B 敬满 docx | 11 | `output/*.docx` | 按需 |
| C 修补 | 12 | 覆写 A 的产物 | ❌ 遇问题才用 |
| 独立工具 | 3 | — | 按需 |
| **合计** | **35** | | |

---

## 七、跨链路的公共约定

### LLM 调用

链路A、B、C 各有脚本会调模型，全部统一走同一组环境变量：

```bash
export ORG_DIAG_LLM_URL="https://your-endpoint/v1/chat/completions"
export ORG_DIAG_LLM_MODEL="gpt-4.1-mini"
export ORG_DIAG_LLM_API_KEY="sk-xxx"
```

不设置时回退到本地 LM Studio 默认地址（`http://127.0.0.1:1234`）。

**调 LLM 的脚本**：
- 链路A：`step1_extract_bg.py`（BP 极性补判）、`generate_source_md_with_llm.py`（生成 1.3）、`generate_html_report.py`（离职原因分析）
- 链路B：`analyze_sentiment_v2.py`、`test_s3_sentiment.py`
- 链路C：`patch_small_dept_wordcloud.py`

### BG 列表

10 个 BG在多处出现，需保持一致：`CDG` `CSIG` `IEG` `PCG` `TEG` `WXG` `S1` `S2` `S3` `OFS`

目录名用小写（`报告/csig/`），脚本参数用大写（`--bgs CSIG`）。

### 章节编号对应

链路C 脚本常按章节号定位，对应关系：

| 编号 | 章节 | 相关脚本 |
|---|---|---|
| 1.3 | 全面反馈开放题总结 | `refresh_section_13` `update_section_13` `update_csig_section_1_3` `fix_cadre_dashed_box` |
| 2.3.2 | 值得关注的细分项 | `patch_232_subdivision` |
| 2.3.3 | 敬满开放题 + 词云 | `update_jingman_open` `update_s3_jingman_open` `patch_s3_233` `fix_jingman_open_low_count` |
| — | 组织架构图 | `fix_org_chart_batch` |
