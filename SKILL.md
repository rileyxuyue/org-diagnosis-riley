---
name: org-diagnosis-riley
description: This skill should be used when reproducing or packaging the organization diagnosis report pipeline that generates final `报告/` HTML and Markdown outputs from local Excel inputs and `【全面反馈】/output/{bg}/*.md` sources, especially when a user needs a standalone workspace, reusable input templates, or OpenAI-compatible model replacement for local LLM steps.
---

## Overview

在需要**独立复现**组织诊断报告工程时使用本 skill。目标是把生成 `报告/` 目录最终 HTML / MD 产物所需的本地程序、输入模板、Prompt、说明文档统一封装进一个独立模板工程，并保持与原仓库隔离。

本 skill **不附带任何原始业务数据**：模板工程中的 `xlsx` 已清空，仅保留结构；示例结果仅放在 `assets_examples/` 和 `【全面反馈】/output/examples/` 供参考。

## Quick Start

### 1. 复制独立模板工程

运行：

```bash
python3 scripts/prepare_workspace.py <目标目录>
```

例如：

```bash
python3 scripts/prepare_workspace.py /tmp/org-diagnosis-workspace
```

复制完成后，在目标目录里工作，不要直接修改 skill 目录本身。

### 2. 安装依赖

进入目标目录后运行：

```bash
python3 -m pip install -r requirements.txt
```

### 3. 填充输入模板

优先按 `references/input_templates.md` 中的说明，把真实 Excel 数据与 `【全面反馈】/output/{bg}/*.md` 放入目标目录对应位置。

### 4. 运行流水线

如果 `【全面反馈】/output/{bg}/*.md` 已经准备好，优先运行：

```bash
python3 run_report_pipeline.py --bgs CSIG --steps build-html convert-md refresh-13
```

如果需要从全面反馈开放题一路重建到源 md，再运行全链路：

```bash
python3 run_report_pipeline.py --bgs CSIG --steps extract-open-feedback step1-extract generate-source-md build-html convert-md refresh-13
```

## Workflow Decision Tree

### 场景 A：只想复刻当前最终报告

满足以下条件时走这个最稳妥路径：
- 已有完整输入 xlsx
- 已有 `【全面反馈】/output/{bg}/*.md`
- 目标是尽量生成与现有工程**完全一致**的 `报告/` html / md

执行：
1. 填充模板工程里的 xlsx
2. 放入 `【全面反馈】/output/{bg}/*.md`
3. 运行 `build-html -> convert-md -> refresh-13`

### 场景 B：需要把 1.3 源 md 也重新生出来

满足以下条件时走模型辅助路径：
- 没有 `【全面反馈】/output/{bg}/*.md`
- 只有全面反馈原始 Excel
- 接受极小概率的模型文本差异

执行：
1. 运行 `【全面反馈】/extract_open_feedback.py`
2. 运行 `【全面反馈】/step1_extract_bg.py`
3. 运行 `【全面反馈】/generate_source_md_with_llm.py`
4. 再继续 `build-html -> convert-md -> refresh-13`

## Core Files

### 模板工程

位于：`assets/project_template/`

包含：
- 根目录主程序：`generate_html_report.py`、`html2md.py`、`refresh_section_13.py`、`run_report_pipeline.py`
- 开放题链路程序：`【全面反馈】/extract_open_feedback.py`、`【全面反馈】/step1_extract_bg.py`、`【全面反馈】/generate_source_md_with_llm.py`
- Prompt：`【全面反馈】/output/报告生成_Prompt2.md`
- 输入模板：各目录下清空数据但保留结构的 xlsx 文件

### 参考文档

按需读取：
- `references/pipeline_overview.md`
- `references/input_templates.md`
- `references/model_and_repro.md`
- `references/architecture_three_pipelines.md`（架构：三条链路与 35 个 py 文件归属、数据流）
- `references/handover_scripts.md`（交接用：全部脚本职责、上下游、已知问题）

## Model Replacement Rules

### 统一使用 OpenAI 兼容接口变量

配置以下环境变量：
- `ORG_DIAG_LLM_URL`
- `ORG_DIAG_LLM_MODEL`
- `ORG_DIAG_LLM_API_KEY`

### 默认策略

- 有现成 `output/{bg}/*.md` 时，优先直接使用，不重新调用模型
- 只有在必须生成源 md 或做 BP 文本补判时才调用模型
- 需要强调“结果一模一样”时，优先要求用户直接提供 `output/{bg}/*.md`

## Guardrails

- 在独立目标目录中运行，不要修改原仓库现有文件
- 刷新 1.3 时，只运行 `refresh_section_13.py`，不要手改 2 / 3 章节
- 需要验证复现边界时，参考 `references/model_and_repro.md`
- 如果用户明确要求“与当前结果完全一致”，优先走**已有源 md + 最终刷新**路径，而不是重新生成源 md
