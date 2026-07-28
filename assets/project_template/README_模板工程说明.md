# 组织诊断报告模板工程说明

这个模板工程用于在独立目录中复现 `报告/` 下最终 html / md 产物，不会修改原仓库。

## 数据打包边界

- 本模板工程**不包含原始业务数据**。
- 所有 `xlsx` 仅保留文件名、sheet、列头与必要格式，内容已清空，需由使用者自行填入真实数据。
- 示例产出仅保留在 `assets_examples/` 与 `【全面反馈】/output/examples/`，用于参考目录结构与结果样式。

## 推荐使用顺序

1. 准备输入数据：将各模板 xlsx 替换为真实数据。
2. 如需先产出 `【全面反馈】/output/{bg}/*.md`：
   - 先运行 `【全面反馈】/extract_open_feedback.py {BG}`
   - 再运行 `【全面反馈】/step1_extract_bg.py --bg {BG}`
   - 再运行 `【全面反馈】/generate_source_md_with_llm.py --bg {BG}`
3. 生成最终报告：
   - `python3 run_report_pipeline.py --bgs CSIG TEG`
4. 如已有 `【全面反馈】/output/{bg}/*.md`，可直接跳过 LLM 步骤，执行 html/md 构建与 1.3 刷新。

## 线上大模型配置

以下环境变量支持本地或线上 OpenAI 兼容接口：

- `ORG_DIAG_LLM_URL`
- `ORG_DIAG_LLM_MODEL`
- `ORG_DIAG_LLM_API_KEY`

例如：

```bash
export ORG_DIAG_LLM_URL="https://your-openai-compatible-endpoint/v1/chat/completions"
export ORG_DIAG_LLM_MODEL="gpt-4.1-mini"
export ORG_DIAG_LLM_API_KEY="sk-xxx"
```

## 结果目录

- HTML：`报告/{bg}/html/`
- MD：`报告/{bg}/md/`

## 说明

- `generate_html_report.py` 是最终报告主生成器。
- `refresh_section_13.py` 只刷新 1.3，不改 1.3 外的其它章节。
- 如需尽量复刻当前仓库结果，建议直接提供 `【全面反馈】/output/{bg}/*.md`。这样 1.3 将完全走结构化源 md，不依赖再次生成。
