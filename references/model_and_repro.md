## 本地模型替换为线上模型

本 skill 内所有与 LLM 相关的脚本都统一支持 OpenAI 兼容接口，通过环境变量配置：

- `ORG_DIAG_LLM_URL`
- `ORG_DIAG_LLM_MODEL`
- `ORG_DIAG_LLM_API_KEY`

例如：

```bash
export ORG_DIAG_LLM_URL="https://your-openai-compatible-endpoint/v1/chat/completions"
export ORG_DIAG_LLM_MODEL="gpt-4.1-mini"
export ORG_DIAG_LLM_API_KEY="sk-xxx"
```

## 哪些脚本会调用模型

### `generate_html_report.py`
- 用于离职原因分析
- 已改成优先读取 `ORG_DIAG_LLM_*` 环境变量
- 不配置时默认回退到本地 LM Studio 风格地址

### `【全面反馈】/step1_extract_bg.py`
- 仅在 BP 文本格式不规范时做极性补判
- 已改成统一读取 `ORG_DIAG_LLM_*`

### `【全面反馈】/generate_source_md_with_llm.py`
- 负责把 `step1_output.xlsx` 生成结构化源 md
- 默认走 OpenAI 兼容接口，适合接线上模型

## 复现边界说明

### 可以做到“结果一模一样”的部分

- 当你直接提供与当前工程一致的：
  - 全部 xlsx 输入
  - `【全面反馈】/output/{bg}/*.md`
- 再运行：
  - `generate_html_report.py`
  - `html2md.py`
  - `refresh_section_13.py`

此时最终 `报告/` 下的 HTML / MD 可以与当前工程保持一致。

### 需要注意的部分

如果你选择重新调用 LLM 生成 `【全面反馈】/output/{bg}/*.md`，虽然：
- prompt 已内置
- 样例 md 已内置
- temperature 设为 0

但不同模型、不同服务端实现仍可能带来**极小文本差异**。因此：

- **绝对严格复刻**：直接提供 `output/{bg}/*.md`
- **全链路再生**：接受极小概率的文本差异
