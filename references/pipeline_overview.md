## 目标

在**独立模板工程**里复现 `报告/` 目录下的最终 HTML / MD 产物，且不修改原仓库。

## 生成链路

### 1. 全面反馈原始拼接

脚本：`【全面反馈】/extract_open_feedback.py`

输入：
- `【组织诊断结果】/组织诊断2025全年最终版.xlsx`
- `【全面反馈】/全面反馈25H2v2.xlsx`
- `【全面反馈】/{BG}部门负责人-1BP观察.xlsx`

输出：
- `【全面反馈】/全面反馈开放题{BG}.xlsx`

### 2. Step1 结构化抽取

脚本：`【全面反馈】/step1_extract_bg.py`

输入：
- `【全面反馈】/全面反馈开放题{BG}.xlsx`

输出：
- `【全面反馈】/output/{bg}/step1_output.xlsx`

### 3. LLM 生成 1.3 源 MD

脚本：`【全面反馈】/generate_source_md_with_llm.py`

输入：
- `【全面反馈】/output/{bg}/step1_output.xlsx`
- `【全面反馈】/output/报告生成_Prompt2.md`
- `【全面反馈】/output/examples/14_示例部门B.md`

输出：
- `【全面反馈】/output/{bg}/*.md`

### 4. 生成最终 HTML

脚本：`generate_html_report.py`

输入：
- 组织诊断、组织架构、异动、敬满、全面反馈干部侧/组织侧等模板工程中的全部 xlsx

输出：
- `报告/{bg}/html/*.html`

### 5. HTML 转 MD

脚本：`html2md.py`

输入：
- `报告/{bg}/html/*.html`

输出：
- `报告/{bg}/md/*.md`

### 6. 只刷新 1.3

脚本：`refresh_section_13.py`

输入：
- `【全面反馈】/output/{bg}/*.md`
- `报告/{bg}/html/*.html`
- `报告/{bg}/md/*.md`

输出：
- 仅更新报告中的 `1.3 开放题总结`
- **1.3 之外的其它章节不应修改**

## 推荐执行方式

直接在模板工程根目录运行：

```bash
python3 run_report_pipeline.py --bgs CSIG
```

如果 `【全面反馈】/output/{bg}/*.md` 已经准备好，则推荐只跑：

```bash
python3 run_report_pipeline.py --bgs CSIG --steps build-html convert-md refresh-13
```

这样最接近当前仓库的最终产物。
