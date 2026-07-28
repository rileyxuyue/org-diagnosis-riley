#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用 OpenAI 兼容模型按 Prompt2 从 `step1_output.xlsx` 生成 `output/{bg}/*.md`。"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

import pandas as pd
import requests

SCRIPT_DIR = Path(__file__).resolve().parent
LLM_URL = os.environ.get("ORG_DIAG_LLM_URL", "http://127.0.0.1:1234/v1/chat/completions")
LLM_MODEL = os.environ.get("ORG_DIAG_LLM_MODEL", "qwen2.5-7b-instruct-mlx")
LLM_API_KEY = os.environ.get("ORG_DIAG_LLM_API_KEY", os.environ.get("LLM_API_KEY", "lm-studio"))


def dept_dump(xlsx_path: Path, dept_idx: int) -> tuple[str, str]:
    xl = pd.ExcelFile(xlsx_path)
    overview = pd.read_excel(xl, "部门概览")
    mgr = pd.read_excel(xl, "管理者小计")
    items = pd.read_excel(xl, "原文清单")
    issues = pd.read_excel(xl, "异常日志")

    ov = overview[overview["dept_idx"] == dept_idx]
    if ov.empty:
        raise ValueError(f"dept_idx={dept_idx} 不存在")

    dept_path = str(ov.iloc[0]["dept_path"])
    dept_name = dept_path.split("/")[-1]
    payload = {
        "部门概览": ov.to_dict(orient="records"),
        "管理者小计": mgr[mgr["dept_idx"] == dept_idx].to_dict(orient="records"),
        "原文清单": items[items["dept_idx"] == dept_idx].to_dict(orient="records"),
        "异常日志": issues[issues["dept_idx"] == dept_idx].to_dict(orient="records"),
    }
    return dept_name, json.dumps(payload, ensure_ascii=False, indent=2)


def call_llm(prompt_text: str, sample_md: str, dept_text: str) -> str:
    messages = [
        {"role": "system", "content": "你是一个严格按格式生成 Markdown 报告的助手。只输出最终 Markdown，不要解释。"},
        {"role": "user", "content": prompt_text},
        {"role": "user", "content": "下面是格式样例，请严格复刻空行、缩进、引号、标题层级。\n\n" + sample_md},
        {"role": "user", "content": "下面是当前 dept 的结构化输入 JSON，请只基于它生成该部门的 .md 报告：\n\n" + dept_text},
    ]
    resp = requests.post(
        LLM_URL,
        headers={"Authorization": f"Bearer {LLM_API_KEY}"},
        json={"model": LLM_MODEL, "messages": messages, "temperature": 0},
        timeout=180,
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"].strip() + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="从 step1_output.xlsx 生成源 md")
    parser.add_argument("--bg", required=True, help="BG，如 CSIG / TEG / OFS")
    parser.add_argument("--prompt-file", default="output/报告生成_Prompt2.md")
    parser.add_argument("--sample-md", default="output/examples/14_示例部门B.md")
    parser.add_argument("--dept-idx", nargs="*", type=int, help="指定 dept_idx；不传则全量")
    args = parser.parse_args()

    bg = args.bg.upper()
    xlsx_path = SCRIPT_DIR / "output" / bg.lower() / "step1_output.xlsx"
    out_dir = SCRIPT_DIR / "output" / bg.lower()
    out_dir.mkdir(parents=True, exist_ok=True)

    prompt_text = (SCRIPT_DIR / args.prompt_file).read_text(encoding="utf-8")
    sample_md = (SCRIPT_DIR / args.sample_md).read_text(encoding="utf-8")
    overview = pd.read_excel(xlsx_path, "部门概览")
    dept_ids = args.dept_idx if args.dept_idx else overview["dept_idx"].tolist()

    for dept_idx in dept_ids:
        dept_name, dept_text = dept_dump(xlsx_path, dept_idx)
        result = call_llm(prompt_text, sample_md, dept_text)
        output_file = out_dir / f"{dept_idx}_{dept_name}.md"
        output_file.write_text(result, encoding="utf-8")
        print(f"✅ 已生成: {output_file}")


if __name__ == "__main__":
    main()
