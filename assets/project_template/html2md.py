#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将 `报告/{bg}/html/*.html` 批量转换为 `报告/{bg}/md/*.md`。"""
from __future__ import annotations

import argparse
import glob
import os
import re
from markdownify import markdownify as md

BGS = ["cdg", "csig", "ieg", "pcg", "teg", "wxg", "s1", "s2", "s3", "ofs"]
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def clean_html_for_md(html: str) -> str:
    html = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.S)
    html = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.S)
    html = re.sub(r"<head[^>]*>.*?</head>", "", html, flags=re.S)
    html = re.sub(r"<div class=\"toolbar[^\"]*\">.*?</div>", "", html, flags=re.S)
    html = re.sub(r"<div class=\"wc-section\">.*?</div>\s*(?=<div class=\"footer\">)", "", html, flags=re.S)
    html = re.sub(r"<div class=\"footer\">.*?</div>", "", html, flags=re.S)
    html = re.sub(r"<div class=\"org-chart-controls\">.*?</div>", "", html, flags=re.S)
    return html


def html_to_markdown(html_content: str) -> str:
    result = md(clean_html_for_md(html_content), heading_style="ATX", bullets="-")
    result = re.sub(r"\n{4,}", "\n\n\n", result)
    result = re.sub(r"[ \t]+\n", "\n", result)
    return result.strip() + "\n"


def convert_file(html_path: str, md_path: str) -> int:
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()
    markdown = html_to_markdown(html)
    os.makedirs(os.path.dirname(md_path), exist_ok=True)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(markdown)
    return len(markdown)


def main() -> None:
    parser = argparse.ArgumentParser(description="HTML 组织诊断报告转 Markdown")
    parser.add_argument("--bgs", nargs="*", default=["all"], help="BG 列表，如 CSIG TEG；默认 all")
    args = parser.parse_args()

    bg_list = BGS if args.bgs == ["all"] else [bg.lower() for bg in args.bgs]
    total = 0
    for bg in bg_list:
        html_dir = os.path.join(BASE_DIR, "报告", bg, "html")
        md_dir = os.path.join(BASE_DIR, "报告", bg, "md")
        if not os.path.isdir(html_dir):
            print(f"跳过不存在目录: {html_dir}")
            continue
        os.makedirs(md_dir, exist_ok=True)
        files = sorted(glob.glob(os.path.join(html_dir, "*.html")))
        print(f"[{bg.upper()}] HTML → MD: {len(files)} 个")
        for fp in files:
            md_path = os.path.join(md_dir, os.path.basename(fp).replace(".html", ".md"))
            convert_file(fp, md_path)
            total += 1
    print(f"完成: {total} 个 MD")


if __name__ == "__main__":
    main()
