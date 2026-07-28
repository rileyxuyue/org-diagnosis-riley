#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按 BG 将 `【全面反馈】/output/{bg}/*.md` 刷新到 `报告/{bg}/html|md` 的 1.3。"""
from __future__ import annotations

import argparse
import re
from html import escape as html_escape
from pathlib import Path

WORKSPACE = Path(__file__).resolve().parent
BG_LOWER_MAP = {
    "CDG": "cdg",
    "CSIG": "csig",
    "IEG": "ieg",
    "PCG": "pcg",
    "TEG": "teg",
    "WXG": "wxg",
    "S1": "s1",
    "S2": "s2",
    "S3": "s3",
    "OFS": "ofs",
}
JUDGE_STYLE = {
    "共性优势": {"color": "#389e0d", "bg": "#f6ffed", "border": "#b7eb8f", "tbl_border": "#b7eb8f"},
    "共性不足": {"color": "#cf1322", "bg": "#fff1f0", "border": "#ffa39e", "tbl_border": "#ffa39e"},
    "局部待关注": {"color": "#d48806", "bg": "#fffbe6", "border": "#ffe58f", "tbl_border": "#ffe58f"},
}


def is_health_dim(dim_name: str) -> bool:
    return any(kw in dim_name for kw in ["身体健康", "健康与可持续", "健康与工作节奏", "健康与节奏", "健康可持续"])


def parse_voice_field(rest: str, key: str):
    m = re.search(rf"{key}\s*(\d+)\s*声音?\s*\(([^)]*?)\)", rest)
    if m:
        return int(m.group(1)), [n.strip() for n in m.group(2).split(",") if n.strip()]
    m = re.search(rf"{key}\s*(\d+)\s*声音?", rest)
    if m:
        return int(m.group(1)), []
    return 0, []


def parse_quote_line(line: str):
    polarity = "pos"
    is_bp = False
    text = line.strip()
    if text.startswith("待关注"):
        polarity = "neg"
        text = re.sub(r"^待关注[：:]\s*", "", text)
    if text.startswith("bp"):
        is_bp = True
        text = re.sub(r"^bp[：:]\s*", "", text)
    tag_match = re.search(r"[（(](bp\s+)?(N\d+)(?:[,，]\s*节选)?[）)]\s*$", text)
    if tag_match:
        if tag_match.group(1):
            is_bp = True
        text = text[: tag_match.start()].rstrip()
    text = re.sub(r'^[“"](.*?)[”"]?$', r"\1", text).strip().strip('"')
    return {"polarity": polarity, "is_bp": is_bp, "text": text}


def parse_dimension_block(block: str):
    lines = block.split("\n")
    if not lines:
        return None
    m = re.match(r"### (维度[^：:]*?[:：]\s*.+)", lines[0])
    if not m:
        return None
    name = m.group(1).strip()
    definition = ""
    judge = ""
    for line in lines[1:]:
        s = line.strip()
        if s.startswith(">") and not definition:
            desc = s.lstrip(">").strip()
            if desc and "整体判断" not in desc and "在该维度无明显评价" not in desc:
                definition = desc
                continue
        jm = re.match(r"\*\*判断[：:]\s*(.+?)\*\*", s)
        if jm:
            judge = re.split(r"[（(]", jm.group(1).strip())[0].strip()
            break
    managers = []
    i = 0
    while i < len(lines):
        mm = re.match(r"^- \*\*([^*]+?)\*\*[：:]\s*(.+)$", lines[i])
        if not mm:
            i += 1
            continue
        mgr_name = mm.group(1).strip()
        rest = mm.group(2).strip()
        pos_count, _ = parse_voice_field(rest, "正向")
        neg_count, _ = parse_voice_field(rest, "待关注")
        pos_quotes = []
        neg_quotes = []
        j = i + 1
        while j < len(lines):
            nl = lines[j]
            stripped = nl.strip()
            if re.match(r"^- \*\*", nl):
                break
            if not stripped or stripped.startswith(">") or re.match(r"^[-—=─_*]+\s*$", stripped):
                j += 1
                continue
            if re.match(r"^\s+-\s", nl):
                q = parse_quote_line(nl.lstrip()[1:].strip())
                (pos_quotes if q["polarity"] == "pos" else neg_quotes).append(q)
                j += 1
                continue
            break
        managers.append({
            "name": mgr_name,
            "pos_count": pos_count,
            "pos_quotes": pos_quotes,
            "neg_count": neg_count,
            "neg_quotes": neg_quotes,
        })
        i = j
    return {"name": name, "definition": definition, "judge": judge, "managers": managers}


def parse_source_md(path: Path):
    text = path.read_text(encoding="utf-8")
    tldr_rows = []
    tldr_match = re.search(r"## TL;DR\s*\n(.+?)(?=\n##|\Z)", text, re.S)
    if tldr_match:
        for line in tldr_match.group(1).splitlines():
            line = line.strip()
            if not line.startswith("|") or line.startswith("|---") or line.startswith("| ---") or "|------" in line:
                continue
            if "维度" in line and "判断" in line and "一句话总结" in line:
                continue
            cells = [c.strip() for c in line.strip("|").split("|")]
            if len(cells) >= 3:
                dim = cells[0].strip().strip("*")
                if not is_health_dim(dim):
                    tldr_rows.append((dim, cells[1].strip(), cells[2].strip()))
    dimensions = []
    detail_match = re.search(r"## 第三步：分维度详情\s*\n(.+?)(?=\n## 报告说明|\Z)", text, re.S)
    if detail_match:
        dim_blocks = re.findall(r"### 维度[^\n]+\n(?:(?!^### 维度).)*", detail_match.group(1), re.S | re.M)
        for blk in dim_blocks:
            dim_obj = parse_dimension_block(blk)
            if dim_obj and not is_health_dim(dim_obj["name"]):
                dimensions.append(dim_obj)
    mgr_order = []
    for dim in dimensions:
        for mgr in dim["managers"]:
            if mgr["name"] not in mgr_order:
                mgr_order.append(mgr["name"])
    return {"tldr": tldr_rows, "dimensions": dimensions, "mgr_order": mgr_order}


def format_cell_md(count: int, quotes: list[dict]) -> str:
    if count <= 0 and not quotes:
        return "—"
    head = f"{count}人" if count > 0 else f"{len(quotes)}人"
    parts = [head]
    for q in quotes[:4]:
        text = q["text"].replace("|", "\\|").replace("\n", " ").strip()
        if not text:
            continue
        if q["is_bp"]:
            text = f"{text} [From BP]"
        parts.append(f'"{text}"')
    return "  ".join(parts)


def format_cell_html(count: int, quotes: list[dict], color: str) -> str:
    if count <= 0 and not quotes:
        return '<td style="padding:8px 10px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:#ccc;">—</td>'
    head = f"{count}人" if count > 0 else f"{len(quotes)}人"
    inner = f'<div style="margin-bottom:6px;font-weight:700;font-size:11px;color:{color};">{head}</div>'
    for q in quotes[:4]:
        text = q["text"].strip()
        if not text:
            continue
        if q["is_bp"]:
            text = f"{text} [From BP]"
        inner += f'<div style="margin:4px 0;padding-left:8px;border-left:2px solid #d9d9d9;color:#555;">"{html_escape(text)}"</div>'
    return f'<td style="padding:8px 10px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:{color};">{inner}</td>'


def render_md_section(parsed) -> str:
    lines = ["📊 最终整体判断（跨维度）", "", "| 维度 | 判断 | 一句话总结 |", "| --- | --- | --- |"]
    for dim, judge, summary in parsed["tldr"]:
        lines.append(f"| **{dim}** | {judge} | {summary} |")
    lines.append("")
    for dim in parsed["dimensions"]:
        lines.extend([dim["name"], ""])
        if dim["judge"]:
            lines.extend([dim["judge"], ""])
        if dim["definition"]:
            lines.extend([dim["definition"], ""])
        lines.extend(["**各管理者情况：**", "", "| 干部 | 正向 | 待关注 |", "| --- | --- | --- |"])
        mgr_map = {m["name"]: m for m in dim["managers"]}
        for mgr_name in parsed["mgr_order"]:
            if mgr_name in mgr_map:
                mgr = mgr_map[mgr_name]
                lines.append(
                    f"| {mgr_name} | {format_cell_md(mgr['pos_count'], mgr['pos_quotes'])} | {format_cell_md(mgr['neg_count'], mgr['neg_quotes'])} |"
                )
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def render_html_section(parsed) -> str:
    parts = [
        '<div style="background:#fff;border:1px solid #e8e8e8;border-radius:8px;padding:20px 24px;margin-bottom:20px;line-height:1.7;">',
        '<div style="margin-bottom:16px;">',
        '<div style="font-size:14px;font-weight:700;color:#1a1a1a;margin-bottom:8px;">📊 最终整体判断（跨维度）</div>',
        '<table style="width:100%;border-collapse:collapse;font-size:13px;line-height:1.6;margin:12px 0;"><thead><tr><th style="padding:8px 12px;border:1px solid #d9d9d9;font-weight:700;text-align:left;background:#f5f5f5;color:#333;">维度</th><th style="padding:8px 12px;border:1px solid #d9d9d9;font-weight:700;text-align:left;background:#f5f5f5;color:#333;">判断</th><th style="padding:8px 12px;border:1px solid #d9d9d9;font-weight:700;text-align:left;background:#f5f5f5;color:#333;">一句话总结</th></tr></thead><tbody>',
    ]
    for dim, judge, summary in parsed["tldr"]:
        color = JUDGE_STYLE.get(judge.strip(), JUDGE_STYLE["局部待关注"])["color"]
        parts.append(
            f'<tr><td style="padding:8px 12px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:#333;"><b>{html_escape(dim)}</b></td><td style="padding:8px 12px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:{color};font-weight:600">{html_escape(judge)}</td><td style="padding:8px 12px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:#333;">{html_escape(summary)}</td></tr>'
        )
    parts.append("</tbody></table></div>")
    for dim in parsed["dimensions"]:
        style = JUDGE_STYLE.get(dim["judge"].strip(), JUDGE_STYLE["局部待关注"])
        parts.append(f'<div style="margin:12px 0;padding:12px 16px;background:{style["bg"]};border:2px solid {style["border"]};border-radius:10px;">')
        parts.append(f'<div style="margin:0 0 6px;font-size:14px;font-weight:700;color:{style["color"]};">{html_escape(dim["name"])}</div>')
        parts.append(f'<div style="margin:6px 0 8px;padding:8px 10px;background:rgba(255,255,255,0.55);border:1px dashed {style["border"]};border-radius:8px;">')
        if dim["judge"]:
            parts.append(f'<div style="margin:0 0 4px;font-size:13px;line-height:1.7;color:#222;"><span style="color:{style["color"]};font-weight:600;">{html_escape(dim["judge"])}</span></div>')
        if dim["definition"]:
            parts.append(f'<div style="margin:0 0 4px;font-size:13px;line-height:1.7;color:#222;">{html_escape(dim["definition"])}</div>')
        parts.append("</div>")
        parts.append('<div style="margin:12px 0 6px;font-size:13px;font-weight:700;color:#444;"><b>各管理者情况：</b></div>')
        parts.append(f'<table style="width:100%;border-collapse:collapse;font-size:12px;line-height:1.6;margin:8px 0;border:2px solid {style["tbl_border"]};border-radius:4px;"><thead><tr><th style="padding:8px 10px;border:1px solid #d9d9d9;font-weight:700;text-align:center;background:#f0f5ff;color:#333;white-space:nowrap;">干部</th><th style="padding:8px 10px;border:1px solid #d9d9d9;font-weight:700;text-align:center;background:#f6ffed;color:#389e0d;white-space:nowrap;">正向</th><th style="padding:8px 10px;border:1px solid #d9d9d9;font-weight:700;text-align:center;background:#fffbe6;color:#d48806;white-space:nowrap;">待关注</th></tr></thead><tbody>')
        mgr_map = {m["name"]: m for m in dim["managers"]}
        for mgr_name in parsed["mgr_order"]:
            if mgr_name in mgr_map:
                mgr = mgr_map[mgr_name]
                parts.append(
                    f'<tr><td style="padding:8px 10px;border:1px solid #e8e8e8;vertical-align:top;font-weight:600;white-space:nowrap;text-align:center;">{html_escape(mgr_name)}</td>{format_cell_html(mgr["pos_count"], mgr["pos_quotes"], "#389e0d")}{format_cell_html(mgr["neg_count"], mgr["neg_quotes"], "#d48806")}</tr>'
                )
        parts.append("</tbody></table></div>")
    parts.append("</div>")
    return "\n".join(parts)


def update_report_md(report_path: Path, new_section_md: str) -> bool:
    text = report_path.read_text(encoding="utf-8")
    pattern = re.compile(r"(#### 1\.3 开放题总结\s*\n)(.+?)(?=\n## 2 异动|\n## 2\s|\Z)", re.S)
    m = pattern.search(text)
    if not m:
        return False
    new_text = text[: m.start()] + m.group(1) + "\n" + new_section_md.rstrip() + "\n\n" + text[m.end() :]
    report_path.write_text(new_text, encoding="utf-8")
    return True


def update_report_html(report_path: Path, new_section_html: str) -> bool:
    text = report_path.read_text(encoding="utf-8")
    pattern = re.compile(r"(<h4>1\.3 开放题总结</h4>\s*)(.+?)(?=\s*<h2>2 异动|\s*<h2>\s*2[\s　]+异动)", re.S)
    m = pattern.search(text)
    if not m:
        return False
    new_text = text[: m.start()] + m.group(1) + "\n" + new_section_html + "\n        " + text[m.end() :]
    report_path.write_text(new_text, encoding="utf-8")
    return True


def find_report_files(report_dir: Path, dept_name: str, ext: str):
    found = []
    for path in report_dir.glob(f"*{ext}"):
        suffix = f"_组织诊断报告{ext}"
        if not path.name.endswith(suffix):
            continue
        body = path.name[: -len(suffix)]
        if body.split("-")[-1] == dept_name:
            found.append(path)
    return found


def process_bg(bg: str) -> None:
    bg_lower = BG_LOWER_MAP[bg]
    src_dir = WORKSPACE / "【全面反馈】" / "output" / bg_lower
    report_md_dir = WORKSPACE / "报告" / bg_lower / "md"
    report_html_dir = WORKSPACE / "报告" / bg_lower / "html"
    if not src_dir.is_dir():
        print(f"跳过 {bg}: 无源目录 {src_dir}")
        return
    md_updated = 0
    html_updated = 0
    for src_path in sorted(src_dir.glob("*.md")):
        m = re.match(r"\d+_(.+)\.md$", src_path.name)
        if not m:
            continue
        dept_name = m.group(1)
        parsed = parse_source_md(src_path)
        if not parsed["tldr"] or not parsed["dimensions"]:
            continue
        new_md_section = render_md_section(parsed)
        new_html_section = render_html_section(parsed)
        for md_file in find_report_files(report_md_dir, dept_name, ".md"):
            if update_report_md(md_file, new_md_section):
                md_updated += 1
        for html_file in find_report_files(report_html_dir, dept_name, ".html"):
            if update_report_html(html_file, new_html_section):
                html_updated += 1
    print(f"[{bg}] 更新完成: HTML={html_updated}, MD={md_updated}")


def main() -> None:
    parser = argparse.ArgumentParser(description="刷新 1.3 开放题总结")
    parser.add_argument("--bgs", nargs="*", default=["CSIG"], help="BG 列表；默认 CSIG")
    args = parser.parse_args()
    for bg in args.bgs:
        process_bg(bg.upper())


if __name__ == "__main__":
    main()
