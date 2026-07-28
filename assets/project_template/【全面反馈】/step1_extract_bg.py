#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用版 Step1：从 `全面反馈开放题{BG}.xlsx` 抽取 `step1_output.xlsx`。"""
from __future__ import annotations

import argparse
import os
import re
from pathlib import Path

import pandas as pd
import requests

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_ROOT = SCRIPT_DIR / "output"
OPEN_FEEDBACK_NAME = {
    "CDG": "全面反馈开放题CDG.xlsx",
    "CSIG": "全面反馈开放题CSIG.xlsx",
    "IEG": "全面反馈开放题IEG.xlsx",
    "PCG": "全面反馈开放题PCG.xlsx",
    "TEG": "全面反馈开放题TEG.xlsx",
    "WXG": "全面反馈开放题WXG.xlsx",
    "S1": "全面反馈开放题S1.xlsx",
    "S2": "全面反馈开放题S2.xlsx",
    "S3": "全面反馈开放题S3.xlsx",
    "OFS": "全面反馈开放题Overseas.xlsx",
    "_template": "全面反馈开放题_模板.xlsx",
}
LLM_URL = os.environ.get("ORG_DIAG_LLM_URL", "http://127.0.0.1:1234/v1/chat/completions")
LLM_MODEL = os.environ.get("ORG_DIAG_LLM_MODEL", "qwen2.5-7b-instruct-mlx")
LLM_API_KEY = os.environ.get("ORG_DIAG_LLM_API_KEY", os.environ.get("LLM_API_KEY", "lm-studio"))
NOISE_SET = {"无", "无。", "NA", "N/A", "暂无", "没有", "-", "/", "空", "NaN", "都挺好", "都挺好的", "|", ""}


def is_noise(text: str) -> bool:
    return text.strip() in NOISE_SET


def clean_prefix(text: str):
    items = []
    if "；本人不可见：" in text or ";本人不可见：" in text:
        parts = re.split(r"[；;]本人不可见[：:]", text)
        for idx, part in enumerate(parts):
            part = re.sub(r"^本人可见[：:]\s*", "", part.strip())
            part = re.sub(r"^本人不可见[：:]\s*", "", part.strip())
            if part and not is_noise(part):
                items.append((part, "本人不可见拆出" if idx > 0 else ""))
    else:
        text = re.sub(r"^本人可见[：:]\s*", "", text).strip()
        text = re.sub(r"^本人不可见[：:]\s*", "", text).strip()
        if text and not is_noise(text):
            items.append((text, ""))
    return items


def classify_bp_segment(text: str) -> str:
    payload = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": "你是一个文本分类器。给定一段bp观察文本，只输出 positive / negative / neutral 三选一。"},
            {"role": "user", "content": text[:500]},
        ],
        "temperature": 0,
        "max_tokens": 5,
    }
    try:
        resp = requests.post(
            LLM_URL,
            headers={"Authorization": f"Bearer {LLM_API_KEY}"},
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        answer = resp.json()["choices"][0]["message"]["content"].strip().lower()
        for kw in ["positive", "negative", "neutral"]:
            if kw in answer:
                return kw
    except Exception:
        pass
    return "neutral"


def parse_bp(bp_text: str):
    if not bp_text or bp_text.strip() in {"无", "空缺", "无数据", "（如无则写：无）"}:
        return [], [], "空缺", [("bp_all_empty", "bp段为空")]

    status_match = re.search(r"BP观察结果.*?\n+\s*\*\*(.+?)\*\*", bp_text, re.S)
    if not status_match:
        status_match = re.search(r"BP观察结果.*?\n+\s*(\S+)", bp_text, re.S)
    bp_status = status_match.group(1).strip() if status_match else "正常"

    positive_items = []
    negative_items = []
    issues = []

    for title, target in [("做得好的地方", positive_items), ("做得不好的地方", negative_items)]:
        match = re.search(rf"{title}[^\n]*\n(.*?)(?=\n###\s*\*\*|\n---\s*\n###|\Z)", bp_text, re.S)
        section = match.group(1).strip() if match else ""
        for quote in re.findall(r"[\"\u201c\u201d]([^\"\u201c\u201d]+)[\"\u201c\u201d]", section):
            quote = quote.strip()
            if quote and not is_noise(quote):
                target.append((quote, "bp观察"))

    if not positive_items and not negative_items and bp_text.strip() not in {"无数据", "空缺"}:
        polarity = classify_bp_segment(bp_text)
        cleaned = bp_text.strip()
        if polarity == "positive":
            positive_items.append((cleaned, "bp观察"))
        elif polarity == "negative":
            negative_items.append((cleaned, "bp观察"))
        else:
            issues.append(("bp_format_irregular", cleaned[:100]))

    return positive_items, negative_items, bp_status, issues


def parse_department(full_text: str):
    mgr_pattern = re.compile(r"【管理者[：:](.+?)】")
    splits = mgr_pattern.split(full_text)
    managers = []
    for i in range(1, len(splits), 2):
        header = splits[i].strip()
        body = splits[i + 1] if i + 1 < len(splits) else ""
        role_match = re.search(r"[（(]([^）)]*负责人[^）)]*)[）)]$", header)
        if role_match:
            role = role_match.group(1)
            name_full = header[: role_match.start()].strip()
        else:
            role = "部门负责人-1"
            name_full = header
        flash_match = re.search(r"<闪光点>\s*(.*?)\s*</闪光点>", body, re.S)
        more_match = re.search(r"<更多期待>\s*(.*?)\s*</更多期待>", body, re.S)
        bp_match = re.search(r"<bp观察>\s*(.*?)\s*</bp观察>", body, re.S)
        flash_items = [s.strip() for s in (flash_match.group(1).strip() if flash_match else "").split("||") if s.strip()]
        more_items = [s.strip() for s in (more_match.group(1).strip() if more_match else "").split("||") if s.strip()]
        bp_text = bp_match.group(1).strip() if bp_match else ""
        managers.append({
            "name_full": name_full,
            "role": role,
            "flash_items": flash_items,
            "more_items": more_items,
            "bp_text": bp_text,
        })
    return managers


def main() -> None:
    parser = argparse.ArgumentParser(description="通用 step1 提取")
    parser.add_argument("--bg", required=True, help="BG，如 CSIG / TEG / OFS")
    args = parser.parse_args()
    bg = args.bg.upper()

    input_file = SCRIPT_DIR / OPEN_FEEDBACK_NAME.get(bg, OPEN_FEEDBACK_NAME["_template"])
    if not input_file.exists():
        # fallback to generic template
        input_file = SCRIPT_DIR / OPEN_FEEDBACK_NAME["_template"]
    output_dir = OUTPUT_ROOT / bg.lower()
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / "step1_output.xlsx"

    df = pd.read_excel(input_file)
    dept_data = []
    for _, row in df.iterrows():
        org_path = str(row.get("组织全路径", "")).strip()
        full_text = "\n".join(
            str(v) for k, v in row.items() if k != "组织全路径" and pd.notna(v) and str(v).strip()
        )
        managers = parse_department(full_text)
        dept_data.append({
            "path": org_path,
            "name": org_path.split("/")[-1],
            "n_managers": len(managers),
            "total_chars": len(full_text),
            "managers": managers,
        })

    dept_data.sort(key=lambda d: (d["n_managers"], d["total_chars"]))
    for idx, dept in enumerate(dept_data):
        dept["dept_idx"] = idx

    all_items = []
    mgr_summaries = []
    dept_overviews = []
    issues_log = []

    for dept in dept_data:
        dept_total_items = 0
        dept_total_pos = 0
        dept_total_neg = 0
        for mgr in dept["managers"]:
            n_counter = 0
            flash_count = 0
            more_count = 0
            bp_positive_items = 0
            bp_negative_items = 0

            for raw in mgr["flash_items"]:
                for text, note in clean_prefix(raw):
                    n_counter += 1
                    flash_count += 1
                    dept_total_items += 1
                    dept_total_pos += 1
                    all_items.append({
                        "dept_idx": dept["dept_idx"],
                        "mgr_name": mgr["name_full"],
                        "mgr_role": mgr["role"],
                        "n_id": f"N{n_counter:02d}",
                        "source": "闪光点",
                        "polarity": "正向",
                        "text": text,
                        "note": note,
                    })

            for raw in mgr["more_items"]:
                for text, note in clean_prefix(raw):
                    n_counter += 1
                    more_count += 1
                    dept_total_items += 1
                    dept_total_neg += 1
                    all_items.append({
                        "dept_idx": dept["dept_idx"],
                        "mgr_name": mgr["name_full"],
                        "mgr_role": mgr["role"],
                        "n_id": f"N{n_counter:02d}",
                        "source": "更多期待",
                        "polarity": "待关注",
                        "text": text,
                        "note": note,
                    })

            bp_pos, bp_neg, bp_status, bp_issues = parse_bp(mgr["bp_text"])
            for text, note in bp_pos:
                n_counter += 1
                bp_positive_items += 1
                dept_total_items += 1
                all_items.append({
                    "dept_idx": dept["dept_idx"],
                    "mgr_name": mgr["name_full"],
                    "mgr_role": mgr["role"],
                    "n_id": f"N{n_counter:02d}",
                    "source": "bp观察",
                    "polarity": "正向",
                    "text": text,
                    "note": note,
                })
            for text, note in bp_neg:
                n_counter += 1
                bp_negative_items += 1
                dept_total_items += 1
                all_items.append({
                    "dept_idx": dept["dept_idx"],
                    "mgr_name": mgr["name_full"],
                    "mgr_role": mgr["role"],
                    "n_id": f"N{n_counter:02d}",
                    "source": "bp观察",
                    "polarity": "待关注",
                    "text": text,
                    "note": note,
                })
            dept_total_pos += 1 if bp_positive_items else 0
            dept_total_neg += 1 if bp_negative_items else 0

            mgr_summaries.append({
                "dept_idx": dept["dept_idx"],
                "mgr_name": mgr["name_full"],
                "mgr_role": mgr["role"],
                "bp_status": bp_status,
                "flash_count": flash_count,
                "more_count": more_count,
                "bp_positive_items": bp_positive_items,
                "bp_negative_items": bp_negative_items,
            })
            for issue_type, detail in bp_issues:
                issues_log.append({
                    "dept_idx": dept["dept_idx"],
                    "mgr_name": mgr["name_full"],
                    "issue_type": issue_type,
                    "detail": detail,
                })

        dept_overviews.append({
            "dept_idx": dept["dept_idx"],
            "dept_path": dept["path"],
            "n_managers": dept["n_managers"],
            "total_items": dept_total_items,
            "positive_voices": dept_total_pos,
            "negative_voices": dept_total_neg,
        })

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        pd.DataFrame(dept_overviews).to_excel(writer, sheet_name="部门概览", index=False)
        pd.DataFrame(mgr_summaries).to_excel(writer, sheet_name="管理者小计", index=False)
        pd.DataFrame(all_items).to_excel(writer, sheet_name="原文清单", index=False)
        pd.DataFrame(issues_log).to_excel(writer, sheet_name="异常日志", index=False)

    print(f"✅ 已生成: {output_file}")


if __name__ == "__main__":
    main()
