"""
module4_bottom_table.py
排名靠后明细表：_fav_bg <= 10 的子题
列：标题 / 题目 / BG内排名（击败率）/ 分值变化（diff + growth）
排序：按 _fav_bg 从低到高（最差在最上）
"""

import pandas as pd
import numpy as np
from data_loader import (
    load_all, get_dept_row, Q47,
    get_q_values, fmt_pct, fmt_diff_growth
)

THRESHOLD = 10  # 末10%判断阈值


def build_bottom_table_data(row: pd.Series, var_map: dict) -> list[dict]:
    """
    返回末10%题目列表，每项包含：
        short       题目简称（remark）
        full        题目全称（remark2）
        fav_bg      击败率（数值，用于排序）
        bg_rank_str 击败率展示字符串
        yoy_change  分值变化展示字符串
    """
    results = []
    for key in Q47:
        vals = get_q_values(row, key)
        fav_bg = vals["fav_bg"]
        if pd.isna(fav_bg) or fav_bg > THRESHOLD:
            continue

        info = var_map.get(key, {})
        results.append({
            "key":          key,
            "short":        info.get("short", key),
            "full":         info.get("full", key),
            "fav_bg":       fav_bg,
            "bg_rank_str":  fmt_pct(fav_bg),
            "yoy_change":   fmt_diff_growth(vals["diff"], vals["growth"]),
        })

    # 按击败率从低到高排序（最差在最上）
    results.sort(key=lambda x: x["fav_bg"])
    return results


# ============================================================
# 验证入口
# ============================================================
if __name__ == "__main__":
    dept_df, bg_df, var_map = load_all()

    for dept_name in dept_df["部门"]:
        row = get_dept_row(dept_df, dept_name)
        bg  = row["所属bg"]
        data = build_bottom_table_data(row, var_map)

        print(f"\n{'='*70}")
        print(f"部门：{dept_name}（BG：{bg}）")
        print(f"末10%题目共 {len(data)} 道")
        print(f"{'='*70}")

        if not data:
            print("  （无末10%题目）")
            continue

        print(f"  {'标题':<12} {'BG内排名':>8} {'分值变化':>18}  题目")
        print(f"  {'-'*70}")
        for r in data:
            print(f"  {r['short']:<12} {r['bg_rank_str']:>8} {r['yoy_change']:>18}  {r['full']}")
