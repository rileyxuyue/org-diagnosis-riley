"""
module1_core_table.py
核心维度表：敬业度 / 满意度，各4列
"""

import pandas as pd
import numpy as np
from data_loader import (
    load_all, get_dept_row,
    fmt_pct, fmt_score, fmt_diff_growth, fmt_rank_change
)


def calc_core_row(row: pd.Series, metric: str) -> dict:
    """
    计算敬业度或满意度的核心维度表一行数据
    metric: '敬业度' 或 '满意度'
    返回 dict：
        in_bg_rank    在BG的排名（展示字符串）
        rank_change   排名变化（展示字符串）
        score         分值（展示字符串）
        yoy_change    较去年变化（展示字符串）
    """
    fav        = row.get(f"{metric}_fav",        np.nan)
    fav_bg     = row.get(f"{metric}_fav_bg",     np.nan)
    fav2024_bg = row.get(f"{metric}_fav2024_bg", np.nan)
    diff       = row.get(f"{metric}_fav_diff",   np.nan)
    growth     = row.get(f"{metric}_fav_growth", np.nan)

    return {
        "metric":       metric,
        "in_bg_rank":   fmt_pct(fav_bg),
        "rank_change":  fmt_rank_change(fav_bg, fav2024_bg),
        "score":        fmt_score(fav),
        "yoy_change":   fmt_diff_growth(diff, growth),
    }


def build_core_table_data(row: pd.Series) -> list[dict]:
    """返回两行：[敬业度行, 满意度行]"""
    return [
        calc_core_row(row, "敬业度"),
        calc_core_row(row, "满意度"),
    ]


# ============================================================
# 验证入口
# ============================================================
if __name__ == "__main__":
    dept_df, bg_df, var_map = load_all()

    for dept_name in dept_df["部门"]:
        row = get_dept_row(dept_df, dept_name)
        bg  = row["所属bg"]
        print(f"\n{'='*60}")
        print(f"部门：{dept_name}（BG：{bg}）")
        print(f"{'='*60}")
        print(f"  {'指标':<6} {'在BG排名':>10} {'排名变化':>20} {'分值':>8} {'较去年变化':>20}")
        print(f"  {'-'*66}")
        for r in build_core_table_data(row):
            print(f"  {r['metric']:<6} {r['in_bg_rank']:>10} {r['rank_change']:>20} {r['score']:>8} {r['yoy_change']:>20}")
