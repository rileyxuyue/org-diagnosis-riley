"""
module2_risk_table.py
风险区间表：BG末20% / BG末10%，各5列
  - 风险区间（固定文字）
  - 题数（自己计算：47道子题中 _fav_bg <= 阈值 的数量）
  - 占比（直接读取 pct_bottom20 / pct_bottom10）
  - Great Boss 题数（gb_ 开头）
  - 其他题数（题数 - Great Boss）
"""

import pandas as pd
import numpy as np
from data_loader import load_all, get_dept_row, Q47, is_gb, get_q_values


def calc_risk_row(row: pd.Series, threshold: int) -> dict:
    """
    计算某个风险区间的一行数据
    threshold: 20 或 10
    """
    # 题数：自己从47道子题的 _fav_bg 计算
    count = 0
    count_gb = 0
    for key in Q47:
        fav_bg = row.get(f"{key}_fav_bg", np.nan)
        if pd.isna(fav_bg):
            continue
        if fav_bg <= threshold:
            count += 1
            if is_gb(key):
                count_gb += 1

    count_other = count - count_gb

    # 占比：直接读取 pct_bottom 字段
    pct_col = f"pct_bottom{threshold}"
    pct_val = row.get(pct_col, np.nan)
    if not pd.isna(pct_val):
        pct_str = f"{pct_val:.0f}%"
    else:
        pct_str = "-"

    return {
        "zone":        f"BG末{threshold}%",
        "count":       count,
        "pct":         pct_str,
        "count_gb":    count_gb,
        "count_other": count_other,
    }


def build_risk_table_data(row: pd.Series) -> list[dict]:
    """返回两行：[末20%行, 末10%行]"""
    return [
        calc_risk_row(row, 20),
        calc_risk_row(row, 10),
    ]


# ============================================================
# 验证入口
# ============================================================
if __name__ == "__main__":
    dept_df, bg_df, var_map = load_all()

    for dept_name in dept_df["部门"]:
        row = get_dept_row(dept_df, dept_name)
        bg  = row["所属bg"]
        print(f"\n{'='*55}")
        print(f"部门：{dept_name}（BG：{bg}）")
        print(f"{'='*55}")
        print(f"  {'风险区间':<10} {'题数':>6} {'占比':>8} {'Great Boss':>12} {'其他':>6}")
        print(f"  {'-'*44}")
        for r in build_risk_table_data(row):
            print(f"  {r['zone']:<10} {r['count']:>6} {r['pct']:>8} {r['count_gb']:>12} {r['count_other']:>6}")

        # 同时列出末10%的具体题目，方便核对
        print(f"\n  末10%具体题目（_fav_bg <= 10）：")
        for key in Q47:
            fav_bg = row.get(f"{key}_fav_bg", np.nan)
            if not pd.isna(fav_bg) and fav_bg <= 10:
                gb_tag = "【GB】" if is_gb(key) else "      "
                print(f"    {gb_tag} {key:<30} fav_bg={fav_bg:.1f}%")
