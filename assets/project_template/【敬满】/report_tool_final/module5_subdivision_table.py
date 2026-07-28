"""
module5_subdivision_table.py
细分项大表：四个分组，7列
前置计算：
  - 部门内排名：该题 _fav 在47道题分值中的排名（1=最高）
  - BG内排名：BG Excel 的 rank_2025
  - 部门增幅：_fav_growth
  - BG增幅：BG Excel 的 growth

四个分组：
  1. BG内排名 - 部门内排名 > 10（BG相对更好）
  2. 部门内排名 - BG内排名 > 10（部门相对更好）
  3. BG增幅 - 部门增幅 > 10%（部门增幅落后）
  4. 部门增幅 - BG增幅 > 10%（部门增幅领先）

7列：分组 / 标题 / 题目 / 部门内排名 / BG内排名 / 排名差值 / 增幅差值
"""

import pandas as pd
import numpy as np
from data_loader import load_all, get_dept_row, Q47, get_q_values

RANK_GAP_THRESHOLD   = 10    # 排名差值阈值
GROWTH_GAP_THRESHOLD = 10.0  # 增幅差值阈值（百分点）

GROUP_LABELS = {
    "bg_higher":   "BG排名较高的题",
    "dept_higher": "部门排名较高的题",
    "dept_lag":    "部门增幅落后BG",
    "dept_lead":   "部门增幅领先BG",
}


def build_subdivision_table_data(
    row: pd.Series,
    bg_df: pd.DataFrame,
    var_map: dict,
) -> list[dict]:
    """
    返回所有分组的题目列表，每项包含：
        group        分组key
        group_label  分组名称
        key          题目key
        short        题目简称
        full         题目全称
        dept_rank    部门内排名（数字）
        bg_rank      BG内排名（数字）
        rank_diff    排名差值 = 部门内排名 - BG内排名
        growth_diff  增幅差值 = 部门增幅 - BG增幅（百分点）
        growth_diff_str  增幅差值展示字符串
    """
    bg = str(row.get("所属bg", ""))

    # ── BG数据：取该BG的47道题数据，建立 key -> {bg_rank, bg_growth} ──
    bg_subset = bg_df[bg_df["bg"] == bg].set_index("question")

    # ── 部门内排名：47道题按 _fav 从高到低排名 ─────────────────────────
    fav_scores = {}
    for key in Q47:
        v = row.get(f"{key}_fav", np.nan)
        fav_scores[key] = v if not pd.isna(v) else -999

    # 排名：分值相同的题并列（使用 dense rank 逻辑）
    sorted_keys = sorted(fav_scores.keys(), key=lambda k: fav_scores[k], reverse=True)
    dept_rank_map = {}
    rank = 1
    for i, k in enumerate(sorted_keys):
        if i > 0 and fav_scores[k] < fav_scores[sorted_keys[i - 1]]:
            rank = i + 1
        dept_rank_map[k] = rank

    # ── 逐题计算，归入分组 ───────────────────────────────────────────
    results = []
    for key in Q47:
        vals        = get_q_values(row, key)
        dept_rank   = dept_rank_map.get(key, 99)
        dept_growth = vals["growth"]  # 已是百分点，如 8.15

        if key not in bg_subset.index:
            continue
        bg_row      = bg_subset.loc[key]
        bg_rank     = int(bg_row["rank_2025"])
        bg_growth   = bg_row["growth"] * 100  # BG Excel 的 growth 是小数，转成百分点

        rank_diff   = dept_rank - bg_rank
        if pd.isna(dept_growth):
            growth_diff = np.nan
        else:
            growth_diff = dept_growth - bg_growth

        info = var_map.get(key, {})

        def growth_diff_str(gd):
            if pd.isna(gd):
                return "-"
            sign = "+" if gd >= 0 else ""
            return f"{sign}{gd:.1f}%"

        base = {
            "key":          key,
            "short":        info.get("short", key),
            "full":         info.get("full", key),
            "dept_rank":    dept_rank,
            "bg_rank":      bg_rank,
            "rank_diff":    rank_diff,
            "growth_diff":  growth_diff,
            "growth_diff_str": growth_diff_str(growth_diff),
        }

        # 分组1：BG排名较高（BG内排名数字更小，即BG更好）
        if dept_rank - bg_rank > RANK_GAP_THRESHOLD:
            results.append({**base, "group": "bg_higher",
                             "group_label": GROUP_LABELS["bg_higher"]})

        # 分组2：部门排名较高（部门内排名数字更小，即部门更好）
        if bg_rank - dept_rank > RANK_GAP_THRESHOLD:
            results.append({**base, "group": "dept_higher",
                             "group_label": GROUP_LABELS["dept_higher"]})

        # 分组3：部门增幅落后BG
        if not pd.isna(growth_diff) and bg_growth - dept_growth > GROWTH_GAP_THRESHOLD:
            results.append({**base, "group": "dept_lag",
                             "group_label": GROUP_LABELS["dept_lag"]})

        # 分组4：部门增幅领先BG
        if not pd.isna(growth_diff) and dept_growth - bg_growth > GROWTH_GAP_THRESHOLD:
            results.append({**base, "group": "dept_lead",
                             "group_label": GROUP_LABELS["dept_lead"]})

    # 按分组顺序排序，组内按排名差值绝对值从大到小
    group_order = list(GROUP_LABELS.keys())
    results.sort(key=lambda x: (
        group_order.index(x["group"]),
        -abs(x["rank_diff"]) if x["group"] in ("bg_higher", "dept_higher")
        else -abs(x["growth_diff"]) if not pd.isna(x["growth_diff"]) else 0
    ))
    return results


# ============================================================
# 验证入口
# ============================================================
if __name__ == "__main__":
    dept_df, bg_df, var_map = load_all()

    for dept_name in dept_df["部门"]:
        row  = get_dept_row(dept_df, dept_name)
        bg   = row["所属bg"]
        data = build_subdivision_table_data(row, bg_df, var_map)

        print(f"\n{'='*80}")
        print(f"部门：{dept_name}（BG：{bg}）  共 {len(data)} 条")
        print(f"{'='*80}")

        if not data:
            print("  （无符合条件的题目）")
            continue

        current_group = None
        for r in data:
            if r["group"] != current_group:
                current_group = r["group"]
                print(f"\n  ── {r['group_label']} ──")
                print(f"  {'标题':<12} {'部门排名':>8} {'BG排名':>8} {'排名差':>8} {'增幅差':>10}  题目")
                print(f"  {'-'*75}")
            print(
                f"  {r['short']:<12} {r['dept_rank']:>8} {r['bg_rank']:>8} "
                f"{r['rank_diff']:>+8} {r['growth_diff_str']:>10}  {r['full'][:30]}"
            )
