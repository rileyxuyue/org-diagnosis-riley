"""
data_loader.py
负责：读取三个Excel、查找部门、定义47道子题、格式化工具函数
"""

import pandas as pd
import numpy as np
from config import DEPT_FILE, BG_FILE, VAR_FILE


# ============================================================
# 47道子题（带 _q、不含维度总分）
# 顺序：先敬业度4道，再满意度43道
# ============================================================
Q47 = [
    # 敬业度子题（4道）
    "say_q1", "stay_q1", "stay_q2", "strive_q1",
    # gb（9道）
    "gb_中干_q1",
    "gb_直接上级_q1", "gb_直接上级_q2", "gb_直接上级_q3", "gb_直接上级_q4",
    "gb_直接上级_q5", "gb_直接上级_q6", "gb_直接上级_q7", "gb_直接上级_q8",
    # gj（5道）
    "gj_job_q1", "gj_job_q2", "gj_job_q3", "gj_job_q4", "gj_job_q5",
    # gr（5道）
    "gr_绩效_q1",
    "gr_薪酬_q1", "gr_薪酬_q2",
    "gr_福利_q1",
    "gr_晋升_q1",
    # gc（24道）
    "gc_卓越团队_q1", "gc_卓越团队_q2", "gc_卓越团队_q3",
    "gc_卓越团队_q4", "gc_卓越团队_q5", "gc_卓越团队_q6",
    "gc_协作信任_q1", "gc_协作信任_q2",
    "gc_工作支持_q1", "gc_工作支持_q2", "gc_工作支持_q3",
    "gc_文化价值观_q1", "gc_文化价值观_q2", "gc_文化价值观_q3",
    "gc_客户导向_q1", "gc_客户导向_q2",
    "gc_创造_q1",
    "gc_沟通渠道_q1", "gc_沟通渠道_q2",
    "gc_公司未来_q1", "gc_公司未来_q2",
    "gc_组织活力_q1",
    "gc_人才管理_q1",
    "gc_多样性_q1",
]

# 敬业度子题 key 集合（用于家族判断）
ENG_Q_KEYS = {"say_q1", "stay_q1", "stay_q2", "strive_q1"}

# Great Boss 题：gb_ 开头
def is_gb(key: str) -> bool:
    return key.startswith("gb_")

# 家族判断
def family(key: str) -> str:
    """返回 '敬业度' 或 '满意度'"""
    if key in ENG_Q_KEYS:
        return "敬业度"
    return "满意度"


# ============================================================
# 数据加载
# ============================================================

def load_all():
    """
    返回：
        dept_df   - 部门数据 DataFrame
        bg_df     - BG数据 DataFrame
        var_map   - dict: key -> {short, category, full}
    """
    print("📂 正在加载数据...")

    dept_df = pd.read_excel(DEPT_FILE)
    bg_df   = pd.read_excel(BG_FILE)
    var_df  = pd.read_excel(VAR_FILE)

    # 构建题目对照字典
    var_map = {}
    for _, row in var_df.iterrows():
        var_map[str(row["key"])] = {
            "short":    str(row["remark"]),
            "category": str(row["remark1"]),
            "full":     str(row["remark2"]),
        }

    print(f"  ✅ 部门数据：{len(dept_df)} 个部门，{len(dept_df.columns)} 列")
    print(f"  ✅ BG数据：{len(bg_df)} 行")
    print(f"  ✅ 题目对照表：{len(var_map)} 道题")
    return dept_df, bg_df, var_map


def get_dept_row(dept_df: pd.DataFrame, dept_name: str) -> pd.Series:
    """按部门名称查找，找不到返回 None"""
    rows = dept_df[dept_df["部门"] == dept_name]
    if rows.empty:
        return None
    return rows.iloc[0]


# ============================================================
# 格式化工具函数
# ============================================================

def fmt_pct(v, default="-"):
    """击败率/百分比：42.0 → '42%'"""
    if pd.isna(v):
        return default
    return f"{v:.0f}%"

def fmt_score(v, default="-"):
    """分值：82.1 → '82.1'"""
    if pd.isna(v):
        return default
    return f"{v:.1f}"

def fmt_diff_growth(diff, growth, default="-"):
    """
    分值变化 + 增幅：+4.1（+5.3%）
    diff   = _fav_diff（如 4.1）
    growth = _fav_growth（已是百分比数值，如 8.15 表示 +8.15%）
    """
    if pd.isna(diff) and pd.isna(growth):
        return default
    parts = []
    if not pd.isna(diff):
        sign = "+" if diff >= 0 else ""
        parts.append(f"{sign}{diff:.1f}")
    if not pd.isna(growth):
        sign = "+" if growth >= 0 else ""
        parts.append(f"（{sign}{growth:.1f}%）")
    return "".join(parts)

def fmt_rank_change(current_bg, prev_bg, default="/"):
    """
    排名变化：+15%（去年57%）
    current_bg = _fav_bg（今年击败率）
    prev_bg    = _fav2024_bg（去年击败率）
    若去年数据缺失，返回 "/"
    """
    if pd.isna(current_bg):
        return default
    if pd.isna(prev_bg):
        return "/"
    change = current_bg - prev_bg
    sign = "+" if change >= 0 else ""
    return f"{sign}{change:.0f}%（去年{prev_bg:.0f}%）"


# ============================================================
# 子题数据提取
# ============================================================

def get_q_values(row: pd.Series, key: str) -> dict:
    """
    提取某道题的所有关键数值：
      fav         今年正向比例
      fav2024     去年正向比例
      fav_bg      今年击败率（0-100）
      fav2024_bg  去年击败率
      diff        分值变化（如 +4.1）
      growth      增幅%（如 8.15 表示 +8.15%）
    列名格式：{key}_fav_diff / {key}_fav_growth
    """
    def v(col):
        val = row.get(col, np.nan)
        return np.nan if pd.isna(val) else val

    return {
        "fav":        v(f"{key}_fav"),
        "fav2024":    v(f"{key}_fav2024"),
        "fav_bg":     v(f"{key}_fav_bg"),
        "fav2024_bg": v(f"{key}_fav2024_bg"),
        "diff":       v(f"{key}_fav_diff"),
        "growth":     v(f"{key}_fav_growth"),
    }


# ============================================================
# 验证入口（直接运行此文件时执行）
# ============================================================

if __name__ == "__main__":
    dept_df, bg_df, var_map = load_all()

    print(f"\n📋 所有部门：")
    for name in dept_df["部门"]:
        print(f"   {name}")

    # 用第一个部门验证
    test_dept = dept_df["部门"].iloc[0]
    row = get_dept_row(dept_df, test_dept)
    print(f"\n🔍 验证部门：{test_dept}（所属BG：{row['所属bg']}）")

    print(f"\n📊 47道子题验证（共{len(Q47)}道）：")
    print(f"   {'题目key':<30} {'fav':>8} {'fav_bg':>8} {'diff':>8} {'growth':>8}")
    print(f"   {'-'*62}")
    for key in Q47:
        vals = get_q_values(row, key)
        fav_str    = fmt_score(vals['fav'])
        bg_str     = fmt_pct(vals['fav_bg'])
        diff_str   = fmt_diff_growth(vals['diff'], vals['growth'])
        print(f"   {key:<30} {fav_str:>8} {bg_str:>8} {diff_str:>16}")

    print(f"\n✅ 格式化示例：")
    row0 = get_dept_row(dept_df, test_dept)
    print(f"   敬业度 在BG排名：{fmt_pct(row0.get('敬业度_fav_bg'))}")
    print(f"   敬业度 排名变化：{fmt_rank_change(row0.get('敬业度_fav_bg'), row0.get('敬业度_fav2024_bg'))}")
    print(f"   敬业度 分值：    {fmt_score(row0.get('敬业度_fav'))}")
    print(f"   敬业度 变化：    {fmt_diff_growth(row0.get('敬业度_fav_diff'), row0.get('敬业度_fav_growth'))}")
