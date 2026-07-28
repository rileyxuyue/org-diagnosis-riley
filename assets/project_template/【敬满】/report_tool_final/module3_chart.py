"""
module3_chart.py
柱状图：47道子题击败率双向柱状图
  - X轴：47道题，按 _fav_bg 从高到低排序（靠前在左）
  - Y轴：(_fav_bg - 50) / 100，中线=0
  - 方向：>= 50 向上（蓝），< 50 向下（橙/红）
  - 颜色：蓝>=50，橙20<x<50，红<=20
  - 三条红线：倒数20%(-0.30)、倒数10%(-0.40)、倒数第一（动态）
  - 动态标题与X轴标注
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io
from data_loader import load_all, get_dept_row, Q47, family

# 中文字体设置
from matplotlib import font_manager
font_path = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
font_manager.fontManager.addfont(font_path)
prop = font_manager.FontProperties(fname=font_path)
plt.rcParams["font.family"] = prop.get_name()
plt.rcParams["axes.unicode_minus"] = False

# 颜色常量
COLOR_BLUE   = "#4472C4"   # 击败率 >= 50
COLOR_ORANGE = "#F4A460"   # 20 < 击败率 < 50
COLOR_RED    = "#C00000"   # 击败率 <= 20
COLOR_LINE   = "#FF0000"   # 红色基准线


def get_bg_min_fav_bg(dept_df: pd.DataFrame, bg: str) -> float:
    """
    找出该BG所有部门中，47道子题 _fav_bg 的最小值
    用于「倒数第一」红线位置
    """
    bg_df = dept_df[dept_df["所属bg"] == bg]
    min_val = 100.0
    for key in Q47:
        col = f"{key}_fav_bg"
        if col in bg_df.columns:
            vals = bg_df[col].dropna()
            if not vals.empty:
                min_val = min(min_val, vals.min())
    return min_val


def make_chart(row: pd.Series, dept_df: pd.DataFrame) -> io.BytesIO:
    """
    生成双向柱状图，返回 BytesIO（PNG格式）
    """
    bg = str(row.get("所属bg", ""))

    # ── 收集47道题数据 ──────────────────────────────
    items = []
    for key in Q47:
        fav_bg = row.get(f"{key}_fav_bg", np.nan)
        if pd.isna(fav_bg):
            fav_bg = 0.0
        items.append({"key": key, "fav_bg": fav_bg, "family": family(key)})

    # 按击败率从高到低排序
    items.sort(key=lambda x: x["fav_bg"], reverse=True)

    # ── 统计数字（动态标题用）──────────────────────
    n_front     = sum(1 for it in items if it["fav_bg"] >= 50)
    n_sat       = sum(1 for it in items if it["family"] == "满意度")
    n_eng       = sum(1 for it in items if it["family"] == "敬业度")
    bg_min      = get_bg_min_fav_bg(dept_df, bg)

    # ── Y轴值转换：(fav_bg - 50) / 100 ────────────
    x_pos   = list(range(len(items)))
    y_vals  = [(it["fav_bg"] - 50) / 100 for it in items]
    colors  = []
    for it in items:
        if it["fav_bg"] >= 50:
            colors.append(COLOR_BLUE)
        elif it["fav_bg"] > 20:
            colors.append(COLOR_ORANGE)
        else:
            colors.append(COLOR_RED)

    # ── 绘图 ───────────────────────────────────────
    fig, ax = plt.subplots(figsize=(11, 7))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # 柱子（基准线=0，向上/向下自动由正负决定）
    ax.bar(x_pos, y_vals, color=colors, width=0.6, zorder=2)

    # 中线（50%分界）
    ax.axhline(y=0, color="#1F3864", linewidth=1.8, zorder=3)

    # 两条红色基准线
    y_b20 = (20 - 50) / 100   # -0.30
    y_b10 = (10 - 50) / 100   # -0.40

    line_cfg = dict(color=COLOR_LINE, linewidth=1.5, linestyle="-", zorder=3)
    ax.axhline(y=y_b20, **line_cfg)
    ax.axhline(y=y_b10, **line_cfg)

    # 红线标签（右侧）
    x_label = len(items) - 0.3
    ax.text(x_label, y_b20 + 0.012, "倒数20%", color=COLOR_LINE,
            fontsize=11, va="bottom", ha="left")
    ax.text(x_label, y_b10 + 0.012, "倒数10%", color=COLOR_LINE,
            fontsize=11, va="bottom", ha="left")

    # Y轴
    ax.set_ylabel("")
    ytick_rates = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
    ax.set_yticks([(r - 50) / 100 for r in ytick_rates])
    ax.set_yticklabels([str(r) for r in ytick_rates], fontsize=12)
    ax.tick_params(axis="y", labelsize=12)

    # X轴
    ax.set_xticks([])
    ax.set_xlim(-0.8, len(items) + 2.5)
    ax.set_ylim(-0.62, 0.65)

    # X轴底部标注
    ax.text(len(items) / 2, -0.60,
            f"{n_sat}道满意度题  +  {n_eng}道敬业度题",
            ha="center", va="bottom", fontsize=12, color="#1F3864")

    # 标题
    ax.set_title(
        f"该部门47道敬满子题仅{n_front}题排名靠前",
        fontsize=14, fontweight="bold", color="#1F3864", pad=16
    )

    # 图例
    legend_handles = [
        mpatches.Patch(color=COLOR_BLUE,   label="排名靠前（击败率 ≥ 50%）"),
        mpatches.Patch(color=COLOR_ORANGE, label="排名靠后（20% ＜ 击败率 ＜ 50%）"),
        mpatches.Patch(color=COLOR_RED,    label="高危（击败率 ≤ 20%）"),
    ]
    ax.legend(handles=legend_handles, loc="upper right",
              fontsize=11, framealpha=0.9, edgecolor="#CCCCCC")

    # 去掉多余边框
    for spine in ["top", "right", "bottom"]:
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color("#AAAAAA")
    ax.spines["left"].set_linewidth(1.2)

    plt.tight_layout()

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor="white")
    plt.close()
    buf.seek(0)
    return buf


# ============================================================
# 验证入口：生成所有部门的图表并保存为PNG
# ============================================================
if __name__ == "__main__":
    import os
    dept_df, bg_df, var_map = load_all()
    os.makedirs("output", exist_ok=True)

    for dept_name in dept_df["部门"]:
        row = get_dept_row(dept_df, dept_name)
        safe = dept_name.replace("/", "_")
        print(f"  生成图表：{dept_name} ...")
        buf = make_chart(row, dept_df)
        path = f"output/chart_{safe}.png"
        with open(path, "wb") as f:
            f.write(buf.read())
        print(f"  ✅ 已保存：{path}")
