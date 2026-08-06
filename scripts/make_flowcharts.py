#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成交接文档用的流程图 PNG。"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from matplotlib import font_manager

names = [f.name for f in font_manager.fontManager.ttflist]
for cand in ["Arial Unicode MS", "Songti SC", "STHeiti"]:
    if cand in names:
        plt.rcParams["font.family"] = cand
        break
plt.rcParams["axes.unicode_minus"] = False

C_INPUT, C_INPUT_E = "#E8F1FA", "#5B9BD5"
C_A, C_A_E = "#E5F3E5", "#4CAF50"
C_B, C_B_E = "#FFF3E0", "#FF9800"
C_C, C_C_E = "#FDE8E8", "#E57373"
C_OUT, C_OUT_E = "#EDE7F6", "#7E57C2"
C_LLM, C_LLM_E = "#FFFDE7", "#FBC02D"


def box(ax, x, y, w, h, text, fc, ec, fs=9, bold=False):
    ax.add_patch(FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.008,rounding_size=0.02",
        facecolor=fc, edgecolor=ec, linewidth=1.6,
    ))
    if text:
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
                fontsize=fs, fontweight="bold" if bold else "normal",
                linespacing=1.5)


def arrow(ax, x1, y1, x2, y2, color="#555555", style="-|>", lw=1.6, ls="-"):
    ax.add_patch(FancyArrowPatch(
        (x1, y1), (x2, y2), arrowstyle=style, mutation_scale=14,
        color=color, linewidth=lw, linestyle=ls, shrinkA=2, shrinkB=2,
    ))


def label(ax, x, y, text, fs=8, color="#666666", ha="center"):
    ax.text(x, y, text, ha=ha, va="center", fontsize=fs, color=color,
            linespacing=1.5)


# ══════════════════════ 图1：总览 ══════════════════════
fig, ax = plt.subplots(figsize=(11.5, 7.6))
ax.set_xlim(0, 100)
ax.set_ylim(0, 68)
ax.axis("off")

ax.text(50, 65.5, "组织诊断报告系统 · 总览", ha="center", fontsize=15, fontweight="bold")
label(ax, 50, 62.6, "35 个 Python 文件，分成三条流水线", fs=9.5, color="#777777")

box(ax, 6, 53.5, 88, 6.4,
    "【原料仓库】5 类 Excel 表格\n组织诊断结果 · 组织架构信息 · 全面反馈 · 异动 · 敬满",
    C_INPUT, C_INPUT_E, fs=9.5, bold=True)

box(ax, 5, 33, 40, 15.5,
    "流水线 A —— 主线（9 个文件）\n\n把 Excel 数据做成\n「组织诊断报告」\n\n产物：网页版 + 文字版",
    C_A, C_A_E, fs=10, bold=True)
label(ax, 25, 30.3, "★ 日常干活就靠这条", fs=9, color="#2E7D32")

box(ax, 55, 33, 40, 15.5,
    "流水线 B —— 敬满专项（11 个文件）\n\n把敬满问卷数据做成\n「敬满分析报告」\n\n产物：Word 文档",
    C_B, C_B_E, fs=10, bold=True)
label(ax, 75, 30.3, "和 A 完全独立，互不干扰", fs=9, color="#E65100")

arrow(ax, 30, 53.5, 25, 48.5, color=C_A_E)
arrow(ax, 70, 53.5, 75, 48.5, color=C_B_E)

box(ax, 5, 19, 40, 6.6, "报告成品\n网页(html) + 文字(md)", C_OUT, C_OUT_E, fs=9.5, bold=True)
box(ax, 55, 19, 40, 6.6, "报告成品\nWord 文档(docx)", C_OUT, C_OUT_E, fs=9.5, bold=True)

arrow(ax, 25, 33, 25, 25.6, color=C_A_E)
arrow(ax, 75, 33, 75, 25.6, color=C_B_E)

box(ax, 5, 6, 40, 8.4,
    "流水线 C —— 修补车间（12 个文件）\n报告印出来发现哪里不对，\n单独去改那一小块，不重做整份",
    C_C, C_C_E, fs=9.5, bold=True)
arrow(ax, 20, 14.4, 20, 19, color=C_C_E, ls="--")
label(ax, 32.5, 16.6, "只改局部", fs=8, color="#C62828")

box(ax, 55, 6, 40, 8.4,
    "独立小工具（3 个文件）\n搭工作台 · 生成空白填写模板 · 报告脱敏\n不属于任何流水线，用到才拿出来",
    "#F5F5F5", "#9E9E9E", fs=9.5, bold=True)

ax.text(50, 1.6, "说明：A 和 B 都从同一个原料仓库取数，但产物不同、代码互不调用；C 依赖 A 的成品",
        ha="center", fontsize=8.5, color="#777777")

plt.tight_layout()
plt.savefig("flow_overview.png", dpi=200, bbox_inches="tight", facecolor="white")
plt.close()
print("OK flow_overview.png")


# ══════════════════════ 图2：主链路══════════════════════
fig, ax = plt.subplots(figsize=(12.2, 10.4))
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis("off")

ax.text(50, 97.5, "流水线 A · 主线 6 道工序", ha="center", fontsize=15, fontweight="bold")
label(ax, 50, 94.2, "从 Excel 原料，到最终的组织诊断报告", fs=9.5, color="#777777")

BX, BW = 19, 46
steps = [
    (81.0, 8.6, "第 1 道工序   extract_open_feedback.py",
     "把「组织诊断表 + 全面反馈表 + BP观察表」\n三本册子，照着人名对齐、抄成一本"),
    (70.0, 8.6, "第 2 道工序   step1_extract_bg.py",
     "把大杂烩的一段话，拆成一条条整齐的评价\n（谁的闪光点、谁的待改进，分门别类）"),
    (59.0, 8.6, "第 3 道工序   generate_source_md_with_llm.py",
     "请 AI 读完所有评价，写成一份总结稿\n（这就是报告里「1.3 开放题总结」那节）"),
    (46.5, 10.2, "第 4 道工序   generate_html_report.py",
     "★ 最核心的大厨\n把所有数据 + 上面的总结稿，\n汇总排版成一整份网页报告"),
    (34.5, 8.6, "第 5 道工序   html2md.py",
     "把网页版另存一份纯文字版\n（方便复制粘贴、发消息）"),
    (23.5, 8.6, "第 6 道工序   refresh_section_13.py",
     "只把「1.3 总结」那一节重新贴一遍\n其余章节一个字都不碰"),
]

for i, (y, h, title, desc) in enumerate(steps):
    box(ax, BX, y, BW, h, "", C_A, C_A_E)
    ax.text(BX + 1.8, y + h - 2.3, title, ha="left", va="center",
            fontsize=9.6, fontweight="bold", color="#1B5E20")
    ax.text(BX + 1.8, y + (h - 3.2) / 2, desc, ha="left", va="center",
            fontsize=9.0, linespacing=1.55)
    if i < len(steps) - 1:
        ny, nh = steps[i + 1][0], steps[i + 1][1]
        arrow(ax, BX + BW / 2, y, BX + BW / 2, ny + nh, color=C_A_E)

label(ax, 8.5, 91.0, "中间产物", fs=8.5, color="#888888")
for y, t in [(76.5, "全面反馈开放题\n{BG}.xlsx"),
             (65.5, "step1_output.xlsx\n（4 张表）"),
             (54.5, "output/{bg}/\n各部门总结.md")]:
    box(ax, 2.0, y, 13, 6.0, t, "#FAFAFA", "#BDBDBD", fs=7.8)
    arrow(ax, 15.0, y + 3.0, BX, y + 3.0, color="#BDBDBD", lw=1.2)

box(ax, 69.0, 41.0, 29, 12.0,
    "第 4 道工序还要读入：\n组织架构信息 · 岗位信息\n异动/离职明细 · 敬满数据\n干部侧/组织侧问题",
    C_INPUT, C_INPUT_E, fs=8.4)
arrow(ax, 69.0, 47.0, BX + BW, 49.5, color=C_INPUT_E, lw=1.3)

box(ax, 69.0, 60.5, 15.5, 4.6, "用到 AI", C_LLM, C_LLM_E, fs=8.6, bold=True)
arrow(ax, 69.0, 62.8, BX + BW, 62.8, color=C_LLM_E, lw=1.3, style="<|-")

box(ax, 69.0, 55.0, 15.5, 4.6, "用到 AI", C_LLM, C_LLM_E, fs=8.6, bold=True)
label(ax, 88.5, 57.3, "（算离职原因）", fs=7.6, color="#B28704")
arrow(ax, 69.0, 57.3, 66.5, 57.3, color=C_LLM_E, lw=1.3, style="<|-")
arrow(ax, 66.5, 57.3, 66.5, 53.5, color=C_LLM_E, lw=1.3)
arrow(ax, 66.5, 53.5, BX + BW, 53.5, color=C_LLM_E, lw=1.3)

box(ax, BX, 11.0, BW, 8.4,
    "最终成品\n报告/{bg}/html/*.html    （网页版）\n报告/{bg}/md/*.md          （文字版）",
    C_OUT, C_OUT_E, fs=9.4, bold=True)
arrow(ax, BX + BW / 2, 23.5, BX + BW / 2, 19.4, color=C_A_E)

arrow(ax, 8.5, 54.5, 8.5, 27.8, color="#9E9E9E", ls="--", lw=1.3)
arrow(ax, 8.5, 27.8, BX, 27.8, color="#9E9E9E", ls="--", lw=1.3)
label(ax, 8.5, 40.0, "第 6 道工序\n回头再读\n一次总结稿", fs=7.6, color="#888888")

box(ax, 6, 1.5, 88, 7.0,
    "两种跑法：\n① 常用：总结稿已经有了 → 只跑第 4、5、6 道工序（结果最稳，能复刻旧报告）\n"
    "② 完整：从头重做 → 六道工序全跑（会请 AI 重写总结稿，措辞可能和上次不完全一样）",
    "#FFFDE7", C_LLM_E, fs=8.8)

plt.tight_layout()
plt.savefig("flow_pipeline_a.png", dpi=200, bbox_inches="tight", facecolor="white")
plt.close()
print("OK flow_pipeline_a.png")
