"""
generate_report.py
主入口：输入部门名称 → 生成完整 docx 报告

用法：
    python3 generate_report.py "KFC/战略发展部"
    python3 generate_report.py --all       # 生成所有部门
"""

import sys
import os
import json
import subprocess
import tempfile
import datetime
from data_loader import load_all, get_dept_row, Q47
from module1_core_table import build_core_table_data
from module2_risk_table import build_risk_table_data
from module3_chart import make_chart
from module4_bottom_table import build_bottom_table_data
from module5_subdivision_table import build_subdivision_table_data


def generate_report(dept_name: str, dept_df, bg_df, var_map) -> str:
    """
    生成单个部门报告，返回输出文件路径
    """
    row = get_dept_row(dept_df, dept_name)
    if row is None:
        print(f"❌ 找不到部门：{dept_name}")
        return None

    bg = str(row["所属bg"])
    today = datetime.date.today().strftime("%Y年%m月%d日")
    print(f"\n📄 生成报告：{dept_name}（{bg}）")

    # ── 1. 生成图表 PNG ────────────────────────────────────────
    print("   [1/6] 生成柱状图...")
    chart_buf = make_chart(row, dept_df)
    chart_b64 = __import__("base64").b64encode(chart_buf.read()).decode()

    # ── 2. 收集各模块数据 ─────────────────────────────────────
    print("   [2/6] 计算核心维度表...")
    core_data = build_core_table_data(row)

    print("   [3/6] 计算风险区间表...")
    risk_data = build_risk_table_data(row)

    print("   [4/6] 计算排名靠后明细表...")
    bottom_data = build_bottom_table_data(row, var_map)

    print("   [5/6] 计算细分项大表...")
    subdiv_data = build_subdivision_table_data(row, bg_df, var_map)

    # ── 3. 组装 payload 传给 Node.js ──────────────────────────
    payload = {
        "dept_name":   dept_name,
        "bg":          bg,
        "today":       today,
        "chart_b64":   chart_b64,
        "core_data":   core_data,
        "risk_data":   risk_data,
        "bottom_data": bottom_data,
        "subdiv_data": subdiv_data,
    }

    safe_name = dept_name.replace("/", "_")
    out_path  = os.path.join("output", f"报告_{safe_name}.docx")

    print("   [6/6] 生成 Word 文档...")
    import math
    import numpy as np

    def clean(obj):
        """递归将 NaN/numpy 类型转成 JSON 兼容类型"""
        if isinstance(obj, dict):
            return {k: clean(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [clean(v) for v in obj]
        if isinstance(obj, float) and math.isnan(obj):
            return None
        if isinstance(obj, (np.floating, np.integer)):
            v = obj.item()
            return None if isinstance(v, float) and math.isnan(v) else v
        return obj

    payload_json = json.dumps(clean(payload), ensure_ascii=False)

    result = subprocess.run(
        ["node", "report_builder.js", out_path],
        input=payload_json,
        capture_output=True,
        text=True,
        cwd=os.path.dirname(os.path.abspath(__file__))
    )

    if result.returncode != 0:
        print(f"❌ 生成失败：\n{result.stderr}")
        return None

    print(f"   ✅ 已保存：{out_path}")
    return out_path


# ── 主入口 ─────────────────────────────────────────────────────
if __name__ == "__main__":
    os.makedirs("output", exist_ok=True)
    dept_df, bg_df, var_map = load_all()

    if len(sys.argv) < 2 or sys.argv[1] == "--all":
        targets = list(dept_df["部门"])
    else:
        targets = [sys.argv[1]]

    results = []
    for dept in targets:
        path = generate_report(dept, dept_df, bg_df, var_map)
        if path:
            results.append(path)

    print(f"\n✅ 全部完成，共生成 {len(results)} 份报告")
    for p in results:
        print(f"   {p}")
