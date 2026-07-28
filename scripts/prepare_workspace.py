#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 skill 自带的模板工程复制到目标目录。"""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_ROOT = SKILL_ROOT / "assets" / "project_template"


def main() -> None:
    parser = argparse.ArgumentParser(description="复制 org-diagnosis-riley 模板工程")
    parser.add_argument("target_dir", help="目标目录，例如 /tmp/org-diagnosis-workspace")
    parser.add_argument("--force", action="store_true", help="目标目录已存在时强制覆盖")
    args = parser.parse_args()

    target_dir = Path(args.target_dir).expanduser().resolve()
    if target_dir.exists():
        if any(target_dir.iterdir()) and not args.force:
            raise SystemExit(f"目标目录非空：{target_dir}，如需覆盖请加 --force")
        if args.force:
            shutil.rmtree(target_dir)

    shutil.copytree(TEMPLATE_ROOT, target_dir)
    print(f"✅ 已复制模板工程到: {target_dir}")
    print("下一步：")
    print("1. 用真实数据替换模板 xlsx / 输出 md")
    print("2. 在目标目录执行 `python3 run_report_pipeline.py --help` 查看流水线参数")


if __name__ == "__main__":
    main()
