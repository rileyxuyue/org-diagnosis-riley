#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""在模板工程内串起最终报告生成链路。"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def run(cmd: list[str], cwd: Path | None = None) -> None:
    print('>>>', ' '.join(cmd))
    subprocess.check_call(cmd, cwd=str(cwd or ROOT))


def main() -> None:
    parser = argparse.ArgumentParser(description='组织诊断报告流水线总控')
    parser.add_argument('--bgs', nargs='*', default=['all'], help='BG 列表，如 CSIG TEG；默认 all')
    parser.add_argument('--steps', nargs='*', default=['build-html', 'convert-md', 'refresh-13'],
                        choices=['extract-open-feedback', 'step1-extract', 'generate-source-md', 'build-html', 'convert-md', 'refresh-13'],
                        help='要执行的步骤')
    parser.add_argument('--sample-md', default='【全面反馈】/output/examples/14_示例部门B.md', help='LLM 生成源 md 时的样例文件')
    parser.add_argument('--prompt-file', default='【全面反馈】/output/报告生成_Prompt2.md', help='LLM 生成源 md 时的 prompt 文件')
    args = parser.parse_args()

    bg_list = ['CDG', 'CSIG', 'IEG', 'PCG', 'TEG', 'WXG', 'S1', 'S2', 'S3', 'OFS'] if args.bgs == ['all'] else [b.upper() for b in args.bgs]

    if 'extract-open-feedback' in args.steps:
        run([sys.executable, '【全面反馈】/extract_open_feedback.py', *bg_list])
    if 'step1-extract' in args.steps:
        for bg in bg_list:
            run([sys.executable, '【全面反馈】/step1_extract_bg.py', '--bg', bg])
    if 'generate-source-md' in args.steps:
        for bg in bg_list:
            run([sys.executable, '【全面反馈】/generate_source_md_with_llm.py', '--bg', bg, '--prompt-file', args.prompt_file, '--sample-md', args.sample_md])
    if 'build-html' in args.steps:
        run([sys.executable, 'generate_html_report.py', '--batch', *bg_list])
    if 'convert-md' in args.steps:
        run([sys.executable, 'html2md.py', '--bgs', *bg_list])
    if 'refresh-13' in args.steps:
        run([sys.executable, 'refresh_section_13.py', '--bgs', *bg_list])

    print('✅ 流水线完成')


if __name__ == '__main__':
    main()
