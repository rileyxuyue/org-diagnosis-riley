#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
为已有的 CDG / S3 HTML 报告补上 2.3.2 值得关注的细分项（敬满逐题数据）。
仅在 <h3>2.3.3 之前插入，不改其他内容。

用法:
  python3 patch_232_subdivision.py          # 处理 CDG + S3
  python3 patch_232_subdivision.py cdg      # 只处理 CDG
  python3 patch_232_subdivision.py s3       # 只处理 S3
"""

import os, sys, re, glob

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)
sys.path.insert(0, '.')

from generate_html_report import JingmanDetailLoader

BG_CONFIG = {
    'cdg': '报告/cdg/html',
    's3':  '报告/s3/html',
}


def extract_org_path(html_content):
    """从 HTML title 提取 org_full_path"""
    m = re.search(r'<title>.*?-\s*(.+?)\s*</title>', html_content)
    return m.group(1).strip() if m else None


def build_232_html(loader, org_path):
    """生成 2.3.2 章节的 HTML 片段，如果无数据则返回 None"""
    detail_row = loader.find_dept(org_path)
    if detail_row is None:
        return None

    subdiv_table = loader.build_subdivision_table(detail_row)
    if not subdiv_table:
        return None

    parts = []
    parts.append('''
        <h3>3.2 值得关注的细分项</h3>
        <p class="note-text">
            以下题目在部门排名与BG排名、或增幅之间存在显著差异（差值 &gt; 10）。
        </p>

        <table>
            <tr>
                <th class="th-blue">标题</th>
                <th class="th-blue">题目</th>
                <th class="th-blue">部门排名</th>
                <th class="th-blue">BG排名</th>
                <th class="th-blue">排名差</th>
                <th class="th-blue">增幅差</th>
            </tr>
''')

    current_group = None
    group_colors = {
        'bg_higher':   ('#FFF2CC', '#E65100'),
        'dept_higher': ('#E8F5E9', '#2E7D32'),
        'dept_lag':    ('#FCE4D6', '#C00000'),
        'dept_lead':   ('#E3F2FD', '#1565C0'),
    }

    for sd in subdiv_table:
        if sd['group'] != current_group:
            current_group = sd['group']
            g_bg, g_color = group_colors.get(current_group, ('#F5F5F5', '#333'))
            parts.append(f'''
            <tr>
                <td colspan="6" style="background: {g_bg}; color: {g_color}; font-weight: bold; font-size: 13px; padding: 8px 10px;">{sd['group_label']}</td>
            </tr>
''')
        rank_diff_val = sd['rank_diff']
        rd_style = ''
        if rank_diff_val > 10:
            rd_style = 'color: #C00000; font-weight: bold;'
        elif rank_diff_val < -10:
            rd_style = 'color: #2E7D32; font-weight: bold;'

        parts.append(f'''
            <tr>
                <td style="font-weight: bold; white-space: nowrap;">{sd['short']}</td>
                <td style="font-size: 13px;">{sd['full'][:50]}</td>
                <td style="text-align: center;">{sd['dept_rank']}</td>
                <td style="text-align: center;">{sd['bg_rank']}</td>
                <td style="text-align: center; {rd_style}">{rank_diff_val:+d}</td>
                <td style="text-align: center;">{sd['growth_diff_str']}</td>
            </tr>
''')

    parts.append('''
        </table>
''')
    return ''.join(parts)


def patch_bg(bg_key, loader):
    html_dir = BG_CONFIG[bg_key]
    if not os.path.isdir(html_dir):
        print(f'  ❌ 目录不存在: {html_dir}')
        return

    files = sorted(glob.glob(os.path.join(html_dir, '*.html')))
    print(f'\n{"="*55}')
    print(f'  {bg_key.upper()} — {len(files)} 个报告')
    print(f'{"="*55}')

    updated = 0
    skipped_exists = 0
    skipped_nodata = 0
    skipped_nomatch = 0

    for filepath in files:
        fname = os.path.basename(filepath)
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        # 已有 2.3.2 则跳过
        if '2.3.2' in content:
            skipped_exists += 1
            continue

        org_path = extract_org_path(content)
        if not org_path:
            skipped_nomatch += 1
            continue

        section_html = build_232_html(loader, org_path)
        if not section_html:
            skipped_nodata += 1
            continue

        # 在 <h3>3.3 之前插入
        insert_pos = content.find('<h3>3.3')
        if insert_pos < 0:
            # 没有2.3.3，试在footer前
            insert_pos = content.find('<div class="footer">')
        if insert_pos < 0:
            skipped_nomatch += 1
            continue

        new_content = content[:insert_pos] + section_html + '\n        ' + content[insert_pos:]

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)

        dept_short = org_path.split('/')[-1] if org_path else fname[:30]
        print(f'  ✅ {dept_short}')
        updated += 1

    print(f'\n  完成: 更新 {updated}, 已有 {skipped_exists}, 无数据 {skipped_nodata}, 无匹配 {skipped_nomatch}')


def main():
    args = [a.lower() for a in sys.argv[1:]]
    if not args:
        args = ['cdg', 's3']

    print('\n加载敬满逐题数据...')
    loader = JingmanDetailLoader(
        '【敬满】/report_tool_final/data/全量敬满数据.xlsx',
        '【敬满】/report_tool_final/data/BG相关数据.xlsx',
        '【敬满】/report_tool_final/data/题目与标题对照表.xlsx',
    )
    print(f'  ✓ {len(loader.dept_df)} 个部门, {len(loader.var_map)} 道题')

    for bg in args:
        if bg in BG_CONFIG:
            patch_bg(bg, loader)
        else:
            print(f'  ⚠ 未知BG: {bg}，支持: {", ".join(BG_CONFIG.keys())}')

    print('\n全部完成！\n')


if __name__ == '__main__':
    main()
