#!/usr/bin/env python3
"""批量修复 HTML 报告中敬满开放题部分：
删除"关键发现概要"累计人数 < 2 的维度（总览行 + 详情卡片），
并重新排序编号。
适用于 WXG 和 CSIG 的所有 HTML 报告。
"""

import os
import re
import glob
from html.parser import HTMLParser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_DIRS = [
    os.path.join(BASE_DIR, '报告', 'csig', 'html'),
    os.path.join(BASE_DIR, '报告', 'wxg', 'html'),
]

CN_NUMS = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
           '十一', '十二', '十三', '十四', '十五']


def extract_total_count_from_summary(summary_text):
    """从概要文本中提取所有 (N人) 的人数并累加"""
    matches = re.findall(r'\((\d+)人\)', summary_text)
    return sum(int(m) for m in matches)


def extract_total_count_from_card(card_html):
    """从卡片 HTML 中提取所有 <span class="jm-open-count ...">N人</span> 并累加"""
    matches = re.findall(r'<span\s+class="jm-open-count\s+\w+">(\d+)人</span>', card_html)
    return sum(int(m) for m in matches)


def fix_html_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        html = f.read()

    # 检查是否有敬满开放题部分
    if '<div class="jm-open-overview">' not in html:
        return None

    # ── 1. 解析总览表 ──
    # 提取整个 overview div
    overview_match = re.search(
        r'(<div class="jm-open-overview">.*?</table></div>)',
        html, re.DOTALL
    )
    if not overview_match:
        return None

    overview_html = overview_match.group(1)
    overview_start = overview_match.start()
    overview_end = overview_match.end()

    # 提取每一行 <tr>...</tr> (在 <tbody> 内)
    tbody_match = re.search(r'<tbody>(.*?)</tbody>', overview_html, re.DOTALL)
    if not tbody_match:
        return None

    rows = re.findall(r'<tr>(.*?)</tr>', tbody_match.group(1), re.DOTALL)

    # 解析每行：序号、维度名、类型、概要
    dim_rows = []
    for row_html in rows:
        # 提取维度名
        dim_name_match = re.search(r'class="jm-open-dim-name">(.*?)</td>', row_html, re.DOTALL)
        dim_name = dim_name_match.group(1).strip() if dim_name_match else ''

        # 提取概要文本
        summary_match = re.search(r'class="jm-open-finding-summary">(.*?)</td>', row_html, re.DOTALL)
        summary_text = summary_match.group(1).strip() if summary_match else ''

        # 计算累计人数
        total_count = extract_total_count_from_summary(summary_text)

        dim_rows.append({
            'dim_name': dim_name,
            'summary': summary_text,
            'total_count': total_count,
            'row_html': row_html,
        })

    # ── 2. 解析所有卡片 ──
    # 卡片格式: <div class="jm-open-card card-xxx">...</div></div>
    # 紧跟在 overview 之后
    after_overview = html[overview_end:]

    # 找出所有卡片
    cards = []
    card_pattern = re.compile(
        r'<div class="jm-open-card card-\w+">\s*'
        r'<div class="jm-open-card-header">【维度([一二三四五六七八九十百]+)】(.*?)</div>\s*'
        r'<div class="jm-open-card-body">(.*?)</div></div>',
        re.DOTALL
    )

    for m in card_pattern.finditer(after_overview):
        cn_num = m.group(1)
        card_dim_name = m.group(2).strip()
        card_body = m.group(3)
        card_full = m.group(0)
        total_count = extract_total_count_from_card(card_full)

        cards.append({
            'cn_num': cn_num,
            'dim_name': card_dim_name,
            'card_html': card_full,
            'total_count': total_count,
            'offset_start': overview_end + m.start(),
            'offset_end': overview_end + m.end(),
        })

    # ── 3. 判断哪些维度需要删除 ──
    # 用维度名匹配总览行和卡片，累计人数取两边最大值（应该一致）
    # 判断依据：卡片中的实际人数（更准确）
    dims_to_remove = set()

    # 先按卡片中实际 count 判断
    card_name_to_count = {}
    for card in cards:
        card_name_to_count[card['dim_name']] = card['total_count']

    for i, dr in enumerate(dim_rows):
        # 优先用卡片的 count，fallback 用概要的 count
        count = card_name_to_count.get(dr['dim_name'], dr['total_count'])
        if count < 2:
            dims_to_remove.add(dr['dim_name'])

    if not dims_to_remove:
        return None  # 没有需要删除的维度

    # ── 4. 构建新的总览表 ──
    kept_rows = [dr for dr in dim_rows if dr['dim_name'] not in dims_to_remove]

    new_tbody_rows = []
    for idx, dr in enumerate(kept_rows, 1):
        # 替换原来的序号
        new_row = re.sub(
            r'<td>\d+</td>',
            f'<td>{idx}</td>',
            '<tr>' + dr['row_html'] + '</tr>',
            count=1
        )
        new_tbody_rows.append(new_row)

    new_tbody = '\n'.join(new_tbody_rows)
    new_overview = re.sub(
        r'<tbody>.*?</tbody>',
        f'<tbody>\n{new_tbody}\n</tbody>',
        overview_html,
        flags=re.DOTALL
    )

    # ── 5. 构建新的卡片序列 ──
    kept_cards = [c for c in cards if c['dim_name'] not in dims_to_remove]

    new_cards_html = []
    for idx, card in enumerate(kept_cards):
        cn_num = CN_NUMS[idx] if idx < len(CN_NUMS) else str(idx + 1)
        # 替换卡片 header 中的维度编号
        new_card = re.sub(
            r'【维度[一二三四五六七八九十百]+】',
            f'【维度{cn_num}】',
            card['card_html'],
            count=1
        )
        new_cards_html.append(new_card)

    # ── 6. 重组 HTML ──
    # 找到从 overview 开始到最后一个卡片结束的范围
    if cards:
        section_end = cards[-1]['offset_end']
    else:
        section_end = overview_end

    # overview 之前 + 新 overview + 新卡片 + overview 到最后卡片之后的内容
    new_html = (
        html[:overview_start] +
        new_overview + '\n' +
        '\n'.join(new_cards_html) + '\n' +
        html[section_end:]
    )

    removed_count = len(dims_to_remove)
    remaining_count = len(kept_rows)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(new_html)

    return {
        'removed': sorted(dims_to_remove),
        'removed_count': removed_count,
        'remaining_count': remaining_count,
    }


def main():
    total_fixed = 0
    total_scanned = 0

    for report_dir in REPORT_DIRS:
        if not os.path.isdir(report_dir):
            print(f'目录不存在，跳过: {report_dir}')
            continue

        files = sorted(glob.glob(os.path.join(report_dir, '*.html')))
        dir_label = os.path.basename(os.path.dirname(report_dir)).upper()
        print(f'\n{"="*60}')
        print(f'扫描 {dir_label}: {report_dir}')
        print(f'共 {len(files)} 个 HTML 文件')
        print(f'{"="*60}')

        for fp in files:
            fname = os.path.basename(fp)
            total_scanned += 1
            result = fix_html_file(fp)
            if result:
                total_fixed += 1
                removed_list = ', '.join(result['removed'])
                print(f'  ✓ {fname}')
                print(f'    删除 {result["removed_count"]} 个维度 (剩余 {result["remaining_count"]}): {removed_list}')
            else:
                print(f'  · {fname} (无需修改)')

    print(f'\n{"="*60}')
    print(f'完成: 扫描 {total_scanned} 个文件, 修复 {total_fixed} 个')


if __name__ == '__main__':
    main()
