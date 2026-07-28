#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基于 【全面反馈】/output/csig/*.md 中的 TL;DR + 第三步分维度详情，
更新 报告/csig/md/*.md 与 报告/csig/html/*.html 中的 1.3 开放题总结部分。

只更新 1.3 部分，其他章节绝对不动。
"""
import os
import re
from html import escape as html_escape

WORKSPACE = '/Users/xuyue/Desktop/workbuddy/task2 report'
SRC_DIR = os.path.join(WORKSPACE, '【全面反馈】/output/csig')
REPORT_MD_DIR = os.path.join(WORKSPACE, '报告/csig/md')
REPORT_HTML_DIR = os.path.join(WORKSPACE, '报告/csig/html')


# ======================================================================
# 1. 解析新MD源文件
# ======================================================================

def parse_source_md(path):
    """解析新MD源文件，返回 {tldr: [(dim, judge, summary), ...], dimensions: [...]} 结构"""
    with open(path, encoding='utf-8') as f:
        text = f.read()

    # === TL;DR 表格 ===
    tldr_match = re.search(r'## TL;DR\s*\n(.+?)(?=\n##|\Z)', text, re.S)
    tldr_rows = []
    if tldr_match:
        for line in tldr_match.group(1).splitlines():
            line = line.strip()
            if not line.startswith('|') or line.startswith('|---') or line.startswith('| ---') or '|------' in line:
                continue
            # 跳过表头
            if '维度' in line and '判断' in line and '一句话总结' in line:
                continue
            cells = [c.strip() for c in line.strip('|').split('|')]
            if len(cells) >= 3:
                dim = cells[0].strip().strip('*')
                judge = cells[1].strip()
                summary = cells[2].strip()
                tldr_rows.append((dim, judge, summary))

    # === 第一步：建 N -> source 映射（区分 bp / peer） ===
    mgr_n_source = {}
    # 用 ### === 管理者：分块
    parts = re.split(r'\n### === 管理者：', text)
    mgr_order = []
    for blk in parts[1:]:
        head_end = blk.find('===')
        if head_end < 0:
            continue
        header = blk[:head_end].strip()
        m = re.match(r'(.+?)[（(][^（()]*?(?:负责人|leader)[^（()]*?[）)]\s*$', header)
        if m:
            mgr_full = m.group(1).strip()
        else:
            # 去掉最后一个括号
            mgr_full = re.sub(r'[（(][^（()]*[）)]\s*$', '', header).strip()
        if mgr_full not in mgr_order:
            mgr_order.append(mgr_full)
        body = blk[head_end + 3:]
        # 取 [N## | source | polarity]
        for nm in re.finditer(r'\[N(\d+)\s*\|\s*([^|]+?)\s*\|\s*([^\]]+?)\]', body):
            nid = f'N{nm.group(1)}'
            src_name = nm.group(2).strip()
            is_bp = 'bp' in src_name.lower() or 'BP' in src_name
            mgr_n_source.setdefault(mgr_full, {})[nid] = 'bp' if is_bp else 'peer'

    # === 第三步：分维度详情 ===
    detail_match = re.search(r'## 第三步：分维度详情\s*\n(.+?)(?=\n## 报告说明|\Z)', text, re.S)
    dimensions = []
    if detail_match:
        detail = detail_match.group(1)
        # 按 ### 维度 切块
        dim_blocks = re.findall(r'### 维度[^\n]+\n(?:(?!^### 维度).)*', detail, re.S | re.M)
        for blk in dim_blocks:
            dim_obj = parse_dimension_block(blk, mgr_n_source)
            if dim_obj and not is_health_dim(dim_obj['name']):
                dimensions.append(dim_obj)

    # 同步过滤 TL;DR 中的健康维度
    tldr_rows = [row for row in tldr_rows if not is_health_dim(row[0])]

    return {
        'tldr': tldr_rows,
        'dimensions': dimensions,
        'mgr_order': mgr_order,
    }


def is_health_dim(dim_name):
    """识别健康维度（按 prompt2 规范应排除）"""
    keywords = ['身体健康', '健康与可持续', '健康与工作节奏', '健康与节奏', '健康可持续']
    return any(kw in dim_name for kw in keywords)


def parse_dimension_block(blk, mgr_n_source):
    """解析单个维度块，返回 dict
    {
        name: '维度一：xxx',
        definition: '一句话定义',
        judge: '共性优势',
        managers: [
            {
                name: 'fcding(丁川达)',
                pos_count: 5,
                pos_nids: ['N02', 'N05', 'bp N23', ...],
                pos_quotes: [(quote, nid_str, is_bp), ...],
                neg_count: 0,
                neg_nids: [],
                neg_quotes: [],
            },
            ...
        ]
    }
    """
    lines = blk.split('\n')
    if not lines:
        return None
    # 第一行: ### 维度一：xxx
    head = lines[0]
    m = re.match(r'### (维度[^：:]*?[:：]\s*.+)', head)
    if not m:
        return None
    name = m.group(1).strip()

    # 找定义行 (> 开头第一行) 和判断行 (**判断：xxx**)
    definition = ''
    judge = ''
    for line in lines[1:]:
        ls = line.strip()
        if ls.startswith('>') and not definition:
            df = ls.lstrip('>').strip()
            if df and not df.startswith('整体判断') and '在该维度无明显评价' not in df:
                definition = df
                continue
        jm = re.match(r'\*\*判断[：:]\s*(.+?)\*\*', ls)
        if jm:
            judge = jm.group(1).strip()
            break

    # 解析管理者条目: - **mgr_name**：正向 X 声音 (Nxx, ...) | 待关注 Y 声音 (...)
    managers = []
    blk_lines = blk.split('\n')
    i = 0
    while i < len(blk_lines):
        line = blk_lines[i]
        m = re.match(r'^- \*\*([^*]+?)\*\*[：:]\s*(.+)$', line)
        if not m:
            i += 1
            continue
        mgr_name = m.group(1).strip()
        rest = m.group(2).strip()
        # 解析正向/待关注
        pos_count, pos_nids = parse_voice_field(rest, '正向')
        neg_count, neg_nids = parse_voice_field(rest, '待关注')

        # 收集后续行的引用 (以 "  - " 或 "    - " 开头)
        # 遇到下一个 "- **" mgr 行才停止；中途的 ">" 注释/空行跳过
        pos_quotes = []
        neg_quotes = []
        j = i + 1
        while j < len(blk_lines):
            nl = blk_lines[j]
            # 顶级 - **xxx**: 下一个 mgr 块
            if re.match(r'^- \*\*', nl):
                break
            stripped = nl.strip()
            if not stripped:
                j += 1
                continue
            # > 注释 / 整体判断 — 跳过但不停
            if stripped.startswith('>'):
                j += 1
                continue
            # 分隔线 ———— 也跳过
            if re.match(r'^[-—=─_*]+\s*$', stripped):
                j += 1
                continue
            # 仅以缩进 - 开头的引用行才算
            if re.match(r'^\s+-\s', nl):
                quote_line = nl.lstrip()[1:].strip()
                quote_obj = parse_quote_line(quote_line, mgr_name, mgr_n_source)
                if quote_obj:
                    if quote_obj['polarity'] == 'pos':
                        pos_quotes.append(quote_obj)
                    else:
                        neg_quotes.append(quote_obj)
                j += 1
                continue
            # 其他不识别的行，停止该 mgr 的引用收集（保守）
            break

        managers.append({
            'name': mgr_name,
            'pos_count': pos_count,
            'pos_nids': pos_nids,
            'pos_quotes': pos_quotes,
            'neg_count': neg_count,
            'neg_nids': neg_nids,
            'neg_quotes': neg_quotes,
        })
        i = j

    return {
        'name': name,
        'definition': definition,
        'judge': judge,
        'managers': managers,
    }


def parse_voice_field(rest, key):
    """从 '正向 X 声音 (N01, N02) | 待关注 Y 声音 (N18)' 解析 X 和 nids 列表"""
    m = re.search(rf'{key}\s*(\d+)\s*声音?\s*\(([^)]*?)\)', rest)
    if m:
        count = int(m.group(1))
        nids = [n.strip() for n in m.group(2).split(',') if n.strip()]
        return count, nids
    # 不带括号
    m = re.search(rf'{key}\s*(\d+)\s*声音?', rest)
    if m:
        return int(m.group(1)), []
    return 0, []


def parse_quote_line(line, mgr_name, mgr_n_source):
    """解析单条引用行，例如:
       "正向引用"（N03，节选）
       待关注："xxx"（N18，节选）
       bp："xxx"（bp N23，节选）
       待关注：bp："xxx"（bp N30，节选）
    """
    polarity = 'pos'
    is_bp = False
    text = line.strip()

    if text.startswith('待关注'):
        polarity = 'neg'
        text = re.sub(r'^待关注[：:]\s*', '', text)
    if text.startswith('bp'):
        is_bp = True
        text = re.sub(r'^bp[：:]\s*', '', text)

    # 取最后的 (Nxx[，节选]) 作为标识
    tag_match = re.search(r'[（(](bp\s+)?(N\d+)(?:[,，]\s*节选)?[）)]\s*$', text)
    nid = ''
    if tag_match:
        if tag_match.group(1):
            is_bp = True
        nid = tag_match.group(2)
        text = text[:tag_match.start()].rstrip()

    # 去掉外部引号
    text = text.strip()
    text = re.sub(r'^[\u201c"](.*?)[\u201d"]?$', r'\1', text)
    text = text.strip().strip('"').strip('"').strip('"')

    # 验证 bp（用 mgr_n_source 复核）
    if nid and mgr_name in mgr_n_source:
        src = mgr_n_source[mgr_name].get(nid)
        if src == 'bp':
            is_bp = True

    return {
        'polarity': polarity,
        'is_bp': is_bp,
        'text': text,
        'nid': nid,
    }


# ======================================================================
# 2. 生成 MD 1.3 内容
# ======================================================================

def render_md_section(parsed, mgr_order_in_table=None):
    """生成 MD 形式的 1.3 节内容（不含 #### 1.3 开放题总结 这一行标题）

    返回字符串。
    """
    lines = []
    # === 整体判断表 ===
    lines.append('📊 最终整体判断（跨维度）')
    lines.append('')
    lines.append('| 维度 | 判断 | 一句话总结 |')
    lines.append('| --- | --- | --- |')
    for dim, judge, summary in parsed['tldr']:
        lines.append(f'| **{dim}** | {judge} | {summary} |')
    lines.append('')

    # === 各维度详情 ===
    # 维度顺序使用第三步顺序
    dim_name_to_chinese_idx = {}
    for i, d in enumerate(parsed['dimensions']):
        dim_name_to_chinese_idx[d['name']] = i

    # 决定 mgr 顺序 — 用第一步出现顺序
    if mgr_order_in_table is None:
        mgr_order_in_table = parsed['mgr_order']

    for d in parsed['dimensions']:
        # 标题（去掉 ### 维度N： 改为 维度N：）
        lines.append(d['name'])
        lines.append('')
        if d['judge']:
            lines.append(d['judge'])
            lines.append('')
        if d['definition']:
            lines.append(d['definition'])
            lines.append('')
        lines.append('**各管理者情况：**')
        lines.append('')
        lines.append('| 干部 | 正向 | 待关注 |')
        lines.append('| --- | --- | --- |')

        # 按 mgr_order_in_table 顺序输出（仅限该维度有发声的）
        mgr_in_dim = {m['name']: m for m in d['managers']}
        for mname in mgr_order_in_table:
            if mname not in mgr_in_dim:
                continue
            mg = mgr_in_dim[mname]
            pos_cell = format_cell_md(mg['pos_count'], mg['pos_quotes'])
            neg_cell = format_cell_md(mg['neg_count'], mg['neg_quotes'])
            lines.append(f'| {mname} | {pos_cell} | {neg_cell} |')
        lines.append('')

    return '\n'.join(lines).rstrip() + '\n'


def format_cell_md(count, quotes):
    """生成 MD 表格单元格内容: '5人  "..."  "..."' 或 '—'
    每个单元格最多展示 3 条引用（声音数仍按原始数）。
    """
    if count <= 0 and not quotes:
        return '—'
    head = f'{count}人' if count > 0 else f'{len(quotes)}人'
    parts = [head]
    for q in quotes[:3]:
        # 用全角引号包裹
        text = q['text'].strip()
        if not text:
            continue
        text = text.replace('|', '\\|')  # 表格内转义
        text = text.replace('\n', ' ')
        if q['is_bp']:
            text = f'{text} [From BP]'
        parts.append(f'"{text}"')
    return '  '.join(parts)


# ======================================================================
# 3. 生成 HTML 1.3 内容
# ======================================================================

JUDGE_STYLE = {
    '共性优势': {
        'color': '#389e0d',
        'bg': '#f6ffed',
        'border': '#b7eb8f',
        'tbl_border': '#b7eb8f',
    },
    '共性不足': {
        'color': '#cf1322',
        'bg': '#fff1f0',
        'border': '#ffa39e',
        'tbl_border': '#ffa39e',
    },
    '局部待关注': {
        'color': '#d48806',
        'bg': '#fffbe6',
        'border': '#ffe58f',
        'tbl_border': '#ffe58f',
    },
}


def get_judge_color(judge):
    return JUDGE_STYLE.get(judge.strip(), JUDGE_STYLE['局部待关注'])['color']


def render_html_section(parsed, mgr_order_in_table=None):
    """生成 HTML 形式的 1.3 节内容（含外层 wrapper div，不含 <h4> 标题）"""
    if mgr_order_in_table is None:
        mgr_order_in_table = parsed['mgr_order']

    parts = []
    parts.append(
        '<div style="background: #fff; border: 1px solid #e8e8e8; border-radius: 8px; '
        'padding: 20px 24px; margin-bottom: 20px; line-height: 1.7;">'
    )

    # === TL;DR 表格 ===
    parts.append('<div style="margin-bottom:16px;">')
    parts.append('<div style="font-size:14px;font-weight:700;color:#1a1a1a;margin-bottom:8px;">📊 最终整体判断（跨维度）</div>')
    parts.append(
        '<table style="width:100%;border-collapse:collapse;font-size:13px;line-height:1.6;margin:12px 0;">'
        '<thead><tr>'
        '<th style="padding:8px 12px;border:1px solid #d9d9d9;font-weight:700;text-align:left;background:#f5f5f5;color:#333;">维度</th>'
        '<th style="padding:8px 12px;border:1px solid #d9d9d9;font-weight:700;text-align:left;background:#f5f5f5;color:#333;">判断</th>'
        '<th style="padding:8px 12px;border:1px solid #d9d9d9;font-weight:700;text-align:left;background:#f5f5f5;color:#333;">一句话总结</th>'
        '</tr></thead><tbody>'
    )
    for dim, judge, summary in parsed['tldr']:
        jc = get_judge_color(judge)
        parts.append(
            f'<tr>'
            f'<td style="padding:8px 12px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:#333;"><b>{html_escape(dim)}</b></td>'
            f'<td style="padding:8px 12px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:{jc};font-weight:600">{html_escape(judge)}</td>'
            f'<td style="padding:8px 12px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:#333;">{html_escape(summary)}</td>'
            f'</tr>'
        )
    parts.append('</tbody></table>')
    parts.append('</div>')

    # === 各维度块 ===
    for d in parsed['dimensions']:
        style = JUDGE_STYLE.get(d['judge'], JUDGE_STYLE['局部待关注'])
        parts.append(
            f'<div style="margin:12px 0;padding:12px 16px;background:{style["bg"]};'
            f'border:2px solid {style["border"]};border-radius:10px;">'
        )
        parts.append(
            f'<div style="margin:0 0 6px;font-size:14px;font-weight:700;color:{style["color"]};">{html_escape(d["name"])}</div>'
        )
        parts.append(
            f'<div style="margin:6px 0 8px;padding:8px 10px;background:rgba(255,255,255,0.55);'
            f'border:1px dashed {style["border"]};border-radius:8px;">'
        )
        if d['judge']:
            parts.append(
                f'<div style="margin:0 0 4px;font-size:13px;line-height:1.7;color:#222;">'
                f'<span style="color:{style["color"]};font-weight:600;">{html_escape(d["judge"])}</span></div>'
            )
        if d['definition']:
            parts.append(
                f'<div style="margin:0 0 4px;font-size:13px;line-height:1.7;color:#222;">{html_escape(d["definition"])}</div>'
            )
        parts.append('</div>')
        parts.append('<div style="margin:12px 0 6px;font-size:13px;font-weight:700;color:#444;"><b>各管理者情况：</b></div>')
        parts.append(
            f'<table style="width:100%;border-collapse:collapse;font-size:12px;line-height:1.6;margin:8px 0;'
            f'border:2px solid {style["tbl_border"]};border-radius:4px;">'
            '<thead><tr>'
            '<th style="padding:8px 10px;border:1px solid #d9d9d9;font-weight:700;text-align:center;background:#f0f5ff;color:#333;white-space:nowrap;">干部</th>'
            '<th style="padding:8px 10px;border:1px solid #d9d9d9;font-weight:700;text-align:center;background:#f6ffed;color:#389e0d;white-space:nowrap;">正向</th>'
            '<th style="padding:8px 10px;border:1px solid #d9d9d9;font-weight:700;text-align:center;background:#fffbe6;color:#d48806;white-space:nowrap;">待关注</th>'
            '</tr></thead><tbody>'
        )

        mgr_in_dim = {m['name']: m for m in d['managers']}
        for mname in mgr_order_in_table:
            if mname not in mgr_in_dim:
                continue
            mg = mgr_in_dim[mname]
            pos_cell = format_cell_html(mg['pos_count'], mg['pos_quotes'], '#389e0d')
            neg_cell = format_cell_html(mg['neg_count'], mg['neg_quotes'], '#d48806')
            parts.append(
                f'<tr>'
                f'<td style="padding:8px 10px;border:1px solid #e8e8e8;vertical-align:top;font-weight:600;white-space:nowrap;text-align:center;">{html_escape(mname)}</td>'
                f'{pos_cell}'
                f'{neg_cell}'
                f'</tr>'
            )
        parts.append('</tbody></table>')
        parts.append('</div>')

    parts.append('</div>')
    return '\n'.join(parts)


def format_cell_html(count, quotes, color):
    """生成 HTML 表格单元格 — 每个单元格最多展示 3 条引用（声音数仍按原始数）"""
    if count <= 0 and not quotes:
        return '<td style="padding:8px 10px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:#ccc;">—</td>'
    head = f'{count}人' if count > 0 else f'{len(quotes)}人'
    inner = f'<div style="margin-bottom:6px;font-weight:700;font-size:11px;color:{color};">{head}</div>'
    for q in quotes[:3]:
        text = q['text'].strip()
        if not text:
            continue
        if q['is_bp']:
            text = f'{text} [From BP]'
        inner += f'<div style="margin:4px 0;padding-left:8px;border-left:2px solid #d9d9d9;color:#555;">"{html_escape(text)}"</div>'
    return f'<td style="padding:8px 10px;border:1px solid #e8e8e8;vertical-align:top;text-align:left;color:{color};">{inner}</td>'


# ======================================================================
# 4. 替换报告中的 1.3 部分
# ======================================================================

def update_report_md(report_path, new_section_md):
    """更新 MD 报告中的 1.3 部分"""
    with open(report_path, encoding='utf-8') as f:
        text = f.read()

    # 找 #### 1.3 开放题总结 ... 到 ## 2 异动 之间
    pattern = re.compile(r'(#### 1\.3 开放题总结\s*\n)(.+?)(?=\n## 2 异动|\n## 2\s|\Z)', re.S)
    m = pattern.search(text)
    if not m:
        return False, '未找到 1.3 开放题总结 标记'

    # 保留原 #### 1.3 开放题总结 标题
    new_text = text[:m.start()] + m.group(1) + '\n' + new_section_md.rstrip() + '\n\n' + text[m.end():]
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(new_text)
    return True, ''


def update_report_html(report_path, new_section_html):
    """更新 HTML 报告中的 1.3 部分"""
    with open(report_path, encoding='utf-8') as f:
        text = f.read()

    # 找 <h4>1.3 开放题总结</h4> 到 下一个 <h2> 或 <h3> 之前
    # HTML 中 1.3 后紧跟一个 <div>...</div> 容器，再后是 <h2>2 异动</h2>
    # 用宽松匹配：从 <h4>1.3 开放题总结</h4> 到 (?=<h2>2 异动)
    pattern = re.compile(r'(<h4>1\.3 开放题总结</h4>\s*)(.+?)(?=\s*<h2>2 异动|\s*<h2>\s*2[\s　]+异动)', re.S)
    m = pattern.search(text)
    if not m:
        return False, '未找到 <h4>1.3 开放题总结</h4> 标记或 <h2>2 异动</h2> 边界'

    new_text = text[:m.start()] + m.group(1) + '\n' + new_section_html + '\n        ' + text[m.end():]
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(new_text)
    return True, ''


# ======================================================================
# 5. 主流程
# ======================================================================

def find_report_files_for_dept(dept_name):
    """根据部门名找到对应的 md / html 报告文件 (短名版 + 全名版)"""
    md_files = []
    html_files = []
    for f in os.listdir(REPORT_MD_DIR):
        if not f.endswith('_组织诊断报告.md'):
            continue
        body = f[:-len('_组织诊断报告.md')]
        last = body.split('-')[-1]
        if last == dept_name:
            md_files.append(os.path.join(REPORT_MD_DIR, f))
    for f in os.listdir(REPORT_HTML_DIR):
        if not f.endswith('_组织诊断报告.html'):
            continue
        body = f[:-len('_组织诊断报告.html')]
        last = body.split('-')[-1]
        if last == dept_name:
            html_files.append(os.path.join(REPORT_HTML_DIR, f))
    return md_files, html_files


def main():
    src_files = sorted([f for f in os.listdir(SRC_DIR) if f.endswith('.md')])

    total = len(src_files)
    md_updated = 0
    html_updated = 0
    no_match = []
    errors = []

    for src_file in src_files:
        # 解析 idx_部门名.md
        m = re.match(r'(\d+)_(.+)\.md$', src_file)
        if not m:
            continue
        dept_idx = m.group(1)
        dept_name = m.group(2)

        src_path = os.path.join(SRC_DIR, src_file)
        try:
            parsed = parse_source_md(src_path)
        except Exception as e:
            errors.append(f'{src_file}: 解析失败 - {e}')
            continue

        if not parsed['tldr'] or not parsed['dimensions']:
            errors.append(f'{src_file}: TL;DR 或维度详情为空')
            continue

        # 找对应报告文件
        md_files, html_files = find_report_files_for_dept(dept_name)
        if not md_files and not html_files:
            no_match.append(dept_name)
            continue

        # 生成 MD 内容
        new_md_section = render_md_section(parsed)
        # 生成 HTML 内容
        new_html_section = render_html_section(parsed)

        for f in md_files:
            ok, err = update_report_md(f, new_md_section)
            if ok:
                md_updated += 1
            else:
                errors.append(f'{f}: {err}')

        for f in html_files:
            ok, err = update_report_html(f, new_html_section)
            if ok:
                html_updated += 1
            else:
                errors.append(f'{f}: {err}')

    print(f'源 MD 文件总数: {total}')
    print(f'MD 报告更新: {md_updated}')
    print(f'HTML 报告更新: {html_updated}')
    if no_match:
        print(f'无对应报告的部门 ({len(no_match)}): {no_match}')
    if errors:
        print(f'错误 ({len(errors)}):')
        for e in errors[:20]:
            print(f'  {e}')


if __name__ == '__main__':
    main()
