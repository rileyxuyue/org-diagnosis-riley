#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
更新单个部门报告的 1.3 开放题总结部分
从【全面反馈】/output/{bg}/{部门名}.md 读取分析结果，
转换为 cadre_to_html 期望的格式，替换 HTML 报告中的 1.3 section。

用法:
  python3 update_section_13.py <html文件路径> <md文件路径>
"""
import os
import re
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from generate_html_report import OpenFeedbackLoader


def md_to_cadre_text(md_path):
    """从 output md 文件提取内容，转换为 cadre_to_html 期望的纯文本格式。"""
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    lines = [l.rstrip('\n') for l in lines]

    # 定位各section
    tldr_start = -1
    conclusion_start = -1
    step3_start = -1
    for i, line in enumerate(lines):
        s = line.strip()
        if s == '## TL;DR':
            tldr_start = i
        elif s.startswith('**总体结论**'):
            conclusion_start = i
        elif s == '## 第三步：分维度详情':
            step3_start = i

    # 1. TL;DR 表格
    table_lines = []
    if tldr_start >= 0:
        for i in range(tldr_start + 1, len(lines)):
            s = lines[i].strip()
            if s == '---':
                break
            if '|' in s:
                table_lines.append(lines[i].strip())

    # 2. 总体结论
    conclusion_text = ''
    if conclusion_start >= 0:
        parts = []
        for i in range(conclusion_start, len(lines)):
            s = lines[i].strip()
            if i > conclusion_start and (s == '---' or s.startswith('## ')):
                break
            parts.append(s)
        conclusion_text = '\n'.join(parts)

    # 3. 分维度详情
    dim_lines_raw = []
    if step3_start >= 0:
        for i in range(step3_start + 1, len(lines)):
            s = lines[i].strip()
            if s.startswith('## ') and not s.startswith('### '):
                break
            dim_lines_raw.append(lines[i])

    # 按维度切分
    dim_blocks = []
    cur_title = None
    cur_judgment = ''
    cur_desc = ''
    cur_lines = []
    for line in dim_lines_raw:
        s = line.strip()
        if s.startswith('### 维度'):
            if cur_title:
                dim_blocks.append((cur_title, cur_judgment, cur_desc, cur_lines))
            cur_title = s.replace('### ', '')
            cur_judgment = ''
            cur_desc = ''
            cur_lines = []
        elif s.startswith('**判断：') or s.startswith('**判断:'):
            cur_judgment = re.sub(r'\*\*判断[：:]\s*', '', s).rstrip('*').strip()
            # 规范化：只取括号前的主判断（三档之一）
            cur_judgment = re.split(r'[（(]', cur_judgment)[0].strip()
            # 去掉可能残留的 "—" 和后面的说明
            cur_judgment = re.split(r'[—\-–]', cur_judgment)[0].strip()
        elif s.startswith('>') and not cur_lines:
            desc = s.lstrip('>').strip()
            if '整体判断' not in desc:
                cur_desc = desc
        elif not s:
            # 空行跳过，不加入cur_lines
            continue
        else:
            cur_lines.append(line)
    if cur_title:
        dim_blocks.append((cur_title, cur_judgment, cur_desc, cur_lines))

    # 过滤掉健康相关维度
    health_keywords = ['身体健康', '健康与可持续', '健康与工作节奏']
    dim_blocks = [b for b in dim_blocks if not any(kw in b[0] for kw in health_keywords)]

    # 同时过滤TL;DR表格中的健康维度行
    if table_lines:
        table_lines = [l for l in table_lines if not any(kw in l for kw in health_keywords)]

    # 解析管理者
    def parse_managers(block_lines):
        managers = []
        cur_mgr = None
        pos_count = att_count = 0
        pos_quotes = []
        att_quotes = []

        for line in block_lines:
            s = line.strip()
            if not s or s == '————————————————————' or s.startswith('>'):
                continue
            mgr_m = re.match(r'^-\s*\*\*(.+?)\*\*[：:]\s*(.*)', s)
            if mgr_m:
                if cur_mgr:
                    managers.append((cur_mgr, pos_count, att_count, pos_quotes, att_quotes))
                cur_mgr = mgr_m.group(1).strip()
                rest = mgr_m.group(2)
                pm = re.search(r'正向\s*(\d+)\s*声音', rest)
                am = re.search(r'待关注\s*(\d+)\s*声音', rest)
                pos_count = int(pm.group(1)) if pm else 0
                att_count = int(am.group(1)) if am else 0
                pos_quotes = []
                att_quotes = []
                continue
            if cur_mgr and s.startswith('- '):
                q = s[2:].strip()
                is_att = q.startswith('待关注') or q.startswith('待关注：')
                if is_att:
                    q = re.sub(r'^待关注[：:]\s*', '', q)
                # 检测是否来自bp
                is_bp = bool(re.match(r'^bp[：:"]', q) or re.search(r'\(bp\s*N\d+', q))
                # 去掉 bp：前缀 或 bp紧跟引号的前缀
                q = re.sub(r'^bp[：:]\s*', '', q)
                q = re.sub(r'^bp(?=["""\u201c])', '', q)
                # 去N编号标注（含bp N21格式）
                q = re.sub(r'[（(]\s*(?:bp\s*)?N\d+[^)）]*[)）]', '', q).strip()
                # 去首尾的；和空格
                q = q.strip('；; \t')
                # 如果包含引号，先尝试提取引号内容
                if '"' in q or '\u201c' in q or '"' in q:
                    sub_quotes = re.findall(r'[""\u201c]([^""\u201d]+)[""\u201d]', q)
                    if sub_quotes:
                        for sq in sub_quotes:
                            sq = sq.strip('；; \t')
                            if not sq or sq in ('bp', 'bp：', 'bp:', '无', '；') or len(sq) <= 3:
                                continue
                            if is_bp:
                                sq = sq + ' [From BP]'
                            if is_att or pos_count == 0:
                                att_quotes.append(sq)
                            else:
                                pos_quotes.append(sq)
                        continue
                # 没有引号包裹的文本：strip引号后处理
                q = q.strip('""\u201c\u201d')
                q = q.strip('；; \t')
                # 跳过空引用或纯bp残留
                if not q or q in ('bp', 'bp：', 'bp:', '无') or len(q) <= 3:
                    continue
                # 加 From BP 标记
                if is_bp:
                    q = q + ' [From BP]'
                if is_att or pos_count == 0:
                    att_quotes.append(q)
                else:
                    pos_quotes.append(q)
        if cur_mgr:
            managers.append((cur_mgr, pos_count, att_count, pos_quotes, att_quotes))
        return managers

    # 组装输出文本
    output = []

    # (A) 表格
    if table_lines:
        output.append('最终整体判断（跨维度）')
        output.append('')
        output.extend(table_lines)
        output.append('')

    # (B) 总体结论
    if conclusion_text:
        output.append(conclusion_text)
        output.append('')

    # (C) 分维度
    for dim_title, dim_judgment, dim_desc, block_lines in dim_blocks:
        output.append(dim_title)
        if dim_judgment:
            output.append(f'**整体判断：{dim_judgment}**')
        if dim_desc:
            output.append(dim_desc)
        output.append('')

        mgrs = parse_managers(block_lines)
        if mgrs:
            output.append('**各管理者情况：**')
            output.append('')
            for name, pc, ac, pq, aq in mgrs:
                output.append(f'- **{name}**：')
                # bp去重计数：多条bp来源只算1个来源人
                def _dedup_count(quotes):
                    non_bp = sum(1 for q in quotes if '[From BP]' not in q)
                    has_bp = any('[From BP]' in q for q in quotes)
                    return non_bp + (1 if has_bp else 0)
                # 正向：每条引用用""包裹（cadre_to_html现在按""分割），最多4条
                if pc > 0 or pq:
                    real_pc = _dedup_count(pq) if pq else pc
                    display_pq = pq[:4]  # 正向最多展示4条
                    quotes_str = ''.join(f'"{q}"' for q in display_pq)
                    output.append(f'  - 正向，{real_pc}人：')
                    output.append(f'    > {quotes_str}')
                # 待关注，最多4条
                if ac > 0 or aq:
                    real_ac = _dedup_count(aq) if aq else ac
                    display_aq = aq[:4]  # 待关注最多展示4条
                    quotes_str = ''.join(f'"{q}"' for q in display_aq)
                    output.append(f'  - 待关注，{real_ac}人：')
                    output.append(f'    > {quotes_str}')
                output.append('')
        output.append('')

    return '\n'.join(output)


def extract_tldr_judgments(md_path):
    """从md提取TL;DR表格中各维度的正确判断值"""
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    judgments = {}
    in_tldr = False
    for line in lines:
        s = line.strip()
        if s == '## TL;DR':
            in_tldr = True
            continue
        if in_tldr and s == '---':
            break
        if in_tldr and '|' in s and '维度' in s:
            cells = [c.strip() for c in s.split('|')]
            cells = [c for c in cells if c]
            if len(cells) >= 2:
                dim_text = re.sub(r'\*+', '', cells[0]).strip()
                dim_key = re.sub(r'^维度\d+[：:]\s*', '', dim_text).strip()
                judgment = cells[1].strip()
                if judgment in ('共性优势', '局部待关注', '共性不足', '局部优势'):
                    judgments[dim_key] = judgment
    return judgments


def fix_table_judgments(html, judgments):
    """修正HTML中跨维度表格的判断列文本和颜色"""
    CLR = {
        '共性优势': '#389e0d',
        '局部优势': '#389e0d',
        '局部待关注': '#d48806',
        '共性不足': '#cf1322',
    }
    for dim_key, correct_label in judgments.items():
        color = CLR.get(correct_label, '#333')
        escaped_key = re.escape(dim_key)
        pattern = re.compile(
            r'(<td[^>]*><b>[^<]*?' + escaped_key + r'[^<]*?</b></td>\s*)'
            r'<td\s+style="([^"]*)">(.*?)</td>',
            re.DOTALL
        )
        def replacer(m, _color=color, _label=correct_label):
            prefix = m.group(1)
            old_style = m.group(2)
            # 去掉旧的color和font-weight（如果有）
            new_style = re.sub(r'color:[^;]+;?', '', old_style)
            new_style = re.sub(r'font-weight:[^;]+;?', '', new_style)
            new_style = new_style.rstrip(';') + f';color:{_color};font-weight:600'
            return f'{prefix}<td style="{new_style}">{_label}</td>'
        html = pattern.sub(replacer, html)
    return html


def replace_section_13(html_path, new_html_content):
    """替换 HTML 报告中的 1.3 section 的 cadre_html 部分（不影响2异动和3敬满）"""
    import re as _re
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    start_marker = '<h4>1.3 开放题总结</h4>'
    start = html.find(start_marker)
    if start < 0:
        print(f'  ❌ 未找到 1.3 section: {html_path}')
        return False

    cadre_div_marker = '<div style="background: #fff; border: 1px solid #e8e8e8; border-radius: 8px; padding: 20px 24px; margin-bottom: 20px; line-height: 1.7;">'
    cadre_start = html.find(cadre_div_marker, start)
    
    if cadre_start > 0 and cadre_start - start < 500:
        # cadre div 已存在，只替换其中的 cadre_html 部分
        content_start = cadre_start + len(cadre_div_marker)
        
        # 找 cadre_html 结束位置：cadre_html 中不会有这些标记
        # 而 section 2/3 的内容（在 cadre div 内部但属于后续 section）会有
        non_cadre_markers = [
            '<h2', '<h3', '<h4',
            '<script', '<!-- ===',
            '<p class="note-text">', '<p class="section-subtitle">',
            '<div class="empty-hint">', '<div class="footer">',
            '<table>\n', '<table>\r\n',  # 原始 section 的 table（无style）
            '<colgroup>',  # 异动表格特征
        ]
        
        cadre_html_end = len(html)
        for marker in non_cadre_markers:
            pos = html.find(marker, content_start)
            if 0 < pos < cadre_html_end:
                cadre_html_end = pos
        
        # 如果找不到任何non_cadre标记（cadre div后面直接就是文件末尾），
        # 则找 cadre div 的模板闭合 "\n        </div>"
        if cadre_html_end >= len(html):
            close_pos = html.find('\n        </div>', content_start)
            if close_pos > 0:
                cadre_html_end = close_pos
        
        # 只替换 content_start 到 cadre_html_end 之间的内容
        # 保留 h4 标题和 cadre div 开始标签不变，保留 cadre_html_end 之后的内容不变
        new_html = html[:content_start] + '\n            ' + new_html_content + '\n        ' + html[cadre_html_end:]
        
    else:
        # 首次插入（无 cadre 容器）：找到1.3后面的第一个非cadre标记
        non_cadre_markers = [
            '<h2', '<h3', '<h4',
            '<script', '<!-- ===',
            '<p class="note-text">', '<p class="section-subtitle">',
            '<div class="empty-hint">', '<div class="footer">',
            '<table>\n', '<table>\r\n', '<colgroup>',
        ]
        end = len(html)
        for marker in non_cadre_markers:
            pos = html.find(marker, start + len(start_marker))
            if 0 < pos < end:
                end = pos

        new_section = f'''        <h4>1.3 开放题总结</h4>
        <div style="background: #fff; border: 1px solid #e8e8e8; border-radius: 8px; padding: 20px 24px; margin-bottom: 20px; line-height: 1.7;">
            {new_html_content}
        </div>
'''
        new_html = html[:start] + new_section + html[end:]

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(new_html)
    return True


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return

    html_path = sys.argv[1]
    md_path = sys.argv[2]

    if not os.path.exists(html_path):
        print(f'❌ HTML 文件不存在: {html_path}')
        return
    if not os.path.exists(md_path):
        print(f'❌ MD 文件不存在: {md_path}')
        return

    print(f'📄 HTML: {os.path.basename(html_path)}')
    print(f'📄 MD:   {os.path.basename(md_path)}')

    # Step 1: md → cadre_to_html 期望的文本
    cadre_text = md_to_cadre_text(md_path)
    if not cadre_text.strip():
        print('  ❌ 从 MD 文件中未提取到有效内容')
        return
    print(f'  ✓ 提取文本: {len(cadre_text)} 字符')

    # Step 2: cadre_to_html 生成 HTML
    cadre_html = OpenFeedbackLoader.cadre_to_html(cadre_text)
    if not cadre_html:
        print('  ❌ cadre_to_html 返回空结果')
        return

    # Step 3: 修正表格判断列（兜底逻辑可能覆盖）
    judgments = extract_tldr_judgments(md_path)
    if judgments:
        cadre_html = fix_table_judgments(cadre_html, judgments)
        print(f'  ✓ 修正表格判断: {len(judgments)} 个维度')

    print(f'  ✓ 生成HTML: {len(cadre_html)} 字符')

    # Step 4: 替换
    ok = replace_section_13(html_path, cadre_html)
    if ok:
        print(f'  ✅ 已更新 1.3 section')
    else:
        print(f'  ❌ 替换失败')


if __name__ == '__main__':
    main()
