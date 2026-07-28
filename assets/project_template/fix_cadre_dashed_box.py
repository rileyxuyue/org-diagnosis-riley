#!/usr/bin/env python3
"""批量修复 CSIG HTML 报告中干部开放题总结部分：
当虚线框内只有标签（如"共性优势"），summary 段落跑到虚线框外面时，
把 summary 段落移入虚线框内。
"""

import os
import re
import glob

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_DIR = os.path.join(BASE_DIR, '报告', 'csig', 'html')

LABEL_KEYWORDS = ['共性优势', '共性不足', '共性短板', '局部待关注', '待关注']


def fix_html_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        html = f.read()

    if '2.1.3 开放题总结' not in html:
        return None

    fixes = 0

    # 用循环方式逐个修复
    while True:
        # 匹配：虚线框 → 短标签 → </div>关闭 → 空margin divs → summary段落
        # summary 段落内可能含 <b> <span> 等子标签，用 .*? 而非 [^<]+
        pattern = re.compile(
            r'(<div style="margin:6px 0 8px;padding:8px 10px;background:rgba\(255,255,255,0\.55\);border:1px dashed [^"]+;border-radius:8px;">\s*'
            r'<div style="[^"]*">(?:<[^>]+>)*[^<]{1,30}(?:<[^>]+>)*</div>\s*'  # 短标签行，可含span
            r')</div>'   # 虚线框关闭
            r'(\s*(?:<div style="margin-top:6px;"></div>\s*)*)'  # 空 margin divs
            r'(<div style="margin:3px 0;font-size:13px;line-height:1\.7;color:#444;">.*?</div>)',  # summary 段落
            re.DOTALL
        )

        m = pattern.search(html)
        if not m:
            break

        dashed_content = m.group(1)
        spacers = m.group(2)
        summary_div = m.group(3)

        # 验证：虚线框内纯文本确实是短标签
        inner_text = re.sub(r'<[^>]+>', '', dashed_content).strip()
        if len(inner_text) > 30 or not any(kw in inner_text for kw in LABEL_KEYWORDS):
            break

        # 把 summary 移入虚线框，调整样式
        summary_inner = summary_div.replace(
            'style="margin:3px 0;font-size:13px;line-height:1.7;color:#444;"',
            'style="margin:0 0 4px;font-size:13px;line-height:1.7;color:#222;"'
        )

        replacement = dashed_content + summary_inner + '\n</div>' + spacers
        html = html[:m.start()] + replacement + html[m.end():]
        fixes += 1

    if fixes > 0:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html)
        return fixes
    return None


def main():
    files = sorted(glob.glob(os.path.join(HTML_DIR, '*.html')))
    print(f'扫描目录: {HTML_DIR}')
    print(f'共 {len(files)} 个 HTML 文件\n')

    fixed = 0
    total_dims = 0
    for fp in files:
        fname = os.path.basename(fp)
        result = fix_html_file(fp)
        if result:
            print(f'  ✓ {fname}  →  {result} 个维度修复')
            fixed += 1
            total_dims += result
        else:
            print(f'  · {fname}  (无需修改)')

    print(f'\n完成: {fixed} 个文件修复, 共 {total_dims} 个维度框修复')


if __name__ == '__main__':
    main()
