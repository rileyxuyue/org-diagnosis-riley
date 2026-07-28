#!/usr/bin/env python3
"""批量修复 CSIG HTML 报告中词云位置和自适应高度问题：
1. 词云 <div class="wc-section">...</div> + <script>词云JS</script>
   从 footer 内部移到 footer 之前
2. CSS: .wc-cloud-container 添加 transition
3. JS: 词云放置完成后，自适应收缩容器高度
"""

import os
import re
import glob

HTML_DIR = os.path.join(os.path.dirname(__file__), '报告', 'csig', 'html')

# 自适应高度 JS 代码片段（将插入到词云放置循环结束后、tooltip 代码之前）
AUTOFIT_JS = """            // ── 自适应高度：计算实际 bounding box 并收缩容器 ──
            if (placed.length > 0) {
                var minY = Infinity, maxY = -Infinity;
                for (var pi = 0; pi < placed.length; pi++) {
                    if (placed[pi].y < minY) minY = placed[pi].y;
                    if (placed[pi].y + placed[pi].h > maxY) maxY = placed[pi].y + placed[pi].h;
                }
                var contentH = maxY - minY;
                var pad = 28;  // 上下留白
                var newH = Math.max(contentH + pad * 2, 160);
                // 垂直偏移：把所有词整体移到新容器正中
                var offsetY = pad - minY + (newH - pad * 2 - contentH) / 2;
                var spans = box.querySelectorAll('.wc-word');
                for (var si = 0; si < spans.length; si++) {
                    var curTop = parseFloat(spans[si].style.top);
                    spans[si].style.top = (curTop + offsetY) + 'px';
                }
                box.style.height = newH + 'px';
                window._wcRenderHeight = newH;
            }"""


def fix_html_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        html = f.read()

    changes = []

    # ── 1. 修复位置：词云从 footer 内移到 footer 前 ──
    # 匹配模式：<div class="footer">...<p>生成时间：...</p>  紧接着 <div class="wc-section">...</script>  然后 </div> (关闭footer)
    # 目标：把 wc-section + script 提到 footer 之前
    pattern_pos = re.compile(
        r'(<div class="footer">\s*'
        r'<p>本报告由组织诊断报告生成系统自动生成</p>\s*'
        r'<p>生成时间：[^<]+</p>)'   # group(1): footer 开头到生成时间
        r'\s*'
        r'(<div class="wc-section">.*?</div>\s*'   # group(2): wc-section div
        r'<script>\s*\(function\(\)\s*\{.*?\}\)\(\);\s*</script>)'  # + 词云 script
        r'\s*'
        r'(</div>)',   # group(3): 关闭 footer 的 </div>
        re.DOTALL
    )

    m = pattern_pos.search(html)
    if m:
        # 重组：wc_block 放在 footer 前
        wc_block = m.group(2)
        footer_open = m.group(1)
        footer_close = m.group(3)
        replacement = f"\n{wc_block}\n        {footer_open}\n        {footer_close}"
        html = html[:m.start()] + replacement + html[m.end():]
        changes.append('位置修复')
    else:
        # 也许词云已经在 footer 前了，检查一下
        if '<div class="wc-section">' in html and '<div class="footer">' in html:
            wc_pos = html.index('<div class="wc-section">')
            footer_pos = html.index('<div class="footer">')
            if wc_pos > footer_pos:
                changes.append('位置异常(非标准模式,跳过)')
        # 如果没有词云就跳过
        if '<div class="wc-section">' not in html:
            pass  # 无词云

    # ── 2. CSS: 添加 transition ──
    old_css = 'height: 480px;\n            border: 1px solid #eee;\n            border-radius: 14px;\n            background: #fff;\n            overflow: hidden;\n        }'
    new_css = 'height: 480px;          /* 初始布局高度，JS 放置完后会自适应收缩 */\n            border: 1px solid #eee;\n            border-radius: 14px;\n            background: #fff;\n            overflow: hidden;\n            transition: height .3s ease;\n        }'
    if old_css in html:
        html = html.replace(old_css, new_css, 1)
        changes.append('CSS transition')

    # ── 3. JS: 插入自适应高度代码 ──
    # 定位：在 tooltip 代码之前插入（ "var tip = document.getElementById('wcTooltip');" 之前）
    # 但只在词云 IIFE 内部
    tip_marker = "            var tip = document.getElementById('wcTooltip');"
    if tip_marker in html and AUTOFIT_JS.strip().split('\n')[0].strip() not in html:
        html = html.replace(tip_marker, AUTOFIT_JS + '\n' + tip_marker, 1)
        changes.append('自适应高度JS')

    if changes:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html)
        return changes
    return None


def main():
    files = sorted(glob.glob(os.path.join(HTML_DIR, '*.html')))
    print(f'扫描目录: {HTML_DIR}')
    print(f'共 {len(files)} 个 HTML 文件\n')

    fixed = 0
    skipped = 0
    for fp in files:
        fname = os.path.basename(fp)
        result = fix_html_file(fp)
        if result:
            print(f'  ✓ {fname}  →  {", ".join(result)}')
            fixed += 1
        else:
            print(f'  · {fname}  (无需修改)')
            skipped += 1

    print(f'\n完成: {fixed} 个文件已修复, {skipped} 个跳过')


if __name__ == '__main__':
    main()
