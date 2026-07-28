#!/usr/bin/env python3
"""
只更新 S3 报告的 2.3.3 敬满开放题部分（HTML + MD）
其余内容完全不动
"""
import os, re, sys, glob
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generate_html_report import JingmanOpenLoader, WordCloudLoader
import json
from update_wordcloud import load_keyword_data, gen_wc_section, match_dept

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_DIR   = os.path.join(SCRIPT_DIR, '报告', 's3', 'html')
MD_DIR     = os.path.join(SCRIPT_DIR, '报告', 's3', 'md')
JINGMAN_DIR = os.path.join(SCRIPT_DIR, '【敬满】')
WC_XLSX    = os.path.join(JINGMAN_DIR, 'S3_敬满开放题关键词分析.xlsx')

H3_MARKER  = '<h3>3.3 敬满开放题</h3>'
FOOTER_MARKER = '<div class="footer">'

# ---------- 加载数据 ----------
print('正在加载敬满开放题分析...')
jm_loader = JingmanOpenLoader(JINGMAN_DIR)

print('正在加载关键词词云数据...')
wc_data = load_keyword_data(WC_XLSX)  # {dept_name: [(word, sentiment, freq), ...]}
print(f'  词云部门数: {len(wc_data)}')

# ---------- 构建新的2.3.3内容 ----------
def build_233_content(org_path, dept_name):
    """生成 <h3>2.3.3 ...> 到 </wc_section> 的完整内容"""
    lines = []
    lines.append(H3_MARKER)
    lines.append('')

    # 1. 开放题分析 HTML
    jm_html = jm_loader.get_html(org_path)
    if jm_html:
        lines.append('        ' + jm_html.replace('\n', '\n        '))
    else:
        lines.append('        <p class="note-text">（该部门暂无敬满开放题分析数据）</p>')

    lines.append('')

    # 2. 词云 section
    wc_dept = match_dept(wc_data, dept_name)
    if wc_dept:
        wc_json_str = json.dumps(
            [{'text': w, 'sentiment': s, 'freq': f} for w, s, f in wc_data[wc_dept]],
            ensure_ascii=False
        )
        wc_html = gen_wc_section(wc_json_str, wc_dept)
        lines.append('        ' + wc_html)
    # 没有词云就不加区块

    lines.append('')
    return '\n'.join(lines)


# ---------- 更新 HTML ----------
html_files = sorted(glob.glob(os.path.join(HTML_DIR, '*.html')))
print(f'\n找到 {len(html_files)} 个S3 HTML报告')

ok = 0
skip = 0

for fp in html_files:
    with open(fp, encoding='utf-8') as f:
        html = f.read()

    h3_pos = html.find(H3_MARKER)
    footer_pos = html.find(FOOTER_MARKER)
    if h3_pos == -1 or footer_pos == -1 or h3_pos >= footer_pos:
        print(f'  ✗ 跳过(结构异常): {os.path.basename(fp)}')
        skip += 1
        continue

    # 从文件名提取组织信息
    # 格式: {id}_S3职能系统－HR与管理线-{dept_name}_组织诊断报告.html
    fn = os.path.basename(fp)
    m = re.match(r'\d+_(.+?)_组织诊断报告\.html', fn)
    if not m:
        print(f'  ✗ 跳过(文件名格式异常): {fn}')
        skip += 1
        continue

    # 路径格式 "S3职能系统－HR与管理线-CDG人力资源中心" -> "S3职能系统－HR与管理线/CDG人力资源中心"
    path_part = m.group(1)
    # 找到最后一个'-'分割（部门路径之间的分隔符）
    # 匹配格式: "S3职能系统－HR与管理线-部门名"
    # 组织全路径固定前缀
    prefix = 'S3职能系统－HR与管理线'
    if path_part.startswith(prefix + '-'):
        dept_name = path_part[len(prefix)+1:]
        org_path = prefix + '/' + dept_name
    else:
        # 尝试直接当dept_name
        dept_name = path_part.split('-')[-1]
        org_path = 'S3职能系统－HR与管理线/' + dept_name

    # 构建新内容
    new_233 = build_233_content(org_path, dept_name)

    # 替换 h3_pos 到 footer_pos 之间的内容
    new_html = html[:h3_pos] + new_233 + '\n        ' + html[footer_pos:]

    with open(fp, 'w', encoding='utf-8') as f:
        f.write(new_html)

    print(f'  ✓ {fn}')
    ok += 1

print(f'\nHTML 更新完成: {ok} 成功, {skip} 跳过')

# ---------- 更新 MD ----------
print('\n正在更新 MD 报告...')
try:
    from markdownify import markdownify as md_convert
except ImportError:
    print('markdownify 未安装，跳过 MD 更新')
    sys.exit(0)

md_files = sorted(glob.glob(os.path.join(MD_DIR, '*.md')))
print(f'找到 {len(md_files)} 个S3 MD报告')

MD_H3_PATTERN = re.compile(r'(#+\s*2\.3\.3.*?)(?=\n#+\s*2\.|$)', re.DOTALL)
MD_FOOTER_PATTERN = re.compile(r'\n+本报告由组织诊断报告生成系统自动生成.*$', re.DOTALL)

ok_md = skip_md = 0
for md_fp in md_files:
    # 找到对应的 HTML
    fn = os.path.basename(md_fp).replace('.md', '.html')
    html_fp = os.path.join(HTML_DIR, fn)
    if not os.path.exists(html_fp):
        print(f'  ✗ 跳过(找不到对应HTML): {os.path.basename(md_fp)}')
        skip_md += 1
        continue

    # 从HTML中提取2.3.3部分转成MD
    with open(html_fp, encoding='utf-8') as f:
        html_content = f.read()

    h3_pos = html_content.find(H3_MARKER)
    footer_pos = html_content.find(FOOTER_MARKER)
    if h3_pos == -1 or footer_pos == -1:
        print(f'  ✗ 跳过(HTML结构异常): {os.path.basename(md_fp)}')
        skip_md += 1
        continue

    # 提取2.3.3的HTML片段（排除词云script）
    section_html = html_content[h3_pos:footer_pos]
    # 移除词云 script 和 wc-section（不适合转MD）
    section_html = re.sub(r'<div class="wc-section">.*?</div>\s*<script>.*?</script>', '', section_html, flags=re.DOTALL)
    section_html = re.sub(r'<script>.*?</script>', '', section_html, flags=re.DOTALL)
    section_html = re.sub(r'<style[^>]*>.*?</style>', '', section_html, flags=re.DOTALL)

    new_233_md = md_convert(section_html, heading_style='atx', bullets='-')
    new_233_md = re.sub(r'\n{3,}', '\n\n', new_233_md).strip()

    # 读取现有MD，替换2.3.3部分
    with open(md_fp, encoding='utf-8') as f:
        md_content = f.read()

    # 找到现有MD中的2.3.3位置
    md_233_m = re.search(r'(#+\s*2\.3\.3[^\n]*\n)', md_content)
    if not md_233_m:
        print(f'  ✗ 跳过(MD中找不到2.3.3标题): {os.path.basename(md_fp)}')
        skip_md += 1
        continue

    # 找到2.3.3起始位置，以及后面的页脚或下一个章节
    start = md_233_m.start()
    # 找到 "本报告由" 或文件末尾
    footer_md = re.search(r'\n+本报告由组织诊断报告生成系统自动生成', md_content[start:])
    if footer_md:
        end = start + footer_md.start()
        new_md = md_content[:start] + new_233_md + '\n\n' + md_content[end:].lstrip('\n')
    else:
        new_md = md_content[:start] + new_233_md + '\n'

    with open(md_fp, 'w', encoding='utf-8') as f:
        f.write(new_md)

    print(f'  ✓ {os.path.basename(md_fp)}')
    ok_md += 1

print(f'\nMD 更新完成: {ok_md} 成功, {skip_md} 跳过')
