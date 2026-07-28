#!/usr/bin/env python3
"""
通用版：只更新指定 BG 报告的 2.3.3 敬满开放题部分（HTML + MD）
其余内容完全不动。

用法:
  python3 update_jingman_open.py s3
  python3 update_jingman_open.py cdg
  python3 update_jingman_open.py all
"""

import os, re, sys, glob
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generate_html_report import JingmanOpenLoader
import json
from update_wordcloud import load_keyword_data, gen_wc_section, match_dept

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
JINGMAN_DIR = os.path.join(SCRIPT_DIR, '【敬满】')

# BG 配置
#   prefix        : 文件名中 BG 前缀，如 "S3职能系统－HR与管理线"
#   sep          : 前缀与部门路径之间的分隔符（文件名中用这个字符）
#   html_dir     : 报告 html 目录
#   md_dir       : 报告 md 目录
#   wc_xlsx_name : 【敬满】目录下对应的关键词分析 xlsx 文件名
BG_CONFIG = {
    's3': {
        'prefix':       'S3职能系统－HR与管理线',
        'sep':          '-',       # 前缀和部门名之间是半角 -
        'html_dir':    os.path.join(SCRIPT_DIR, '报告', 's3', 'html'),
        'md_dir':      os.path.join(SCRIPT_DIR, '报告', 's3', 'md'),
        'wc_xlsx_name': 'S3_敬满开放题关键词分析.xlsx',
    },
    'cdg': {
        'prefix':       'CDG企业发展事业群',
        'sep':          '-',       # 多级部门名之间也是半角 -
        'html_dir':    os.path.join(SCRIPT_DIR, '报告', 'cdg', 'html'),
        'md_dir':      os.path.join(SCRIPT_DIR, '报告', 'cdg', 'md'),
        'wc_xlsx_name': 'CDG_敬满开放题关键词分析.xlsx',
    },
    'ieg': {
        'prefix':       'IEG',
        'sep':          '-',
        'html_dir':    os.path.join(SCRIPT_DIR, '报告', 'ieg', 'html'),
        'md_dir':      os.path.join(SCRIPT_DIR, '报告', 'ieg', 'md'),
        'wc_xlsx_name': 'IEG_敬满开放题关键词分析.xlsx',
    },
}

H3_MARKER   = '<h3>3.3 敬满开放题</h3>'
H2_MARKER   = '<h2>3 敬满</h2>'
FOOTER_MARKER = '<div class="footer">'


# ── 从文件名解析 org_path ──────────────────────────────────
def parse_filename(fn, prefix, sep):
    """
    fn 如:
      S3 : 数字_S3职能系统－HR与管理线-企业文化部_组织诊断报告.html
      CDG: 数字_CDG企业发展事业群-支付基础平台与金融应用线-支付应用产品部_组织诊断报告.html
    返回 (org_path, dept_name)
      org_path = prefix + '/' + 部门路径（所有 sep 替换为 /）
      dept_name = 末级部门名
    """
    m = re.match(r'\d+_(.+?)_组织诊断报告\.html$', fn)
    if not m:
        return None, None

    path_part = m.group(1)          # 如 "S3职能系统－HR与管理线-企业文化部"
    sep_full = prefix + sep
    if not path_part.startswith(sep_full):
        # fallback：取最后一段
        dept_name = path_part.split(sep)[-1]
        return prefix + '/' + dept_name, dept_name

    rest = path_part[len(sep_full):]      # 如 "企业文化部" 或 "支付基础平台与金融应用线-支付应用产品部"
    # rest 中的 sep 全部替换为 / 得到部门路径
    dept_path = rest.replace(sep, '/')
    dept_name = rest.split(sep)[-1]
    org_path  = prefix + '/' + dept_path
    return org_path, dept_name


# ── 构建新的 2.3.3 内容 ────────────────────────────────
def build_233_content(org_path, dept_name, jm_loader, wc_data):
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

    lines.append('')
    return '\n'.join(lines)


# ── 更新单个 BG ─────────────────────────────────────────
def update_bg(bg_key):
    cfg = BG_CONFIG[bg_key]
    label      = bg_key.upper()
    prefix     = cfg['prefix']
    sep        = cfg['sep']
    html_dir   = cfg['html_dir']
    md_dir     = cfg['md_dir']
    wc_xlsx    = os.path.join(JINGMAN_DIR, cfg['wc_xlsx_name'])

    print('\n' + '=' * 50)
    print(f'  处理 BG: {label}')
    print('=' * 50)

    # ---------- 加载数据 ----------
    print('正在加载敬满开放题分析...')
    jm_loader = JingmanOpenLoader(JINGMAN_DIR)

    print('正在加载关键词词云数据...')
    if os.path.exists(wc_xlsx):
        wc_data = load_keyword_data(wc_xlsx)
        print(f'  词云部门数: {len(wc_data)}')
    else:
        wc_data = {}
        print(f'  ⚠ 词云文件不存在: {cfg["wc_xlsx_name"]}')

    # ---------- 更新 HTML ----------
    html_files = sorted(glob.glob(os.path.join(html_dir, '*.html')))
    print(f'\n找到 {len(html_files)} 个 {label} HTML 报告')

    ok = skip = 0
    for fp in html_files:
        fn = os.path.basename(fp)
        org_path, dept_name = parse_filename(fn, prefix, sep)
        if not org_path:
            print(f'  ✗ 跳过(文件名格式异常): {fn}')
            skip += 1
            continue

        with open(fp, encoding='utf-8') as f:
            html = f.read()

        h3_pos    = html.find(H3_MARKER)
        h2_pos    = html.find(H2_MARKER)
        footer_pos = html.find(FOOTER_MARKER)

        new_233 = build_233_content(org_path, dept_name, jm_loader, wc_data)

        # 模式1：已有 h3 → 替换 h3 到 footer 之间的内容
        if h3_pos != -1 and footer_pos != -1 and h3_pos < footer_pos:
            new_html = html[:h3_pos] + new_233 + '\n        ' + html[footer_pos:]

        # 模式2：没有 h3 但有 h2 → 在 h2 结束后插入
        elif h2_pos != -1 and footer_pos != -1 and h2_pos < footer_pos:
            h2_end = html.find('</h2>', h2_pos)
            if h2_end == -1:
                print(f'  ✗ 跳过(h2 结构异常): {fn}')
                skip += 1
                continue
            insert_pos = h2_end + len('</h2>')
            new_html   = html[:insert_pos] + '\n' + new_233 + '\n        ' + html[insert_pos:]

        else:
            print(f'  ✗ 跳过(结构异常): {fn}')
            skip += 1
            continue

        with open(fp, 'w', encoding='utf-8') as f:
            f.write(new_html)

        print(f'  ✓ {fn}')
        ok += 1

    print(f'HTML 更新完成: {ok} 成功, {skip} 跳过')

    # ---------- 更新 MD ----------
    print('\n正在更新 MD 报告...')
    try:
        from markdownify import markdownify as md_convert
    except ImportError:
        print('  markdownify 未安装，跳过 MD 更新')
        return

    md_files = sorted(glob.glob(os.path.join(md_dir, '*.md')))
    print(f'找到 {len(md_files)} 个 {label} MD 报告')

    ok_md = skip_md = 0
    for md_fp in md_files:
        fn = os.path.basename(md_fp)
        # 找对应 HTML
        html_fn = fn.replace('.md', '.html')
        html_fp = os.path.join(html_dir, html_fn)
        if not os.path.exists(html_fp):
            print(f'  ✗ 跳过(找不到对应HTML): {fn}')
            skip_md += 1
            continue

        # 从 HTML 提取最新 2.3.3 片段
        with open(html_fp, encoding='utf-8') as f:
            html_content = f.read()

        h3_pos    = html_content.find(H3_MARKER)
        footer_pos = html_content.find(FOOTER_MARKER)
        if h3_pos == -1 or footer_pos == -1:
            print(f'  ✗ 跳过(HTML结构异常): {fn}')
            skip_md += 1
            continue

        section_html = html_content[h3_pos:footer_pos]
        # 移除词云 script / style（不适合转 MD）
        section_html = re.sub(
            r'<div class="wc-section">.*?</div>\s*<script>.*?</script>',
            '', section_html, flags=re.DOTALL
        )
        section_html = re.sub(r'<script>.*?</script>', '', section_html, flags=re.DOTALL)
        section_html = re.sub(r'<style[^>]*>.*?</style>', '', section_html, flags=re.DOTALL)

        new_233_md = md_convert(section_html, heading_style='atx', bullets='-')
        new_233_md = re.sub(r'\n{3,}', '\n\n', new_233_md).strip()

        with open(md_fp, encoding='utf-8') as f:
            md_content = f.read()

        md_233_m = re.search(r'(#+\s*2\.3\.3[^\n]*\n)', md_content)
        if not md_233_m:
            print(f'  ✗ 跳过(MD中找不到2.3.3标题): {fn}')
            skip_md += 1
            continue

        start = md_233_m.start()
        footer_md = re.search(r'\n+本报告由组织诊断报告生成系统自动生成', md_content[start:])
        if footer_md:
            end = start + footer_md.start()
            new_md = md_content[:start] + new_233_md + '\n\n' + md_content[end:].lstrip('\n')
        else:
            new_md = md_content[:start] + new_233_md + '\n'

        with open(md_fp, 'w', encoding='utf-8') as f:
            f.write(new_md)

        print(f'  ✓ {fn}')
        ok_md += 1

    print(f'MD 更新完成: {ok_md} 成功, {skip_md} 跳过')


# ── 主入口 ─────────────────────────────────────────────
if __name__ == '__main__':
    args = [a.lower() for a in sys.argv[1:]]

    if not args or args[0] in ('-h', '--help'):
        print(__doc__)
        sys.exit(0)

    targets = []
    if 'all' in args:
        targets = list(BG_CONFIG.keys())
    else:
        for a in args:
            if a in BG_CONFIG:
                targets.append(a)
            else:
                print(f'未知 BG: {a}，支持: {list(BG_CONFIG.keys())}')
                sys.exit(1)

    for bg in targets:
        update_bg(bg)

    print('\n全部完成！')
