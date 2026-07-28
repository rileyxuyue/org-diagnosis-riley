#!/usr/bin/env python3
"""批量刷新所有 HTML + MD 报告中的组织架构图部分（增量替换）"""

import os, sys, re, glob

sys.path.insert(0, os.path.dirname(__file__))
from generate_html_report import build_org_chart_data, ExcelReader

# ── 读取诊断筛查表，构建 org_full_path -> leader_name 映射 ──
diag_reader = ExcelReader('【组织诊断结果】/组织诊断2025全年最终版.xlsx', '亮灯明细')
org_to_leader = {}
for row in diag_reader.rows[1:]:
    org_path = str(row[2]).strip() if len(row) > 2 and row[2] else ''
    leader = str(row[5]).strip() if len(row) > 5 and row[5] else ''
    if org_path and leader:
        org_to_leader[org_path] = leader


def render_org_node_html(node, is_root=False):
    """递归渲染组织节点 HTML（与 generate_html_report.py 中一致）"""
    is_virtual = node.get('is_virtual', False)
    if is_virtual:
        node_html = f'''
            <div class="org-child-container">
                <div class="org-connector"></div>
                <div class="org-node virtual {node['light_color']}">
                    <div class="org-node-leader">{node['leader']}</div>
                </div>
            '''
    else:
        emp_display = f"{node['employee_count']}人" if node.get('employee_count') not in ('', None) else ''
        node_html = f'''
            <div class="org-child-container">
                {'<div class="org-connector"></div>' if not is_root else ''}
                <div class="org-node {node['light_color']}">
                    <div class="org-node-name">{node['org_name']}</div>
                    <div class="org-node-leader">{node['leader']}</div>
                    <div class="org-node-count">{emp_display}</div>
                </div>
            '''

    if node.get('children'):
        node_html += '''
                <div class="org-children-wrapper">
                    <div class="org-down-connector"></div>
                    <div class="org-children">'''
        for child in node['children']:
            node_html += render_org_node_html(child, False)
        node_html += '''
                    </div>
                </div>'''

    node_html += '</div>'
    return node_html


def build_full_org_section(org_chart_data):
    """构建完整的组织架构图 HTML section"""
    chart_html = render_org_node_html(org_chart_data, True)
    
    return f'''
        <p class="section-subtitle">组织架构</p>
        <p style="color: #999; font-size: 12px; margin-top: -8px; margin-bottom: 12px;">*组织架构信息为年底</p>
        
        <div class="org-chart-wrapper">
            <div style="margin-bottom: 15px; font-size: 13px; color: #888;">
                <strong>图例说明：</strong>
                <span style="display: inline-block; margin-left: 10px;">
                    <span style="color: #ff4d4f;">●</span> 一级/二级预警
                </span>
                <span style="display: inline-block; margin-left: 10px;">
                    <span style="color: #faad14;">●</span> 三级预警
                </span>
                <span style="display: inline-block; margin-left: 10px;">
                    <span style="color: #52c41a;">●</span> 正常
                </span>
            </div>
            
            <div class="org-chart-controls no-print">
                <button class="zoom-btn" onclick="zoomOut()">🔍− 缩小</button>
                <span class="zoom-level" id="zoomLevel">100%</span>
                <button class="zoom-btn" onclick="zoomIn()">🔍+ 放大</button>
                <button class="zoom-btn" onclick="resetZoom()">↻ 重置</button>
                <button class="zoom-btn" onclick="fitToScreen()">⤢ 适应屏幕</button>
            </div>
            
            <div class="org-chart-container" id="orgChartContainer">
                <div class="org-chart" id="orgChart">
                    {chart_html}
                </div>
            </div>
        </div>
        '''


def extract_org_path_from_html(html):
    """从 HTML 报告中提取 org_full_path"""
    # 找 <meta name="org_full_path" content="...">
    m = re.search(r'<meta\s+name="org_full_path"\s+content="([^"]+)"', html)
    if m:
        return m.group(1)
    
    # 从文件名推断: 报告_BG_部门_日期.html
    return None


def extract_org_path_from_filename(fname):
    """从文件名反向推断 org_full_path（根据诊断筛查表匹配）"""
    # 报告_WXG_搜索应用部_20260327_180000.html
    # 报告_CSIG云与智慧产业事业群_安全产品一部_20260409.html
    # 报告_S1职能系统－职能线_CSIG公共事务部_20260409.html
    base = os.path.basename(fname).replace('.html', '').replace('.md', '')
    # 去掉 "报告_" 前缀和日期后缀
    base = re.sub(r'^报告_', '', base)
    base = re.sub(r'_\d{8}(_\d{6})?$', '', base)
    
    # 尝试在诊断表中精确匹配（将 _ 替换为 /）
    trial_path = base.replace('_', '/')
    if trial_path in org_to_leader:
        return trial_path
    
    # 模糊匹配：文件名中的最后一个部分（部门名）
    parts = base.split('_')
    if len(parts) >= 2:
        dept_name = parts[-1]
        for org_path in org_to_leader:
            if org_path.endswith('/' + dept_name) or org_path == dept_name:
                return org_path
    
    return None


# ── CSS 补丁：确保 .org-node.virtual 样式存在 ──
VIRTUAL_CSS = """.org-node.virtual {
            border-style: dashed;
            border-width: 2px;
            min-width: 100px;
            padding: 8px 12px;
        }"""


def patch_html_file(filepath):
    """增量替换单个 HTML 文件的组织架构图部分"""
    with open(filepath, 'r') as f:
        html = f.read()
    
    # 提取 org_full_path
    org_path = extract_org_path_from_html(html)
    if not org_path:
        org_path = extract_org_path_from_filename(filepath)
    if not org_path:
        return 'skip_no_path'
    
    leader_name = org_to_leader.get(org_path, '')
    if not leader_name:
        return 'skip_no_leader'
    
    # 重新生成组织架构图数据
    org_chart = build_org_chart_data(org_path, leader_name_override=leader_name)
    if not org_chart:
        return 'skip_no_data'
    
    new_section = build_full_org_section(org_chart)
    
    # 替换架构图区域
    # 匹配从 <p class="section-subtitle">组织架构</p> 到 </div>\s*</div>\s*</div> 结束
    pattern = re.compile(
        r'<p class="section-subtitle">组织架构</p>.*?'
        r'</div>\s*</div>\s*</div>\s*</div>\s*(?=\s*<(?:h4|p class="section-subtitle"|div class="footer"))',
        re.DOTALL
    )
    
    if not pattern.search(html):
        # 可能没有架构图部分——找插入点（在 2.1.3 之前）
        return 'skip_no_match'
    
    html = pattern.sub(new_section.strip() + '\n        ', html)
    
    # 确保 virtual CSS 存在
    if '.org-node.virtual' not in html:
        # 在 .org-node.gray 样式后面插入
        html = html.replace(
            '.org-node.gray {\n            border-color: #d9d9d9;\n            background: #fafafa;\n        }',
            '.org-node.gray {\n            border-color: #d9d9d9;\n            background: #fafafa;\n        }\n        \n        ' + VIRTUAL_CSS
        )
    
    with open(filepath, 'w') as f:
        f.write(html)
    
    return 'ok'


def render_org_tree_md(node, depth=0, prefix=""):
    """递归渲染组织树为 Markdown 文本格式"""
    lines = []
    is_virtual = node.get('is_virtual', False)
    
    if depth == 0:
        # 根节点
        emp = f" ({node['employee_count']}人)" if node.get('employee_count') not in ('', None) else ''
        lines.append(f"**{node['org_name']}**{emp}")
        lines.append(f"负责人：{node['leader']}")
        lines.append("")
    else:
        indent = "  " * (depth - 1)
        bullet = "- "
        if is_virtual:
            lines.append(f"{indent}{bullet}**{node['leader']}** (分管)")
        else:
            emp = f" ({node['employee_count']}人)" if node.get('employee_count') not in ('', None) else ''
            lines.append(f"{indent}{bullet}{node['org_name']} - {node['leader']}{emp}")
    
    for child in node.get('children', []):
        lines.extend(render_org_tree_md(child, depth + 1))
    
    return lines


def patch_md_file(md_path, org_path, leader_name):
    """增量替换 MD 文件中的组织架构部分"""
    if not os.path.exists(md_path):
        return 'skip_no_file'
    
    with open(md_path, 'r') as f:
        content = f.read()
    
    org_chart = build_org_chart_data(org_path, leader_name_override=leader_name)
    if not org_chart:
        return 'skip_no_data'
    
    tree_lines = render_org_tree_md(org_chart)
    new_section = "组织架构\n\n*组织架构信息为年底\n\n" + "\n".join(tree_lines)
    
    # 替换 "组织架构" 到下一个 ## 标题之间的内容
    pattern = re.compile(
        r'组织架构\n.*?(?=\n##|\n\*\*2\.1\.3|\n2\.1\.3|\nAll In|\n####|\Z)',
        re.DOTALL
    )
    
    if pattern.search(content):
        content = pattern.sub(new_section + '\n\n', content)
        with open(md_path, 'w') as f:
            f.write(content)
        return 'ok'
    
    return 'skip_no_match'


# ── 主流程 ──
if __name__ == '__main__':
    html_dirs = [
        '报告/csig/html',
        '报告/wxg/html',
    ]
    md_dirs = [
        '报告/csig/md',
        '报告/wxg/md',
    ]
    
    stats = {'ok': 0, 'skip': 0, 'fail': 0}
    
    # 处理 HTML 文件
    print("=" * 60)
    print("处理 HTML 报告")
    print("=" * 60)
    for d in html_dirs:
        files = sorted(glob.glob(os.path.join(d, '*.html')))
        for fp in files:
            result = patch_html_file(fp)
            short = os.path.basename(fp)[:60]
            if result == 'ok':
                print(f"  ✓ {short}")
                stats['ok'] += 1
            else:
                print(f"  ✗ {short} ({result})")
                stats['skip'] += 1
    
    # 处理 MD 文件
    print()
    print("=" * 60)
    print("处理 MD 报告")
    print("=" * 60)
    md_stats = {'ok': 0, 'skip': 0}
    for d in md_dirs:
        if not os.path.exists(d):
            continue
        files = sorted(glob.glob(os.path.join(d, '*.md')))
        for fp in files:
            org_path = extract_org_path_from_filename(fp)
            if not org_path:
                print(f"  ✗ {os.path.basename(fp)[:60]} (skip_no_path)")
                md_stats['skip'] += 1
                continue
            leader = org_to_leader.get(org_path, '')
            if not leader:
                print(f"  ✗ {os.path.basename(fp)[:60]} (skip_no_leader)")
                md_stats['skip'] += 1
                continue
            result = patch_md_file(fp, org_path, leader)
            short = os.path.basename(fp)[:60]
            if result == 'ok':
                print(f"  ✓ {short}")
                md_stats['ok'] += 1
            else:
                print(f"  ✗ {short} ({result})")
                md_stats['skip'] += 1
    
    print()
    print(f"HTML: {stats['ok']} 成功, {stats['skip']} 跳过")
    print(f"MD:   {md_stats['ok']} 成功, {md_stats['skip']} 跳过")
