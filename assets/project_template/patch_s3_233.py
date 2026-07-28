#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
仅更新 S3 HTML 报告的 2.3.3 节（敬满开放题 + 词云），
其余部分（样式、其他章节、组织架构图、footer）不动。
"""

import os
import re
import sys
import json
import shutil
from datetime import datetime

os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, '.')

from generate_html_report import JingmanOpenLoader, WordCloudLoader

HTML_DIR = '报告/s3/html'
BACKUP_DIR = '报告/s3/html_backup_2.3.3'

# ============================================================
# 1. 备份
# ============================================================
if os.path.exists(BACKUP_DIR):
    print(f'备份目录已存在: {BACKUP_DIR}')
else:
    shutil.copytree(HTML_DIR, BACKUP_DIR)
    print(f'已备份到: {BACKUP_DIR}')

# ============================================================
# 2. 加载数据
# ============================================================
print('\n加载数据...')
jm_open_loader = JingmanOpenLoader('【敬满】')
wc_loader = WordCloudLoader('【敬满】')

# ============================================================
# 3. 构造 2.3.3 替换内容
# ============================================================

def build_section_233(org_full_path):
    """生成 2.3.3 节的 HTML 内容（h3 标题 + 分析卡片 + 词云）"""
    parts = []
    parts.append('<h3>3.3 敬满开放题</h3>')

    # 敬满开放题分析卡片
    jm_open_html = jm_open_loader.get_html(org_full_path)
    if jm_open_html:
        parts.append(jm_open_html)
    else:
        parts.append('\n        <p class="note-text">（该部门暂无敬满开放题分析数据）</p>')

    # 词云
    wc_json, wc_dept = wc_loader.get_wordcloud_data(org_full_path)
    if wc_json:
        wc_dept_safe = (wc_dept or '').replace('"', '&quot;').replace("'", '&#39;')
        parts.append(f'''
        <div class="wc-section">
            <div class="wc-section-title">{wc_dept_safe} · 敬满开放题关键词云</div>
            <div class="wc-section-subtitle">字号 = 词频 &nbsp;&middot;&nbsp; 颜色深浅 = 词频强度 &nbsp;&middot;&nbsp; 色系 = 情感倾向</div>
            <div class="wc-cloud-container" id="wcCloudBox"></div>
            <div id="wcTooltip" class="wc-tooltip"></div>
            <div class="wc-legend">
                <div class="wc-legend-item"><div class="wc-legend-bar positive"></div>正向词</div>
                <div class="wc-legend-item"><div class="wc-legend-bar neutral"></div>中性词</div>
                <div class="wc-legend-item"><div class="wc-legend-bar negative"></div>负向词</div>
            </div>
        </div>
        <script>
        (function() {{
            var wcData = {wc_json};
            function lerpColor(c1, c2, t) {{
                return [
                    Math.round(c1[0] + (c2[0] - c1[0]) * t),
                    Math.round(c1[1] + (c2[1] - c1[1]) * t),
                    Math.round(c1[2] + (c2[2] - c1[2]) * t),
                ];
            }}
            function sentimentColor(sentiment, freq, maxF, minF) {{
                var logMax = Math.log(maxF), logMin = Math.log(minF);
                var t = logMax === logMin ? 0.5 : (Math.log(freq) - logMin) / (logMax - logMin);
                var rgb;
                if (sentiment === 'positive') {{
                    rgb = lerpColor([183, 235, 143], [35, 120, 4], t);
                }} else if (sentiment === 'negative') {{
                    rgb = lerpColor([255, 163, 158], [168, 7, 26], t);
                }} else {{
                    rgb = lerpColor([255, 229, 143], [173, 104, 0], t);
                }}
                return 'rgb(' + rgb[0] + ',' + rgb[1] + ',' + rgb[2] + ')';
            }}
            var sentimentLabel = {{ positive: '正向', negative: '负向', neutral: '中性' }};
            var maxFreq = Math.max.apply(null, wcData.map(function(w){{ return w.freq; }}));
            var minFreq = Math.min.apply(null, wcData.map(function(w){{ return w.freq; }}));
            function fontSize(freq) {{
                var logMax = Math.log(maxFreq), logMin = Math.log(minFreq);
                var t = logMax === logMin ? 0.5 : (Math.log(freq) - logMin) / (logMax - logMin);
                return 13 + t * 44;
            }}
            var box = document.getElementById('wcCloudBox');
            var W = box.clientWidth, H = box.clientHeight;
            window._wcRenderWidth = W;
            var cx = W / 2, cy = H / 2;
            var _mc = document.createElement('canvas').getContext('2d');
            wcData.sort(function(a, b) {{ return b.freq - a.freq; }});
            var placed = [];
            function measure(text, size) {{
                _mc.font = '700 ' + size + 'px "PingFang SC","Microsoft YaHei",sans-serif';
                return {{ w: _mc.measureText(text).width + 4, h: size * 1.25 + 2 }};
            }}
            function collides(r) {{
                if (r.x < 2 || r.y < 2 || r.x + r.w > W - 2 || r.y + r.h > H - 2) return true;
                for (var i = 0; i < placed.length; i++) {{
                    var p = placed[i];
                    if (!(r.x >= p.x + p.w || r.x + r.w <= p.x || r.y >= p.y + p.h || r.y + r.h <= p.y)) return true;
                }}
                return false;
            }}
            for (var wi = 0; wi < wcData.length; wi++) {{
                var word = wcData[wi];
                var size = fontSize(word.freq);
                var m = measure(word.text, size);
                var ok = false;
                var a = Math.random() * Math.PI * 2;
                for (var i = 0; i < 12000 && !ok; i++) {{
                    var r = i * 0.04;
                    var x = cx + r * Math.cos(a) - m.w / 2;
                    var y = cy + r * Math.sin(a) - m.h / 2;
                    a += 0.12;
                    var rect = {{ x: x, y: y, w: m.w, h: m.h }};
                    if (!collides(rect)) {{
                        placed.push(rect);
                        var el = document.createElement('span');
                        el.className = 'wc-word';
                        el.textContent = word.text;
                        el.setAttribute('data-freq', word.freq);
                        el.setAttribute('data-sentiment', sentimentLabel[word.sentiment]);
                        el.style.cssText = 'left:' + x + 'px;top:' + y + 'px;font-size:' + size + 'px;color:' + sentimentColor(word.sentiment, word.freq, maxFreq, minFreq) + ';';
                        box.appendChild(el);
                        ok = true;
                    }}
                }}
            }}
            // ── 自适应高度：计算实际 bounding box 并收缩容器 ──
            if (placed.length > 0) {{
                var minY = Infinity, maxY = -Infinity;
                for (var pi = 0; pi < placed.length; pi++) {{
                    if (placed[pi].y < minY) minY = placed[pi].y;
                    if (placed[pi].y + placed[pi].h > maxY) maxY = placed[pi].y + placed[pi].h;
                }}
                var contentH = maxY - minY;
                var pad = 28;  // 上下留白
                var newH = Math.max(contentH + pad * 2, 160);
                // 垂直偏移：把所有词整体移到新容器正中
                var offsetY = pad - minY + (newH - pad * 2 - contentH) / 2;
                var spans = box.querySelectorAll('.wc-word');
                for (var si = 0; si < spans.length; si++) {{
                    var curTop = parseFloat(spans[si].style.top);
                    spans[si].style.top = (curTop + offsetY) + 'px';
                }}
                box.style.height = newH + 'px';
                window._wcRenderHeight = newH;
            }}
            var tip = document.getElementById('wcTooltip');
            box.addEventListener('mousemove', function(e) {{
                var t = e.target.closest('.wc-word');
                if (t) {{
                    tip.textContent = t.textContent + '  ×' + t.getAttribute('data-freq') + '  ' + t.getAttribute('data-sentiment');
                    tip.style.display = 'block';
                    tip.style.left = (e.clientX + 12) + 'px';
                    tip.style.top  = (e.clientY - 30) + 'px';
                }} else {{
                    tip.style.display = 'none';
                }}
            }});
            box.addEventListener('mouseleave', function() {{ tip.style.display = 'none'; }});
        }})();
        </script>''')

    return '\n'.join(parts)


def find_section_233_end(content, h3_start):
    """找到 2.3.3 节的结束位置。

    策略：
    - 如果有词云 script（包含 wcData 变量），替换到该 </script> 为止
    - 如果没有词云，替换到 footer/div 闭合之前
    """
    section = content[h3_start:]

    # 找词云 script（包含 var wcData = 的 script 块）
    wc_script_match = re.search(r'<script>\s*\(function\(\)\s*\{\s*var\s+wcData\s*=', section)
    if wc_script_match:
        # 找到对应的 </script>
        script_end = section.find('</script>', wc_script_match.start())
        if script_end >= 0:
            return h3_start + script_end + len('</script>')

    # 没有词云 script → 替换到 footer 之前
    footer_match = re.search(r'\s*<div\s+class="footer">', section)
    if footer_match:
        return h3_start + footer_match.start()

    # 兜底：找到下一个 <script>（组织架构图）之前
    org_script = re.search(r'\s*<script>\s*\n\s*//\s*组织架构图缩放控制', section)
    if org_script:
        return h3_start + org_script.start()

    # 都找不到 → 替换到文件末尾
    return len(content)


def extract_org_path_from_filename(fname):
    """从文件名提取 org_full_path，用于匹配敬满数据。

    文件名格式: 1623_S3职能系统－HR与管理线-企业文化部_组织诊断报告.html
    → S3职能系统－HR与管理线/企业文化部
    """
    # 去掉前后缀
    base = fname.replace('_组织诊断报告.html', '')
    # 去掉开头的数字ID
    base = re.sub(r'^\d+_', '', base)
    # 把 - 替换为 /（S3职能系统－HR与管理线-企业文化部 → S3职能系统－HR与管理线/企业文化部）
    # 但注意：前缀部分 "S3职能系统－HR与管理线" 不需要 /
    # 实际上 JingmanOpenLoader._match_dept 只用末段匹配，所以直接返回整个字符串即可
    return base


# ============================================================
# 4. 逐文件替换
# ============================================================
print('\n' + '=' * 60)
print('  仅更新 2.3.3 节')
print('=' * 60)

html_files = sorted(f for f in os.listdir(HTML_DIR) if f.endswith('.html'))
success = 0
skip = 0
fail = 0

for fname in html_files:
    fpath = os.path.join(HTML_DIR, fname)
    with open(fpath, 'r', encoding='utf-8') as f:
        content = f.read()

    h3_start = content.find('<h3>3.3 敬满开放题</h3>')
    if h3_start < 0:
        print(f'  ⊘ {fname}: 无 2.3.3 节，跳过')
        skip += 1
        continue

    section_end = find_section_233_end(content, h3_start)
    after_section = content[section_end:]

    # 生成新 2.3.3 内容
    org_path = extract_org_path_from_filename(fname)
    try:
        new_section = build_section_233(org_path)
    except Exception as e:
        print(f'  ✗ {fname}: 构建内容失败 ({e})')
        fail += 1
        continue

    # 拼接
    new_content = content[:h3_start] + new_section + '\n' + after_section

    with open(fpath, 'w', encoding='utf-8') as f:
        f.write(new_content)

    # 简要统计变化
    old_len = section_end - h3_start
    new_len = len(new_section)
    print(f'  ✓ {fname}: {old_len} → {new_len} 字符')
    success += 1

print(f'\n完成！成功 {success}，跳过 {skip}，失败 {fail}')
