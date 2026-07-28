#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step1: 从 CSIG 全面反馈开放题 Excel 抽取每位管理者的结构化原文清单
输出: output/csig/step1_output.xlsx（4个Sheet）
"""
import os
import re
import sys
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, '全面反馈开放题CSIG.xlsx')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output', 'csig')
OUTPUT_FILE = os.path.join(OUTPUT_DIR, 'step1_output.xlsx')

LLM_URL = 'http://127.0.0.1:1234/v1/chat/completions'
LLM_MODEL = 'qwen2.5-7b-instruct-mlx'

# 噪声列表
NOISE_SET = {'无', '无。', 'NA', 'N/A', '暂无', '没有', '-', '/', '空', 'NaN', '都挺好', '都挺好的', '|', ''}

# ─── Parser ──────────────────────────────────────────────

def get_full_text(row):
    parts = []
    for col, val in row.items():
        if col == '组织全路径':
            continue
        if pd.notna(val) and str(val).strip():
            parts.append(str(val))
    return '\n'.join(parts)


def parse_department(full_text):
    """切分出每位管理者的三段内容（适配新格式）"""
    # 按 ————————————————————\n【管理者：xxx】 分隔
    # 或直接按 【管理者：xxx】 分隔
    mgr_pattern = re.compile(r'【管理者[：:](.+?)】')
    splits = mgr_pattern.split(full_text)
    
    managers = []
    for i in range(1, len(splits), 2):
        header = splits[i].strip()
        body = splits[i+1] if i+1 < len(splits) else ''
        
        # 解析 name 和 role: "loriwu(吴祖榕)（组织负责人）"
        role_match = re.search(r'[（(]([^）)]*负责人[^）)]*)[）)]$', header)
        if role_match:
            role = role_match.group(1)
            name_full = header[:role_match.start()].strip()
        else:
            role = '部门负责人-1'
            name_full = header
        
        # 提取 <闪光点>...</闪光点>
        flash_match = re.search(r'<闪光点>\s*(.*?)\s*</闪光点>', body, re.S)
        more_match = re.search(r'<更多期待>\s*(.*?)\s*</更多期待>', body, re.S)
        bp_match = re.search(r'<bp观察>\s*(.*?)\s*</bp观察>', body, re.S)
        
        flash_text = flash_match.group(1).strip() if flash_match else ''
        more_text = more_match.group(1).strip() if more_match else ''
        bp_text = bp_match.group(1).strip() if bp_match else ''
        
        flash_items = [s.strip() for s in flash_text.split('||') if s.strip()]
        more_items = [s.strip() for s in more_text.split('||') if s.strip()]
        
        managers.append({
            'name_full': name_full,
            'role': role,
            'flash_items': flash_items,
            'more_items': more_items,
            'bp_text': bp_text,
        })
    
    return managers


# ─── 噪声过滤 ──────────────────────────────────────────────

def is_noise(text):
    return text.strip() in NOISE_SET


def clean_prefix(text):
    """去前缀，处理拆分"""
    items = []
    # 检查是否含 ；本人不可见： 拆分
    if '；本人不可见：' in text or ';本人不可见：' in text:
        parts = re.split(r'[；;]本人不可见[：:]', text)
        for j, p in enumerate(parts):
            p = p.strip()
            p = re.sub(r'^本人可见[：:]\s*', '', p)
            p = re.sub(r'^本人不可见[：:]\s*', '', p)
            if p and not is_noise(p):
                note = '本人不可见拆出' if j > 0 else ''
                items.append((p, note))
    else:
        text = re.sub(r'^本人可见[：:]\s*', '', text)
        text = re.sub(r'^本人不可见[：:]\s*', '', text)
        if text and not is_noise(text):
            items.append((text, ''))
    return items


# ─── BP 解析 ──────────────────────────────────────────────

def parse_bp(bp_text):
    """解析 bp 段，返回 (positive_items, negative_items, bp_status, issues)
    每个 item: (text, note)
    """
    if not bp_text or bp_text.strip() in ('无', '空缺', '无数据', '（如无则写：无）'):
        return [], [], '空缺', [('bp_all_empty', 'bp段为空')]
    
    # 检查是否整段标记空缺
    if re.search(r'BP观察结果.*?空缺', bp_text, re.S) and '做得好' not in bp_text:
        return [], [], '空缺', [('bp_all_empty', 'BP观察结果标记空缺')]
    
    # 判断bp_status: 匹配 ### **BP观察结果** 后面紧跟的 **xxx** 或直接文字
    status_match = re.search(r'BP观察结果.*?\n+\s*\*\*(.+?)\*\*', bp_text, re.S)
    if not status_match:
        status_match = re.search(r'BP观察结果.*?\n+\s*(\S+)', bp_text, re.S)
    bp_status = status_match.group(1).strip() if status_match else '正常'
    
    # 检查是否有规范小标题
    has_good = bool(re.search(r'做得好的地方', bp_text))
    has_bad = bool(re.search(r'做得不好的地方', bp_text))
    
    positive_items = []
    negative_items = []
    issues = []
    
    if has_good or has_bad:
        # 规范格式：按小标题切分
        # 使用 ### 或 --- 作为 section 结束标志
        # 提取"做得好"段：从标题行之后到下一个 ### 标题行或文档末尾
        good_section = ''
        good_match = re.search(r'###\s*\*\*做得好的地方\*\*[^\n]*\n(.*?)(?=\n###\s*\*\*|\Z)', bp_text, re.S)
        if not good_match:
            good_match = re.search(r'做得好的地方[^\n]*\n(.*?)(?=\n###\s*\*\*|\n---\s*\n###|\Z)', bp_text, re.S)
        if good_match:
            good_section = good_match.group(1).strip()
            # 去掉末尾的 --- 分隔线
            good_section = re.sub(r'\n---\s*$', '', good_section).strip()
        
        # 提取"做得不好"段
        bad_section = ''
        bad_match = re.search(r'###\s*\*\*做得不好的地方\*\*[^\n]*\n(.*?)(?=\n###\s*\*\*|\Z)', bp_text, re.S)
        if not bad_match:
            bad_match = re.search(r'做得不好的地方[^\n]*\n(.*?)(?=\n###\s*\*\*|\n---\s*\n###|\Z)', bp_text, re.S)
        if bad_match:
            bad_section = bad_match.group(1).strip()
            bad_section = re.sub(r'\n---\s*$', '', bad_section).strip()
        
        # 解析小项
        positive_items = parse_bp_section(good_section)
        negative_items = parse_bp_section(bad_section)
        
        # 过滤"无数据"
        positive_items = [(t, n) for t, n in positive_items if t.strip() not in ('无数据', '空缺', '无', '空')]
        negative_items = [(t, n) for t, n in negative_items if t.strip() not in ('无数据', '空缺', '无', '空')]
    else:
        # 无规范小标题，需要调用本地模型（或跳过）
        # 简单处理：如果全是"无数据"就跳过
        if all(w in bp_text for w in ['无数据']) or bp_text.strip() == '无数据':
            return [], [], bp_status, [('bp_all_empty', 'bp段全为无数据')]
        issues.append(('bp_format_irregular', f'bp段无规范小标题，需人工检查: {bp_text[:100]}'))
    
    return positive_items, negative_items, bp_status, issues


def parse_bp_section(section_text):
    """解析bp的一个section（做得好/做得不好），按编号小项拆分"""
    if not section_text or section_text.strip() in ('无数据', '空缺', '无', '空', '*无数据支持*', '**空缺**'):
        return []
    # 检查开头是否为空缺标记
    first_line = section_text.strip().split('\n')[0].strip()
    if first_line in ('**空缺**', '空缺', '无数据', '*无数据支持*'):
        return []
    
    items = []
    # 引号正则：匹配 ASCII双引号、中文弯引号 \u201c \u201d
    QUOTE_RE = re.compile(r'[\u0022\u201c\u201d\uff02]([^\u0022\u201c\u201d\uff02]+)[\u0022\u201c\u201d\uff02]')
    
    # 按 "数字. **标题**" 拆分小项（支持行首或换行后）
    parts = re.split(r'(?:^|\n)\s*\d+\.\s*\*\*', section_text)
    
    if len(parts) <= 1:
        # 没有编号格式，可能整段就是一条
        text = section_text.strip()
        text = re.sub(r'^\d+\.\s*', '', text)
        quotes = QUOTE_RE.findall(text)
        if quotes:
            combined = '\uff1b'.join(q.strip() for q in quotes if q.strip())
            if combined:
                items.append((combined, 'bp观察'))
        elif text and not is_noise(text):
            items.append((text, 'bp观察'))
        return items
    
    for part in parts[1:]:  # 跳过第一个空部分
        # 提取小项标题
        title_match = re.match(r'(.+?)\*\*', part)
        title = title_match.group(1).strip() if title_match else ''
        
        # 提取引号内原文
        quotes = QUOTE_RE.findall(part)
        if quotes:
            combined = '\uff1b'.join(q.strip() for q in quotes if q.strip() and not is_noise(q.strip()))
            if combined:
                note = f'bp观察:小项-{title}' if title else 'bp观察'
                items.append((combined, note))
    
    # 处理 parts[0]：如果原文以 "1. **" 开头，split 会导致第一项内容留在 parts[0]
    if parts[0].strip():
        first = parts[0].strip()
        # 检查是否含有小项内容（有标题格式残留）
        title_match = re.match(r'(.+?)\*\*', first)
        if title_match:
            title = title_match.group(1).strip()
            # 去掉前面可能残留的 "1. " 或 "1. **title**"
            title = re.sub(r'^\d+\.\s*', '', title)
        else:
            title = ''
        quotes = QUOTE_RE.findall(first)
        if quotes:
            combined = '\uff1b'.join(q.strip() for q in quotes if q.strip() and not is_noise(q.strip()))
            if combined:
                note = f'bp观察:小项-{title}' if title else 'bp观察'
                items.insert(0, (combined, note))
    
    return items


# ─── LLM 调用（仅bp格式不规范时使用）──────────────────────

def classify_bp_segment(text):
    """调用本地模型判断bp段倾向"""
    system_prompt = """你是一个文本分类器。给定一段 360 评价中"bp观察"的文本片段，判断其语义类型。
类型定义：
- positive: 明确的正向评价、肯定、赞扬、能力描述
- negative: 明确的待关注、改进建议、不足、批评
- neutral: 中性的工作职责、岗位说明、纯事实描述
输出仅一个词：positive / negative / neutral"""
    
    payload = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"文本：\n{text[:500]}"}
        ],
        "temperature": 0,
        "max_tokens": 5,
        "stream": False
    }
    try:
        resp = requests.post(LLM_URL, json=payload, timeout=30)
        resp.raise_for_status()
        answer = resp.json()["choices"][0]["message"]["content"].strip().lower()
        for kw in ["positive", "negative", "neutral"]:
            if kw in answer:
                return kw
        return "neutral"
    except:
        return "neutral"


# ─── 主流程 ──────────────────────────────────────────────

def main():
    print('📂 读取数据...')
    df = pd.read_excel(INPUT_FILE)
    print(f'  {len(df)} 个部门')
    
    # 解析所有部门
    dept_data = []
    for _, row in df.iterrows():
        path = str(row['组织全路径']).strip()
        full_text = get_full_text(row)
        managers = parse_department(full_text)
        dept_data.append({
            'path': path,
            'name': path.split('/')[-1],
            'n_managers': len(managers),
            'total_chars': len(full_text),
            'managers': managers,
        })
    
    # 按管理者数升序排序
    dept_data.sort(key=lambda d: (d['n_managers'], d['total_chars']))
    for idx, d in enumerate(dept_data):
        d['dept_idx'] = idx
    
    print(f'  解析完成，共 {sum(d["n_managers"] for d in dept_data)} 位管理者')
    
    # 抽取
    print('\n📊 开始抽取...')
    all_items = []       # Sheet 1
    mgr_summaries = []   # Sheet 2
    dept_overviews = []  # Sheet 3
    issues_log = []      # Sheet 4
    
    for dept in dept_data:
        dept_total_pos = 0
        dept_total_neg = 0
        dept_total_items = 0
        
        for mgr in dept['managers']:
            n_counter = 0
            flash_count = 0
            more_count = 0
            bp_pos_items_count = 0
            bp_neg_items_count = 0
            bp_has_pos = False
            bp_has_neg = False
            bp_status = ''
            mgr_notes = []
            
            # 闪光点
            for item_text in mgr['flash_items']:
                cleaned_list = clean_prefix(item_text)
                for text, note in cleaned_list:
                    if is_noise(text):
                        continue
                    n_counter += 1
                    flash_count += 1
                    all_items.append({
                        'dept_idx': dept['dept_idx'],
                        'dept_path': dept['path'],
                        'dept_name': dept['name'],
                        'n_managers': dept['n_managers'],
                        'mgr_name': mgr['name_full'],
                        'mgr_role': mgr['role'],
                        'n_id': f'N{n_counter:02d}',
                        'source': '闪光点',
                        'polarity': '正向',
                        'text': text,
                        'note': note,
                    })
            
            # 更多期待
            for item_text in mgr['more_items']:
                cleaned_list = clean_prefix(item_text)
                for text, note in cleaned_list:
                    if is_noise(text):
                        continue
                    n_counter += 1
                    more_count += 1
                    all_items.append({
                        'dept_idx': dept['dept_idx'],
                        'dept_path': dept['path'],
                        'dept_name': dept['name'],
                        'n_managers': dept['n_managers'],
                        'mgr_name': mgr['name_full'],
                        'mgr_role': mgr['role'],
                        'n_id': f'N{n_counter:02d}',
                        'source': '更多期待',
                        'polarity': '待关注',
                        'text': text,
                        'note': note,
                    })
            
            # BP观察
            pos_bp, neg_bp, bp_status, bp_issues = parse_bp(mgr['bp_text'])
            for issue_type, detail in bp_issues:
                issues_log.append({
                    'dept_idx': dept['dept_idx'],
                    'dept_name': dept['name'],
                    'mgr_name': mgr['name_full'],
                    'issue_type': issue_type,
                    'detail': detail,
                })
            
            bp_pos_items_count = len(pos_bp)
            bp_neg_items_count = len(neg_bp)
            bp_has_pos = bp_pos_items_count > 0
            bp_has_neg = bp_neg_items_count > 0
            
            for text, note in pos_bp:
                n_counter += 1
                all_items.append({
                    'dept_idx': dept['dept_idx'],
                    'dept_path': dept['path'],
                    'dept_name': dept['name'],
                    'n_managers': dept['n_managers'],
                    'mgr_name': mgr['name_full'],
                    'mgr_role': mgr['role'],
                    'n_id': f'N{n_counter:02d}',
                    'source': 'bp观察',
                    'polarity': '正向',
                    'text': text,
                    'note': note,
                })
            
            for text, note in neg_bp:
                n_counter += 1
                all_items.append({
                    'dept_idx': dept['dept_idx'],
                    'dept_path': dept['path'],
                    'dept_name': dept['name'],
                    'n_managers': dept['n_managers'],
                    'mgr_name': mgr['name_full'],
                    'mgr_role': mgr['role'],
                    'n_id': f'N{n_counter:02d}',
                    'source': 'bp观察',
                    'polarity': '待关注',
                    'text': text,
                    'note': note,
                })
            
            # 管理者小计
            voice_pos = flash_count + (1 if bp_has_pos else 0)
            voice_neg = more_count + (1 if bp_has_neg else 0)
            
            if flash_count == 0 and more_count == 0 and not bp_has_pos and not bp_has_neg:
                mgr_notes.append('全部反馈均为"无"，无可抽取条目')
                issues_log.append({
                    'dept_idx': dept['dept_idx'],
                    'dept_name': dept['name'],
                    'mgr_name': mgr['name_full'],
                    'issue_type': 'mgr_all_empty',
                    'detail': '三段全部为无/噪声',
                })
            
            mgr_summaries.append({
                'dept_idx': dept['dept_idx'],
                'dept_name': dept['name'],
                'mgr_name': mgr['name_full'],
                'mgr_role': mgr['role'],
                'flash_count': flash_count,
                'more_count': more_count,
                'bp_positive_items': bp_pos_items_count,
                'bp_negative_items': bp_neg_items_count,
                'voice_positive': voice_pos,
                'voice_negative': voice_neg,
                'bp_status': bp_status,
                'notes': '; '.join(mgr_notes) if mgr_notes else '',
            })
            
            dept_total_pos += voice_pos
            dept_total_neg += voice_neg
            dept_total_items += n_counter
        
        dept_overviews.append({
            'dept_idx': dept['dept_idx'],
            'dept_path': dept['path'],
            'dept_name': dept['name'],
            'n_managers': dept['n_managers'],
            'total_chars': dept['total_chars'],
            'total_voice_positive': dept_total_pos,
            'total_voice_negative': dept_total_neg,
            'total_items': dept_total_items,
        })
    
    # 写入 Excel
    print(f'\n📝 写入 Excel...')
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        pd.DataFrame(all_items).to_excel(writer, sheet_name='原文清单', index=False)
        pd.DataFrame(mgr_summaries).to_excel(writer, sheet_name='管理者小计', index=False)
        pd.DataFrame(dept_overviews).to_excel(writer, sheet_name='部门概览', index=False)
        pd.DataFrame(issues_log).to_excel(writer, sheet_name='异常日志', index=False)
    
    print(f'\n✅ 完成! {OUTPUT_FILE}')
    print(f'  Sheet 1 原文清单: {len(all_items)} 行')
    print(f'  Sheet 2 管理者小计: {len(mgr_summaries)} 行')
    print(f'  Sheet 3 部门概览: {len(dept_overviews)} 行')
    print(f'  Sheet 4 异常日志: {len(issues_log)} 行')
    
    # 抽样验证
    print(f'\n📋 抽样验证（3个部门）:')
    sample_depts = [dept_overviews[0], dept_overviews[len(dept_overviews)//2], dept_overviews[-1]]
    for d in sample_depts:
        mgrs = [m for m in mgr_summaries if m['dept_idx'] == d['dept_idx']]
        print(f'  [{d["dept_idx"]}] {d["dept_name"]} ({d["n_managers"]}人):')
        for m in mgrs[:3]:
            print(f'    {m["mgr_name"]}: 正向{m["voice_positive"]} 待关注{m["voice_negative"]} (闪光{m["flash_count"]}+bp正{m["bp_positive_items"]} / 期待{m["more_count"]}+bp负{m["bp_negative_items"]})')
    
    # 异常汇总
    if issues_log:
        from collections import Counter
        issue_counts = Counter(i['issue_type'] for i in issues_log)
        print(f'\n⚠️  异常汇总:')
        for t, c in issue_counts.most_common():
            print(f'  {t}: {c} 次')


if __name__ == '__main__':
    main()
