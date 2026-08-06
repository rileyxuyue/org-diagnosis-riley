#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从「XXG_敬满开放题关键词汇总.xlsx」
生成「XXG_敬满开放题关键词分析.xlsx」

- 不筛选词频，所有词均保留
- 每个部门最多输出词频前 200 个（并列第 200 名全部保留）
- 每个词的情感单独判断
- 已判断过的词写入本地缓存，跨 BG 复用
- 每处理完一个部门就写一次 xlsx（中断不丢数据）
"""

import os
import re
import json
import time
import requests
import pandas as pd
from collections import Counter

# ── 配置 ──────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, 'sentiment_cache.json')

LLM_API_URL = os.environ.get('ORG_DIAG_LLM_URL', 'http://127.0.0.1:1234/v1/chat/completions')
LLM_KEY     = os.environ.get('ORG_DIAG_LLM_API_KEY', 'lm-studio')
LLM_MODEL   = os.environ.get('ORG_DIAG_LLM_MODEL', 'qwen2.5-7b-instruct-mlx')

# 当前目标 BG
BG_LIST = [
    ('CDG', 'CDG_敬满开放题关键词汇总.xlsx', 'CDG_敬满开放题关键词分析.xlsx'),
]

# ── 分词分隔符 ────────────────────────────────────
# 用户约定：&, /, ／, +, ＋ 也视为并列分隔符
SPLIT_RE = re.compile(
    r'[、，,;；；：:&/／+＋]+'    # 标点与并列连接符
    r'|\s+'                     # 空白（空格/制表符等）
    r'|["\'"\u2018\u2019\u201c\u201d]+'  # 各种引号
    r'|[\n\r]+'                # 换行
)

# ── 无意义词过滤 ─────────────────────────────────
MEANINGLESS = {
    '无','没有','暂无','不知道','暂无意见','没有意见',
    'na','NA','N/A','n/a','none','None','NaN','nan',
}

# 英文停用词（词云中无意义的虚词/介词/连词等）
EN_STOPWORDS = {
    'a', 'an', 'the', 'to', 'and', 'or', 'of', 'in', 'on', 'at',
    'is', 'it', 'be', 'for', 'as', 'by', 'no', 'not', 'n', 'but',
    'with', 'that', 'this', 'are', 'was', 'were', 'has', 'have', 'had',
    'if', 'so', 'do', 'up', 'we', 'my', 'me', 'he', 'she', 'i', 'its',
    'can', 'will', 'just', 'should', 'would', 'could', 'may', 'might',
    'am', 'been', 'being', 'did', 'does', 'doing', 'each', 'few',
    'from', 'get', 'got', 'her', 'him', 'his', 'how', 'into',
    'more', 'most', 'much', 'must', 'nor', 'our', 'out', 'own',
    'same', 'some', 'such', 'than', 'them', 'then', 'too', 'very',
    'what', 'when', 'who', 'whom', 'why', 'you', 'your',
    'c',  # 单字母无意义
}

def is_meaningless(kw):
    if kw.lower() in MEANINGLESS:
        return True
    if kw.lower() in EN_STOPWORDS:
        return True
    if re.fullmatch(r'\d+\.?', kw):
        return True
    if re.fullmatch(r'[。.\-—！!?,;|/\\+=（）()【】\[\]{}"\'·■□☆★◆◇●○►▼▽▻▿\u200f]+', kw):
        return True
    return False


# ── 从单元格提取词频（|| 分隔人，每人内部去重） ──────────
def extract_counter(cell_value):
    raw = str(cell_value).strip()
    if raw in ('无','nan','NaN',''):
        return Counter()
    persons = re.split(r'\|\|', raw)
    counter = Counter()
    for p in persons:
        p = p.strip()
        if not p or p in ('无','nan','NaN'):
            continue
        # 按分隔符切分
        parts = [x.strip() for x in SPLIT_RE.split(p) if x.strip()]
        # 过滤无意义词 + 转小写 + 每人内部去重
        valid = set()
        for kw in parts:
            kw = kw.lower()
            if not kw: continue
            if is_meaningless(kw): continue
            valid.add(kw)
        # 计数
        for kw in valid:
            counter[kw] += 1
    return counter


# ── 情感分析（带本地 JSON 缓存）────────────────────
def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                cache = json.load(f)
            # 统一缓存键为小写
            cache = {k.lower(): v for k, v in cache.items()}
            print('  已加载缓存：{} 个词'.format(len(cache)), flush=True)
            return cache
        except Exception as e:
            print('  缓存加载失败：{}，从头开始'.format(e), flush=True)
    return {}


def save_cache(cache):
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    print('  缓存已保存：{} 个词'.format(len(cache)), flush=True)


def analyze_sentiment(kw, cache):
    """查缓存 → 请求 LLM → 写缓存"""
    kw = kw.lower()
    if kw in cache:
        return cache[kw]

    prompt = (
        '请判断关键词「' + kw + '」的情感倾向，'
        '只回答"正向"或"中性"或"负向"，不要解释，不要加其他内容。'
    )
    valid = {'正向', '中性', '负向'}

    for _ in range(2):
        try:
            resp = requests.post(
                LLM_API_URL,
                headers={
                    'Authorization': 'Bearer ' + LLM_KEY,
                    'Content-Type': 'application/json',
                },
                json={
                    'model': LLM_MODEL,
                    'messages': [{'role': 'user', 'content': prompt}],
                    'temperature': 0.1,
                    'max_tokens': 10,
                },
                timeout=30,
            )
            if resp.status_code == 200:
                content = resp.json()['choices'][0]['message']['content'].strip()
                if content in valid:
                    cache[kw] = content
                    return content
                for v in valid:
                    if v in content:
                        cache[kw] = v
                        return v
        except Exception:
            pass
        time.sleep(0.1)

    cache[kw] = '中性'
    return '中性'


# ── 主流程 ──────────────────────────────────────────
def process_bg(bg_name, input_name, output_name, cache):
    input_path  = os.path.join(BASE_DIR, input_name)
    output_path = os.path.join(BASE_DIR, output_name)

    print('', flush=True)
    print('=' * 52, flush=True)
    print('  ' + bg_name, flush=True)
    print('=' * 52, flush=True)

    if not os.path.exists(input_path):
        print('  skipped: 输入文件不存在 (' + input_name + ')', flush=True)
        return

    df = pd.read_excel(input_path)

    # 找关键词列
    keyword_col = None
    for col in df.columns:
        if '关键词' in col or '详情' in col:
            keyword_col = col
            break
    if keyword_col is None:
        print('  skipped: 未找到关键词列', flush=True)
        return

    print('  读取 ' + str(len(df)) + ' 个部门，列="' + keyword_col + '"', flush=True)

    # 逐部门提取词频，同时收集所有不重复的词
    dept_data  = []       # [(dept_name, Counter), ...]
    all_kw_set = set()

    for _, row in df.iterrows():
        dept_name = str(row['部门名称']).strip()
        counter   = extract_counter(row[keyword_col])
        dept_data.append((dept_name, counter))
        all_kw_set.update(counter.keys())

    # 去掉缓存中已有的词
    need_llm = [kw for kw in all_kw_set if kw not in cache]
    print('  共 ' + str(len(all_kw_set)) + ' 个不重复词，需请求 LLM：' + str(len(need_llm)) + ' 个', flush=True)

    # 逐词请求（有缓存则跳过），每 200 词保存一次
    for i, kw in enumerate(need_llm, 1):
        analyze_sentiment(kw, cache)
        if i % 20 == 0 or i == len(need_llm):
            print('    进度：' + str(i) + '/' + str(len(need_llm)) + ' (' + kw + ' → ' + cache.get(kw, '?') + ')', flush=True)
        if i % 200 == 0:
            save_cache(cache)
        time.sleep(0.04)   # 限速，约 25 词/秒

    print('  情感分析完成，开始生成输出…', flush=True)

    # 生成每个部门的输出行（前 200，并列全部保留）
    output_rows = []
    for dept_name, counter in dept_data:
        if not counter:
            kw_text = '无关键词'
        else:
            sorted_items = sorted(counter.items(), key=lambda x: -x[1])
            top_n = 200
            if len(sorted_items) > top_n:
                cutoff_freq = sorted_items[top_n - 1][1]
                items_to_show = [(kw, cnt) for kw, cnt in sorted_items if cnt >= cutoff_freq]
            else:
                items_to_show = sorted_items

            parts = []
            for kw, cnt in items_to_show:
                sent = cache.get(kw, '中性')
                parts.append(kw + ' ' + sent + ' ' + str(cnt))
            kw_text = '；'.join(parts)

        output_rows.append({
            '部门名称': dept_name,
            '高频关键词（关键词 词性 频次）': kw_text,
        })

    out_df = pd.DataFrame(output_rows)
    out_df.to_excel(output_path, index=False)
    print('  输出：' + output_name + '（' + str(len(out_df)) + ' 行）', flush=True)
    print('  正在保存缓存…', flush=True)
    save_cache(cache)


def main():
    cache = load_cache()
    total_bgs = len(BG_LIST)

    for i, (bg_name, inp, out) in enumerate(BG_LIST, 1):
        print('', flush=True)
        print('[' + str(i) + '/' + str(total_bgs) + ']', flush=True)
        process_bg(bg_name, inp, out, cache)

    print('', flush=True)
    print('=' * 52, flush=True)
    print('  全部完成！缓存共 ' + str(len(cache)) + ' 个词', flush=True)
    print('=' * 52, flush=True)
    save_cache(cache)


if __name__ == '__main__':
    main()
