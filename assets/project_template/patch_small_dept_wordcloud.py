#!/usr/bin/env python3
"""
补丁脚本：仅对 ≤20 人的部门重新生成关键词分析（门槛 ≥1），
更新对应 BG 的关键词分析 xlsx，然后只更新这些部门的 HTML 词云。
其余部门和其余报告内容不受任何影响。
"""
import os, sys, re, json
import pandas as pd
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
JINGMAN_DIR = os.path.join(SCRIPT_DIR, '【敬满】')

PERSON_THRESHOLD = 20   # ≤ 此人数的部门，词频门槛降为 ≥1

INNER_SEP = re.compile(r'\n|\\n|、|，|,|；|;')
MEANINGLESS_KW = {
    '无','没有','暂无','不知道','暂无意见','没有意见',
    'na','NA','N/A','n/a','none','None',
}

LLM_API_URL = os.environ.get('ORG_DIAG_LLM_URL', 'http://127.0.0.1:1234/v1/chat/completions')
LLM_KEY     = os.environ.get('ORG_DIAG_LLM_API_KEY', 'lm-studio')
LLM_MODEL   = os.environ.get('ORG_DIAG_LLM_MODEL', 'qwen2.5-7b-instruct-mlx')

FILE_PAIRS = [
    ('CSIG', 'CSIG_敬满开放题关键词汇总.xlsx', 'CSIG_敬满开放题关键词分析.xlsx'),
    ('WXG',  'WXG_敬满开放题关键词汇总.xlsx',  'WXG_敬满开放题关键词分析.xlsx'),
    ('TEG',  'TEG_敬满开放题关键词汇总.xlsx',  'TEG_敬满开放题关键词分析.xlsx'),
    ('CDG',  'CDG_敬满开放题关键词汇总.xlsx',  'CDG_敬满开放题关键词分析.xlsx'),
    ('PCG',  'PCG_敬满开放题关键词汇总.xlsx',  'PCG_敬满开放题关键词分析.xlsx'),
    ('IEG',  'IEG_敬满开放题关键词汇总.xlsx',  'IEG_敬满开放题关键词分析.xlsx'),
    ('OFS',  'OFS_敬满开放题关键词汇总.xlsx',  'OFS_敬满开放题关键词分析.xlsx'),
    ('S1',   'S1_敬满开放题关键词汇总.xlsx',   'S1_敬满开放题关键词分析.xlsx'),
    ('S2',   'S2_敬满开放题关键词汇总.xlsx',   'S2_敬满开放题关键词分析.xlsx'),
    ('S3',   'S3_敬满开放题关键词汇总.xlsx',   'S3_敬满开放题关键词分析.xlsx'),
]


def is_meaningless(kw):
    if kw in MEANINGLESS_KW:
        return True
    if re.fullmatch(r'\d+', kw):
        return True
    if re.fullmatch(r'[\s。.\-—！!，、；;|/\\+=()（）【】\[\]{}]+', kw):
        return True
    return False


def extract_keywords(cell_value):
    """提取词频。返回 (n_persons, counter)"""
    raw = str(cell_value).strip()
    if raw in ('无', 'nan', 'NaN', ''):
        return 0, Counter()
    persons = re.split(r'\|\|', raw)
    valid_persons = [p.strip() for p in persons if p.strip() and p.strip() not in ('无', 'nan', 'NaN')]
    counter = Counter()
    for p in valid_persons:
        parts = [x.strip() for x in INNER_SEP.split(p) if x.strip()]
        parts = [x for x in parts if not is_meaningless(x)]
        for kw in set(parts):
            counter[kw] += 1
    return len(valid_persons), counter


def batch_analyze_sentiment(keywords, batch_size=30):
    import requests
    results = {}
    for i in range(0, len(keywords), batch_size):
        batch = keywords[i:i+batch_size]
        prompt = '\n'.join([
            '请判断以下中文关键词的情感倾向，每个词只回答"正向""中性"或"负向"，',
            '一行一个，顺序严格对应，不要编号不要解释。',
            '',
            '关键词列表（每行一个）：',
            '\n'.join(batch),
            '',
            '输出（严格每行一个）：',
        ])
        try:
            resp = requests.post(
                LLM_API_URL,
                headers={'Authorization': 'Bearer ' + LLM_KEY, 'Content-Type': 'application/json'},
                json={'model': LLM_MODEL, 'messages': [{'role': 'user', 'content': prompt}],
                      'temperature': 0.1, 'max_tokens': 500},
                timeout=60
            )
            if resp.status_code == 200:
                content = resp.json()['choices'][0]['message']['content'].strip()
                lines = [l.strip() for l in content.split('\n') if l.strip()]
                valid = {'正向', '中性', '负向'}
                sentiments = [l for l in lines if l in valid]
                for kw, sent in zip(batch, sentiments):
                    results[kw] = sent
            else:
                for kw in batch:
                    results[kw] = '中性'
        except Exception:
            for kw in batch:
                results[kw] = '中性'
    return results


def main():
    print('\n' + '=' * 60)
    print('  小部门词云补丁（≤%d 人门槛降为 ≥1）' % PERSON_THRESHOLD)
    print('=' * 60)

    # 收集所有需要更新的部门
    all_updates = {}  # bg -> [(dept_name, new_kw_text), ...]

    for bg_name, input_name, output_name in FILE_PAIRS:
        input_path  = os.path.join(JINGMAN_DIR, input_name)
        output_path = os.path.join(JINGMAN_DIR, output_name)

        if not os.path.exists(input_path) or not os.path.exists(output_path):
            continue

        df = pd.read_excel(input_path)
        keyword_col = None
        for col in ['关键词', '详情']:
            if col in df.columns:
                keyword_col = col
                break
        if keyword_col is None:
            continue

        # 读取现有的分析结果
        out_df = pd.read_excel(output_path)
        out_dict = {str(r['部门名称']).strip(): r.get('高频关键词（关键词 词性 频次）', '') for _, r in out_df.iterrows()}

        updates = []
        for _, row in df.iterrows():
            dept_name = str(row['部门名称']).strip()
            n_persons, counter = extract_keywords(row[keyword_col])

            if n_persons == 0 or n_persons > PERSON_THRESHOLD:
                continue  # 跳过：无数据 或 人数充足（保持原有 ≥2 逻辑）

            # ≤20 人：用 ≥1 门槛
            old_high_freq = {k: v for k, v in counter.items() if v >= 2}
            new_all_freq  = {k: v for k, v in counter.items() if v >= 1}

            # 只有当旧逻辑无高频词、但新逻辑有词时，才需要更新
            old_val = str(out_dict.get(dept_name, '')).strip()
            if old_val not in ('无高频关键词', '无', 'nan', '') and old_high_freq:
                continue  # 原来就有高频词，不需要改

            if not new_all_freq:
                continue  # 连 ≥1 的词都没有

            updates.append((dept_name, n_persons, new_all_freq))

        if updates:
            all_updates[bg_name] = (output_path, out_df, updates)

    # 汇总
    total_depts = sum(len(v[2]) for v in all_updates.values())
    print(f'\n共 {total_depts} 个小部门需要更新词云：')
    for bg, (_, _, ups) in all_updates.items():
        for dept, n, _ in ups:
            print(f'  {bg}/{dept} ({n}人)')

    if total_depts == 0:
        print('\n无需更新，退出。')
        return

    # 收集所有需要情感分析的新关键词
    all_new_kw = set()
    for bg, (_, _, ups) in all_updates.items():
        for _, _, freq_map in ups:
            all_new_kw.update(freq_map.keys())

    print(f'\n待情感分析关键词: {len(all_new_kw)} 个')
    print('开始情感分析...')
    sentiment_map = batch_analyze_sentiment(list(all_new_kw), batch_size=30)
    print(f'  完成，已分析 {len(sentiment_map)} 个')

    # 更新各 BG 的关键词分析 xlsx
    for bg, (output_path, out_df, updates) in all_updates.items():
        print(f'\n更新 {bg} 关键词分析...')
        for dept_name, n_persons, freq_map in updates:
            sorted_items = sorted(freq_map.items(), key=lambda x: -x[1])
            parts = []
            for kw, cnt in sorted_items:
                sent = sentiment_map.get(kw, '中性')
                parts.append(f'{kw} {sent} {cnt}')
            kw_text = '；'.join(parts)

            # 更新 DataFrame 中对应行
            mask = out_df['部门名称'].astype(str).str.strip() == dept_name
            if mask.any():
                out_df.loc[mask, '高频关键词（关键词 词性 频次）'] = kw_text
                print(f'  ✓ {dept_name}: {len(sorted_items)} 个词')
            else:
                print(f'  ⚠ {dept_name}: 在分析文件中未找到对应行')

        out_df.to_excel(output_path, index=False)
        print(f'  已保存: {os.path.basename(output_path)}')

    # 更新词云
    print('\n' + '=' * 50)
    print('  更新受影响部门的 HTML 词云')
    print('=' * 50)

    from update_wordcloud import load_keyword_data, gen_wc_section, match_dept, replace_wc_section

    # 收集所有受影响部门名
    affected_depts = set()
    for bg, (_, _, ups) in all_updates.items():
        for dept_name, _, _ in ups:
            affected_depts.add(dept_name)

    # 遍历所有有报告的 BG
    report_base = os.path.join(SCRIPT_DIR, '报告')
    bg_report_dirs = {
        'CDG': 'cdg', 'CSIG': 'csig', 'WXG': 'wxg', 'S3': 's3',
        'TEG': 'teg', 'PCG': 'pcg', 'IEG': 'ieg', 'OFS': 'ofs',
        'S1': 's1', 'S2': 's2',
    }

    wc_updated = 0
    for bg, (_, _, updates) in all_updates.items():
        # 加载该 BG 的最新关键词分析数据
        xlsx_name = f'{bg}_敬满开放题关键词分析.xlsx'
        xlsx_path = os.path.join(JINGMAN_DIR, xlsx_name)
        if not os.path.exists(xlsx_path):
            continue

        wc_data = load_keyword_data(xlsx_path)

        # 找对应的 HTML 目录
        dir_name = bg_report_dirs.get(bg)
        if not dir_name:
            continue
        html_dir = os.path.join(report_base, dir_name, 'html')
        if not os.path.isdir(html_dir):
            continue

        affected_in_bg = {d for d, _, _ in updates}

        for fn in sorted(os.listdir(html_dir)):
            if not fn.endswith('.html'):
                continue

            # 从文件名提取部门名
            m = re.match(r'\d+_(.+?)_组织诊断报告\.html$', fn)
            if not m:
                continue
            path_part = m.group(1)
            dept_name = path_part.split('-')[-1]

            # 只处理受影响的部门
            if dept_name not in affected_in_bg:
                # 也尝试模糊匹配
                matched = False
                for ad in affected_in_bg:
                    if ad in dept_name or dept_name in ad:
                        matched = True
                        dept_name = ad
                        break
                if not matched:
                    continue

            # 查找词云数据
            wc_dept = match_dept(wc_data, dept_name)
            if not wc_dept:
                continue

            entries = wc_data[wc_dept]
            top_n = min(80, len(entries))
            words_data = [{'text': w, 'freq': f, 'sentiment': s} for w, s, f in entries[:top_n]]
            wc_json = json.dumps(words_data, ensure_ascii=False)
            new_wc = gen_wc_section(wc_json, wc_dept)

            filepath = os.path.join(html_dir, fn)
            with open(filepath, 'r', encoding='utf-8') as f:
                html = f.read()

            new_html, err = replace_wc_section(html, new_wc)
            if err:
                print(f'  ✗ {fn}: {err}')
                continue

            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(new_html)

            print(f'  ✓ 词云已更新: {fn}')
            wc_updated += 1

    print(f'\n词云更新完成: {wc_updated} 个文件')
    print('\n全部完成！')


if __name__ == '__main__':
    main()
