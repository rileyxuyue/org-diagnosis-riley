#!/usr/bin/env python3
"""
脱敏报告生成器
将组织诊断报告中的人名、部门名、业务名替换为匿名标识
"""

import re
import string

# ── 源文件 ──
SRC = "报告_示例BG_示例部门_20260101_000000.html"

# ── 1. 人名映射表（英文ID → 匿名代号） ──
# 从报告中提取的所有人名
PERSON_MAP = {}
_person_list = [
    # 英文ID, 中文名(可选), 昵称/简称(可选)
    # 示例：填入真实报告中出现的人名即可自动脱敏
    # ("engid", "中文名", ["昵称1", "昵称2"]),
]

# 生成匿名代号: 管理者A, 管理者B, ...
_alpha = list(string.ascii_uppercase)
for idx, (eid, cname, aliases) in enumerate(_person_list):
    label = f"管理者{_alpha[idx]}"
    PERSON_MAP[eid] = label
    if cname:
        PERSON_MAP[cname] = label
    for a in aliases:
        PERSON_MAP[a] = label

# ── 2. 部门/组织名映射 ──
ORG_MAP = {
    # 示例：填入真实组织名即可自动脱敏
    # "BG名称": "XX事业群",
    # "BG/部门名": "XX事业群/XX部门",
}

# ── 3. 业务/外部名称映射 ──
BIZ_MAP = {
    # 示例：填入真实业务/客户名即可自动脱敏
    # "竞品名": "竞品A",
    # "客户名": "客户A",
}


def desensitize(html: str) -> str:
    """对HTML内容进行脱敏"""
    
    # 标题替换
    html = html.replace(
        "<title>组织诊断报告 - 示例BG/示例部门</title>",
        "<title>组织诊断报告 - XX事业群/XX部门（脱敏版）</title>"
    )
    
    # 先替换较长的组合形式，避免短串先被替换后长串匹配不上
    # 人名：先替换 "engID(中文名)" 形式
    for eid, cname, aliases in _person_list:
        label = PERSON_MAP[eid]
        if cname:
            # engid(中文名) → 管理者A
            html = html.replace(f"{eid}({cname})", label)
        # engID(engID) — 有些人中文名就是ID
        html = html.replace(f"{eid}({eid})", label)
    
    # 替换部门/组织名（按长度降序，避免短串先被替换）
    for old, new in sorted(ORG_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        html = html.replace(old, new)
    
    # 替换业务名（按长度降序）
    for old, new in sorted(BIZ_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        html = html.replace(old, new)
    
    # 替换剩余的人名引用（各种格式）
    for eid, cname, aliases in _person_list:
        label = PERSON_MAP[eid]
        # 【engID】格式
        html = html.replace(f"【{eid}】", f"【{label}】")
        if cname:
            html = html.replace(f"【{cname}】", f"【{label}】")
        for a in aliases:
            html = html.replace(f"【{a}】", f"【{label}】")
    
    # 替换 <b>engID</b> 格式
    for eid, cname, aliases in _person_list:
        label = PERSON_MAP[eid]
        html = html.replace(f"<b>{eid}</b>", f"<b>{label}</b>")
        if cname:
            html = html.replace(f"<b>{cname}</b>", f"<b>{label}</b>")
        for a in aliases:
            html = html.replace(f"<b>{a}</b>", f"<b>{label}</b>")
    
    # 替换表格单元格中独立出现的人名（如 <td ...>engID</td>）
    for eid, cname, aliases in _person_list:
        label = PERSON_MAP[eid]
        # 精确替换 ">engID<" 形式
        html = html.replace(f">{eid}<", f">{label}<")
        if cname:
            html = html.replace(f">{cname}<", f">{label}<")
        for a in aliases:
            html = html.replace(f">{a}<", f">{label}<")
    
    # 替换引号中出现的人名/昵称
    for eid, cname, aliases in _person_list:
        label = PERSON_MAP[eid]
        # 纯文本中的人名（前后非字母数字）
        for name in [eid] + ([cname] if cname else []) + aliases:
            if len(name) >= 3:  # 只替换3字符以上的，避免误伤
                # 使用正则替换，确保是独立词（非HTML标签属性）
                pattern = re.compile(r'(?<![a-zA-Z\-_/])' + re.escape(name) + r'(?![a-zA-Z\-_])', re.IGNORECASE)
                html = pattern.sub(label, html)
    
    # 兜底：清理所有 "管理者X(任意中文名)" 残留 — 防止遗漏
    html = re.sub(r'(管理者[A-Z])[（(][\u4e00-\u9fff]{1,6}[)）]', r'\1', html)
    
    # 特殊处理"翻书"策略 — 这是竞品的谐音竞争策略名，替换掉
    html = html.replace('\u201c翻书\u201d', '\u201c竞争策略\u201d')
    
    # 在头部添加脱敏标记水印
    html = html.replace(
        '<div class="header">',
        '<div class="header"><div style="background:#fff3cd;color:#856404;padding:8px 16px;border-radius:6px;font-size:13px;margin-bottom:12px;border:1px solid #ffeeba;">⚠️ 本报告为脱敏版本，所有人名、部门名称、业务名称均已替换为匿名标识</div>'
    )
    
    return html


def main():
    print(f"📖 读取源文件: {SRC}")
    with open(SRC, 'r', encoding='utf-8') as f:
        html = f.read()
    
    print("🔒 正在脱敏...")
    result = desensitize(html)
    
    # 验证残留
    residuals = []
    check_words = (
        [eid for eid, _, _ in _person_list] + 
        [cn for _, cn, _ in _person_list if cn] +
        [a for _, _, als in _person_list for a in als if len(a) >= 4] +
        # 示例：填入需要替换的组织名/产品名/客户名
        []
    )
    for w in check_words:
        count = result.lower().count(w.lower())
        if count > 0:
            residuals.append(f"  ⚠️  '{w}' 仍出现 {count} 次")
    
    if residuals:
        print("⚠️  检测到残留敏感词：")
        for r in residuals:
            print(r)
    else:
        print("✅ 未检测到残留敏感词")
    
    # 写出HTML
    out_html = "报告_脱敏版_组织诊断报告.html"
    with open(out_html, 'w', encoding='utf-8') as f:
        f.write(result)
    print(f"✅ 脱敏HTML已生成: {out_html}")
    
    return out_html


if __name__ == '__main__':
    main()
