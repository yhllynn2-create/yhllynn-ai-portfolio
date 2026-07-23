#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
强排除校验：候选公司名 vs 你团队的在库/存量客户名单。

⚠️ 机密：在库名单含商业机密，绝不提交进公开仓库。
   本脚本默认读取 示例/incumbents_sample.json（明显虚构的演示数据）。
   真实名单放仓库外本地路径，用 --list 传入。

用法：
  python strong_exclusion.py --candidates candidates.txt
  python strong_exclusion.py --candidates candidates.txt --list /你的/本地/在库名单.json
  python strong_exclusion.py --candidates candidates.txt --list a.json --list b.xlsx   # 多来源并集

依赖（仅当名单为 xlsx 时）：pip install openpyxl
"""
import argparse
import csv
import json
import os
import re
import difflib


def norm(s):
    """归一化：去括号、去城市前缀、去法律后缀、去冗余空格。"""
    s = str(s)
    s = re.sub(r'[\(（][^()（）]*[\)）]', '', s)          # 去括号及内容
    for p in ['深圳市', '深圳', '广东', '广东省', '中国']:
        s = s.replace(p, '')
    s = re.sub(r'(股份)?有限公司$', '', s)                # 末尾法律后缀
    s = re.sub(r'股份公司$', '', s)
    s = s.replace('（深圳）', '').replace('(深圳)', '')
    s = re.sub(r'\s+', '', s)
    return s.strip()


# ── 别名/母子公司穿透表（本地保管，勿入库）──────────────────────────
# 按需替换为你公司的真实关系。键=别名/简称，值=该主体在库名单里可能出现的写法。
ALIAS = {
    # '示例集团': ['示例科技', '示例智能'],
}


def load_incumbents(path, col=2):
    """从 json({'names':[...]}) 或 xlsx 读取公司名列表。"""
    if path.endswith('.xlsx'):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        return [r[col] for r in rows if r and r[col]]
    with open(path, encoding='utf-8') as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    for v in data.values():
        if isinstance(v, list):
            return v
    raise ValueError('无法从 %s 解析公司名列表（需为 list 或 {"names": [...]}）' % path)


def check(candidate, incumbents, threshold=0.85):
    """返回命中列表（空=通过）。"""
    nc = norm(candidate)
    hits = []
    for n in incumbents:
        nn = norm(n)
        if not nc or not nn:
            continue
        if nc in nn or nn in nc:                 # 双向子串
            hits.append(n)
        elif difflib.SequenceMatcher(None, nc, nn).ratio() >= threshold:
            hits.append(n)
    # 别名/母子公司穿透
    for key, aliases in ALIAS.items():
        if key in candidate:
            for a in aliases:
                for n in incumbents:
                    if a in norm(n) and n not in hits:
                        hits.append(n + ' [别名命中:%s]' % a)
    return hits


def main():
    ap = argparse.ArgumentParser(description='强排除校验：候选 vs 在库名单')
    ap.add_argument('--candidates', required=True, help='候选公司名文件，每行一个')
    ap.add_argument('--list', action='append', default=None,
                    help='在库名单 json/xlsx（可多次传入做并集）；默认用示例假名单')
    ap.add_argument('--col', type=int, default=2, help='xlsx 名单的公司名列索引(0起)，默认2')
    ap.add_argument('--threshold', type=float, default=0.85, help='difflib 模糊阈值')
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    list_paths = args.list or [os.path.join(here, '..', '示例', 'incumbents_sample.json')]

    incumbents = []
    for lp in list_paths:
        incumbents.extend(load_incumbents(lp, args.col))

    with open(args.candidates, encoding='utf-8') as f:
        cands = [l.strip() for l in f if l.strip()]

    print('在库名单条数: %d | 候选条数: %d' % (len(incumbents), len(cands)))
    any_hit = False
    for c in cands:
        h = check(c, incumbents, args.threshold)
        if h:
            any_hit = True
            print('[命中] %s' % c)
            for x in h[:5]:
                print('       在库: %s' % x)
        else:
            print('[通过] %s' % c)
    print('\n结果: %s' % ('存在撞库，需剔除命中项' if any_hit else '全部通过强排除 ✅'))


if __name__ == '__main__':
    main()
