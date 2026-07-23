#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
value_score.py —— 客户商业价值打分引擎

把一次调研得到的定性事实，按「商业价值评估维度与打分卡」转成量化层级(S/A/B/C)。

用法：
  # 单文件：scores.json = {"示例智能装备集团股份有限公司": {"经营稳定性":5, "资本化":3, ...}, ...}
  python value_score.py --json scores.json

  # 或 stdin
  cat scores.json | python value_score.py

  # 指定输出排序后的文本报告
  python value_score.py --json scores.json --out 价值分层.txt

依赖：无第三方库（仅标准库）。Python 3.8+

打分维度与权重（须与 references/商业价值评估维度与打分卡.md 保持一致）：
"""
import argparse
import json
import sys

# ── 维度 → 权重（合计 100%）──────────────────────────────────────
WEIGHTS = {
    "经营稳定性": 0.08,   # 成立时间
    "资本化": 0.10,       # 是否上市/融资阶段
    "人员规模": 0.12,
    "业务匹配": 0.13,     # 主营业务与卖方匹配度
    "盈利情况": 0.13,
    "工厂产能": 0.10,     # 工厂/产能分布
    "行业地位": 0.12,
    "近期动能": 0.07,     # 近期发展动能
    "未来趋势": 0.07,     # 未来赛道趋势
    "近期招聘": 0.08,
}
assert abs(sum(WEIGHTS.values()) - 1.0) < 1e-9, "权重合计必须等于 100%"

DIM_KEYS = list(WEIGHTS.keys())


def tier_of(score):
    if score >= 4.0:
        return "S"
    if score >= 3.2:
        return "A"
    if score >= 2.4:
        return "B"
    return "C"


TIER_ACTION = {
    "S": "战略/重点攻坚：高层直触、定制化方案、优先资源倾斜、专人对接",
    "A": "优先跟进：主力销售跟进、标准方案+批量折扣、季度复盘",
    "B": "常规培育：纳入培育池、定期触达、轻量方案、事件触发再升级",
    "C": "低优先/观察：仅留存信息、低优先、重大信号出现再激活",
}


def score_one(name, scores):
    """返回 (加权总分, 维度分列表, 红旗列表)。"""
    total = 0.0
    red_flags = []
    detail = []
    for k in DIM_KEYS:
        v = scores.get(k)
        if v is None:
            raise ValueError("公司 [%s] 缺少维度 [%s] 的分数" % (name, k))
        v = float(v)
        if not (1 <= v <= 5):
            raise ValueError("公司 [%s] 维度 [%s] 分数 %s 须在 1–5" % (name, k, v))
        total += v * WEIGHTS[k]
        detail.append((k, v))
        if v <= 1.0:
            red_flags.append(k)
    return round(total, 2), detail, red_flags


def main():
    ap = argparse.ArgumentParser(description="客户商业价值打分引擎（S/A/B/C）")
    ap.add_argument("--json", help="scores.json 路径；省略则从 stdin 读取")
    ap.add_argument("--out", help="输出排序报告到文件（否则打印到 stdout）")
    args = ap.parse_args()

    raw = args.json and open(args.json, encoding="utf-8").read() or sys.stdin.read()
    data = json.loads(raw)

    results = []
    for name, scores in data.items():
        total, detail, flags = score_one(name, scores)
        t = tier_of(total)
        # 红旗：不得直接判 S
        if flags and t == "S":
            t = "A(红旗复核)"
        results.append((name, total, t, flags, detail))

    # 按加权分降序
    results.sort(key=lambda x: x[1], reverse=True)

    lines = []
    lines.append("=" * 64)
    lines.append("客户商业价值分层结果（按加权总分降序）")
    lines.append("=" * 64)
    for name, total, t, flags, detail in results:
        lines.append("")
        lines.append("● %s   总分 %.2f / 5.0   层级：%s" % (name, total, t))
        lines.append("  维度分：" + "  ".join("%s=%g" % (k, v) for k, v in detail))
        if flags:
            lines.append("  ⚠ 红旗维度(=1分，需人工复核)：%s" % "、".join(flags))
        lines.append("  资源建议：%s" % TIER_ACTION.get(t.split("(")[0], ""))
    lines.append("")
    lines.append("维度权重：" + "  ".join("%s=%.0f%%" % (k, w * 100)
                                          for k, w in WEIGHTS.items()))
    lines.append("层级阈值：S≥4.0  A 3.2–3.99  B 2.4–3.19  C<2.4")

    out = "\n".join(lines)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(out)
        print("已写出 %d 家公司分层结果 → %s" % (len(results), args.out))
    else:
        print(out)


if __name__ == "__main__":
    main()
