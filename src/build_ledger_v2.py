#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI 票据智能台账 v0.2（业财融合实践 · 小微餐饮）
在 v0.1 基础上补三块：
  1. 数据质量校验与「待复核清单」（重复票号 / 疑似重复 / 金额与日期异常 / 一图多笔）
  2. 现金日结与库存现金校验（现金收付逐日结平，负数报警）
  3. 月度汇总与环比、毛利率（经营分析口径）

输入：data/raw_tickets.csv（AI 从票据 / 收款截图识别出的结构化数据）
输出：out/台账_v2.xlsx（明细台账 / 待复核清单 / 科目汇总 / 月度汇总 / 现金日结 / 经营小结）
用法：python3 src/build_ledger_v2.py [--src data/raw_tickets.csv] [--out out/台账_v2.xlsx]
"""
import argparse
import csv
import os
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------- 科目映射：小微餐饮最小科目集 ----------
EXPENSE_RULES = [
    (("面粉", "酵母", "芝麻", "豆沙", "馅", "糖", "食用油", "调和油", "牛奶", "黄油", "原料", "食材"), "主营业务成本-食材"),
    (("包装", "纸袋", "餐盒", "一次性", "袋子"), "主营业务成本-包装耗材"),
    (("电费", "水费", "燃气", "煤气", "充电"), "管理费用-水电燃气"),
    (("房租", "租金", "摊位", "物业"), "销售费用-房租"),
    (("工资", "临时工", "兼职", "人工"), "销售费用-人工"),
    (("美团", "饿了么", "佣金", "平台服务费"), "销售费用-平台佣金"),
    (("运费", "快递", "配送"), "销售费用-配送费"),
    (("设备", "维修", "烤炉", "周转箱"), "管理费用-设备维修"),
    (("税", "发票"), "税金及附加"),
]
OTHER_EXPENSE = "其他支出"
INCOME_ACCOUNT = "主营业务收入"
CASH_KEYS = ("现金",)


def classify(direction, summary):
    if direction == "收":
        return INCOME_ACCOUNT
    for keys, acc in EXPENSE_RULES:
        for k in keys:
            if k in summary:
                return acc
    return OTHER_EXPENSE


def parse_amount(text):
    try:
        v = float(str(text).replace(",", "").strip())
        return v
    except Exception:
        return None


def load(src):
    """读入并逐行校验；返回 (rows, issues)。issues: [(行号, 问题类型, 说明)]"""
    rows, issues = [], []
    with open(src, "r", encoding="utf-8-sig", newline="") as f:
        for idx, r in enumerate(csv.DictReader(f), start=2):
            raw_date = (r.get("日期") or "").strip()
            summary = (r.get("摘要") or "").strip()
            direction = (r.get("收付") or "").strip()
            amt_text = (r.get("金额") or "").strip()
            pay = (r.get("支付方式") or "").strip()
            img = (r.get("来源图片") or "").strip()
            tno = (r.get("票据号") or "").strip()
            if not raw_date and not summary:
                continue

            if raw_date:
                try:
                    datetime.strptime(raw_date, "%Y-%m-%d")
                except Exception:
                    issues.append((idx, "日期格式异常", raw_date))
            else:
                issues.append((idx, "日期为空", ""))

            if direction not in ("收", "支"):
                issues.append((idx, "收付方向异常", direction or "空"))

            amt = parse_amount(amt_text)
            if amt is None:
                issues.append((idx, "金额无法解析", amt_text or "空"))
                amt = 0.0
            elif amt <= 0:
                issues.append((idx, "金额非正数", amt_text))

            rows.append({
                "行号": idx,
                "日期": raw_date,
                "月份": raw_date[:7].replace("/", "-") if len(raw_date) >= 7 else "未知",
                "摘要": summary,
                "收付": direction,
                "金额": amt,
                "支付方式": pay,
                "来源图片": img,
                "票据号": tno,
                "科目": classify(direction, summary),
                "校验": "OK",
            })

    # 重复票号
    by_no = defaultdict(list)
    for r in rows:
        if r["票据号"]:
            by_no[r["票据号"]].append(r["行号"])
    for no, lines in by_no.items():
        if len(lines) > 1:
            issues.append((lines[1], "票据号重复", "%s（另见第 %s 行）" % (no, lines[0])))

    # 疑似重复（无票号时按 日期+金额+摘要）
    by_key = defaultdict(list)
    for r in rows:
        if not r["票据号"]:
            by_key[(r["日期"], r["金额"], r["摘要"])].append(r["行号"])
    for k, lines in by_key.items():
        if len(lines) > 1:
            issues.append((lines[1], "疑似重复记录", "%s %s %s 元" % (k[0], k[2], k[1])))

    # 一图多笔
    by_img = defaultdict(list)
    for r in rows:
        if r["来源图片"]:
            by_img[r["来源图片"]].append(r["行号"])
    for img, lines in by_img.items():
        if len(lines) > 1:
            issues.append((lines[1], "同一张图对应多笔", img))

    flag = {i[0] for i in issues}
    for r in rows:
        r["校验"] = "待复核" if r["行号"] in flag else "OK"
    return rows, issues


def cash_check(rows, issues):
    """现金日结：逐日结平，累计不能为负"""
    days = defaultdict(lambda: [0.0, 0.0])
    for r in rows:
        if any(k in r["支付方式"] for k in CASH_KEYS):
            if r["收付"] == "收":
                days[r["日期"]][0] += r["金额"]
            elif r["收付"] == "支":
                days[r["日期"]][1] += r["金额"]
    out, balance = [], 0.0
    for d in sorted(days):
        inc, exp = days[d]
        balance += inc - exp
        out.append({"日期": d, "现金收入": inc, "现金支出": exp, "日净额": inc - exp, "累计库存": balance})
        if balance < 0:
            issues.append((0, "库存现金出现负数", "%s 累计 %s 元" % (d, round(balance, 2))))
    return out


def monthly(rows):
    m = defaultdict(lambda: [0.0, 0.0, 0.0])  # 收 / 支 / 食材成本
    for r in rows:
        if r["收付"] == "收":
            m[r["月份"]][0] += r["金额"]
        else:
            m[r["月份"]][1] += r["金额"]
            if r["科目"] == "主营业务成本-食材":
                m[r["月份"]][2] += r["金额"]
    out, prev = [], None
    for mo in sorted(m):
        inc, exp, food = m[mo]
        net = inc - exp
        gross_rate = (inc - food) / inc if inc else 0.0
        mom = None
        if prev is not None and prev != 0:
            mom = (net - prev) / abs(prev)
        out.append({"月份": mo, "收入": inc, "支出": exp, "净额": net,
                    "食材成本": food, "毛利率": gross_rate, "净额环比": mom})
        prev = net
    return out


def write_excel(path, rows, issues, cash, months, kpi):
    wb = Workbook()
    head_font = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="DCE6F1")

    def sheet(title, headers, data, widths=None):
        ws = wb.create_sheet(title)
        ws.append(headers)
        for c in ws[1]:
            c.font = head_font
            c.fill = head_fill
        for d in data:
            ws.append(d)
        ws.freeze_panes = "A2"
        for i, w in enumerate(widths or [], start=1):
            ws.column_dimensions[chr(64 + i)].width = w
        return ws

    # 明细台账
    ws = wb.active
    ws.title = "明细台账"
    headers = ["行号", "日期", "摘要", "收付", "金额", "支付方式", "科目", "票据号", "来源图片", "校验"]
    ws.append(headers)
    for c in ws[1]:
        c.font = head_font
        c.fill = head_fill
    for r in rows:
        ws.append([r["行号"], r["日期"], r["摘要"], r["收付"], round(r["金额"], 2), r["支付方式"],
                   r["科目"], r["票据号"], r["来源图片"], r["校验"]])
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFGHIJ", [6, 12, 22, 6, 10, 12, 22, 20, 14, 10]):
        ws.column_dimensions[col].width = w

    # 待复核清单
    sheet("待复核清单", ["行号", "问题类型", "说明"],
          [[a, b, c] for a, b, c in issues], [6, 20, 46])

    # 科目汇总
    by_acc = defaultdict(float)
    for r in rows:
        if r["收付"] == "支":
            by_acc[r["科目"]] += r["金额"]
    total_exp = sum(by_acc.values()) or 1.0
    sheet("科目汇总", ["支出科目", "金额", "占比"],
          [[k, round(v, 2), round(v / total_exp, 4)] for k, v in sorted(by_acc.items(), key=lambda x: -x[1])],
          [24, 12, 10])

    # 月度汇总
    sheet("月度汇总", ["月份", "收入", "支出", "净额", "食材成本", "毛利率", "净额环比"],
          [[m["月份"], round(m["收入"], 2), round(m["支出"], 2), round(m["净额"], 2),
            round(m["食材成本"], 2), round(m["毛利率"], 4),
            ("-" if m["净额环比"] is None else round(m["净额环比"], 4))] for m in months],
          [10, 12, 12, 12, 12, 10, 12])

    # 现金日结
    sheet("现金日结", ["日期", "现金收入", "现金支出", "日净额", "累计库存"],
          [[c["日期"], round(c["现金收入"], 2), round(c["现金支出"], 2),
            round(c["日净额"], 2), round(c["累计库存"], 2)] for c in cash],
          [12, 12, 12, 12, 12])

    # 经营小结
    ws = wb.create_sheet("经营小结")
    ws.append(["项目", "数值"])
    for c in ws[1]:
        c.font = head_font
        c.fill = head_fill
    for k, v in kpi:
        ws.append([k, v])
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 40
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(vertical="center")

    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=os.path.join(BASE, "data", "raw_tickets.csv"))
    ap.add_argument("--out", default=os.path.join(BASE, "out", "台账_v2.xlsx"))
    args = ap.parse_args()

    rows, issues = load(args.src)
    cash = cash_check(rows, issues)
    months = monthly(rows)

    income = sum(r["金额"] for r in rows if r["收付"] == "收")
    expense = sum(r["金额"] for r in rows if r["收付"] == "支")
    food = sum(r["金额"] for r in rows if r["收付"] == "支" and r["科目"] == "主营业务成本-食材")
    days = len({r["日期"] for r in rows if r["日期"]})
    review = len({i[0] for i in issues if i[0] > 0})
    worst = max(rows, key=lambda r: r["金额"]) if rows else None

    kpi = [
        ("票据笔数", len(rows)),
        ("有交易天数", days),
        ("收入合计", round(income, 2)),
        ("支出合计", round(expense, 2)),
        ("净额", round(income - expense, 2)),
        ("食材成本", round(food, 2)),
        ("毛利率", round((income - food) / income, 4) if income else 0),
        ("日均支出", round(expense / days, 2) if days else 0),
        ("最大单笔", "%s 元（%s）" % (round(worst["金额"], 2), worst["摘要"]) if worst else "-"),
        ("待复核条目", review),
        ("校验问题条数", len(issues)),
        ("校验通过率", round(1 - review / len(rows), 4) if rows else 0),
    ]
    write_excel(args.out, rows, issues, cash, months, kpi)

    print("已生成：%s" % args.out)
    print("票据 %d 笔｜收入 %.2f｜支出 %.2f｜净额 %.2f｜毛利率 %.2f%%"
          % (len(rows), income, expense, income - expense, ((income - food) / income * 100) if income else 0))
    print("待复核 %d 条，校验问题 %d 条：" % (review, len(issues)))
    for line, kind, note in issues:
        print("  第%s行 %s %s" % (line or "-", kind, note))
    for m in months:
        print("  %s 收入 %.2f 支出 %.2f 净额 %.2f 毛利率 %.2f%% 环比 %s"
              % (m["月份"], m["收入"], m["支出"], m["净额"], m["毛利率"] * 100,
                 "-" if m["净额环比"] is None else "%.2f%%" % (m["净额环比"] * 100)))


if __name__ == "__main__":
    main()
