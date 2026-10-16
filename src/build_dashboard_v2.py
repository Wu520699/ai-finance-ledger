#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI 票据智能台账 v0.2 经营看板（单文件 HTML，ECharts 本地引用）
读 data/raw_tickets_v2.csv（或 --src 指定），输出 web/dashboard_v2.html
用法：python3 src/build_dashboard_v2.py [--src data/raw_tickets_v2.csv] [--out web/dashboard_v2.html]
"""
import argparse
import json
import os
import sys
from collections import defaultdict
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_ledger_v2 as core  # noqa: E402

BASE = core.BASE

HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI 票据智能台账 · 经营看板 v0.2</title>
<script src="./vendor/echarts.min.js"></script>
<style>
  :root { --ink:#1f2937; --muted:#6b7280; --line:#e5e7eb; --bg:#f6f7f9; --accent:#2f6fed; }
  * { box-sizing: border-box; }
  body { margin:0; background:var(--bg); color:var(--ink);
         font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif; }
  header { padding:22px 26px 6px; }
  h1 { margin:0; font-size:21px; letter-spacing:.5px; }
  .sub { margin:6px 0 0; color:var(--muted); font-size:13px; }
  .kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr));
          gap:12px; padding:16px 26px 4px; }
  .kpi { background:#fff; border:1px solid var(--line); border-radius:12px; padding:14px 16px; }
  .kpi .label { color:var(--muted); font-size:12px; }
  .kpi .value { font-size:20px; font-weight:600; margin-top:6px; }
  .kpi .value.warn { color:#c2410c; }
  .grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(360px,1fr));
          gap:12px; padding:12px 26px; }
  .card { background:#fff; border:1px solid var(--line); border-radius:12px; padding:14px 16px 6px; }
  .card h2 { margin:0 0 8px; font-size:14px; font-weight:600; color:var(--ink); }
  .chart { width:100%; height:280px; }
  .wide { grid-column:1/-1; }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  th, td { text-align:left; padding:8px 6px; border-bottom:1px solid var(--line); }
  th { color:var(--muted); font-weight:500; font-size:12px; }
  .tag { display:inline-block; padding:1px 8px; border-radius:999px;
         background:#fff1e6; color:#c2410c; font-size:12px; }
  footer { padding:8px 26px 26px; color:var(--muted); font-size:12px; }
</style>
</head>
<body>
<header>
  <h1>AI 票据智能台账 · 经营看板 v0.2</h1>
  <p class="sub">小微餐饮门店（业财融合实践）· 演示数据 · 生成于 __TIME__</p>
</header>
__KPI__
<div class="grid">
  <div class="card"><h2>月度经营对比（收入 / 支出 / 净额）</h2><div id="c1" class="chart"></div></div>
  <div class="card"><h2>支出结构（科目占比）</h2><div id="c2" class="chart"></div></div>
  <div class="card"><h2>每日收支</h2><div id="c3" class="chart"></div></div>
  <div class="card"><h2>现金日结与库存现金</h2><div id="c4" class="chart"></div></div>
  <div class="card wide">
    <h2>数据质量：待复核清单</h2>
    <table><thead><tr><th>行号</th><th>问题类型</th><th>说明</th></tr></thead>
    <tbody id="tb"></tbody></table>
  </div>
</div>
<footer>票据数据由 AI 从收款截图与票据识别后导出，本页只做汇总与校验展示。</footer>
<script>
const D = __DATA__;
const money = v => v.toLocaleString('zh-CN', {minimumFractionDigits:2, maximumFractionDigits:2});
const axisStyle = {axisLine:{lineStyle:{color:'#d1d5db'}}, axisLabel:{color:'#6b7280', fontSize:11},
                   splitLine:{lineStyle:{color:'#f0f1f3'}}};

function showEmpty(el, msg){
  if (!el) return;
  el.innerHTML = '<div style="height:100%;display:flex;align-items:center;justify-content:center;color:#6b7280;font-size:13px">' + (msg || '暂无数据') + '</div>';
}
function initChart(id, hasData){
  const el = document.getElementById(id);
  if (!hasData) { showEmpty(el); return null; }
  return echarts.init(el);
}

const charts = [];
if (typeof echarts === 'undefined') {
  document.querySelectorAll('.chart').forEach(function(el){ showEmpty(el, 'ECharts 资源未加载'); });
} else {
  // 1. 月度对比
  const c1 = initChart('c1', D.months.length);
  if (c1) {
    c1.setOption({
      tooltip:{trigger:'axis'},
      legend:{data:['收入','支出','净额'], bottom:0, textStyle:{fontSize:11}},
      grid:{left:52, right:16, top:16, bottom:44},
      xAxis:{type:'category', data:D.months.map(m=>m.月), ...axisStyle},
      yAxis:{type:'value', ...axisStyle},
      series:[
        {name:'收入', type:'bar', barWidth:18, itemStyle:{color:'#2f6fed'}, data:D.months.map(m=>m.收入)},
        {name:'支出', type:'bar', barWidth:18, itemStyle:{color:'#f59e0b'}, data:D.months.map(m=>m.支出)},
        {name:'净额', type:'line', smooth:true, symbolSize:7, itemStyle:{color:'#059669'},
         label:{show:true, fontSize:10, formatter:p=>money(p.value)}, data:D.months.map(m=>m.净额)}
      ]
    });
    charts.push(c1);
  }

  // 2. 支出结构
  const c2 = initChart('c2', D.accounts.length);
  if (c2) {
    c2.setOption({
      tooltip:{trigger:'item', formatter:p=>p.name+'：'+money(p.value)+' 元（'+p.percent+'%）'},
      legend:{bottom:0, type:'scroll', textStyle:{fontSize:11}},
      series:[{type:'pie', radius:['38%','62%'], center:['50%','44%'],
        itemStyle:{borderColor:'#fff', borderWidth:2},
        label:{fontSize:11, formatter:'{b}\\n{d}%'},
        data:D.accounts}]
    });
    charts.push(c2);
  }

  // 3. 每日收支
  const c3 = initChart('c3', D.daily.length);
  if (c3) {
    c3.setOption({
      tooltip:{trigger:'axis'},
      legend:{data:['收入','支出','当日净额'], bottom:0, textStyle:{fontSize:11}},
      grid:{left:52, right:16, top:16, bottom:44},
      xAxis:{type:'category', data:D.daily.map(d=>d.日期.slice(5)), ...axisStyle},
      yAxis:{type:'value', ...axisStyle},
      series:[
        {name:'收入', type:'bar', stack:'io', barWidth:14, itemStyle:{color:'#2f6fed'}, data:D.daily.map(d=>d.收入)},
        {name:'支出', type:'bar', stack:'io', barWidth:14, itemStyle:{color:'#f59e0b'}, data:D.daily.map(d=>d.支出)},
        {name:'当日净额', type:'line', smooth:true, symbolSize:5, itemStyle:{color:'#059669'},
         data:D.daily.map(d=>+(d.收入-d.支出).toFixed(2))}
      ]
    });
    charts.push(c3);
  }

  // 4. 现金日结与库存
  const c4 = initChart('c4', D.cash.length);
  if (c4) {
    c4.setOption({
      tooltip:{trigger:'axis'},
      legend:{data:['现金收入','现金支出','累计库存'], bottom:0, textStyle:{fontSize:11}},
      grid:{left:52, right:16, top:16, bottom:44},
      xAxis:{type:'category', data:D.cash.map(c=>c.日期.slice(5)), ...axisStyle},
      yAxis:{type:'value', ...axisStyle},
      series:[
        {name:'现金收入', type:'bar', barWidth:14, itemStyle:{color:'#93c5fd'}, data:D.cash.map(c=>c.现金收入)},
        {name:'现金支出', type:'bar', barWidth:14, itemStyle:{color:'#fdba74'}, data:D.cash.map(c=>c.现金支出)},
        {name:'累计库存', type:'line', smooth:true, symbolSize:6, itemStyle:{color:'#dc2626'},
         areaStyle:{color:'rgba(220,38,38,.08)'},
         markLine:{silent:true, symbol:'none', lineStyle:{color:'#dc2626', type:'dashed'},
           data:[{yAxis:0, label:{formatter:'现金为零', fontSize:10}}]},
         data:D.cash.map(c=>c.累计库存)}
      ]
    });
    charts.push(c4);
  }
}

// 5. 待复核清单
const tb = document.getElementById('tb');
if (!D.issues.length) {
  const tr = document.createElement('tr');
  tr.innerHTML = '<td colspan="3" style="color:#6b7280">暂无待复核记录</td>';
  tb.appendChild(tr);
} else {
  D.issues.forEach(function(i){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td>'+(i.行号||'—')+'</td><td><span class="tag">'+i.类型+'</span></td><td>'+(i.说明||'')+'</td>';
    tb.appendChild(tr);
  });
}
window.addEventListener('resize', function(){
  charts.forEach(function(c){ c.resize(); });
});
</script>
</body>
</html>
"""

KPI_TPL = """<section class="kpis">
{cards}
</section>"""


def kpi_card(label, value, warn=False):
    cls = "value warn" if warn else "value"
    return '  <div class="kpi"><div class="label">%s</div><div class="%s">%s</div></div>' % (label, cls, value)


def main():
    ap = argparse.ArgumentParser(
        description="生成 AI 票据智能台账 v0.2 经营看板（单页 HTML + 本地 ECharts）。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    ap.add_argument("--src", default=os.path.join(BASE, "data", "raw_tickets_v2.csv"), help="票据流水 CSV 路径")
    ap.add_argument("--out", default=os.path.join(BASE, "web", "dashboard_v2.html"), help="看板 HTML 输出路径")
    ap.add_argument("--cash-begin", type=float, default=0.0, help="期初库存现金（现金日结累计起点）")
    args = ap.parse_args()

    if not os.path.isfile(args.src):
        ap.error("找不到数据文件：%s" % args.src)

    rows, issues = core.load(args.src)
    cash = core.cash_check(rows, issues, args.cash_begin)
    months = core.monthly(rows)

    income = sum(r["金额"] for r in rows if r["收付"] == "收")
    expense = sum(r["金额"] for r in rows if r["收付"] == "支")
    food = sum(r["金额"] for r in rows if r["收付"] == "支" and r["科目"] == "主营业务成本-食材")
    days = len({r["日期"] for r in rows if r["日期"]})
    review = len({i[0] for i in issues if i[0] > 0})
    gross = (income - food) / income if income else 0

    by_acc = defaultdict(float)
    for r in rows:
        if r["收付"] == "支":
            by_acc[r["科目"]] += r["金额"]

    by_day = defaultdict(lambda: [0.0, 0.0])
    for r in rows:
        if not r["日期"][:4].isdigit():
            continue
        idx = 0 if r["收付"] == "收" else 1
        by_day[r["日期"]][idx] += r["金额"]

    data = {
        "kpi": {
            "income": round(income, 2), "expense": round(expense, 2),
            "net": round(income - expense, 2), "gross": round(gross, 4),
        },
        "months": [{"月": m["月份"], "收入": round(m["收入"], 2), "支出": round(m["支出"], 2),
                    "净额": round(m["净额"], 2), "毛利率": round(m["毛利率"], 4),
                    "环比": None if m["净额环比"] is None else round(m["净额环比"], 4)} for m in months],
        "accounts": [{"name": k, "value": round(v, 2)} for k, v in sorted(by_acc.items(), key=lambda x: -x[1])],
        "daily": [{"日期": d, "收入": round(v[0], 2), "支出": round(v[1], 2)} for d, v in sorted(by_day.items())],
        "cash": [{"日期": c["日期"], "现金收入": round(c["现金收入"], 2), "现金支出": round(c["现金支出"], 2),
                  "累计库存": round(c["累计库存"], 2)} for c in cash],
        "issues": [{"行号": a if a else "", "类型": b, "说明": c} for a, b, c in issues],
    }

    cards = [
        kpi_card("收入合计", "%.2f" % income),
        kpi_card("支出合计", "%.2f" % expense),
        kpi_card("净额", "%.2f" % (income - expense)),
        kpi_card("毛利率", "%.2f%%" % (gross * 100)),
        kpi_card("票据笔数", "%d 笔" % len(rows)),
        kpi_card("有交易天数", "%d 天" % days),
        kpi_card("日均支出", "%.2f" % (expense / days if days else 0)),
        kpi_card("待复核条目", "%d 条" % review, warn=review > 0),
    ]
    now = (datetime.now(timezone.utc) + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M")

    html = (HTML.replace("__TIME__", now)
                .replace("__KPI__", KPI_TPL.format(cards="\n".join(cards)))
                .replace("__DATA__", json.dumps(data, ensure_ascii=False)))

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)
    print("已生成：%s（%d 字节）" % (args.out, len(html.encode("utf-8"))))
    print("收入 %.2f｜支出 %.2f｜净额 %.2f｜毛利率 %.2f%%｜待复核 %d 条"
          % (income, expense, income - expense, gross * 100, review))


if __name__ == "__main__":
    main()