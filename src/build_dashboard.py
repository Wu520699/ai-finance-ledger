#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成门店经营可视化看板（单页 HTML + 本地 echarts，离线可用）"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from build_ledger import load, summarize  # noqa: E402

OUT_DIR = os.path.join(BASE, "web")
OUT_HTML = os.path.join(OUT_DIR, "dashboard.html")

TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>门店经营看板 · AI 票据智能台账</title>
<script src="vendor/echarts.min.js"></script>
<style>
  :root{--bg:#f4f6fa;--card:#fff;--ink:#1f2d3d;--muted:#7c8798;}
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;padding:14px 12px 40px}
  header h1{font-size:19px;margin:6px 0 4px}
  header p{margin:0 0 12px;color:var(--muted);font-size:12px}
  .kpis{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin-bottom:12px}
  .kpi{background:var(--card);border-radius:12px;padding:10px 12px;box-shadow:0 1px 3px rgba(20,40,80,.06)}
  .kpi .t{font-size:12px;color:var(--muted)}
  .kpi .v{font-size:19px;font-weight:600;margin-top:4px}
  .card{background:var(--card);border-radius:14px;padding:12px;margin-bottom:12px;box-shadow:0 1px 3px rgba(20,40,80,.06)}
  .card h2{font-size:14px;margin:2px 0 8px;font-weight:600}
  .chart{width:100%;height:250px}
  footer{color:var(--muted);font-size:11px;text-align:center;margin-top:16px;line-height:1.7}
</style>
</head>
<body>
<header>
  <h1 id="shop">门店经营看板</h1>
  <p id="sub">数据来源：票据 AI 识别 + 自动入账</p>
</header>
<section class="kpis" id="kpis"></section>
<div class="card"><h2>每日营业额与支出</h2><div id="c1" class="chart"></div></div>
<div class="card"><h2>成本费用结构</h2><div id="c2" class="chart" style="height:280px"></div></div>
<div class="card"><h2>每日净额与毛利率</h2><div id="c3" class="chart"></div></div>
<div class="card"><h2>总收入 / 总支出 / 净现金流</h2><div id="c4" class="chart" style="height:220px"></div></div>
<footer>AI 票据智能台账 · 业财融合实践<br>数据已脱敏，仅用于教学与竞赛展示</footer>
<script>
const D = __DATA__;
const money = n => n.toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2});
document.getElementById('shop').textContent = D.shop + ' 经营看板';
document.getElementById('sub').textContent = '统计区间 ' + D.range + ' ｜ 共 ' + D.kpi.days + ' 天 ｜ 票据 AI 识别 + 自动入账';
const kpiList = [
  ['营业额合计', money(D.kpi.income) + ' 元'],
  ['支出合计', money(D.kpi.expense) + ' 元'],
  ['净现金流', money(D.kpi.net) + ' 元'],
  ['毛利率', D.kpi.gross_rate.toFixed(1) + ' %'],
  ['日均营业额', money(D.kpi.avg_income) + ' 元'],
  ['食材占营业额', D.kpi.food_rate.toFixed(1) + ' %']
];
document.getElementById('kpis').innerHTML = kpiList.map(function(k){
  return '<div class="kpi"><div class="t">' + k[0] + '</div><div class="v">' + k[1] + '</div></div>';
}).join('');

const base = {color:['#3b6cf6','#e8663c','#22a06b','#8b5cf6','#f4b740','#12b5cb'], textStyle:{fontSize:11}};
const c1 = echarts.init(document.getElementById('c1'));
c1.setOption(Object.assign({}, base, {
  tooltip:{trigger:'axis'}, legend:{top:0,itemWidth:10,itemHeight:8,textStyle:{fontSize:11}},
  grid:{left:46,right:12,top:34,bottom:26},
  xAxis:{type:'category',data:D.days.map(function(d){return d.slice(5)}),axisLabel:{fontSize:10}},
  yAxis:{type:'value',axisLabel:{fontSize:10}},
  series:[
    {name:'营业额',type:'bar',data:D.income,itemStyle:{borderRadius:[4,4,0,0]}},
    {name:'支出',type:'bar',data:D.expense,itemStyle:{borderRadius:[4,4,0,0]}}
  ]
}));

const c2 = echarts.init(document.getElementById('c2'));
c2.setOption(Object.assign({}, base, {
  tooltip:{trigger:'item',formatter:'{b}<br/>{c} 元 ({d}%)'},
  legend:{bottom:0,itemWidth:10,itemHeight:8,textStyle:{fontSize:10}},
  series:[{type:'pie',radius:['42%','68%'],center:['50%','42%'],
    label:{fontSize:10,formatter:'{b}\n{d}%'},
    data:D.acc}]
}));

const c3 = echarts.init(document.getElementById('c3'));
c3.setOption(Object.assign({}, base, {
  tooltip:{trigger:'axis'}, legend:{top:0,itemWidth:10,itemHeight:8,textStyle:{fontSize:11}},
  grid:{left:46,right:44,top:34,bottom:26},
  xAxis:{type:'category',data:D.days.map(function(d){return d.slice(5)}),axisLabel:{fontSize:10}},
  yAxis:[{type:'value',name:'净额',axisLabel:{fontSize:10}},
         {type:'value',name:'毛利率%',axisLabel:{fontSize:10}}],
  series:[
    {name:'当日净额',type:'line',smooth:true,areaStyle:{opacity:.12},data:D.net},
    {name:'毛利率%',type:'line',yAxisIndex:1,smooth:true,data:D.gross_rate}
  ]
}));

const c4 = echarts.init(document.getElementById('c4'));
c4.setOption(Object.assign({}, base, {
  tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},
  grid:{left:56,right:14,top:20,bottom:26},
  xAxis:{type:'category',data:['总收入','总支出','净现金流'],axisLabel:{fontSize:11}},
  yAxis:{type:'value',axisLabel:{fontSize:10}},
  series:[{type:'bar',barWidth:'42%',
    label:{show:true,position:'top',fontSize:11},
    data:[
      {value:D.kpi.income,itemStyle:{color:'#3b6cf6'}},
      {value:D.kpi.expense,itemStyle:{color:'#e8663c'}},
      {value:D.kpi.net,itemStyle:{color:'#22a06b'}}
    ]}]
}));
window.addEventListener('resize',function(){[c1,c2,c3,c4].forEach(function(c){c.resize()})});
</script>
</body>
</html>
"""


def build(shop="烤馒头店"):
    rows = load()
    s = summarize(rows)
    days = sorted(s["by_day"])
    income = s["income"]
    gross_rate = []
    for d in days:
        inc = s["by_day"][d][0]
        day_food = sum(r["金额"] for r in rows if r["日期"] == d and r["科目"] == "主营业务成本-食材")
        gross_rate.append(round((inc - day_food) / inc * 100, 1) if inc else 0.0)
    data = {
        "shop": shop,
        "range": ("%s ~ %s" % (days[0], days[-1])) if days else "",
        "days": days,
        "income": [round(s["by_day"][d][0], 2) for d in days],
        "expense": [round(s["by_day"][d][1], 2) for d in days],
        "net": [round(s["by_day"][d][0] - s["by_day"][d][1], 2) for d in days],
        "gross_rate": gross_rate,
        "acc": [{"name": k, "value": round(v, 2)} for k, v in sorted(s["by_acc"].items(), key=lambda x: -x[1])],
        "kpi": {
            "days": s["days"],
            "income": round(income, 2),
            "expense": round(s["expense"], 2),
            "net": round(s["profit"], 2),
            "gross_rate": round(s["gross"] / income * 100, 1) if income else 0,
            "avg_income": round(income / max(s["days"], 1), 2),
            "food_rate": round(s["food"] / income * 100, 1) if income else 0,
        },
    }
    os.makedirs(OUT_DIR, exist_ok=True)
    html = TEMPLATE.replace("__DATA__", json.dumps(data, ensure_ascii=False))
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print("看板已生成：", OUT_HTML)
    print("KPI：", json.dumps(data["kpi"], ensure_ascii=False))
    return OUT_HTML


if __name__ == "__main__":
    build()
