import openpyxl
import json

wb = openpyxl.load_workbook(r"D:\AI分析\税费分析\2026年税费统计分析表_终版7.23_v2.xlsx", data_only=True)

ws_yr = wb['年度总览-2026']
cities = []
for r in range(2, ws_yr.max_row):
    label = ws_yr.cell(r, 1).value
    if label and '小计' in str(label):
        name = str(label).replace('【','').replace('小计】','').replace('】','')
        income = float(ws_yr.cell(r,4).value or 0)
        vat = float(ws_yr.cell(r,5).value or 0)
        surtax = float(ws_yr.cell(r,6).value or 0)
        stamp = float(ws_yr.cell(r,9).value or 0)
        cultural = float(ws_yr.cell(r,10).value or 0)
        total_tax = float(ws_yr.cell(r,15).value or 0)
        vat_rate = float(ws_yr.cell(r,16).value or 0)
        total_rate = float(ws_yr.cell(r,18).value or 0)
        cities.append({'name':name,'income':round(income/10000,2),'vat':round(vat/10000,2),
                       'total_tax':round(total_tax/10000,2),'vat_rate':round(vat_rate*100,4),'total_rate':round(total_rate*100,4)})

# Grand total
gt_income = float(ws_yr.cell(ws_yr.max_row,4).value or 0)
gt_tax = float(ws_yr.cell(ws_yr.max_row,15).value or 0)
gt_rate = float(ws_yr.cell(ws_yr.max_row,18).value or 0)

ws_city = wb['累计同比-城市']
months_order = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
monthly = {}
for r in range(2, ws_city.max_row+1):
    m = ws_city.cell(r,1).value
    city = ws_city.cell(r,2).value
    if city == '【全集团合计】':
        inc26 = float(ws_city.cell(r,3).value or 0)
        inc25 = float(ws_city.cell(r,4).value or 0)
        vat26 = float(ws_city.cell(r,6).value or 0)
        monthly[m] = {'inc26':round(inc26/10000,2),'inc25':round(inc25/10000,2),'vat26':round(vat26/10000,2)}

# Alerts
ws_warn = wb['增值税税负预警']
alerts = {}
for r in range(2, ws_warn.max_row+1):
    sig = ws_warn.cell(r,11).value
    name = ws_warn.cell(r,2).value
    detail = ws_warn.cell(r,12).value
    if sig and sig != '正常' and name:
        if name not in alerts:
            alerts[name] = {'count':0,'details':set()}
        alerts[name]['count'] += 1
        if detail and detail != '正常': alerts[name]['details'].add(detail)

wb.close()

output = {
    'cities': cities,
    'grand_total': {'income': round(gt_income/10000,2), 'tax': round(gt_tax/10000,2), 'rate': round(gt_rate*100,4)},
    'monthly': monthly,
    'alerts': {n: {'count':v['count'], 'details': list(v['details'])[:3]} for n,v in sorted(alerts.items(), key=lambda x:-x[1]['count'])[:15]}
}

with open(r'D:\AI分析\BI看板\data.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print("Data extracted:", len(cities), "cities,", len(monthly), "months,", len(alerts), "alerts")
