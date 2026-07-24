import openpyxl, json

wb = openpyxl.load_workbook(r"D:\AI分析\税费分析\2026年税费统计分析表_终版7.23_v2.xlsx", data_only=True)

# ===== 年度总览 =====
ws_yr = wb['年度总览-2026']
cities = []; companies = []
for r in range(2, ws_yr.max_row):
    label = ws_yr.cell(r,1).value
    name = ws_yr.cell(r,2).value
    inc = float(ws_yr.cell(r,4).value or 0)
    vat = float(ws_yr.cell(r,5).value or 0)
    st = float(ws_yr.cell(r,6).value or 0)
    stamp = float(ws_yr.cell(r,9).value or 0)
    cult = float(ws_yr.cell(r,10).value or 0)
    tt = float(ws_yr.cell(r,15).value or 0)
    vr = float(ws_yr.cell(r,16).value or 0)
    tr = float(ws_yr.cell(r,18).value or 0)
    
    if label and '小计' in str(label):
        city = str(label).replace('【','').replace('小计】','').replace('】','')
        cities.append({'name':city,'income':round(inc/10000,2),'vat':round(vat/10000,2),
                       'total_tax':round(tt/10000,2),'vat_rate':round(vr*100,4),'total_rate':round(tr*100,4)})
    elif name and '小计' not in str(name) and '合计' not in str(name):
        companies.append({'name':name,'city':label or '','income':round(inc/10000,2),
                          'vat':round(vat/10000,2),'total_tax':round(tt/10000,2),
                          'vat_rate':round(vr*100,4),'total_rate':round(tr*100,4)})

# 全集团合计
lr = ws_yr.max_row
gt = {'income':round(float(ws_yr.cell(lr,4).value or 0)/10000,2),
      'tax':round(float(ws_yr.cell(lr,15).value or 0)/10000,2),
      'rate':round(float(ws_yr.cell(lr,18).value or 0)*100,4)}

# ===== 月度趋势 =====
ws_mo = wb['累计同比-城市']
months_order = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
monthly = {}
for r in range(2, ws_mo.max_row+1):
    m = ws_mo.cell(r,1).value; cy = ws_mo.cell(r,2).value
    if cy == '【全集团合计】':
        monthly[m] = {'inc26':round(float(ws_mo.cell(r,3).value or 0)/10000,2),
                      'inc25':round(float(ws_mo.cell(r,4).value or 0)/10000,2),
                      'vat26':round(float(ws_mo.cell(r,6).value or 0)/10000,2)}

# ===== 公司月度明细 =====
ws_md = wb['月度明细-2026']
company_monthly = {}
for r in range(2, ws_md.max_row+1):
    m = ws_md.cell(r,1).value; n = ws_md.cell(r,3).value
    inc = float(ws_md.cell(r,4).value or 0)
    vat = float(ws_md.cell(r,5).value or 0)
    st = float(ws_md.cell(r,6).value or 0)
    if n and '全集团' not in str(n):
        if n not in company_monthly: company_monthly[n] = {}
        company_monthly[n][m] = {'inc':round(inc/10000,4),'vat':round(vat/10000,4)}

# ===== 增值税预警 =====
ws_warn = wb['增值税税负预警']
alerts = {}
for r in range(2, ws_warn.max_row+1):
    sig = ws_warn.cell(r,11).value
    n = ws_warn.cell(r,2).value; m = ws_warn.cell(r,1).value
    det = ws_warn.cell(r,12).value
    inc = float(ws_warn.cell(r,5).value or 0)
    if sig and sig != '正常' and n and inc > 0:
        if n not in alerts: alerts[n] = {'months':[],'details':set()}
        alerts[n]['months'].append(m)
        if det and det != '正常': alerts[n]['details'].add(det)

# ===== 按城市的月度明细 =====
city_monthly = {}
for r in range(2, ws_md.max_row+1):
    m = ws_md.cell(r,1).value; n = ws_md.cell(r,3).value
    cy = ws_md.cell(r,2).value
    inc = float(ws_md.cell(r,4).value or 0)
    if n and '全集团' not in str(n) and cy and cy != '—':
        if cy not in city_monthly: city_monthly[cy] = {}
        if m not in city_monthly[cy]: city_monthly[cy][m] = 0
        city_monthly[cy][m] += round(inc/10000,4)

wb.close()

output = {
    'grand_total': gt,
    'cities': cities,
    'companies': companies,
    'monthly': monthly,
    'alerts': {n:{'count':len(v['months']),'details':list(v['details'])[:3],'months':v['months']}
               for n,v in sorted(alerts.items(), key=lambda x:-len(x[1]['months']))[:20]},
    'company_monthly': company_monthly,
    'city_monthly': city_monthly,
    'months': months_order
}

with open(r'D:\AI分析\BI看板\data.json','w',encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"Cities:{len(cities)} Companies:{len(companies)} Alerts:{len(alerts)} CompanyMonthly:{len(company_monthly)}")
