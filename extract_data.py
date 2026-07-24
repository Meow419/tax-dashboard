import openpyxl, json

wb = openpyxl.load_workbook(r"D:\AI分析\税费分析\2026年税费统计分析表_终版7.23_v2.xlsx", data_only=True)

ws_yr = wb['年度总览-2026']
cities = []; companies = []; all_raw = []
for r in range(2, ws_yr.max_row):
    label = ws_yr.cell(r,1).value
    name = ws_yr.cell(r,2).value
    nature = ws_yr.cell(r,3).value
    inc = float(ws_yr.cell(r,4).value or 0)
    vat = float(ws_yr.cell(r,5).value or 0)
    surtax = float(ws_yr.cell(r,6).value or 0)      # 附加税
    cit = float(ws_yr.cell(r,7).value or 0)          # 企业所得税
    iit = float(ws_yr.cell(r,8).value or 0)          # 个税
    stamp = float(ws_yr.cell(r,9).value or 0)        # 印花税
    cultural = float(ws_yr.cell(r,10).value or 0)    # 文建费
    tt = float(ws_yr.cell(r,15).value or 0)          # 税费合计
    vr = float(ws_yr.cell(r,16).value or 0)
    tr = float(ws_yr.cell(r,18).value or 0)
    
    if label and '小计' in str(label):
        city = str(label).replace('【','').replace('小计】','').replace('】','')
        cities.append({'name':city,'income':round(inc/10000,2),'vat':round(vat/10000,2),
                       'surtax':round(surtax/10000,2),'cit':round(cit/10000,2),
                       'iit':round(iit/10000,2),'stamp':round(stamp/10000,2),
                       'cultural':round(cultural/10000,2),'total_tax':round(tt/10000,2),
                       'vat_rate':round(vr*100,4),'total_rate':round(tr*100,4)})
    elif name and '小计' not in str(name) and '合计' not in str(name):
        companies.append({'name':name,'city':label or '','nature':nature or '',
                          'income':round(inc/10000,2),'vat':round(vat/10000,2),
                          'surtax':round(surtax/10000,2),'cit':round(cit/10000,2),
                          'iit':round(iit/10000,2),'stamp':round(stamp/10000,2),
                          'cultural':round(cultural/10000,2),'total_tax':round(tt/10000,2),
                          'vat_rate':round(vr*100,4),'total_rate':round(tr*100,4)})
        all_raw.append({'name':name,'city':label or '','nature':nature or '',
                        'income':inc,'vat':vat,'surtax':surtax,'cit':cit,'iit':iit,
                        'stamp':stamp,'cultural':cultural,'total_tax':tt})

lr = ws_yr.max_row
gt = {'income':round(float(ws_yr.cell(lr,4).value or 0)/10000,2),
      'vat':round(float(ws_yr.cell(lr,5).value or 0)/10000,2),
      'surtax':round(float(ws_yr.cell(lr,6).value or 0)/10000,2),
      'cit':round(float(ws_yr.cell(lr,7).value or 0)/10000,2),
      'iit':round(float(ws_yr.cell(lr,8).value or 0)/10000,2),
      'stamp':round(float(ws_yr.cell(lr,9).value or 0)/10000,2),
      'cultural':round(float(ws_yr.cell(lr,10).value or 0)/10000,2),
      'tax':round(float(ws_yr.cell(lr,15).value or 0)/10000,2),
      'rate':round(float(ws_yr.cell(lr,18).value or 0)*100,4)}

# 月度趋势
ws_mo = wb['累计同比-城市']
months_order = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
monthly = {}
for r in range(2, ws_mo.max_row+1):
    m = ws_mo.cell(r,1).value; cy = ws_mo.cell(r,2).value
    if cy == '【全集团合计】':
        monthly[m] = {'inc26':round(float(ws_mo.cell(r,3).value or 0)/10000,2),
                      'inc25':round(float(ws_mo.cell(r,4).value or 0)/10000,2),
                      'vat26':round(float(ws_mo.cell(r,6).value or 0)/10000,2)}

# 小规模纳税人分析
small_raw = [c for c in all_raw if '小规模' in str(c['nature'])]
general_raw = [c for c in all_raw if '一般' in str(c['nature'])]
small_count = len(small_raw); general_count = len(general_raw)
si = sum(c['income'] for c in small_raw)
gi = sum(c['income'] for c in general_raw)

# 滚动12个月收入预警（2025全年+2026上半年）
# 加载2025年度数据
wb25 = openpyxl.load_workbook(r"D:\AI分析\税费分析\税费统计表2025 .xlsx", data_only=True)
ws25 = wb25['25-汇总']
inc_2025 = {}  # {name: income}
for r in range(4, ws25.max_row+1):
    n = ws25.cell(r,4).value; cy = ws25.cell(r,2).value
    if not n: continue
    ns = str(n).strip()
    if '小计' in ns or '合计' in ns or 'check' in ns: continue
    if ns in ['上海','北京','南京','广州','成都','杭州','武汉','深圳']: continue
    if cy is None or str(cy).strip() == '': continue
    inc_2025[ns] = float(ws25.cell(r,16).value or 0)

# 2026年已过月份数 (1-6月=6个月)
months_elapsed_2026 = 6

small_detail = []
for c in small_raw:
    name = c['name']
    inc_2026_ytd = c['income']  # 万元
    inc_2025_full = inc_2025.get(name, 0)  # 元
    inc_2025_h2 = inc_2025_full / 2  # 估算下半年
    
    # 滚动12个月 = 2025下半年 + 2026上半年（均以万元计）
    rolling_12m = inc_2025_h2/10000 + inc_2026_ytd
    # 年化收入（如果仅有部分月份数据）
    annualized = inc_2026_ytd * 12 / months_elapsed_2026
    
    rate = round(c['total_tax']/c['income']*100,4) if c['income']>0 else 0
    vat_rate = round(c['vat']/c['income']*100,4) if c['income']>0 else 0
    
    warn_500 = rolling_12m > 500 or annualized > 500
    warn_text = ''
    if rolling_12m > 500:
        warn_text = f'滚动12月收入超500万(¥{rolling_12m:.0f}万)'
    elif annualized > 500:
        warn_text = f'年化收入超500万(¥{annualized:.0f}万)'
    
    small_detail.append({
        'name':name,'city':c['city'],
        'income':round(c['income'],2),'vat':round(c['vat'],2),
        'tax':round(c['total_tax'],2),'vat_rate':vat_rate,
        'total_rate':rate,
        'rolling_12m':round(rolling_12m,0),
        'annualized':round(annualized,0),
        'warn_500':warn_500,'warn_text':warn_text
    })

wb25.close()
wb.close()

# 小规模按城市分布
small_by_city = {}
for c in small_raw:
    cy = c['city']
    if cy not in small_by_city: small_by_city[cy]={'count':0,'income':0,'tax':0,'warn':0}
    small_by_city[cy]['count']+=1
    small_by_city[cy]['income']+=c['income']
    small_by_city[cy]['tax']+=c['total_tax']
    detail = next((x for x in small_detail if x['name']==c['name']), None)
    if detail and detail['warn_500']: small_by_city[cy]['warn']+=1

output = {
    'grand_total': gt,
    'cities': cities,
    'companies': companies,
    'monthly': monthly,
    'small_taxpayer': {
        'small_count':small_count,'general_count':general_count,
        'small_income':round(si/10000,2),'general_income':round(gi/10000,2),
        'small_tax':round(sum(c['total_tax'] for c in small_raw)/10000,2),
        'general_tax':round(sum(c['total_tax'] for c in general_raw)/10000,2),
        'small_rate':round(sum(c['total_tax'] for c in small_raw)/si*100,4) if si>0 else 0,
        'general_rate':round(sum(c['total_tax'] for c in general_raw)/gi*100,4) if gi>0 else 0,
        'by_city':{cy:{'count':v['count'],'income':round(v['income']/10000,2),'tax':round(v['tax']/10000,2),'warn':v['warn']}
                   for cy,v in sorted(small_by_city.items(), key=lambda x:-x[1]['income'])},
        'detail': small_detail
    },
    'months': months_order
}
# 加回alerts
# 从增值税税负预警sheet读取
ws_warn = wb['增值税税负预警']  
alerts = {}
for r in range(2, ws_warn.max_row+1):
    sig = ws_warn.cell(r,11).value; n = ws_warn.cell(r,2).value
    m = ws_warn.cell(r,1).value; det = ws_warn.cell(r,12).value
    inc = float(ws_warn.cell(r,5).value or 0)
    if sig and sig != '正常' and n and inc > 0:
        if n not in alerts: alerts[n] = {'months':[],'details':set()}
        alerts[n]['months'].append(m)
        if det and det != '正常': alerts[n]['details'].add(det)
output['alerts'] = {n:{'count':len(v['months']),'details':list(v['details'])[:3],'months':v['months']}
                    for n,v in sorted(alerts.items(), key=lambda x:-len(x[1]['months']))[:20]}

with open(r'D:\AI分析\BI看板\data.json','w',encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"OK: {len(cities)} cities, {len(companies)} companies, {len(small_detail)} small taxpayers")
print(f"Small rolling >500w: {sum(1 for x in small_detail if x['warn_500'])} companies")
