import json

with open(r'D:\AI分析\税费分析\2026年税费分析看板\data.json','r',encoding='utf-8') as f:
    data = f.read()

with open(r'D:\AI分析\税费分析\2026年税费分析看板\index.html','r',encoding='utf-8') as f:
    html = f.read()

# Replace fetch+init with embedded data
old_init = """async function init(){try{
  document.getElementById('panel-overview').innerHTML='<div style="padding:40px;text-align:center;color:#666">加载中...</div>';
  const r=await fetch('data.json?v=20260724');D=await r.json();
  console.log('Data loaded:',D.grand_total.income);
  document.getElementById('updateTime').textContent='数据更新：2026-07-23';
  renderOverview();renderCityPanel();renderCompanyPanel();renderSmallPanel();initPivot();renderAlerts();
  console.log('All panels rendered:',typeof echarts);
}catch(e){console.error(e);document.getElementById('panel-overview').innerHTML='<div style="padding:40px;text-align:center;color:#c00">数据加载失败：'+e.message+'</div>';}}"""

new_init = """function init(){
  D=EMBEDDED_DATA;
  document.getElementById('updateTime').textContent='数据更新：2026-07-23';
  renderOverview();renderCityPanel();renderCompanyPanel();renderSmallPanel();initPivot();renderAlerts();
}"""

html = html.replace(old_init, new_init)

# Insert EMBEDDED_DATA before the script section
data_line = f'const EMBEDDED_DATA={data};\n'
html = html.replace('<script>\nlet D=null;', f'<script>\n{data_line}let D=null;')

with open(r'D:\AI分析\税费分析\2026年税费分析看板\index.html','w',encoding='utf-8') as f:
    f.write(html)

print(f'Done! HTML size: {len(html)} chars')
