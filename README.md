# 2026年税费分析BI看板

基于税费统计分析表自动生成的交互式看板，部署在 GitHub Pages。

## 目录结构

```
├── index.html        # 看板主页面（ECharts图表）
├── data.json          # 从Excel提取的分析数据
├── extract_data.py    # 数据提取脚本
└── .github/workflows/ # GitHub Actions自动部署配置
```

## 更新数据

1. 更新Excel文件后，运行：
   ```
   python extract_data.py
   ```
2. 将 data.json 提交到 GitHub，Actions自动部署更新

## 看板内容

- 月度收入趋势（2026 vs 2025同比）
- 各城市收入与税费对比
- 税负率排名
- 月度增值税趋势
- 增值税预警提醒
