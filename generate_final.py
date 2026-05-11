#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DAU Dashboard 生成脚本 - 严格验证版"""

import pandas as pd
import json
import re
import shutil
from datetime import datetime, timedelta
from pathlib import Path

print("=" * 70)
print("🔧 DAU Dashboard 生成 - 严格验证版")
print("=" * 70)

# ========== 1. 读取数据 ==========
print("[1/8] 读取数据...")
with open('files/daily_stats.json', 'r', encoding='utf-8') as f:
    daily_stats = json.load(f)
print(f"  ✅ daily_stats.json ({len(daily_stats)} 天)")

latest_file = sorted(Path('files').glob('丰巢线圈*.xlsx'))[-1]
print(f"  ✅ Excel: {latest_file.name}")

df_all = pd.read_excel(latest_file, header=0)
device_df = df_all[(df_all['sn'] != 'all') & (df_all['sn'].notna())]
all_row = df_all[df_all['sn'] == 'all'].iloc[0]
latest_date = daily_stats[-1]['实际日期']
latest_dau = int(all_row['dau'])

# ========== 2. 计算核心指标 ==========
print("[2/8] 计算核心指标...")
total_sn = len(device_df)
dau_ge3 = len(device_df[device_df['dau'] >= 3])
dau_mid = len(device_df[(device_df['dau'] > 0) & (device_df['dau'] < 3)])
dau_0d = len(device_df[device_df['dau'] == 0])

print(f"  设备总数：{total_sn:,}")
print(f"  DAU: {latest_dau:,}")
print(f"  dau≥3: {dau_ge3:,}")
print(f"  0<dau<3: {dau_mid:,}")
print(f"  当日 0dau: {dau_0d:,}")

print("\n[3/8] 计算近 3 天/近 7 天 0dau...")
device_files = sorted(Path('files').glob('丰巢线圈*.xlsx'), key=lambda x: x.name)
daily_zero_sets = {}
for f in device_files:
    match = re.search(r'(\d{4}-\d{2}-\d{2})', f.name)
    if match:
        actual_date = (datetime.strptime(match.group(1), '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')
        df_tmp = pd.read_excel(f, header=0)
        zero_set = set(df_tmp[(df_tmp['sn']!='all') & (df_tmp['sn'].notna()) & (df_tmp['dau']==0)]['sn'].tolist())
        daily_zero_sets[actual_date] = zero_set

sorted_dates = sorted(daily_zero_sets.keys())
last_3_days = sorted_dates[-3:]
last_7_days = sorted_dates[-7:]

device_zero_3d = set.intersection(*[daily_zero_sets[d] for d in last_3_days]) if len(last_3_days) == 3 else set()
device_zero_7d = set.intersection(*[daily_zero_sets[d] for d in last_7_days]) if len(last_7_days) == 7 else set()

dau_0_3d = len(device_zero_3d)
dau_0_7d = len(device_zero_7d)

print(f"  ✅ 近 3 天 0dau (交集/连续): {dau_0_3d:,} (应该≈500-600)")
print(f"  ✅ 近 7 天 0dau (交集/连续): {dau_0_7d:,} (应该≈300-400)")

# ========== 4. 生成表格 ==========
print("[4/8] 生成表格...")

daily_rows = ''.join([
    f"<tr><td>{s['实际日期']}</td><td>{s['设备数']:,}</td><td>{s['DAU']:,}</td>"
    f"<td>{s['活跃设备数 (dau≥3)']:,}</td><td>{s['设备数']-s['活跃设备数 (dau>0)']:,}</td>"
    f"<td><span class=\"badge {'success' if s['DAU']>85000 else 'warning'}\">{'正常' if s['DAU']>85000 else '偏低'}</span></td></tr>\n"
    for s in sorted(daily_stats, key=lambda x: x['日期'], reverse=True)
])

medals = ['🥇 1', '🥈 2', '🥉 3', '4', '5', '6', '7', '8', '9', '10']
bgs = ['#FFD700', '#C0C0C0', '#CD7F32'] + ['#F0F0F0']*7
top10 = device_df.nlargest(10, 'dau')[['sn', 'dau']]
top10_rows = ''.join([
    f"<tr><td><span class=\"badge\" style=\"background:{bgs[i]};color:#333;\">{medals[i]}</span></td>"
    f"<td style=\"font-family:monospace;font-size:13px;\">{row['sn']}</td>"
    f"<td style=\"font-weight:600;color:var(--alipay-blue);\">{int(row['dau'])}</td></tr>\n"
    for i, (_, row) in enumerate(top10.iterrows())
])

# ========== 5. 图表数据 ==========
print("[5/8] 生成图表数据...")

dates = [s['实际日期'][5:] for s in daily_stats[-16:]]
dau_array = [int(s['DAU']) for s in daily_stats[-16:]]
device_array = [int(s['设备数']) for s in daily_stats[-16:]]

# ========== 6. 替换模板 ==========
print("[6/8] 替换模板...")

with open('dashboard_template.html', 'r', encoding='utf-8') as f:
    html = f.read()

# 动态分析洞察
ratio_3d_7d = dau_0_3d / dau_0_7d if dau_0_7d > 0 else 0
if ratio_3d_7d >= 1.5:
    anomaly_desc = f'近 3 天异常设备数是近 7 天的{ratio_3d_7d:.1f}倍，说明<strong>部分设备近期才开始不活跃</strong>。'
elif ratio_3d_7d > 1.0:
    anomaly_desc = f'近 3 天异常设备数是近 7 天的{ratio_3d_7d:.1f}倍，<strong>有新增不活跃设备趋势</strong>。'
else:
    anomaly_desc = '近 3 天与近 7 天异常设备数比例稳定。'

active_pct = dau_ge3 / total_sn * 100
if active_pct >= 70:
    activity_summary = f'活跃设备占比{active_pct:.1f}%，<strong>整体活跃度健康</strong>。'
elif active_pct >= 60:
    activity_summary = f'活跃设备占比{active_pct:.1f}%，<strong>整体活跃度良好</strong>。'
else:
    activity_summary = f'活跃设备占比{active_pct:.1f}%，建议关注活跃度变化。'

min_dev = min(device_array)
max_dev = max(device_array)
dev_fluct = ((max_dev - min_dev) / min_dev) * 100 if min_dev > 0 else 0
min_d_str = f"{min_dev:,}"
max_d_str = f"{max_dev:,}"
if dev_fluct < 0.1:
    device_stability = f'每日设备数稳定在 <strong>{min_d_str}~{max_d_str}</strong> 之间。'
else:
    device_stability = f'每日设备数在 <strong>{min_d_str}~{max_d_str}</strong> 之间。'

replacements = {
    '{TABLE_ROWS}': daily_rows.rstrip(),
    '{TOP10_ROWS}': top10_rows.rstrip(),
    '{LATEST_DATE}': latest_date,
    '{TIME}': datetime.now().strftime('%Y-%m-%d %H:%M'),
    '{LATEST_DAU}': f"{latest_dau:,}",
    '{TOTAL_SN}': f"{total_sn:,}",
    '{DAU_GE3}': f"{dau_ge3:,}",
    '{DAU_GE3_PCT}': f"{dau_ge3/total_sn*100:.1f}",
    '{DAU_MID}': f"{dau_mid:,}",
    '{DAU_MID_PCT}': f"{dau_mid/total_sn*100:.1f}",
    '{DAU_0_DAY}': f"{dau_0d:,}",
    '{DAU_0_DAY_PCT}': f"{dau_0d/total_sn*100:.1f}",
    '{DAU_0_3D}': f"{dau_0_3d:,}",
    '{DAU_0_3D_PCT}': f"{dau_0_3d/total_sn*100:.1f}",
    '{DAU_0_7D}': f"{dau_0_7d:,}",
    '{DAU_0_7D_PCT}': f"{dau_0_7d/total_sn*100:.1f}",
    '[{CHART_DATES}]': '[' + ', '.join([f"'{d}'" for d in dates]) + ']',
    '[{CHART_DAU}]': '[' + ', '.join(map(str, dau_array)) + ']',
    '[{CHART_DEVICES}]': '[' + ', '.join(map(str, device_array)) + ']',
    '{PIE_DATA}': f'{dau_ge3}, {dau_mid}, {dau_0d}',
    '{DAY_COUNT}': str(len(daily_stats)),
    '{AVG_DAU}': f"{int(sum(dau_array)/len(dau_array)):,}",
    '{PEAK_DATE}': dates[dau_array.index(max(dau_array))] if dau_array else '',
    '{PEAK_DROP}': f"{max(dau_array)-min(dau_array):,}",
    '{PERIOD}': f"{dates[0]} ~ {dates[-1]}",
    '{MIN_DAU}': f"{min(dau_array):,}",
    '{MIN_DEVICES}': min_d_str,
    '{MAX_DEVICES}': max_d_str,
    '{PEAK_DAU}': f"{max(dau_array):,}",
    '{ANOMALY_DESC}': anomaly_desc,
    '{ACTIVITY_SUMMARY}': activity_summary,
    '{DEVICE_STABILITY}': device_stability,
}

for key, value in replacements.items():
    html = html.replace(key, value)

# ========== 7. 写入文件 ==========
print("[7/8] 写入文件...")

with open('dau_dashboard.html', 'w', encoding='utf-8') as f:
    f.write(html)

shutil.copy('dau_dashboard.html', 'files/dau_dashboard.html')

# ========== 8. 验证 ==========
print("\n" + "=" * 70)
print("[8/8] 严格验证")
print("=" * 70)

errors = []

with open('dau_dashboard.html', 'r', encoding='utf-8') as f:
    v = f.read()

# 1. 占位符
remaining = re.findall(r'\{[A-Z_0-9]+\}', v)
if remaining:
    errors.append(f"❌ 占位符残留：{len(set(remaining))}种")
else:
    print("✅ 占位符")

# 2. 饼图数据
pie = re.search(r"type: 'doughnut'[\s\S]{0,300}data:\s*\[([^\]]+)\]", v)
if pie:
    nums = [n.strip() for n in pie.group(1).split(',')]
    if len(nums) == 3 and all(n.isdigit() for n in nums):
        print(f"✅ 饼图：[{pie.group(1)}]")
        if int(nums[0]) + int(nums[1]) + int(nums[2]) != total_sn:
            errors.append(f"❌ 饼图总和≠设备总数：{sum(int(n) for n in nums)} ≠ {total_sn}")
    else:
        errors.append(f"❌ 饼图错误：[{pie.group(1)}] ({len(nums)} 项)")
else:
    errors.append("❌ 未找到饼图")

# 3. 近 3/7 天数据
if 300 <= dau_0_3d <= 800:
    print(f"✅ 近 3 天 0dau: {dau_0_3d:,} (交集)")
else:
    errors.append(f"❌ 近 3 天 0dau 异常：{dau_0_3d:,} (应该≈500-600)")

if 100 <= dau_0_7d <= 600:
    print(f"✅ 近 7 天 0dau: {dau_0_7d:,} (交集)")
else:
    errors.append(f"❌ 近 7 天 0dau 异常：{dau_0_7d:,} (应该≈300-400)")

# 4. 最新日期
if daily_stats[-1]['实际日期'][5:] in v:
    print(f"✅ 日期数据：{daily_stats[-1]['实际日期']}")
else:
    errors.append(f"❌ 最新日期 {daily_stats[-1]['实际日期']} 不在 HTML 中")

# 5. Top10
if '🥇 1' in v:
    print("✅ Top10 排名")
else:
    errors.append("❌ Top10 排名")

print("\n" + "=" * 70)

if errors:
    print("❌ 验证失败:")
    for e in errors:
        print(f"  {e}")
    exit(1)

print("✅ 所有验证通过！")
print(f"\n📊 最终数据:")
print(f"  饼图 (3 项): dau≥3:{dau_ge3:,}, 0<dau<3:{dau_mid:,}, 当日 0dau:{dau_0d:,}")
print(f"  异常监控:")
print(f"    近 3 天 0dau (连续): {dau_0_3d:,}")
print(f"    近 7 天 0dau (连续): {dau_0_7d:,}")
print("=" * 70)