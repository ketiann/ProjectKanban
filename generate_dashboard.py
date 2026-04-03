#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目数据可视化看板 - 一键生成脚本
==================================
使用方法：
  1. 将本脚本与4个源数据文件放在同一目录下
  2. 运行: python generate_dashboard.py
  3. 生成的 看板.html 即为最新看板

源数据文件（格式必须与原始模板一致）：
  - 项目数据表.xlsx
  - 人力投入表.xlsx
  - 本期人力投入项目情况(YYYY-MM-DD开始截至YYYY-MM-DD).xlsx
  - 本期项目变化情况(YYYY-MM-DD开始截至YYYY-MM-DD).xlsx
"""

import os
import sys
import json
import re
import glob
import pandas as pd

# ══════════════════════════════════════════════════════════════
# 配置区
# ══════════════════════════════════════════════════════════════
OUTPUT_FILENAME = '看板.html'          # 输出文件名
CENTER_FILTER = '数智技术服务中心'       # 人员过滤中心
EXPIRING_DAYS = 60                      # 到期预警天数阈值
TOP_N = 10                             # TOP排名数量
TOP3_N = 3                             # 部门TOP3数量

# 部门名称映射：人力投入表部门名 → 项目数据表可能的部门名
DEPT_NAME_MAP = {
    '智能部': ['智能部', '智能技术一部', '智能技术二部', '智能物联网部'],
    '政务部': ['政务部'],
    '大数据部': ['大数据部', '大数据技术部'],
    '产品创新部': ['产品创新部'],
}

# ══════════════════════════════════════════════════════════════
# 文件自动发现
# ══════════════════════════════════════════════════════════════
def find_files(work_dir):
    """自动发现源数据文件"""
    files = {}

    # 1. 项目数据表.xlsx
    pattern1 = os.path.join(work_dir, '项目数据表.xlsx')
    if os.path.exists(pattern1):
        files['project'] = pattern1
    else:
        print(f'[错误] 未找到 项目数据表.xlsx')
        return None

    # 2. 人力投入表.xlsx
    pattern2 = os.path.join(work_dir, '人力投入表.xlsx')
    if os.path.exists(pattern2):
        files['labor'] = pattern2
    else:
        print(f'[错误] 未找到 人力投入表.xlsx')
        return None

    # 3. 本期人力投入项目情况(日期).xlsx
    matches3 = glob.glob(os.path.join(work_dir, '本期人力投入项目情况*.xlsx'))
    if matches3:
        files['month_project'] = matches3[0]
    else:
        print(f'[错误] 未找到 本期人力投入项目情况*.xlsx')
        return None

    # 4. 本期项目变化情况(日期).xlsx
    matches4 = glob.glob(os.path.join(work_dir, '本期项目变化情况*.xlsx'))
    if matches4:
        files['change'] = matches4[0]
    else:
        print(f'[错误] 未找到 本期项目变化情况*.xlsx')
        return None

    return files


def extract_period(files):
    """从文件名中提取统计周期"""
    # 尝试从 本期人力投入项目情况 文件名提取日期
    fname = os.path.basename(files['month_project'])
    dates = re.findall(r'(\d{4}-\d{2}-\d{2})', fname)
    if len(dates) >= 2:
        return f"{dates[0]}至{dates[1]}"
    # 尝试从 本期项目变化情况 文件名提取
    fname2 = os.path.basename(files['change'])
    dates2 = re.findall(r'(\d{4}-\d{2}-\d{2})', fname2)
    if len(dates2) >= 2:
        return f"{dates2[0]}至{dates2[1]}"
    return '未指定周期'


# ══════════════════════════════════════════════════════════════
# 数据读取与处理
# ══════════════════════════════════════════════════════════════
def load_and_process(files):
    """读取所有源数据并处理为看板所需格式"""

    # ── 读取项目数据表 ──
    print('  读取 项目数据表.xlsx ...')
    df = pd.read_excel(files['project'], sheet_name='项目状态信息')
    df.columns = df.columns.str.strip()
    num_cols = ['合同金额（元）', '合同回款情况', '已回款金额（元）', '预估实施成本（元）',
                '项目进度', '截至上月底投入工作量（人月）', '本月投入人力（人月）',
                '累计投入工作量（人月）', '已发生用人成本（元）', '已发生差旅报销费用（元）',
                '已发生实施成本（元）', '实施成本剩余金额（元）']
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    # ── 读取人力投入表 ──
    print('  读取 人力投入表.xlsx ...')
    labor_df = pd.read_excel(files['labor'], sheet_name='人员档案台账', header=1)
    labor_df.columns = labor_df.columns.str.strip()
    labor_df['实际投入工时'] = pd.to_numeric(labor_df['实际投入工时'], errors='coerce').fillna(0)
    labor_df['应投入工时'] = pd.to_numeric(labor_df['应投入工时'], errors='coerce').fillna(0)
    labor_df['工作饱和度'] = pd.to_numeric(labor_df['工作饱和度'], errors='coerce').fillna(0)

    # ── 读取本期人力投入项目情况 ──
    print('  读取 本期人力投入项目情况 ...')
    month_proj = pd.read_excel(files['month_project'], sheet_name='本期人力投入项目情况', header=1)
    month_proj.columns = month_proj.columns.str.strip()
    for c in ['中心人力投入（人月）', '智能部', '政务部', '大数据部', '产品创新部', '其他部门投入', '本期总投入（人月）']:
        month_proj[c] = pd.to_numeric(month_proj[c], errors='coerce').fillna(0)

    # ── 读取本期项目变化情况 ──
    print('  读取 本期项目变化情况 ...')
    change_df = pd.read_excel(files['change'], sheet_name='项目变化', header=1)
    change_df.columns = change_df.columns.str.strip()

    # ── 过滤数智技术服务中心人员 ──
    labor_sz = labor_df[labor_df['所属中心'] == CENTER_FILTER].copy()

    # 构建数智技术服务中心对应的部门名列表（用于项目表过滤）
    sz_project_dept_names = []
    for short, longs in DEPT_NAME_MAP.items():
        sz_project_dept_names.extend(longs)

    data = {}

    # ══════════════════════════════════════════════
    # KPI 指标卡
    # ══════════════════════════════════════════════
    total_contract = df['合同金额（元）'].sum()
    total_cost = df['预估实施成本（元）'].sum()
    total_paid = df['已回款金额（元）'].sum()
    total_month_labor = df['本月投入人力（人月）'].sum()
    cost_ratio = total_cost / total_contract if total_contract > 0 else 0
    pay_rate = total_paid / total_contract if total_contract > 0 else 0

    data['kpi'] = {
        'total_projects': int(len(df)),
        'total_contract': float(total_contract),
        'total_paid': float(total_paid),
        'total_cost': float(total_cost),
        'cost_ratio': float(cost_ratio),
        'total_month_labor': float(total_month_labor),
        'pay_rate': float(pay_rate),
    }

    # ══════════════════════════════════════════════
    # 本月人力投入情况
    # ══════════════════════════════════════════════
    has_labor = df[df['本月投入人力（人月）'] > 0]
    no_labor = df[df['本月投入人力（人月）'] == 0]
    data['month_labor'] = {
        'has_labor_count': int(len(has_labor)),
        'no_labor_count': int(len(no_labor)),
        'has_labor_amount': float(has_labor['本月投入人力（人月）'].sum()),
    }

    # ══════════════════════════════════════════════
    # 合同类型分布
    # ══════════════════════════════════════════════
    ct = df.groupby('合同签订状态').agg(
        count=('项目名称', 'count'), amount=('合同金额（元）', 'sum')
    ).reset_index()
    data['contract_type'] = ct.to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 项目类型分布
    # ══════════════════════════════════════════════
    ptype = df.groupby('项目类型').agg(
        count=('项目名称', 'count'), amount=('合同金额（元）', 'sum')
    ).reset_index()
    ptype = ptype.sort_values('count', ascending=False)
    data['project_type'] = ptype.to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 回款率分布（仅合同项目）
    # ══════════════════════════════════════════════
    contract_df = df[df['合同签订状态'] == '合同项目']
    bins_pay = [
        (1.0, 1.001, '全额回款'), (0.8, 1.0, '80%~100%'), (0.5, 0.8, '50%~80%'),
        (0.1, 0.5, '10%~50%'), (0.0, 0.1, '0%~10%'), (-0.001, 0.0, '未回款')
    ]
    pay_dist = []
    for lo, hi, label in bins_pay:
        sub = contract_df[(contract_df['合同回款情况'] >= lo) & (contract_df['合同回款情况'] < hi)]
        pay_dist.append({'name': label, 'value': int(len(sub))})
    data['pay_distribution'] = pay_dist

    # ══════════════════════════════════════════════
    # 合同金额TOP10
    # ══════════════════════════════════════════════
    top10 = df.nlargest(TOP_N, '合同金额（元）')[
        ['项目名称', '合同金额（元）', '已回款金额（元）', '合同回款情况', '所属部门', '项目阶段', '项目进度']
    ].reset_index(drop=True)
    data['top10'] = top10.to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 人力投入累计TOP10
    # ══════════════════════════════════════════════
    top_labor = df.nlargest(TOP_N, '累计投入工作量（人月）')[
        ['项目名称', '累计投入工作量（人月）', '所属部门', '项目经理']
    ].reset_index(drop=True)
    data['top_labor'] = top_labor.to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 挂起项目
    # ══════════════════════════════════════════════
    suspended = df[df['项目阶段'] == '挂起'][
        ['项目名称', '合同金额（元）', '所属部门', '项目经理', '项目进度']
    ].reset_index(drop=True)
    data['suspended'] = suspended.to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 人员工作饱和度分析（仅数智技术服务中心）
    # ══════════════════════════════════════════════
    active_sz = labor_sz[labor_sz['实际投入工时'] > 0]
    inactive_sz = labor_sz[labor_sz['实际投入工时'] == 0]

    sat_bins = [
        (0.9, 1.001, '高饱和(≥90%)'), (0.7, 0.9, '较高(70%~90%)'), (0.5, 0.7, '中等(50%~70%)'),
        (0.1, 0.5, '较低(10%~50%)'), (0.0, 0.1, '低(<10%)'), (-0.001, 0.0, '无投入')
    ]
    sat_dist = []
    for lo, hi, label in sat_bins:
        sub = labor_sz[(labor_sz['工作饱和度'] >= lo) & (labor_sz['工作饱和度'] < hi)]
        sat_dist.append({'name': label, 'value': int(len(sub))})
    data['saturation_dist'] = sat_dist

    data['staff_summary'] = {
        'total_staff': int(len(labor_sz)),
        'active_staff': int(len(active_sz)),
        'inactive_staff': int(len(inactive_sz)),
        'overall_sat': float(
            active_sz['实际投入工时'].sum() / labor_sz['应投入工时'].sum()
        ) if labor_sz['应投入工时'].sum() > 0 else 0
    }

    # ══════════════════════════════════════════════
    # 部门工时投入对比（仅数智技术服务中心）
    # ══════════════════════════════════════════════
    dept_hours_sz = labor_sz.groupby('部门').agg(
        count=('姓名', 'count'),
        actual=('实际投入工时', 'sum'),
        expected=('应投入工时', 'sum')
    ).reset_index()
    dept_hours_sz['avg_sat'] = dept_hours_sz.apply(
        lambda r: r['actual'] / r['expected'] if r['expected'] > 0 else 0, axis=1
    )
    dept_hours_sz = dept_hours_sz.sort_values('actual', ascending=False)
    data['dept_hours'] = dept_hours_sz.to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 部门人均工作饱和度（仅数智技术服务中心）
    # ══════════════════════════════════════════════
    dept_sat = labor_sz.groupby('部门').agg(
        count=('姓名', 'count'),
        actual=('实际投入工时', 'sum'),
        expected=('应投入工时', 'sum')
    ).reset_index()
    dept_sat['avg_saturation'] = dept_sat.apply(
        lambda r: r['actual'] / r['expected'] if r['expected'] > 0 else 0, axis=1
    )
    dept_sat = dept_sat.sort_values('avg_saturation', ascending=False)
    data['dept_avg_saturation'] = dept_sat[['部门', 'count', 'avg_saturation']].to_dict(orient='records')

    # 按部门人均饱和度的部门顺序重新排列 dept_hours，保持两个模块一致
    dept_sat_order = dept_sat['部门'].tolist()
    dept_hours_sorted = sorted(data['dept_hours'], key=lambda d: dept_sat_order.index(d['部门']) if d['部门'] in dept_sat_order else 999)
    data['dept_hours'] = dept_hours_sorted

    # ══════════════════════════════════════════════
    # 低饱和度人员（仅数智技术服务中心）
    # ══════════════════════════════════════════════
    low_sat_sz = labor_sz[
        (labor_sz['工作饱和度'] < 0.5) & (labor_sz['实际投入工时'] > 0)
    ].sort_values('工作饱和度')
    data['low_sat_staff'] = low_sat_sz[
        ['姓名', '部门', '实际投入工时', '应投入工时', '工作饱和度', '岗位']
    ].to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 本期项目变化
    # ══════════════════════════════════════════════
    new_projects = change_df[change_df['类型'] == '新增项目'][
        ['操作日期', '项目名称', '项目经理', '计划开始日期', '计划完成日期']
    ].reset_index(drop=True)
    delayed_projects = change_df[change_df['类型'] == '延期项目'][
        ['操作日期', '项目名称', '项目经理', '计划开始日期', '计划完成日期']
    ].reset_index(drop=True)
    data['new_projects'] = new_projects.to_dict(orient='records')
    data['delayed_projects'] = delayed_projects.to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 项目超期/到期预警
    # ══════════════════════════════════════════════
    overdue = month_proj[
        month_proj['禅道是否超期/到期'].astype(str).str.contains('超期')
    ][['项目名称', '所属部门', '禅道是否超期/到期', '中心人力投入（人月）']].reset_index(drop=True)

    expiring = month_proj[
        month_proj['禅道是否超期/到期'].astype(str).str.contains('还有')
    ].copy()
    expiring['days_left'] = expiring['禅道是否超期/到期'].str.extract(r'还有(\d+)天').astype(float)
    expiring = expiring[expiring['days_left'] <= EXPIRING_DAYS].sort_values('days_left')[
        ['项目名称', '所属部门', '禅道是否超期/到期', '中心人力投入（人月）']
    ].reset_index(drop=True)

    data['overdue_projects'] = overdue.to_dict(orient='records')
    data['expiring_projects'] = expiring.to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 本期人力投入项目明细
    # ══════════════════════════════════════════════
    month_proj_sorted = month_proj.sort_values('本期总投入（人月）', ascending=False)
    data['month_project_detail'] = month_proj_sorted[[
        '项目名称', '所属部门', '中心人力投入（人月）', '智能部', '政务部', '大数据部',
        '产品创新部', '其他部门投入', '本期总投入（人月）', '禅道是否超期/到期'
    ]].to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 本期人力投入TOP项目（堆叠柱状图）
    # ══════════════════════════════════════════════
    top_month = month_proj_sorted.head(TOP_N).copy()

    def get_center(row):
        if row['智能部'] > 0: return '智能部'
        if row['政务部'] > 0: return '政务部'
        if row['大数据部'] > 0: return '大数据部'
        if row['产品创新部'] > 0: return '产品创新部'
        return '其他'

    top_month['主投入部门'] = top_month.apply(get_center, axis=1)
    data['top_month_chart'] = top_month[[
        '项目名称', '智能部', '政务部', '大数据部', '产品创新部',
        '其他部门投入', '本期总投入（人月）', '主投入部门'
    ]].to_dict(orient='records')

    # ══════════════════════════════════════════════
    # 各部门人力投入TOP3项目
    # ══════════════════════════════════════════════
    sz_dept_cols = {'智能部': '智能部', '政务部': '政务部', '大数据部': '大数据部', '产品创新部': '产品创新部'}
    dept_top3 = {}
    for dept_key, col in sz_dept_cols.items():
        sub = month_proj_sorted[month_proj_sorted[col] > 0].nlargest(TOP3_N, col)
        if len(sub) > 0:
            dept_top3[dept_key] = sub[[
                '项目名称', '所属部门', col, '本期总投入（人月）'
            ]].rename(columns={col: '部门投入（人月）'}).to_dict(orient='records')
        else:
            dept_top3[dept_key] = []
    data['dept_top3_projects'] = dept_top3

    return data


# ══════════════════════════════════════════════════════════════
# HTML 模板
# ══════════════════════════════════════════════════════════════
HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>项目数据可视化大屏</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#0a0e27;color:#e0e6f0;font-family:'Microsoft YaHei','PingFang SC',sans-serif;overflow-x:hidden}
.dashboard{width:100%;min-height:100vh;padding:14px 18px}
.header{text-align:center;padding:8px 0 14px;position:relative}
.header h1{font-size:26px;background:linear-gradient(90deg,#00d4ff,#7b61ff,#00d4ff);-webkit-background-clip:text;-webkit-text-fill-color:transparent;letter-spacing:6px;font-weight:700}
.header .sub{color:#5a6a8a;font-size:12px;margin-top:3px}
.kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:12px 18px;margin-bottom:14px}
.kpi-card{background:linear-gradient(135deg,rgba(20,30,60,0.9),rgba(10,14,39,0.95));border:1px solid rgba(0,212,255,0.15);border-radius:10px;padding:14px 18px;position:relative;overflow:hidden;min-width:0}
.kpi-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,var(--accent),transparent)}
.kpi-card .label{font-size:12px;color:#7a8ba8;margin-bottom:5px}
.kpi-card .value{font-size:22px;font-weight:700;color:#fff;white-space:nowrap}
.kpi-card .sub-info{font-size:11px;color:#7a8ba8;margin-top:3px}
.kpi-card:nth-child(1){--accent:#00d4ff}
.kpi-card:nth-child(2){--accent:#7b61ff}
.kpi-card:nth-child(3){--accent:#00e396}
.kpi-card:nth-child(4){--accent:#4ecdc4}
.grid{display:grid;gap:12px;margin-bottom:12px}
.grid-2{grid-template-columns:1fr 1fr}
.grid-3{grid-template-columns:1fr 1fr 1fr}
.panel{background:linear-gradient(135deg,rgba(20,30,60,0.85),rgba(10,14,39,0.9));border:1px solid rgba(0,212,255,0.1);border-radius:10px;padding:14px;position:relative}
.panel-title{font-size:13px;font-weight:600;color:#00d4ff;margin-bottom:10px;padding-left:10px;border-left:3px solid #00d4ff;display:flex;align-items:center;gap:6px}
.panel-title .dot{width:6px;height:6px;border-radius:50%;background:#00d4ff;animation:pulse 2s infinite}
.panel-title .badge{font-size:10px;padding:1px 6px;border-radius:8px;margin-left:4px}
.badge-red{background:rgba(255,107,107,0.2);color:#ff6b6b}
.badge-green{background:rgba(0,227,150,0.2);color:#00e396}
.badge-orange{background:rgba(255,179,71,0.2);color:#ffb347}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}
.chart-box{width:100%;height:260px}
.chart-box-sm{width:100%;height:220px}
.scroll-table{max-height:280px;overflow-y:auto;position:relative}
.scroll-table table{border-collapse:collapse;font-size:11px;width:100%}
.scroll-table thead{position:sticky;top:0;z-index:2}
.scroll-table thead th{background:rgba(10,14,39,0.98);color:#7ab8e0;padding:6px 5px;text-align:left;font-weight:600;white-space:nowrap;border-bottom:1px solid rgba(0,212,255,0.15)}
.scroll-table tbody td{padding:5px;border-bottom:1px solid rgba(255,255,255,0.04);color:#c0c8d8;white-space:nowrap}
.scroll-table tbody tr:hover{background:rgba(0,212,255,0.05)}
.tag{display:inline-block;padding:1px 5px;border-radius:3px;font-size:10px}
.tag-green{background:rgba(0,227,150,0.15);color:#00e396}
.tag-blue{background:rgba(0,212,255,0.15);color:#00d4ff}
.tag-orange{background:rgba(255,179,71,0.15);color:#ffb347}
.tag-red{background:rgba(255,107,107,0.15);color:#ff6b6b}
.tag-purple{background:rgba(123,97,255,0.15);color:#7b61ff}
.scroll-table::-webkit-scrollbar{width:4px}
.scroll-table::-webkit-scrollbar-thumb{background:rgba(0,212,255,0.2);border-radius:2px}
.summary-row{display:flex;gap:16px;justify-content:center;align-items:center;margin-bottom:10px}
.summary-stat{text-align:center}
.summary-stat .num{font-size:28px;font-weight:700}
.summary-stat .lbl{font-size:11px;color:#7a8ba8;margin-top:2px}
.summary-divider{width:1px;height:40px;background:rgba(255,255,255,0.1)}
.section-label{font-size:12px;color:#5a6a8a;margin:8px 0 4px;padding-left:6px;border-left:2px solid rgba(0,212,255,0.3)}
</style>
</head>
<body>
<div class="dashboard">
  <div class="header">
    <h1>项 目 数 据 可 视 化 大 屏</h1>
    <div class="sub">数据统计周期：__PERIOD_PLACEHOLDER__ ｜ 项目总数：<span id="totalBadge">-</span> ｜ 数智技术服务中心在册人员：<span id="staffBadge">-</span></div>
  </div>

  <!-- KPI -->
  <div class="kpi-row">
    <div class="kpi-card"><div class="label">项目总数</div><div class="value" id="kpi1">-</div></div>
    <div class="kpi-card"><div class="label">合同金额合计</div><div class="value" id="kpi2">-</div></div>
    <div class="kpi-card"><div class="label">已回款金额</div><div class="value" id="kpi3">-</div><div class="sub-info" id="kpi3_sub"></div></div>
    <div class="kpi-card"><div class="label">预估实施成本</div><div class="value" id="kpi4">-</div><div class="sub-info" id="kpi4_sub"></div></div>
  </div>

  <!-- Row 1: 本月人力投入 + 合同类型 + 项目类型 -->
  <div class="grid grid-3">
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>本月人力投入情况</div>
      <div id="labor_summary_area"></div>
      <div class="chart-box-sm" id="chart_month_labor"></div>
    </div>
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>合同类型分布</div>
      <div class="chart-box" id="chart_contract_type"></div>
    </div>
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>项目类型分布</div>
      <div class="chart-box" id="chart_ptype"></div>
    </div>
  </div>

  <!-- Row 2: 回款率分布 + TOP10堆叠图 -->
  <div class="grid grid-2">
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>回款率分布</div>
      <div class="chart-box" id="chart_pay_dist"></div>
    </div>
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>本期人力投入TOP10项目（按部门着色）</div>
      <div class="chart-box" id="chart_top_month_stacked"></div>
    </div>
  </div>

  <!-- Row 3: 人员工作饱和度 + 部门工时对比 + 部门人均饱和度 -->
  <div class="grid grid-3">
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>人员工作饱和度分析<span class="badge badge-green">数智技术服务中心</span></div>
      <div id="staff_summary_area"></div>
      <div class="chart-box-sm" id="chart_saturation"></div>
    </div>
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>部门工时投入对比<span class="badge badge-green">数智技术服务中心</span></div>
      <div class="chart-box-sm" id="chart_dept_hours"></div>
    </div>
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>部门人均工作饱和度<span class="badge badge-green">数智技术服务中心</span></div>
      <div class="chart-box-sm" id="chart_dept_avg_sat"></div>
    </div>
  </div>

  <!-- Row 3.6: 各部门人力投入TOP3 -->
  <div class="grid" style="grid-template-columns:1fr">
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>各部门人力投入TOP3项目<span class="badge badge-green">数智技术服务中心</span></div>
      <div id="dept_top3_area"></div>
    </div>
  </div>

  <!-- Row 4: 本期项目变化动态 + 项目超期预警 -->
  <div class="grid grid-2">
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>本期项目变化动态</div>
      <div id="change_summary_area"></div>
      <div class="scroll-table" id="table_changes"></div>
    </div>
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>项目超期/到期预警</div>
      <div id="alert_summary_area"></div>
      <div class="scroll-table" id="table_alerts"></div>
    </div>
  </div>

  <!-- Row 5: TOP10 + 人力累计TOP10 -->
  <div class="grid grid-2">
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>合同金额TOP10项目</div>
      <div class="scroll-table" id="table_top10"></div>
    </div>
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>人力投入累计TOP10项目</div>
      <div class="scroll-table" id="table_top_labor"></div>
    </div>
  </div>

  <!-- Row 6: 本期人力投入项目明细 -->
  <div class="grid" style="grid-template-columns:1fr">
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>本期人力投入项目明细（按部门拆分）</div>
      <div class="scroll-table" id="table_month_detail"></div>
    </div>
  </div>

  <!-- Row 7: 挂起项目 + 低饱和度人员 -->
  <div class="grid grid-2">
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>挂起项目</div>
      <div class="scroll-table" id="table_suspended"></div>
    </div>
    <div class="panel">
      <div class="panel-title"><span class="dot"></span>低饱和度人员（<50%）<span class="badge badge-green">数智技术服务中心</span></div>
      <div class="scroll-table" id="table_low_sat"></div>
    </div>
  </div>
</div>

<script>
const RAW = __DATA_PLACEHOLDER__;
const fmt = n => { if(n>=1e8) return (n/1e8).toFixed(2)+'亿'; if(n>=1e4) return (n/1e4).toFixed(1)+'万'; return n.toFixed(0); };
const pct = n => (n*100).toFixed(1)+'%';

// ── KPI ──
const k = RAW.kpi;
document.getElementById('kpi1').innerHTML = k.total_projects+'<span class="unit" style="font-size:12px;color:#7a8ba8;font-weight:400;margin-left:3px">个</span>';
document.getElementById('kpi2').innerHTML = '¥'+fmt(k.total_contract);
document.getElementById('kpi3').innerHTML = '¥'+fmt(k.total_paid);
document.getElementById('kpi3_sub').innerHTML = '<span style="color:#00e396">回款率 '+pct(k.pay_rate)+'</span>';
document.getElementById('kpi4').innerHTML = '¥'+fmt(k.total_cost);
document.getElementById('kpi4_sub').innerHTML = '<span style="color:#ffb347">合同占比 '+pct(k.cost_ratio)+'</span>';
document.getElementById('staffBadge').textContent = RAW.staff_summary.total_staff;
document.getElementById('totalBadge').textContent = k.total_projects;

// ── 本月人力投入 ──
const ml = RAW.month_labor;
document.getElementById('labor_summary_area').innerHTML = `
  <div class="summary-row">
    <div class="summary-stat"><div class="num" style="color:#00e396">${ml.has_labor_count}</div><div class="lbl">有人力投入</div></div>
    <div class="summary-divider"></div>
    <div class="summary-stat"><div class="num" style="color:#ffb347">${ml.no_labor_count}</div><div class="lbl">无人力投入</div></div>
    <div class="summary-divider"></div>
    <div class="summary-stat"><div class="num" style="color:#00d4ff">${ml.has_labor_amount.toFixed(2)}</div><div class="lbl">本月投入合计（人月）</div></div>
  </div>`;

// ── 人员饱和度（不显示实际投入工时） ──
const ss = RAW.staff_summary;
document.getElementById('staff_summary_area').innerHTML = `
  <div class="summary-row">
    <div class="summary-stat"><div class="num" style="color:#00d4ff">${ss.active_staff}</div><div class="lbl">在岗人员</div></div>
    <div class="summary-divider"></div>
    <div class="summary-stat"><div class="num" style="color:#ffb347">${ss.inactive_staff}</div><div class="lbl">未投入人员</div></div>
    <div class="summary-divider"></div>
    <div class="summary-stat"><div class="num" style="color:#00e396">${pct(ss.overall_sat)}</div><div class="lbl">整体饱和度</div></div>
  </div>`;

// ── 项目变化 ──
const np = RAW.new_projects.length, dp = RAW.delayed_projects.length;
document.getElementById('change_summary_area').innerHTML = `
  <div class="summary-row">
    <div class="summary-stat"><div class="num" style="color:#00e396">${np}</div><div class="lbl">新增项目</div></div>
    <div class="summary-divider"></div>
    <div class="summary-stat"><div class="num" style="color:#ff6b6b">${dp}</div><div class="lbl">延期项目</div></div>
  </div>`;

// ── 超期预警 ──
const od = RAW.overdue_projects.length, ep = RAW.expiring_projects.length;
document.getElementById('alert_summary_area').innerHTML = `
  <div class="summary-row">
    <div class="summary-stat"><div class="num" style="color:#ff6b6b">${od}</div><div class="lbl">已超期项目</div></div>
    <div class="summary-divider"></div>
    <div class="summary-stat"><div class="num" style="color:#ffb347">${ep}</div><div class="lbl">60天内到期</div></div>
  </div>`;

const COLORS = ['#00d4ff','#7b61ff','#00e396','#ffb347','#ff6b6b','#4ecdc4','#ff9ff3','#54a0ff','#5f27cd','#01a3a4','#f368e0','#ee5a24','#009432','#0652DD'];

function initPie(id, data, rose=false) {
  const ch = echarts.init(document.getElementById(id));
  ch.setOption({
    tooltip:{trigger:'item',formatter:'{b}: {c} ({d}%)',backgroundColor:'rgba(10,14,39,0.9)',borderColor:'rgba(0,212,255,0.3)',textStyle:{color:'#e0e6f0'}},
    color: COLORS,
    legend:{type:'scroll',bottom:0,textStyle:{color:'#a0aec0',fontSize:10},pageTextStyle:{color:'#a0aec0'}},
    series:[{type:'pie',radius:rose?['25%','62%']:['35%','62%'],center:['50%','46%'],roseType:rose?'area':false,
      label:{color:'#a0aec0',fontSize:10,formatter:'{b}\n{d}%'},
      labelLine:{lineStyle:{color:'rgba(255,255,255,0.2)'}},
      itemStyle:{borderColor:'#0a0e27',borderWidth:2,borderRadius:6},
      emphasis:{itemStyle:{shadowBlur:15,shadowColor:'rgba(0,212,255,0.4)'},label:{fontSize:12,fontWeight:'bold'}},
      data:data.map(d=>({name:d.name,value:d.value}))}]
  });
  window.addEventListener('resize',()=>ch.resize());
}

function initBarH(id, names, values) {
  const ch = echarts.init(document.getElementById(id));
  ch.setOption({
    tooltip:{trigger:'axis',axisPointer:{type:'shadow'},backgroundColor:'rgba(10,14,39,0.9)',borderColor:'rgba(0,212,255,0.3)',textStyle:{color:'#e0e6f0'}},
    grid:{left:100,right:30,top:8,bottom:16},
    xAxis:{type:'value',axisLabel:{color:'#7a8ba8',fontSize:9},splitLine:{lineStyle:{color:'rgba(255,255,255,0.04)'}},axisLine:{show:false}},
    yAxis:{type:'category',data:names,axisLabel:{color:'#a0aec0',fontSize:9},axisLine:{lineStyle:{color:'rgba(255,255,255,0.08)'}},axisTick:{show:false}},
    series:[{type:'bar',data:values,barWidth:12,
      itemStyle:{borderRadius:[0,4,4,0],color:new echarts.graphic.LinearGradient(0,0,1,0,[{offset:0,color:'#7b61ff'},{offset:1,color:'#00d4ff'}])},
      label:{show:true,position:'right',color:'#a0aec0',fontSize:9,formatter:p=>p.value},
      emphasis:{itemStyle:{color:new echarts.graphic.LinearGradient(0,0,1,0,[{offset:0,color:'#9b81ff'},{offset:1,color:'#40e0ff'}])}}}]
  });
  window.addEventListener('resize',()=>ch.resize());
}

function initBarV(id, names, values) {
  const ch = echarts.init(document.getElementById(id));
  ch.setOption({
    tooltip:{trigger:'axis',axisPointer:{type:'shadow'},backgroundColor:'rgba(10,14,39,0.9)',borderColor:'rgba(0,212,255,0.3)',textStyle:{color:'#e0e6f0'}},
    grid:{left:45,right:16,top:16,bottom:36},
    xAxis:{type:'category',data:names,axisLabel:{color:'#a0aec0',fontSize:9,rotate:names.length>6?20:0},axisLine:{lineStyle:{color:'rgba(255,255,255,0.08)'}},axisTick:{show:false}},
    yAxis:{type:'value',axisLabel:{color:'#7a8ba8',fontSize:9},splitLine:{lineStyle:{color:'rgba(255,255,255,0.04)'}},axisLine:{show:false}},
    series:[{type:'bar',data:values,barWidth:names.length>8?14:20,
      itemStyle:{borderRadius:[4,4,0,0],color:new echarts.graphic.LinearGradient(0,1,0,0,[{offset:0,color:'#7b61ff'},{offset:1,color:'#00d4ff'}])},
      label:{show:true,position:'top',color:'#a0aec0',fontSize:9}}]
  });
  window.addEventListener('resize',()=>ch.resize());
}

function initBarHGroup(id, names, series1, series2, name1, name2) {
  const ch = echarts.init(document.getElementById(id));
  ch.setOption({
    tooltip:{trigger:'axis',axisPointer:{type:'shadow'},backgroundColor:'rgba(10,14,39,0.9)',borderColor:'rgba(0,212,255,0.3)',textStyle:{color:'#e0e6f0'}},
    legend:{data:[name1,name2],textStyle:{color:'#a0aec0',fontSize:10},top:0},
    grid:{left:100,right:20,top:30,bottom:16},
    xAxis:{type:'value',axisLabel:{color:'#7a8ba8',fontSize:9},splitLine:{lineStyle:{color:'rgba(255,255,255,0.04)'}},axisLine:{show:false}},
    yAxis:{type:'category',data:names,axisLabel:{color:'#a0aec0',fontSize:9},axisLine:{lineStyle:{color:'rgba(255,255,255,0.08)'}},axisTick:{show:false}},
    series:[
      {name:name1,type:'bar',data:series1,barWidth:8,itemStyle:{borderRadius:[0,3,3,0],color:'rgba(0,212,255,0.7)'}},
      {name:name2,type:'bar',data:series2,barWidth:8,itemStyle:{borderRadius:[0,3,3,0],color:'rgba(123,97,255,0.7)'}}
    ]
  });
  window.addEventListener('resize',()=>ch.resize());
}

// ── Charts ──
initPie('chart_contract_type', RAW.contract_type.map(d=>({name:d['合同签订状态'],value:d.count})), true);
initPie('chart_ptype', RAW.project_type.map(d=>({name:d['项目类型'],value:d.count})), true);
initPie('chart_pay_dist', RAW.pay_distribution);
initPie('chart_month_labor', [{name:'有人力投入',value:ml.has_labor_count},{name:'无人力投入',value:ml.no_labor_count}]);

// V2: 饱和度分布
initPie('chart_saturation', RAW.saturation_dist, true);

// V2: 部门工时对比（分组柱状图）
const dh = RAW.dept_hours.filter(d=>d.actual>0);
initBarHGroup('chart_dept_hours', dh.map(d=>d['部门']), dh.map(d=>Math.round(d.actual)), dh.map(d=>Math.round(d.expected)), '实际工时', '应投入工时');

// V2.1: 部门人均工作饱和度
const das = RAW.dept_avg_saturation;
const das_ch = echarts.init(document.getElementById('chart_dept_avg_sat'));
das_ch.setOption({
  tooltip:{trigger:'axis',axisPointer:{type:'shadow'},backgroundColor:'rgba(10,14,39,0.9)',borderColor:'rgba(0,212,255,0.3)',textStyle:{color:'#e0e6f0'},
    formatter:p=>`${p[0].name}<br/>人均饱和度: <b>${(p[0].value*100).toFixed(1)}%</b><br/>人数: ${das[p[0].dataIndex].count}人`},
  grid:{left:90,right:36,top:12,bottom:12},
  xAxis:{type:'value',max:1.2,axisLabel:{color:'#7a8ba8',fontSize:8,formatter:v=>(v*100)+'%'},splitLine:{lineStyle:{color:'rgba(255,255,255,0.04)'}},axisLine:{show:false}},
  yAxis:{type:'category',data:das.map(d=>d['部门']).reverse(),axisLabel:{color:'#a0aec0',fontSize:9},axisLine:{lineStyle:{color:'rgba(255,255,255,0.08)'}},axisTick:{show:false}},
  series:[{type:'bar',data:das.map(d=>parseFloat(d.avg_saturation.toFixed(3))).reverse(),barWidth:12,
    itemStyle:{borderRadius:[0,6,6,0],
      color:p=>{const v=p.value;if(v>=0.9)return new echarts.graphic.LinearGradient(0,0,1,0,[{offset:0,color:'#00e396'},{offset:1,color:'#00b894'}]);if(v>=0.7)return new echarts.graphic.LinearGradient(0,0,1,0,[{offset:0,color:'#00d4ff'},{offset:1,color:'#0984e3'}]);if(v>=0.5)return new echarts.graphic.LinearGradient(0,0,1,0,[{offset:0,color:'#ffb347'},{offset:1,color:'#fdcb6e'}]);return new echarts.graphic.LinearGradient(0,0,1,0,[{offset:0,color:'#ff6b6b'},{offset:1,color:'#ee5a24'}]);}},
    label:{show:true,position:'right',color:'#a0aec0',fontSize:10,formatter:p=>(p.value*100).toFixed(1)+'%'},
    markLine:{data:[]}
  }]
});
window.addEventListener('resize',()=>das_ch.resize());

// V2.1: 本期人力投入TOP10（堆叠柱状图，按部门着色）
const tmc = RAW.top_month_chart;
const dept_colors = {'智能部':'#00d4ff','政务部':'#7b61ff','大数据部':'#00e396','产品创新部':'#ffb347','其他':'#4ecdc4'};
const stack_series = ['智能部','政务部','大数据部','产品创新部','其他部门投入'].map((col,i)=>{
  const color = dept_colors[['智能部','政务部','大数据部','产品创新部','其他'][i]];
  return {name:col,type:'bar',stack:'total',data:tmc.map(d=>parseFloat(d[col].toFixed(2))),barWidth:18,
    itemStyle:{color:color,borderRadius:i===0?[4,4,0,0]:[0,0,0,0]},
    emphasis:{focus:'series'}};
});
const tmc_ch = echarts.init(document.getElementById('chart_top_month_stacked'));
tmc_ch.setOption({
  tooltip:{trigger:'axis',axisPointer:{type:'shadow'},backgroundColor:'rgba(10,14,39,0.9)',borderColor:'rgba(0,212,255,0.3)',textStyle:{color:'#e0e6f0'}},
  legend:{data:['智能部','政务部','大数据部','产品创新部','其他'],textStyle:{color:'#a0aec0',fontSize:10},top:0},
  grid:{left:140,right:20,top:36,bottom:16},
  xAxis:{type:'value',axisLabel:{color:'#7a8ba8',fontSize:9},splitLine:{lineStyle:{color:'rgba(255,255,255,0.04)'}},axisLine:{show:false}},
  yAxis:{type:'category',data:tmc.map(d=>d['项目名称'].length>16?d['项目名称'].slice(0,16)+'…':d['项目名称']),axisLabel:{color:'#a0aec0',fontSize:9},axisLine:{lineStyle:{color:'rgba(255,255,255,0.08)'}},axisTick:{show:false}},
  series:stack_series
});
window.addEventListener('resize',()=>tmc_ch.resize());

// V2.1: 各部门人力投入TOP3项目
const dt3 = RAW.dept_top3_projects;
const dept_color_map = {'智能部':'#00d4ff','政务部':'#7b61ff','大数据部':'#00e396','产品创新部':'#ffb347'};
let html_dt3 = '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px">';
for (const [dept, projects] of Object.entries(dt3)) {
  const clr = dept_color_map[dept] || '#a0aec0';
  html_dt3 += `<div style="background:rgba(255,255,255,0.02);border:1px solid rgba(${clr==='#00d4ff'?'0,212,255':clr==='#7b61ff'?'123,97,255':clr==='#00e396'?'0,227,150':'255,179,71'},0.15);border-radius:8px;padding:10px">
    <div style="font-size:12px;font-weight:600;color:${clr};margin-bottom:8px;border-left:3px solid ${clr};padding-left:6px">${dept}</div>
    <table style="width:100%;border-collapse:collapse;font-size:10px">
      <thead><tr><th style="color:#7ab8e0;padding:4px;text-align:left;border-bottom:1px solid rgba(255,255,255,0.06)">项目名称</th><th style="color:#7ab8e0;padding:4px;text-align:right;border-bottom:1px solid rgba(255,255,255,0.06)">投入(人月)</th></tr></thead><tbody>`;
  projects.forEach(d=>{
    html_dt3 += `<tr><td style="padding:4px;color:#c0c8d8;max-width:120px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td style="padding:4px;color:${clr};text-align:right;font-weight:600">${d['部门投入（人月）'].toFixed(2)}</td></tr>`;
  });
  html_dt3 += '</tbody></table></div>';
}
html_dt3 += '</div>';
document.getElementById('dept_top3_area').innerHTML = html_dt3;

// ── Tables ──
// TOP10
let h1='<table><thead><tr><th>排名</th><th>项目名称</th><th>合同金额</th><th>已回款</th><th>回款率</th><th>阶段</th></tr></thead><tbody>';
RAW.top10.forEach((d,i)=>{const r=d['合同回款情况'];const c=r>=0.9?'tag-green':r>=0.5?'tag-blue':r>0?'tag-orange':'tag-red';h1+=`<tr><td>${i+1}</td><td style="max-width:150px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td>¥${fmt(d['合同金额（元）'])}</td><td>¥${fmt(d['已回款金额（元）'])}</td><td><span class="tag ${c}">${pct(r)}</span></td><td>${d['项目阶段']}</td></tr>`;});
h1+='</tbody></table>';document.getElementById('table_top10').innerHTML=h1;

// 人力累计TOP10
let h2='<table><thead><tr><th>排名</th><th>项目名称</th><th>投入人月</th><th>所属部门</th><th>项目经理</th></tr></thead><tbody>';
RAW.top_labor.forEach((d,i)=>{h2+=`<tr><td>${i+1}</td><td style="max-width:150px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td>${d['累计投入工作量（人月）'].toFixed(1)}</td><td>${d['所属部门']}</td><td>${d['项目经理']}</td></tr>`;});
h2+='</tbody></table>';document.getElementById('table_top_labor').innerHTML=h2;

// V2: 本期项目变化
let hc='<div class="section-label">新增项目</div><table><thead><tr><th>日期</th><th>项目名称</th><th>项目经理</th><th>计划周期</th></tr></thead><tbody>';
RAW.new_projects.forEach(d=>{hc+=`<tr><td>${d['操作日期'].slice(0,10)}</td><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td>${d['项目经理']}</td><td>${d['计划开始日期']} ~ ${d['计划完成日期']}</td></tr>`;});
hc+='</tbody></table>';
hc+='<div class="section-label">延期项目</div><table><thead><tr><th>日期</th><th>项目名称</th><th>项目经理</th><th>计划周期</th></tr></thead><tbody>';
RAW.delayed_projects.forEach(d=>{hc+=`<tr><td>${d['操作日期'].slice(0,10)}</td><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td>${d['项目经理']}</td><td>${d['计划开始日期']} ~ ${d['计划完成日期']}</td></tr>`;});
hc+='</tbody></table>';document.getElementById('table_changes').innerHTML=hc;

// V2: 超期预警
let ha='<div class="section-label" style="border-left-color:#ff6b6b">已超期项目</div><table><thead><tr><th>项目名称</th><th>所属部门</th><th>超期状态</th><th>本月投入（人月）</th></tr></thead><tbody>';
RAW.overdue_projects.forEach(d=>{ha+=`<tr><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td>${d['所属部门']}</td><td><span class="tag tag-red">${d['禅道是否超期/到期']}</span></td><td>${d['中心人力投入（人月）'].toFixed(2)}</td></tr>`;});
ha+='</tbody></table>';
ha+='<div class="section-label" style="border-left-color:#ffb347">60天内到期项目</div><table><thead><tr><th>项目名称</th><th>所属部门</th><th>到期状态</th><th>本月投入（人月）</th></tr></thead><tbody>';
RAW.expiring_projects.forEach(d=>{ha+=`<tr><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td>${d['所属部门']}</td><td><span class="tag tag-orange">${d['禅道是否超期/到期']}</span></td><td>${d['中心人力投入（人月）'].toFixed(2)}</td></tr>`;});
ha+='</tbody></table>';document.getElementById('table_alerts').innerHTML=ha;

// V2: 本期人力投入项目明细（非零数据高亮）
let hd='<table><thead><tr><th>项目名称</th><th>所属部门</th><th>中心投入</th><th>智能部</th><th>政务部</th><th>大数据部</th><th>产品创新部</th><th>其他</th><th>总投入</th><th>状态</th></tr></thead><tbody>';
RAW.month_project_detail.forEach(d=>{
  const st = d['禅道是否超期/到期'];
  const tagCls = st && st.includes('超期') ? 'tag-red' : st && st.includes('还有') ? 'tag-orange' : 'tag-blue';
  const hl = v => v > 0 ? `style="color:#00d4ff;font-weight:600"` : 'style="color:#3a4a6a"';
  hd+=`<tr><td style="max-width:140px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td>${d['所属部门']}</td><td ${hl(d['中心人力投入（人月）'])}>${d['中心人力投入（人月）'].toFixed(2)}</td><td ${hl(d['智能部'])}>${d['智能部'].toFixed(2)}</td><td ${hl(d['政务部'])}>${d['政务部'].toFixed(2)}</td><td ${hl(d['大数据部'])}>${d['大数据部'].toFixed(2)}</td><td ${hl(d['产品创新部'])}>${d['产品创新部'].toFixed(2)}</td><td ${hl(d['其他部门投入'])}>${d['其他部门投入'].toFixed(2)}</td><td style="font-weight:700;color:#e0e6f0">${d['本期总投入（人月）'].toFixed(2)}</td><td><span class="tag ${tagCls}">${st||''}</span></td></tr>`;
});
hd+='</tbody></table>';document.getElementById('table_month_detail').innerHTML=hd;

// 挂起项目
let h3='<table><thead><tr><th>项目名称</th><th>合同金额</th><th>所属部门</th><th>项目经理</th><th>进度</th></tr></thead><tbody>';
RAW.suspended.forEach(d=>{const p=d['项目进度'];h3+=`<tr><td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${d['项目名称']}</td><td>¥${fmt(d['合同金额（元）'])}</td><td>${d['所属部门']}</td><td>${d['项目经理']}</td><td>${pct(p)}</td></tr>`;});
if(!RAW.suspended.length) h3+='<tr><td colspan="5" style="text-align:center;color:#5a6a8a">暂无挂起项目</td></tr>';
h3+='</tbody></table>';document.getElementById('table_suspended').innerHTML=h3;

// V2: 低饱和度人员
let hs='<table><thead><tr><th>姓名</th><th>部门</th><th>实际工时</th><th>应投入工时</th><th>饱和度</th><th>岗位</th></tr></thead><tbody>';
RAW.low_sat_staff.forEach(d=>{const s=d['工作饱和度'];const tc=s<0.3?'tag-red':'tag-orange';hs+=`<tr><td>${d['姓名']}</td><td>${d['部门']}</td><td>${d['实际投入工时']}</td><td>${d['应投入工时']}</td><td><span class="tag ${tc}">${pct(s)}</span></td><td>${d['岗位']||''}</td></tr>`;});
if(!RAW.low_sat_staff.length) hs+='<tr><td colspan="6" style="text-align:center;color:#5a6a8a">无低饱和度人员</td></tr>';
hs+='</tbody></table>';document.getElementById('table_low_sat').innerHTML=hs;
</script>
</body>
</html>'''


# ══════════════════════════════════════════════════════════════
# 主流程
# ══════════════════════════════════════════════════════════════
def main():
    work_dir = os.path.dirname(os.path.abspath(__file__))

    print('=' * 60)
    print('  项目数据可视化看板 - 一键生成工具')
    print('=' * 60)
    print(f'工作目录: {work_dir}')
    print()

    # 1. 发现文件
    print('[1/4] 发现源数据文件 ...')
    files = find_files(work_dir)
    if files is None:
        print('\n[失败] 缺少必要的源数据文件，请检查文件是否齐全。')
        print('  需要: 项目数据表.xlsx, 人力投入表.xlsx,')
        print('        本期人力投入项目情况(日期).xlsx,')
        print('        本期项目变化情况(日期).xlsx')
        sys.exit(1)

    for key, path in files.items():
        print(f'  ✓ {os.path.basename(path)}')

    # 2. 提取统计周期
    print('\n[2/4] 提取统计周期 ...')
    period = extract_period(files)
    print(f'  统计周期: {period}')

    # 3. 数据处理
    print('\n[3/4] 处理数据 ...')
    data = load_and_process(files)
    print(f'  ✓ 项目总数: {data["kpi"]["total_projects"]}')
    print(f'  ✓ 在册人员: {data["staff_summary"]["total_staff"]}')
    print(f'  ✓ 数据维度: {len(data)} 个')

    # 4. 生成看板
    print('\n[4/4] 生成看板 ...')
    html = HTML_TEMPLATE.replace('__DATA_PLACEHOLDER__', json.dumps(data, ensure_ascii=False, default=str))
    html = html.replace('__PERIOD_PLACEHOLDER__', period)

    output_path = os.path.join(work_dir, OUTPUT_FILENAME)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'  ✓ 看板已生成: {output_path}')
    print(f'  ✓ 文件大小: {len(html):,} 字节')
    print()
    print('=' * 60)
    print('  生成完成！请用浏览器打开 看板.html 查看效果。')
    print('=' * 60)


if __name__ == '__main__':
    main()
