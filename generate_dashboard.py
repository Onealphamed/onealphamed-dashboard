#!/usr/bin/env python3
"""
OnealphaMed — Dashboard Generator
===================================
Reads  : Project Profit.xlsx  (the source data file)
Outputs: OnealphaMed_Dashboard.html  (the interactive dashboard)

Run locally  : python generate_dashboard.py
GitHub Action: triggered automatically when you push the Excel file
"""

import pandas as pd
import re
import json
import sys
import os
from pathlib import Path

# ─── PATHS ────────────────────────────────────────────────────────────────────
SCRIPT_DIR   = Path(__file__).parent
EXCEL_FILE   = SCRIPT_DIR / "Project Profit.xlsx"
OUTPUT_HTML  = SCRIPT_DIR / "OnealphaMed_Dashboard.html"

# ─── CONSTANTS ────────────────────────────────────────────────────────────────
MONTHS       = ["Aug'25","Sept'25","Oct'25","Nov'25","Dec'25","Jan'26","Feb'26","Mar'26"]
MONTH_ORDER  = {m: i+1 for i, m in enumerate(MONTHS)}
COMPANIES    = ['Hetero','Bayer','Lupin','P&G','Cipla','Aurobindo','NovoNordisk',
                'Zydus','KOITA','Amneal','Resmed','Sun Pharma']

CO_COLORS = {
    'Hetero':      '#C0392B',
    'Bayer':       '#1A6E37',
    'Lupin':       '#1F618D',
    'P&G':         '#D35400',
    'Cipla':       '#6C3483',
    'Aurobindo':   '#117A65',
    'NovoNordisk': '#7D3C98',
    'Zydus':       '#1565C0',
    'KOITA':       '#00796B',
    'Amneal':      '#E65100',
    'Resmed':      '#37474F',
    'Sun Pharma':  '#F57F17',
}
CO_LIGHT = {
    'Hetero':      '#FADBD8',
    'Bayer':       '#D5F5E3',
    'Lupin':       '#D6EAF8',
    'P&G':         '#FAE5D3',
    'Cipla':       '#E8DAEF',
    'Aurobindo':   '#D1F2EB',
    'NovoNordisk': '#F5EEF8',
    'Zydus':       '#DDEEFF',
    'KOITA':       '#E0F2F1',
    'Amneal':      '#FFF3E0',
    'Resmed':      '#ECEFF1',
    'Sun Pharma':  '#FFFDE7',
}

# ─── STEP 1: READ & CLEAN RAW DATA ────────────────────────────────────────────
def clean_num(val):
    if pd.isna(val) or val in ['-', 'nan', 'None', '']:
        return 0.0
    val = str(val).strip().replace('INR', '').replace(',', '').strip()
    val = re.sub(r'\+GST\..*', '', val)
    try:
        return float(val)
    except:
        return 0.0

def get_company(client, project):
    if pd.isna(client) or str(client).strip() == 'nan':
        return None
    c = str(client).strip()
    p = str(project).strip() if not pd.isna(project) else ''
    cl = c.lower()
    if p == 'TOTAL' or 'datetime' in cl:
        return None
    try:
        float(c)
        return None
    except:
        pass
    if 'hetero' in cl or (p.lower() == 'hetero' and 'translation' in cl) or c == 'French Translation':
        return 'Hetero'
    if c == 'Bayer':       return 'Bayer'
    if c == 'Lupin':       return 'Lupin'
    if c in ['P&G', 'P & G']:  return 'P&G'
    if c == 'Cipla':       return 'Cipla'
    if c == 'Aurobindo' or 'kiosq' in cl:  return 'Aurobindo'
    if 'novo' in cl or 'nordisk' in cl or c == 'Oscar':  return 'NovoNordisk'
    if c == 'Zydus':       return 'Zydus'
    if c == 'KOITA':       return 'KOITA'
    if c == 'Amneal':      return 'Amneal'
    if c == 'Resmed':      return 'Resmed'
    if c in ['Sun Pharma', 'Headon']:  return 'Sun Pharma'
    return None

def get_vendor_cat(vendor):
    """Returns None for Prof/Dr (doctor honorariums - excluded per business logic)."""
    if pd.isna(vendor) or str(vendor).strip() in ['nan', '-', '', 'None']:
        return 'Other'
    v = str(vendor).strip().lower()
    if v.startswith('prof ') or v.startswith('dr ') or v.startswith('dr.'):
        return None   # Doctor honorariums — excluded from vendor analysis
    for k in ['safiya', 'jss edit', 'medical writer', 'combird', 'karishma',
              'shivani', 'nandita', 'vandana', 'ashish', 'vedprakash', 'pervedu', 'mandar']:
        if k in v:
            return 'Medical Writing'
    for k in ['dynamic', 'hotel', 'av set', 'invite print', 'event',
              'radisson', 'taj vivanta', 'flight', 'jyotiba', 'cruzr', 'printer']:
        if k in v:
            return 'Events'
    for k in ['swarnim', 'elevan', 'neeraj', 'arvind', 'st team',
              'wa dissem', 'indonesia', 'internal']:
        if k in v:
            return 'Webinar/Tech'
    return 'Other'

def load_and_process(excel_path):
    print(f"  Reading: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name=0, dtype=str)
    df.columns = ['Project Name', 'Client', 'Month', 'Vendor', 'Tasked Assigned',
                  'Cost', 'Total Cost', 'PO Value', 'Invoice Amount',
                  'Project Profit', 'PO', 'Invoice Raised Date', 'Project Profitability']

    # Project-level rows (rows that carry an Invoice Amount)
    project_rows = []
    for _, row in df.iterrows():
        pname  = str(row['Project Name']).strip()
        client = str(row['Client']).strip()
        month  = str(row['Month']).strip()
        if pname in ['TOTAL', 'nan', 'None'] or month in ['nan', 'None']:
            continue
        inv = clean_num(row['Invoice Amount'])
        if inv == 0:
            continue
        company = get_company(client, pname)
        if company is None:
            continue
        cost   = clean_num(row['Total Cost'])
        profit = clean_num(row['Project Profit'])
        if profit == 0 and inv > 0:
            profit = inv - cost
        pct = profit / inv if inv > 0 else 0
        project_rows.append({
            'Project': pname, 'Client': client, 'Company': company,
            'Month': month, 'MonthOrder': MONTH_ORDER.get(month, 0),
            'Invoice': inv, 'Cost': cost, 'Profit': profit, 'ProfitPct': pct
        })

    # Vendor-level rows (individual cost lines, excluding Prof/Dr)
    vendor_rows = []
    cur_proj = cur_client = cur_month = cur_co = None
    for _, row in df.iterrows():
        pname  = str(row['Project Name']).strip()
        client = str(row['Client']).strip()
        month  = str(row['Month']).strip()
        vendor = str(row['Vendor']).strip()
        cost   = clean_num(row['Cost'])
        if pname not in ['nan', 'None', 'TOTAL'] and month not in ['nan', 'None']:
            comp = get_company(client, pname)
            if comp:
                cur_proj = pname
                cur_month = month
                cur_co = comp
        if cost > 0 and vendor not in ['nan', 'None', '-', ''] and cur_month:
            cat = get_vendor_cat(vendor)
            if cat:
                vendor_rows.append({
                    'Project': cur_proj, 'Company': cur_co,
                    'Month': cur_month, 'Vendor': vendor,
                    'Category': cat, 'Cost': cost
                })

    proj_df = pd.DataFrame(project_rows)
    vend_df = pd.DataFrame(vendor_rows)
    print(f"  Projects loaded: {len(proj_df)} | Vendor entries: {len(vend_df)}")
    return proj_df, vend_df

# ─── STEP 2: BUILD DATA OBJECT ────────────────────────────────────────────────
def build_data_object(proj_df, vend_df):
    data = {}
    data['months']    = MONTHS
    data['companies'] = COMPANIES

    data['monthly'] = {
        'invoice': [float(proj_df[proj_df['Month']==m]['Invoice'].sum()) for m in MONTHS],
        'cost':    [float(proj_df[proj_df['Month']==m]['Cost'].sum())    for m in MONTHS],
        'profit':  [float(proj_df[proj_df['Month']==m]['Profit'].sum())  for m in MONTHS],
    }

    data['company_totals'] = {}
    for co in COMPANIES:
        sub = proj_df[proj_df['Company'] == co]
        inv = float(sub['Invoice'].sum())
        data['company_totals'][co] = {
            'invoice': inv,
            'cost':    float(sub['Cost'].sum()),
            'profit':  float(sub['Profit'].sum()),
            'pct':     float(sub['Profit'].sum() / inv) if inv > 0 else 0,
        }

    data['company_monthly'] = {}
    for co in COMPANIES:
        inv_l=[]; cost_l=[]; prof_l=[]; pct_l=[]
        for m in MONTHS:
            sub = proj_df[(proj_df['Company']==co) & (proj_df['Month']==m)]
            i = float(sub['Invoice'].sum())
            c = float(sub['Cost'].sum())
            p = float(sub['Profit'].sum())
            inv_l.append(i); cost_l.append(c); prof_l.append(p)
            pct_l.append(p/i if i>0 else 0)
        data['company_monthly'][co] = {
            'invoice': inv_l, 'cost': cost_l, 'profit': prof_l, 'pct': pct_l
        }

    data['projects'] = {}
    for co in COMPANIES:
        sub = proj_df[proj_df['Company']==co][
            ['Project','Month','Invoice','Cost','Profit','ProfitPct']
        ].copy()
        data['projects'][co] = (
            sub.sort_values(['Month','Invoice'], ascending=[True, False])
            .to_dict('records')
        )

    CATS = ['Medical Writing', 'Webinar/Tech', 'Events']
    data['vendor'] = {
        'monthly': {
            cat: [float(vend_df[(vend_df['Category']==cat)&(vend_df['Month']==m)]['Cost'].sum())
                  for m in MONTHS]
            for cat in CATS
        },
        'totals': {
            cat: float(vend_df[vend_df['Category']==cat]['Cost'].sum())
            for cat in CATS
        },
        'by_company': {
            co: {
                cat: float(vend_df[(vend_df['Company']==co)&(vend_df['Category']==cat)]['Cost'].sum())
                for cat in CATS
            }
            for co in COMPANIES
        },
    }

    total_inv = float(proj_df['Invoice'].sum())
    data['grand'] = {
        'invoice': total_inv,
        'cost':    float(proj_df['Cost'].sum()),
        'profit':  float(proj_df['Profit'].sum()),
        'pct':     float(proj_df['Profit'].sum() / total_inv) if total_inv > 0 else 0,
    }
    return data

# ─── STEP 3: BUILD HTML ────────────────────────────────────────────────────────
def build_html(data):
    raw_data = json.dumps(data, separators=(',', ':'))

    # ---------- sidebar nav items ----------
    nav_items = ""
    for co in COMPANIES:
        pg = f"co_{co.lower().replace('&','and').replace(' ','_')}"
        nav_items += (
            f'    <div class="nav-item" onclick="showPage(\'{pg}\')">\n'
            f'      <span class="nav-dot" style="background:{CO_COLORS[co]}"></span> {co}\n'
            f'    </div>\n'
        )

    # ---------- company pages ----------
    company_pages = ""
    for co in COMPANIES:
        pg  = f"co_{co.lower().replace('&','and').replace(' ','_')}"
        col = CO_COLORS[co]
        company_pages += f"""
<div id="page-{pg}" class="page">
  <div class="co-header" style="background:{col}">
    <h2>{co}</h2>
    <p>Month-wise project profit analysis &nbsp;|&nbsp; Aug 2025 — Mar 2026</p>
  </div>
  <div class="kpi-grid">
    <div class="kpi-card" style="--kpi-color:{col}">
      <div class="kpi-icon">💰</div><div class="kpi-label">Total Invoiced</div>
      <div class="kpi-value" id="kpi-{pg}-invoice">—</div><div class="kpi-sub">&nbsp;</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#E74C3C">
      <div class="kpi-icon">📦</div><div class="kpi-label">Total Cost</div>
      <div class="kpi-value" id="kpi-{pg}-cost">—</div><div class="kpi-sub">&nbsp;</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#27AE60">
      <div class="kpi-icon">📈</div><div class="kpi-label">Total Profit</div>
      <div class="kpi-value" id="kpi-{pg}-profit">—</div><div class="kpi-sub">&nbsp;</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#F39C12">
      <div class="kpi-icon">🎯</div><div class="kpi-label">Avg Profit %</div>
      <div class="kpi-value" id="kpi-{pg}-pct">—</div><div class="kpi-sub">&nbsp;</div>
    </div>
  </div>
  <div class="chart-grid-2">
    <div class="chart-card tall">
      <h3>📅 Monthly Invoice vs Cost vs Profit (₹)</h3>
      <canvas id="chart-{pg}-bar"></canvas>
    </div>
    <div class="chart-card tall">
      <h3>🍕 Revenue by Project</h3>
      <canvas id="chart-{pg}-pie"></canvas>
    </div>
  </div>
  <div class="table-card">
    <h3>📅 Month-wise Summary</h3>
    <table id="tbl-{pg}-monthly"></table>
  </div>
  <div class="table-card">
    <h3>📋 Project Details</h3>
    <table id="tbl-{pg}-projects"></table>
  </div>
</div>
"""

    # ---------- JS: buildCompanyPage calls ----------
    build_calls = "\n".join([
        f"buildCompanyPage('{co}');" for co in COMPANIES
    ])

    # ---------- full HTML ----------
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OnealphaMed — Project Profit Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --primary:#1A3A52; --primary-light:#2E6DA4; --accent:#E74C3C;
    --green:#27AE60; --orange:#F39C12; --bg:#F0F4F8; --card:#FFFFFF;
    --border:#DDE2E8; --text:#1A252F; --muted:#7F8C8D; --sidebar-w:220px;
  }}
  *{{margin:0;padding:0;box-sizing:border-box;}}
  body{{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--text);display:flex;min-height:100vh;}}
  #sidebar{{width:var(--sidebar-w);min-height:100vh;background:var(--primary);position:fixed;top:0;left:0;z-index:100;overflow-y:auto;display:flex;flex-direction:column;}}
  .sidebar-logo{{padding:20px 16px 12px;border-bottom:1px solid rgba(255,255,255,0.12);}}
  .sidebar-logo h1{{color:#fff;font-size:16px;font-weight:700;line-height:1.3;}}
  .sidebar-logo p{{color:rgba(255,255,255,0.5);font-size:11px;margin-top:2px;}}
  .nav-section{{padding:10px 0 4px;}}
  .nav-label{{padding:6px 16px;color:rgba(255,255,255,0.4);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;}}
  .nav-item{{display:flex;align-items:center;gap:10px;padding:9px 16px;color:rgba(255,255,255,0.75);cursor:pointer;transition:all 0.15s;font-size:13px;border-left:3px solid transparent;}}
  .nav-item:hover{{background:rgba(255,255,255,0.07);color:#fff;}}
  .nav-item.active{{background:rgba(255,255,255,0.12);color:#fff;border-left-color:#4FC3F7;font-weight:600;}}
  .nav-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0;}}
  #main{{margin-left:var(--sidebar-w);flex:1;padding:24px;}}
  .page{{display:none;}} .page.active{{display:block;}}
  .page-header{{background:var(--primary);color:#fff;border-radius:12px;padding:24px 28px;margin-bottom:20px;position:relative;overflow:hidden;}}
  .page-header::before{{content:'';position:absolute;top:-30px;right:-30px;width:180px;height:180px;border-radius:50%;background:rgba(255,255,255,0.04);}}
  .page-header h2{{font-size:22px;font-weight:700;margin-bottom:4px;}}
  .page-header p{{font-size:13px;color:rgba(255,255,255,0.65);}}
  .page-header .period-badge{{display:inline-block;background:rgba(255,255,255,0.15);padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;margin-top:6px;}}
  .kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px;}}
  .kpi-card{{background:var(--card);border-radius:10px;padding:18px 20px;border:1px solid var(--border);position:relative;overflow:hidden;}}
  .kpi-card::after{{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--kpi-color,#1A3A52);}}
  .kpi-label{{font-size:11px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;}}
  .kpi-value{{font-size:22px;font-weight:700;color:var(--text);margin:6px 0 2px;}}
  .kpi-sub{{font-size:11px;color:var(--muted);}}
  .kpi-icon{{position:absolute;top:16px;right:16px;font-size:20px;opacity:0.2;}}
  .chart-grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;}}
  .chart-full{{margin-bottom:16px;}}
  .chart-card{{background:var(--card);border-radius:10px;padding:18px 20px;border:1px solid var(--border);}}
  .chart-card h3{{font-size:13px;font-weight:700;color:var(--text);margin-bottom:14px;}}
  .chart-card canvas{{max-height:260px;}} .chart-card.tall canvas{{max-height:320px;}}
  .table-card{{background:var(--card);border-radius:10px;padding:18px 20px;border:1px solid var(--border);margin-bottom:16px;overflow-x:auto;}}
  .table-card h3{{font-size:13px;font-weight:700;color:var(--text);margin-bottom:12px;}}
  table{{width:100%;border-collapse:collapse;font-size:12px;}}
  th{{background:#F7F9FC;padding:8px 10px;text-align:left;font-weight:700;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:0.4px;border-bottom:2px solid var(--border);white-space:nowrap;}}
  td{{padding:7px 10px;border-bottom:1px solid #F0F4F8;color:var(--text);}}
  tr:hover td{{background:#FAFBFC;}}
  .num{{text-align:right;font-variant-numeric:tabular-nums;}}
  .badge{{display:inline-block;padding:2px 7px;border-radius:10px;font-size:10px;font-weight:700;}}
  .badge-green{{background:#D5F5E3;color:#1A6E37;}} .badge-orange{{background:#FDEBD0;color:#D35400;}} .badge-red{{background:#FADBD8;color:#C0392B;}}
  .co-header{{border-radius:12px;padding:22px 28px;margin-bottom:20px;color:#fff;position:relative;overflow:hidden;}}
  .co-header::before{{content:'';position:absolute;right:-20px;bottom:-20px;width:140px;height:140px;border-radius:50%;background:rgba(255,255,255,0.07);}}
  .co-header h2{{font-size:20px;font-weight:700;margin-bottom:4px;}}
  .co-header p{{font-size:12px;opacity:0.7;}}
  .cat-mw{{background:#D5F5E3;color:#1A6E37;}} .cat-wt{{background:#D6EAF8;color:#1F618D;}} .cat-ev{{background:#FDEBD0;color:#D35400;}}
  ::-webkit-scrollbar{{width:5px;height:5px;}} ::-webkit-scrollbar-track{{background:#f0f4f8;}} ::-webkit-scrollbar-thumb{{background:#B2BEC3;border-radius:3px;}}
  .footer{{text-align:center;padding:20px;color:var(--muted);font-size:11px;margin-top:10px;}}
  @media(max-width:1100px){{.kpi-grid{{grid-template-columns:repeat(2,1fr);}} .chart-grid-2{{grid-template-columns:1fr;}}}}
</style>
</head>
<body>

<nav id="sidebar">
  <div class="sidebar-logo">
    <h1>OnealphaMed</h1>
    <p>Project Profit Dashboard</p>
    <p style="margin-top:6px;color:rgba(255,255,255,0.3);font-size:10px;">Aug 2025 — Mar 2026</p>
  </div>
  <div class="nav-section">
    <div class="nav-label">Overview</div>
    <div class="nav-item active" onclick="showPage('overview')"><span>📊</span> Dashboard Overview</div>
  </div>
  <div class="nav-section">
    <div class="nav-label">Companies</div>
{nav_items}  </div>
  <div class="nav-section">
    <div class="nav-label">Cost Analysis</div>
    <div class="nav-item" onclick="showPage('vendors')"><span>🔧</span> Vendor Analysis</div>
  </div>
</nav>

<main id="main">

<!-- ═══ OVERVIEW ═══ -->
<div id="page-overview" class="page active">
  <div class="page-header">
    <h2>📊 Project Profit Dashboard</h2>
    <p>Consolidated performance across all clients and projects</p>
    <span class="period-badge">Aug 2025 — Mar 2026</span>
  </div>
  <div class="kpi-grid">
    <div class="kpi-card" style="--kpi-color:#1F618D">
      <div class="kpi-icon">💰</div><div class="kpi-label">Total Invoiced</div>
      <div class="kpi-value" id="kpi-invoice">—</div><div class="kpi-sub">Across 12 companies</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#E74C3C">
      <div class="kpi-icon">📦</div><div class="kpi-label">Total Cost</div>
      <div class="kpi-value" id="kpi-cost">—</div><div class="kpi-sub">All vendor &amp; operational costs</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#27AE60">
      <div class="kpi-icon">📈</div><div class="kpi-label">Total Profit</div>
      <div class="kpi-value" id="kpi-profit">—</div><div class="kpi-sub">Net earnings</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#F39C12">
      <div class="kpi-icon">🎯</div><div class="kpi-label">Avg Profit Margin</div>
      <div class="kpi-value" id="kpi-pct">—</div><div class="kpi-sub">Blended across all projects</div>
    </div>
  </div>
  <div class="chart-full">
    <div class="chart-card tall">
      <h3>📅 Monthly Performance — Invoice vs Cost vs Profit (₹)</h3>
      <canvas id="chart-monthly-bar"></canvas>
    </div>
  </div>
  <div class="chart-grid-2">
    <div class="chart-card">
      <h3>🏢 Revenue Share by Company</h3>
      <canvas id="chart-company-pie"></canvas>
    </div>
    <div class="chart-card">
      <h3>📈 Profit Margin Trend by Company (%)
        <span style="font-size:10px;font-weight:400;color:#7F8C8D"> — click to filter</span>
      </h3>
      <div id="margin-filters" style="display:flex;flex-wrap:wrap;gap:5px;margin-bottom:10px"></div>
      <canvas id="chart-margin-trend"></canvas>
    </div>
  </div>
  <div class="table-card">
    <h3>📋 Company Performance Summary</h3>
    <table id="tbl-company-summary"></table>
  </div>
</div>

{company_pages}

<!-- ═══ VENDOR ═══ -->
<div id="page-vendors" class="page">
  <div class="page-header" style="background:#2C3E50">
    <h2>🔧 Vendor Cost Analysis</h2>
    <p>Medical Writing · Webinar/Tech · Events &nbsp;|&nbsp; KOL/Doctor honorariums excluded</p>
    <span class="period-badge">Aug 2025 — Mar 2026</span>
  </div>
  <div class="kpi-grid">
    <div class="kpi-card" style="--kpi-color:#1A6E37">
      <div class="kpi-icon">✍️</div><div class="kpi-label">Medical Writing</div>
      <div class="kpi-value" id="kpi-vend-mw">—</div><div class="kpi-sub">&nbsp;</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#1F618D">
      <div class="kpi-icon">🖥️</div><div class="kpi-label">Webinar / Tech</div>
      <div class="kpi-value" id="kpi-vend-wt">—</div><div class="kpi-sub">&nbsp;</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#C0392B">
      <div class="kpi-icon">🎪</div><div class="kpi-label">Events</div>
      <div class="kpi-value" id="kpi-vend-ev">—</div><div class="kpi-sub">&nbsp;</div>
    </div>
    <div class="kpi-card" style="--kpi-color:#2C3E50">
      <div class="kpi-icon">💸</div><div class="kpi-label">Total Vendor Cost</div>
      <div class="kpi-value" id="kpi-vend-total">—</div><div class="kpi-sub">&nbsp;</div>
    </div>
  </div>
  <div class="chart-grid-2">
    <div class="chart-card tall">
      <h3>📅 Monthly Vendor Cost — Stacked by Category (₹)</h3>
      <canvas id="chart-vend-monthly"></canvas>
    </div>
    <div class="chart-card tall">
      <h3>🍕 Overall Vendor Category Split</h3>
      <canvas id="chart-vend-pie"></canvas>
    </div>
  </div>
  <div class="chart-full">
    <div class="chart-card">
      <h3>🏢 Vendor Costs by Company</h3>
      <canvas id="chart-vend-company" style="max-height:220px"></canvas>
    </div>
  </div>
  <div class="table-card">
    <h3>📅 Month-wise Vendor Cost Breakdown</h3>
    <table id="tbl-vend-monthly"></table>
  </div>
  <div class="table-card">
    <h3>🏢 Company-wise Vendor Cost Breakdown</h3>
    <table id="tbl-vend-company"></table>
  </div>
</div>

<div class="footer">OnealphaMed · Project Profit Dashboard · Aug 2025 – Mar 2026 · All figures in INR (₹)</div>
</main>

<script>
const D = {raw_data};
const MONTHS = D.months;
const COMPANIES = D.companies;
const CO_COLORS = {json.dumps(CO_COLORS)};
const CO_LIGHT  = {json.dumps(CO_LIGHT)};

// ─── Utilities ────────────────────────────────────────────────────────────────
function fmtL(n){{
  if(!n) return '—';
  if(n>=10000000) return '₹'+(n/10000000).toFixed(2)+'Cr';
  if(n>=100000)   return '₹'+(n/100000).toFixed(2)+'L';
  return '₹'+n.toLocaleString('en-IN',{{maximumFractionDigits:0}});
}}
function fmtPct(n){{ return n?(n*100).toFixed(1)+'%':'—'; }}
function pctBadge(p){{
  const cls=p>=0.7?'badge-green':p>=0.4?'badge-orange':'badge-red';
  return `<span class="badge ${{cls}}">${{(p*100).toFixed(1)}}%</span>`;
}}
function coId(co){{ return 'co_'+co.toLowerCase().replace(/&/g,'and').replace(/ /g,'_'); }}
function setKpi(id,v){{ const el=document.getElementById(id); if(el) el.textContent=v; }}

// ─── Navigation ───────────────────────────────────────────────────────────────
function showPage(name){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const pg=document.getElementById('page-'+name);
  if(pg) pg.classList.add('active');
  document.querySelectorAll('.nav-item').forEach(n=>{{
    if(n.getAttribute('onclick')&&n.getAttribute('onclick').includes("'"+name+"'"))
      n.classList.add('active');
  }});
}}

// ─── Chart helpers ─────────────────────────────────────────────────────────────
const chartInst={{}};
function makeChart(id,cfg){{
  if(chartInst[id]) chartInst[id].destroy();
  const el=document.getElementById(id);
  if(el) chartInst[id]=new Chart(el,cfg);
}}
function yTick(v){{ return v>=1e5?'₹'+(v/1e5).toFixed(0)+'L':'₹'+v.toLocaleString(); }}

function barConfig(labels,datasets){{
  return {{type:'bar',data:{{labels,datasets}},options:{{
    responsive:true,maintainAspectRatio:true,
    plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}},boxWidth:12}}}}}},
    scales:{{
      x:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}},
      y:{{grid:{{color:'#F0F4F8'}},ticks:{{font:{{size:10}},callback:yTick}}}}
    }}
  }}}};
}}
function stackedBarConfig(labels,datasets){{
  return {{type:'bar',data:{{labels,datasets}},options:{{
    responsive:true,maintainAspectRatio:true,
    plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}},boxWidth:12}}}}}},
    scales:{{
      x:{{stacked:true,grid:{{display:false}},ticks:{{font:{{size:10}}}}}},
      y:{{stacked:true,grid:{{color:'#F0F4F8'}},ticks:{{font:{{size:10}},callback:yTick}}}}
    }}
  }}}};
}}
function lineConfig(labels,datasets){{
  return {{type:'line',data:{{labels,datasets}},options:{{
    responsive:true,maintainAspectRatio:true,
    plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}},boxWidth:12}}}}}},
    scales:{{
      x:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}},
      y:{{grid:{{color:'#F0F4F8'}},min:0,max:1,ticks:{{font:{{size:10}},callback:v=>(v*100).toFixed(0)+'%'}}}}
    }}
  }}}};
}}
function pieConfig(labels,data,colors){{
  return {{type:'doughnut',data:{{labels,datasets:[{{data,backgroundColor:colors,borderWidth:2,borderColor:'#fff',hoverOffset:6}}]}},options:{{
    responsive:true,maintainAspectRatio:true,cutout:'55%',
    plugins:{{
      legend:{{position:'right',labels:{{font:{{size:11}},boxWidth:12,padding:10}}}},
      tooltip:{{callbacks:{{label:ctx=>{{
        const tot=ctx.dataset.data.reduce((a,b)=>a+b,0);
        return ` ₹${{ctx.parsed.toLocaleString('en-IN',{{maximumFractionDigits:0}})}} (${{(ctx.parsed/tot*100).toFixed(1)}}%)`;
      }}}}}}
    }}
  }}}};
}}

// ─── Overview ─────────────────────────────────────────────────────────────────
function buildOverview(){{
  const g=D.grand;
  setKpi('kpi-invoice',fmtL(g.invoice));
  setKpi('kpi-cost',fmtL(g.cost));
  setKpi('kpi-profit',fmtL(g.profit));
  setKpi('kpi-pct',fmtPct(g.pct));

  makeChart('chart-monthly-bar', barConfig(MONTHS,[
    {{label:'Invoice',data:D.monthly.invoice,backgroundColor:'rgba(46,134,171,0.85)',borderRadius:4}},
    {{label:'Cost',   data:D.monthly.cost,   backgroundColor:'rgba(228,92,58,0.85)', borderRadius:4}},
    {{label:'Profit', data:D.monthly.profit, backgroundColor:'rgba(39,174,96,0.85)', borderRadius:4}},
  ]));

  makeChart('chart-company-pie', pieConfig(
    COMPANIES,
    COMPANIES.map(c=>D.company_totals[c].invoice),
    COMPANIES.map(c=>CO_COLORS[c])
  ));

  // Filterable margin trend
  let activeCompanies=new Set(COMPANIES);
  function buildMarginChart(){{
    const ds=COMPANIES.map(co=>({{
      label:co, data:D.company_monthly[co].pct,
      borderColor:CO_COLORS[co], backgroundColor:CO_COLORS[co]+'22',
      tension:0.3, fill:false, pointRadius:4, pointHoverRadius:6, borderWidth:2.5,
      hidden:!activeCompanies.has(co)
    }}));
    makeChart('chart-margin-trend', lineConfig(MONTHS,ds));
  }}
  const fd=document.getElementById('margin-filters');
  const allBtn=document.createElement('button');
  allBtn.textContent='All';
  allBtn.style.cssText='padding:3px 10px;border-radius:12px;border:2px solid #1A3A52;background:#1A3A52;color:#fff;font-size:10px;font-weight:700;cursor:pointer;';
  allBtn.onclick=()=>{{
    activeCompanies=new Set(COMPANIES);
    document.querySelectorAll('.co-filter-btn').forEach(b=>{{b.style.opacity='1';b.style.background=b.dataset.color;b.style.color='#fff';}});
    allBtn.style.background='#1A3A52'; allBtn.style.color='#fff';
    buildMarginChart();
  }};
  fd.appendChild(allBtn);
  COMPANIES.forEach(co=>{{
    const btn=document.createElement('button');
    btn.className='co-filter-btn'; btn.dataset.color=CO_COLORS[co]; btn.dataset.co=co;
    btn.textContent=co;
    btn.style.cssText=`padding:3px 10px;border-radius:12px;border:2px solid ${{CO_COLORS[co]}};background:${{CO_COLORS[co]}};color:#fff;font-size:10px;font-weight:700;cursor:pointer;transition:all 0.15s;`;
    btn.onclick=()=>{{
      if(activeCompanies.has(co)&&activeCompanies.size===1) return;
      if(activeCompanies.size===COMPANIES.length){{
        activeCompanies=new Set([co]);
        document.querySelectorAll('.co-filter-btn').forEach(b=>{{
          if(b.dataset.co!==co){{b.style.opacity='0.35';b.style.background='#fff';b.style.color=b.dataset.color;}}
        }});
        allBtn.style.background='transparent'; allBtn.style.color='#1A3A52';
      }} else {{
        if(activeCompanies.has(co)){{ activeCompanies.delete(co); btn.style.opacity='0.35';btn.style.background='#fff';btn.style.color=CO_COLORS[co]; }}
        else {{ activeCompanies.add(co); btn.style.opacity='1';btn.style.background=CO_COLORS[co];btn.style.color='#fff'; }}
      }}
      buildMarginChart();
    }};
    fd.appendChild(btn);
  }});
  buildMarginChart();

  // Summary table
  let h=`<thead><tr><th>Company</th><th class="num">Invoice</th><th class="num">Cost</th><th class="num">Profit</th><th class="num">Profit %</th></tr></thead><tbody>`;
  COMPANIES.forEach(co=>{{
    const t=D.company_totals[co];
    const dot=`<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${{CO_COLORS[co]}};margin-right:6px"></span>`;
    h+=`<tr><td>${{dot}}${{co}}</td><td class="num">${{fmtL(t.invoice)}}</td><td class="num">${{fmtL(t.cost)}}</td><td class="num" style="color:#27AE60;font-weight:600">${{fmtL(t.profit)}}</td><td class="num">${{pctBadge(t.pct)}}</td></tr>`;
  }});
  const g2=D.grand;
  h+=`<tr style="background:#F7F9FC;font-weight:700;border-top:2px solid #DDE2E8"><td>Grand Total</td><td class="num">${{fmtL(g2.invoice)}}</td><td class="num">${{fmtL(g2.cost)}}</td><td class="num" style="color:#27AE60">${{fmtL(g2.profit)}}</td><td class="num">${{pctBadge(g2.pct)}}</td></tr></tbody>`;
  document.getElementById('tbl-company-summary').innerHTML=h;
}}

// ─── Company Page ──────────────────────────────────────────────────────────────
function buildCompanyPage(co){{
  const pid=coId(co), t=D.company_totals[co], cm=D.company_monthly[co], color=CO_COLORS[co];
  setKpi(`kpi-${{pid}}-invoice`,fmtL(t.invoice));
  setKpi(`kpi-${{pid}}-cost`,fmtL(t.cost));
  setKpi(`kpi-${{pid}}-profit`,fmtL(t.profit));
  setKpi(`kpi-${{pid}}-pct`,fmtPct(t.pct));

  makeChart(`chart-${{pid}}-bar`, barConfig(MONTHS,[
    {{label:'Invoice',data:cm.invoice,backgroundColor:color+'CC',borderRadius:4}},
    {{label:'Cost',   data:cm.cost,   backgroundColor:'rgba(228,92,58,0.8)',borderRadius:4}},
    {{label:'Profit', data:cm.profit, backgroundColor:'rgba(39,174,96,0.8)',borderRadius:4}},
  ]));

  const projs=D.projects[co], projTotals={{}};
  projs.forEach(p=>{{ projTotals[p.Project]=(projTotals[p.Project]||0)+p.Invoice; }});
  const sorted=Object.entries(projTotals).sort((a,b)=>b[1]-a[1]).slice(0,8);
  const pal=['#2E86AB','#E45C3A','#2ECC71','#F39C12','#9B59B6','#1ABC9C','#E74C3C','#95A5A6'];
  makeChart(`chart-${{pid}}-pie`, pieConfig(
    sorted.map(x=>x[0].length>18?x[0].slice(0,18)+'…':x[0]),
    sorted.map(x=>x[1]), pal
  ));

  let h=`<thead><tr><th>Month</th><th class="num">Invoiced (₹)</th><th class="num">Cost (₹)</th><th class="num">Profit (₹)</th><th class="num">Profit %</th></tr></thead><tbody>`;
  MONTHS.forEach((m,i)=>{{
    const inv=cm.invoice[i],cost=cm.cost[i],prof=cm.profit[i],pct=cm.pct[i];
    if(!inv) return;
    h+=`<tr><td><strong>${{m}}</strong></td><td class="num">${{fmtL(inv)}}</td><td class="num">${{fmtL(cost)}}</td><td class="num" style="color:#27AE60;font-weight:600">${{fmtL(prof)}}</td><td class="num">${{pctBadge(pct)}}</td></tr>`;
  }});
  h+=`<tr style="background:#F7F9FC;font-weight:700;border-top:2px solid #DDE2E8"><td>Total</td><td class="num">${{fmtL(t.invoice)}}</td><td class="num">${{fmtL(t.cost)}}</td><td class="num" style="color:#27AE60">${{fmtL(t.profit)}}</td><td class="num">${{pctBadge(t.pct)}}</td></tr></tbody>`;
  document.getElementById(`tbl-${{pid}}-monthly`).innerHTML=h;

  const mOrd={{"Aug'25":1,"Sept'25":2,"Oct'25":3,"Nov'25":4,"Dec'25":5,"Jan'26":6,"Feb'26":7,"Mar'26":8}};
  const sP=[...projs].sort((a,b)=>(mOrd[a.Month]-mOrd[b.Month])||b.Invoice-a.Invoice);
  let h2=`<thead><tr><th>Month</th><th>Project</th><th class="num">Invoiced</th><th class="num">Cost</th><th class="num">Profit</th><th class="num">Profit %</th></tr></thead><tbody>`;
  sP.forEach(p=>{{
    h2+=`<tr><td style="white-space:nowrap">${{p.Month}}</td><td>${{p.Project}}</td><td class="num">${{fmtL(p.Invoice)}}</td><td class="num">${{fmtL(p.Cost)}}</td><td class="num" style="color:#27AE60;font-weight:600">${{fmtL(p.Profit)}}</td><td class="num">${{pctBadge(Math.min(p.ProfitPct,1))}}</td></tr>`;
  }});
  document.getElementById(`tbl-${{pid}}-projects`).innerHTML=h2+'</tbody>';
}}

// ─── Vendor Page ───────────────────────────────────────────────────────────────
function buildVendorPage(){{
  const v=D.vendor, totMW=v.totals['Medical Writing'], totWT=v.totals['Webinar/Tech'], totEV=v.totals['Events'], totAll=totMW+totWT+totEV;
  setKpi('kpi-vend-mw',fmtL(totMW)); setKpi('kpi-vend-wt',fmtL(totWT));
  setKpi('kpi-vend-ev',fmtL(totEV)); setKpi('kpi-vend-total',fmtL(totAll));

  makeChart('chart-vend-monthly', stackedBarConfig(MONTHS,[
    {{label:'Medical Writing',data:v.monthly['Medical Writing'],backgroundColor:'rgba(26,110,55,0.85)',borderRadius:2}},
    {{label:'Webinar/Tech',   data:v.monthly['Webinar/Tech'],   backgroundColor:'rgba(31,97,141,0.85)',borderRadius:2}},
    {{label:'Events',         data:v.monthly['Events'],         backgroundColor:'rgba(192,57,43,0.85)',borderRadius:2}},
  ]));
  makeChart('chart-vend-pie', pieConfig(['Medical Writing','Webinar/Tech','Events'],[totMW,totWT,totEV],['#1A6E37','#1F618D','#C0392B']));
  makeChart('chart-vend-company', stackedBarConfig(COMPANIES,[
    {{label:'Medical Writing',data:COMPANIES.map(c=>v.by_company[c]['Medical Writing']),backgroundColor:'rgba(26,110,55,0.8)',borderRadius:3}},
    {{label:'Webinar/Tech',   data:COMPANIES.map(c=>v.by_company[c]['Webinar/Tech']),   backgroundColor:'rgba(31,97,141,0.8)',borderRadius:3}},
    {{label:'Events',         data:COMPANIES.map(c=>v.by_company[c]['Events']),         backgroundColor:'rgba(192,57,43,0.8)',borderRadius:3}},
  ]));

  let h=`<thead><tr><th>Month</th><th class="num">Medical Writing</th><th class="num">Webinar/Tech</th><th class="num">Events</th><th class="num">Total</th></tr></thead><tbody>`;
  MONTHS.forEach((m,i)=>{{
    const mw=v.monthly['Medical Writing'][i],wt=v.monthly['Webinar/Tech'][i],ev=v.monthly['Events'][i],tt=mw+wt+ev;
    if(!tt) return;
    h+=`<tr><td><strong>${{m}}</strong></td><td class="num"><span class="badge cat-mw">${{fmtL(mw)||'—'}}</span></td><td class="num"><span class="badge cat-wt">${{fmtL(wt)||'—'}}</span></td><td class="num"><span class="badge cat-ev">${{fmtL(ev)||'—'}}</span></td><td class="num"><strong>${{fmtL(tt)}}</strong></td></tr>`;
  }});
  h+=`<tr style="background:#F7F9FC;font-weight:700;border-top:2px solid #DDE2E8"><td>Total</td><td class="num">${{fmtL(totMW)}}</td><td class="num">${{fmtL(totWT)}}</td><td class="num">${{fmtL(totEV)}}</td><td class="num">${{fmtL(totAll)}}</td></tr></tbody>`;
  document.getElementById('tbl-vend-monthly').innerHTML=h;

  let h2=`<thead><tr><th>Company</th><th class="num">Medical Writing</th><th class="num">Webinar/Tech</th><th class="num">Events</th><th class="num">Total</th><th class="num">% of Total</th></tr></thead><tbody>`;
  COMPANIES.forEach(co=>{{
    const vc=v.by_company[co], mw=vc['Medical Writing'],wt=vc['Webinar/Tech'],ev=vc['Events'],tt=mw+wt+ev;
    if(!tt) return;
    const dot=`<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${{CO_COLORS[co]}};margin-right:6px"></span>`;
    h2+=`<tr><td>${{dot}}${{co}}</td><td class="num">${{mw?fmtL(mw):'—'}}</td><td class="num">${{wt?fmtL(wt):'—'}}</td><td class="num">${{ev?fmtL(ev):'—'}}</td><td class="num"><strong>${{fmtL(tt)}}</strong></td><td class="num">${{pctBadge(tt/totAll)}}</td></tr>`;
  }});
  document.getElementById('tbl-vend-company').innerHTML=h2+'</tbody>';
}}

// ─── Init ──────────────────────────────────────────────────────────────────────
buildOverview();
{build_calls}
buildVendorPage();
</script>
</body>
</html>"""
    return html

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    excel_path = EXCEL_FILE

    # Allow passing a custom Excel path as argument
    if len(sys.argv) > 1:
        excel_path = Path(sys.argv[1])

    if not excel_path.exists():
        print(f"ERROR: Excel file not found at: {excel_path}")
        print(f"       Place your Excel file here: {EXCEL_FILE}")
        sys.exit(1)

    print(f"\n{'='*55}")
    print("  OnealphaMed Dashboard Generator")
    print(f"{'='*55}")
    print(f"\n[1/3] Loading data from Excel...")
    proj_df, vend_df = load_and_process(excel_path)

    print(f"\n[2/3] Building data model...")
    data = build_data_object(proj_df, vend_df)
    print(f"  Grand total invoice : ₹{data['grand']['invoice']:,.0f}")
    print(f"  Companies           : {len(COMPANIES)}")
    print(f"  Months covered      : {len(MONTHS)}")

    print(f"\n[3/3] Generating HTML dashboard...")
    html = build_html(data)
    OUTPUT_HTML.write_text(html, encoding='utf-8')
    size_kb = OUTPUT_HTML.stat().st_size / 1024
    print(f"  Output: {OUTPUT_HTML.name}  ({size_kb:.0f} KB)")

    print(f"\n{'='*55}")
    print("  ✅  Dashboard updated successfully!")
    print(f"{'='*55}\n")

if __name__ == "__main__":
    main()
