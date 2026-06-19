"""
Picker Efficiency Automation
=============================
Analyzes G2 picker performance (picking, stacking, bin-audit, empty-binning)
at India -> City -> FC -> Picker hierarchy.

Scope: G2 staff only. G1 is excluded from this automation entirely (G1's
Assigned_*/Declined_* job-tracking columns are 0 for every row even when
real output exists, so G1 has no usable acceptance/decline signal and is
out of scope per project decision).

Pipeline: load CSV(s) -> clean -> compute KPIs -> rank/flag exceptions ->
write single Excel workbook (multi-sheet) -> generate HTML dashboard ->
push summary sheets to Google Sheets -> deploy dashboard to GitHub Pages.
"""

import json
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter

try:
    import gspread
    from google_auth_oauthlib.flow import InstalledAppFlow
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False


# ──────────────────────────────────────────────────────────────────────────
# Dashboard HTML template (single-file, self-contained; data injected via
# __DASHBOARD_DATA__ marker by write_dashboard())
# ──────────────────────────────────────────────────────────────────────────

_DASHBOARD_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Picker Efficiency Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{
  --bg:#0F1419; --bg-2:#161D26; --bg-3:#1D2630; --card:#182029;
  --border:#27323D; --border-soft:#1F2933;
  --amber:#F5A623; --amber-soft:#F5A62333; --cyan:#22D3EE; --cyan-soft:#22D3EE26;
  --green:#3DD68C; --red:#FF5C72; --violet:#A78BFA;
  --text:#E8EDF2; --text-dim:#8B9AAB; --text-faint:#5A6B7A;
  --mono: 'SF Mono', 'JetBrains Mono', Consolas, monospace;
  --sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  --r: 10px; --r-lg:14px;
  --shadow: 0 8px 24px rgba(0,0,0,0.35);
}
[data-theme="light"]{
  --bg:#F4F6F8; --bg-2:#FFFFFF; --bg-3:#EEF1F4;
  --card:#FFFFFF; --border:#DCE3E9; --border-soft:#E7ECF0;
  --text:#1A2530; --text-dim:#5C6B7A; --text-faint:#94A3B2;
  --shadow: 0 4px 16px rgba(20,30,40,0.08);
}
*{box-sizing:border-box; margin:0; padding:0;}
body{
  font-family:var(--sans); background:var(--bg); color:var(--text);
  min-height:100vh; transition:background .25s ease, color .25s ease;
}
::selection{ background:var(--amber-soft); color:var(--text); }

/* ── Top bar ── */
.topbar{
  position:sticky; top:0; z-index:100; display:flex; align-items:center; justify-content:space-between;
  padding:14px 28px; background:var(--bg-2); border-bottom:1px solid var(--border);
  backdrop-filter:blur(10px);
}
.brand{ display:flex; align-items:center; gap:12px; }
.brand-mark{
  width:34px; height:34px; border-radius:8px; background:linear-gradient(135deg,var(--amber),#E08A0F);
  display:flex; align-items:center; justify-content:center; font-weight:800; color:#1A1300; font-size:15px;
}
.brand-text h1{ font-size:15px; font-weight:700; letter-spacing:.2px; }
.brand-text span{ font-size:11px; color:var(--text-dim); font-family:var(--mono); }
.topbar-right{ display:flex; align-items:center; gap:10px; }
.pill{
  font-family:var(--mono); font-size:11px; color:var(--text-dim); background:var(--bg-3);
  border:1px solid var(--border); padding:6px 12px; border-radius:20px;
}
.icon-btn{
  width:36px; height:36px; border-radius:9px; border:1px solid var(--border); background:var(--bg-3);
  color:var(--text-dim); display:flex; align-items:center; justify-content:center; cursor:pointer;
  transition:all .15s ease; font-size:16px;
}
.icon-btn:hover{ border-color:var(--amber); color:var(--amber); }

/* ── Lens toggle ── */
.lens-toggle{
  display:flex; background:var(--bg-3); border:1px solid var(--border); border-radius:20px; padding:3px; gap:2px;
}
.lens-toggle button{
  border:none; background:transparent; color:var(--text-dim); font-family:var(--sans); font-size:12px;
  font-weight:600; padding:6px 14px; border-radius:16px; cursor:pointer; transition:all .15s ease;
}
.lens-toggle button.active{ background:var(--amber); color:#1A1300; }

/* ── Layout ── */
.wrap{ max-width:1440px; margin:0 auto; padding:24px 28px 80px; }
.section{ margin-bottom:36px; }
.section-head{ display:flex; align-items:baseline; justify-content:space-between; margin-bottom:16px; }
.section-head h2{ font-size:18px; font-weight:700; }
.section-head .eyebrow{
  font-family:var(--mono); font-size:11px; color:var(--amber); text-transform:uppercase; letter-spacing:1.2px;
  display:block; margin-bottom:4px;
}
.section-sub{ font-size:12px; color:var(--text-dim); }

/* ── Breadcrumb ── */
.breadcrumb{
  display:flex; align-items:center; gap:6px; font-family:var(--mono); font-size:12px;
  color:var(--text-dim); margin-bottom:14px; flex-wrap:wrap;
}
.breadcrumb .crumb{
  cursor:pointer; padding:4px 10px; border-radius:6px; transition:all .15s ease; border:1px solid transparent;
}
.breadcrumb .crumb:hover{ background:var(--bg-3); border-color:var(--border); color:var(--text); }
.breadcrumb .crumb.current{ color:var(--cyan); font-weight:600; }
.breadcrumb .sep{ color:var(--text-faint); }

/* ── KPI cards ── */
.kpi-grid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:14px; }
.kpi-card{
  background:var(--card); border:1px solid var(--border); border-radius:var(--r-lg);
  padding:18px 18px 16px; position:relative; overflow:hidden; box-shadow:var(--shadow);
  transition:transform .2s ease, border-color .2s ease;
}
.kpi-card:hover{ transform:translateY(-2px); border-color:var(--border); }
.kpi-card::before{
  content:''; position:absolute; top:0; left:0; width:3px; height:100%; background:var(--accent,var(--amber));
}
.kpi-card .kpi-icon{ font-size:18px; opacity:.85; margin-bottom:10px; display:block; }
.kpi-card .kpi-label{ font-size:11.5px; color:var(--text-dim); text-transform:uppercase; letter-spacing:.6px; font-weight:600; }
.kpi-card .kpi-value{
  font-family:var(--mono); font-size:28px; font-weight:700; margin-top:6px; color:var(--text); line-height:1.1;
}
.kpi-card .kpi-sub{ font-size:11.5px; color:var(--text-faint); margin-top:6px; }
.kpi-card.alert{ --accent:var(--red); }
.kpi-card.warn{ --accent:var(--amber); }
.kpi-card.good{ --accent:var(--green); }
.kpi-card.info{ --accent:var(--cyan); }

/* ── Cards / panels ── */
.panel{
  background:var(--card); border:1px solid var(--border); border-radius:var(--r-lg);
  padding:20px; box-shadow:var(--shadow);
}
.grid-2{ display:grid; grid-template-columns:1.3fr 1fr; gap:16px; }
.grid-3{ display:grid; grid-template-columns:repeat(3,1fr); gap:16px; }
@media (max-width:980px){ .grid-2,.grid-3{ grid-template-columns:1fr; } }

.chart-box{ position:relative; height:280px; }
.chart-box.tall{ height:360px; }

/* ── Tables ── */
.table-toolbar{ display:flex; gap:10px; align-items:center; margin-bottom:12px; flex-wrap:wrap; }
.search-input{
  flex:1; min-width:180px; background:var(--bg-3); border:1px solid var(--border); border-radius:8px;
  padding:9px 12px; color:var(--text); font-size:13px; font-family:var(--sans);
}
.search-input:focus{ outline:none; border-color:var(--cyan); }
.btn{
  font-family:var(--sans); font-size:12.5px; font-weight:600; padding:8px 14px; border-radius:8px;
  border:1px solid var(--border); background:var(--bg-3); color:var(--text); cursor:pointer; transition:all .15s ease;
  white-space:nowrap;
}
.btn:hover{ border-color:var(--cyan); color:var(--cyan); }
.btn.primary{ background:var(--amber); color:#1A1300; border-color:var(--amber); }
.btn.primary:hover{ filter:brightness(1.08); color:#1A1300; }

.table-scroll{ overflow:auto; border-radius:8px; border:1px solid var(--border-soft); max-height:480px; }
table{ width:100%; border-collapse:collapse; font-size:12.5px; }
thead th{
  position:sticky; top:0; background:var(--bg-3); color:var(--text-dim); text-align:left; padding:10px 12px;
  font-weight:600; font-size:11px; text-transform:uppercase; letter-spacing:.4px; cursor:pointer;
  border-bottom:1px solid var(--border); white-space:nowrap; user-select:none;
}
thead th:hover{ color:var(--amber); }
thead th .arrow{ font-size:9px; margin-left:4px; opacity:.6; }
tbody td{ padding:9px 12px; border-bottom:1px solid var(--border-soft); font-family:var(--mono); white-space:nowrap; }
tbody td.name-cell{ font-family:var(--sans); }
tbody tr{ cursor:pointer; transition:background .1s ease; }
tbody tr:hover{ background:var(--bg-3); }
.badge{ display:inline-block; padding:2px 8px; border-radius:12px; font-size:11px; font-weight:700; font-family:var(--mono); }
.badge.good{ background:#3DD68C22; color:var(--green); }
.badge.warn{ background:#F5A62322; color:var(--amber); }
.badge.bad{ background:#FF5C7222; color:var(--red); }
.pager{ display:flex; align-items:center; justify-content:space-between; margin-top:10px; font-size:12px; color:var(--text-dim); }
.pager-btns{ display:flex; gap:6px; }
.pager-btns button{
  background:var(--bg-3); border:1px solid var(--border); color:var(--text); width:30px; height:28px; border-radius:6px;
  cursor:pointer; font-size:12px;
}
.pager-btns button:disabled{ opacity:.35; cursor:not-allowed; }
.pager-btns button:hover:not(:disabled){ border-color:var(--cyan); }

/* ── Ranking lists ── */
.rank-list{ display:flex; flex-direction:column; gap:8px; }
.rank-row{
  display:flex; align-items:center; gap:10px; padding:9px 10px; border-radius:8px; background:var(--bg-3);
  border:1px solid var(--border-soft); cursor:pointer; transition:all .15s ease;
}
.rank-row:hover{ border-color:var(--cyan); }
.rank-num{ font-family:var(--mono); font-size:11px; color:var(--text-faint); width:18px; }
.rank-bar-wrap{ flex:1; }
.rank-name{ font-size:12.5px; font-weight:600; display:flex; justify-content:space-between; margin-bottom:4px; }
.rank-name .v{ font-family:var(--mono); color:var(--text-dim); font-weight:700; }
.rank-bar{ height:6px; border-radius:4px; background:var(--border-soft); overflow:hidden; }
.rank-bar i{ display:block; height:100%; border-radius:4px; background:var(--amber); }
.rank-row.bottom .rank-bar i{ background:var(--red); }

/* ── Exception tabs ── */
.exc-tabs{ display:flex; gap:8px; margin-bottom:16px; flex-wrap:wrap; }
.exc-tab{
  font-family:var(--sans); font-size:12.5px; font-weight:600; padding:9px 16px; border-radius:9px;
  border:1px solid var(--border); background:var(--bg-3); color:var(--text-dim); cursor:pointer; transition:all .15s ease;
}
.exc-tab.active{ background:var(--card); border-color:var(--amber); color:var(--amber); }
.exc-tab .count{ font-family:var(--mono); margin-left:6px; opacity:.8; }

/* ── KPI count-up animation handled in JS ── */

/* ── Drilldown explorer ── */
.explorer-cards{ display:grid; grid-template-columns:repeat(auto-fill,minmax(220px,1fr)); gap:12px; }
.explorer-card{
  background:var(--bg-3); border:1px solid var(--border-soft); border-radius:10px; padding:14px;
  cursor:pointer; transition:all .15s ease;
}
.explorer-card:hover{ border-color:var(--cyan); transform:translateY(-1px); }
.explorer-card .ec-title{ font-size:13px; font-weight:700; margin-bottom:8px; }
.explorer-card .ec-stats{ display:flex; gap:14px; font-family:var(--mono); font-size:11px; color:var(--text-dim); }
.explorer-card .ec-stats b{ color:var(--text); display:block; font-size:14px; }

/* ── Footer ── */
.footer{ text-align:center; padding:30px 0 10px; color:var(--text-faint); font-size:11px; font-family:var(--mono); }

/* ── Scrollbar ── */
::-webkit-scrollbar{ width:9px; height:9px; }
::-webkit-scrollbar-track{ background:transparent; }
::-webkit-scrollbar-thumb{ background:var(--border); border-radius:6px; }
::-webkit-scrollbar-thumb:hover{ background:var(--text-faint); }

/* ── Responsive ── */
@media (max-width:680px){
  .wrap{ padding:16px; }
  .topbar{ padding:12px 16px; flex-wrap:wrap; gap:10px; }
  .kpi-grid{ grid-template-columns:repeat(2,1fr); }
}

/* Reduced motion */
@media (prefers-reduced-motion: reduce){
  *{ transition:none !important; animation:none !important; }
}
</style>
</head>
<body data-theme="dark">

<div class="topbar">
  <div class="brand">
    <div class="brand-mark">PE</div>
    <div class="brand-text">
      <h1>Picker Efficiency</h1>
      <span id="dateRangeLabel">—</span>
    </div>
  </div>
  <div class="topbar-right">
    <span class="pill" id="scopePill">G2 STAFF ONLY</span>
    <button class="icon-btn" id="themeToggle" title="Toggle theme">◐</button>
  </div>
</div>

<div class="wrap">

  <!-- ═══ Executive Summary ═══ -->
  <div class="section" id="sec-summary">
    <div class="section-head">
      <div>
        <span class="eyebrow">India · Overview</span>
        <h2>Executive Summary</h2>
      </div>
      <span class="section-sub" id="execSub"></span>
    </div>
    <div class="kpi-grid" id="kpiGrid"></div>
  </div>

  <!-- ═══ India Overview Charts ═══ -->
  <div class="section">
    <div class="section-head">
      <div><span class="eyebrow">Visual Breakdown</span><h2>India Overview</h2></div>
    </div>
    <div class="grid-2">
      <div class="panel">
        <div class="section-sub" style="margin-bottom:10px;">Utilization % vs Acceptance Rate % — every G2 picker (outliers visible at a glance)</div>
        <div class="chart-box tall"><canvas id="scatterChart"></canvas></div>
      </div>
      <div class="panel">
        <div class="section-sub" style="margin-bottom:10px;">Activity Mix — time share across task types</div>
        <div class="chart-box tall"><canvas id="mixDonut"></canvas></div>
      </div>
    </div>
  </div>

  <!-- ═══ City Ranking ═══ -->
  <div class="section" id="sec-city">
    <div class="section-head">
      <div><span class="eyebrow">Hierarchy · Level 1</span><h2>City Ranking</h2></div>
      <span class="section-sub">Click a city to drill into its FCs</span>
    </div>
    <div class="grid-2">
      <div class="panel">
        <div class="section-sub" style="margin-bottom:10px; font-weight:700; color:var(--green);">Top 10 Cities</div>
        <div class="rank-list" id="topCitiesList"></div>
      </div>
      <div class="panel">
        <div class="section-sub" style="margin-bottom:10px; font-weight:700; color:var(--red);">Bottom 10 Cities</div>
        <div class="rank-list" id="bottomCitiesList"></div>
      </div>
    </div>
  </div>

  <!-- ═══ FC Ranking ═══ -->
  <div class="section" id="sec-fc">
    <div class="section-head">
      <div><span class="eyebrow">Hierarchy · Level 2</span><h2>FC Ranking</h2></div>
      <span class="section-sub">Click an FC to see its pickers</span>
    </div>
    <div class="grid-2">
      <div class="panel">
        <div class="section-sub" style="margin-bottom:10px; font-weight:700; color:var(--green);">Top 10 FCs</div>
        <div class="rank-list" id="topFCsList"></div>
      </div>
      <div class="panel">
        <div class="section-sub" style="margin-bottom:10px; font-weight:700; color:var(--red);">Bottom 10 FCs</div>
        <div class="rank-list" id="bottomFCsList"></div>
      </div>
    </div>
  </div>

  <!-- ═══ Top / Bottom Pickers ═══ -->
  <div class="section" id="sec-pickers">
    <div class="section-head">
      <div><span class="eyebrow">Hierarchy · Level 3</span><h2>Picker Performance</h2></div>
    </div>
    <div class="grid-2">
      <div class="panel">
        <div class="section-sub" style="margin-bottom:10px; font-weight:700; color:var(--green);">Top 10 Pickers</div>
        <div class="rank-list" id="topPickersList"></div>
      </div>
      <div class="panel">
        <div class="section-sub" style="margin-bottom:10px; font-weight:700; color:var(--red);">Bottom 10 Pickers</div>
        <div class="rank-list" id="bottomPickersList"></div>
      </div>
    </div>
  </div>

  <!-- ═══ Exception Reports ═══ -->
  <div class="section" id="sec-exceptions">
    <div class="section-head">
      <div><span class="eyebrow">Needs Attention</span><h2>Exception Reports</h2></div>
    </div>
    <div class="exc-tabs">
      <button class="exc-tab active" data-exc="lowUtil">Low Utilization <span class="count" id="cnt-lowUtil"></span></button>
      <button class="exc-tab" data-exc="highDecline">High Decline <span class="count" id="cnt-highDecline"></span></button>
      <button class="exc-tab" data-exc="zeroProd">Zero Productivity <span class="count" id="cnt-zeroProd"></span></button>
    </div>
    <div class="panel">
      <div class="table-toolbar">
        <input class="search-input" id="excSearch" placeholder="Search picker, FC, or city…">
        <button class="btn" id="excExport">⬇ Export CSV</button>
      </div>
      <div class="table-scroll"><table id="excTable"><thead></thead><tbody></tbody></table></div>
      <div class="pager" id="excPager"></div>
    </div>
  </div>

  <!-- ═══ Drilldown Explorer ═══ -->
  <div class="section" id="sec-explorer">
    <div class="section-head">
      <div><span class="eyebrow">Full Hierarchy</span><h2>Drilldown Explorer</h2></div>
    </div>
    <div class="panel">
      <div class="breadcrumb" id="explorerCrumb"></div>
      <div id="explorerBody"></div>
    </div>
  </div>

  <div class="footer">Picker Efficiency Automation — generated <span id="genTime"></span> — G2 staff only</div>
</div>

<script>
const DATA = __DASHBOARD_DATA__;

/* ───────────────────────── Theme ───────────────────────── */
const themeBtn = document.getElementById('themeToggle');
themeBtn.addEventListener('click', () => {
  const cur = document.body.getAttribute('data-theme');
  const next = cur === 'dark' ? 'light' : 'dark';
  document.body.setAttribute('data-theme', next);
  themeBtn.textContent = next === 'dark' ? '◐' : '◑';
  Chart.defaults.color = next === 'dark' ? '#8B9AAB' : '#5C6B7A';
  rebuildCharts();
});

/* ───────────────────────── Helpers ───────────────────────── */
function fmtNum(n, d=1){ if(n===null||n===undefined||isNaN(n)) return '—'; return Number(n).toLocaleString(undefined,{maximumFractionDigits:d}); }
function fmtPct(n, d=1){ if(n===null||n===undefined||isNaN(n)) return '—'; return Number(n).toFixed(d)+'%'; }
function cssVar(name){ return getComputedStyle(document.body).getPropertyValue(name).trim(); }

function animateCountUp(el, target, isPct, duration=900){
  const start = 0; const startTime = performance.now();
  function tick(now){
    const p = Math.min((now-startTime)/duration, 1);
    const eased = 1 - Math.pow(1-p, 3);
    const val = start + (target-start)*eased;
    el.textContent = isPct ? fmtPct(val) : fmtNum(val,0);
    if(p<1) requestAnimationFrame(tick);
    else el.textContent = isPct ? fmtPct(target) : fmtNum(target,0);
  }
  requestAnimationFrame(tick);
}

function exportCSV(rows, headers, keys, filename){
  const lines = [headers.join(',')];
  rows.forEach(r => { lines.push(keys.map(k => JSON.stringify(r[k] ?? '')).join(',')); });
  const blob = new Blob([lines.join('\\n')], {type:'text/csv'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; a.click();
}

function badge(val, goodMin, warnMin){
  if(val===null||val===undefined||isNaN(val)) return '<span class="badge">—</span>';
  let cls = val>=goodMin ? 'good' : (val>=warnMin ? 'warn' : 'bad');
  return `<span class="badge ${cls}">${fmtPct(val,0)}</span>`;
}

/* ───────────────────────── Header / KPIs ───────────────────────── */
document.getElementById('dateRangeLabel').textContent =
  DATA.kpis.dateMin === DATA.kpis.dateMax ? DATA.kpis.dateMin : `${DATA.kpis.dateMin} → ${DATA.kpis.dateMax}`;
document.getElementById('genTime').textContent = new Date().toLocaleString();
document.getElementById('execSub').textContent = `${fmtNum(DATA.kpis.totalPickers)} pickers · ${fmtNum(DATA.kpis.totalFCs)} FCs · ${fmtNum(DATA.kpis.totalCities)} cities`;

const kpiCards = [
  {icon:'👥', label:'Total G2 Pickers', value:DATA.kpis.totalPickers, sub:`${DATA.kpis.totalFCs} FCs · ${DATA.kpis.totalCities} cities`, cls:'info'},
  {icon:'⚡', label:'Avg Utilization', value:DATA.kpis.avgUtilization, isPct:true, sub:'Busy time / Login time', cls:'good'},
  {icon:'✅', label:'Avg Acceptance Rate', value:DATA.kpis.avgAcceptance, isPct:true, sub:'Overall, all job types', cls:'good'},
  {icon:'🏆', label:'Avg Performance Score', value:DATA.kpis.avgScore, sub:'Out of 100', cls:'warn'},
  {icon:'📦', label:'Total Orders Picked', value:DATA.kpis.totalOrders, sub:'Picking activity only', cls:'info'},
  {icon:'📊', label:'Total Jobs Assigned', value:DATA.kpis.totalJobsAssigned, sub:`${fmtNum(DATA.kpis.totalJobsDeclined)} declined`, cls:'info'},
  {icon:'🔻', label:'Low Utilization', value:DATA.kpis.lowUtilCount, sub:`< ${DATA.config.lowUtilThreshold}% utilization`, cls:'alert'},
  {icon:'⚠️', label:'High Decline', value:DATA.kpis.highDeclineCount, sub:`< ${DATA.config.highDeclineThreshold}% acceptance`, cls:'alert'},
  {icon:'🚫', label:'Zero Productivity', value:DATA.kpis.zeroProdCount, sub:'Logged in, zero jobs', cls:'alert'},
];
const kpiGrid = document.getElementById('kpiGrid');
kpiCards.forEach(k => {
  const card = document.createElement('div');
  card.className = `kpi-card ${k.cls}`;
  card.innerHTML = `<span class="kpi-icon">${k.icon}</span><div class="kpi-label">${k.label}</div><div class="kpi-value">0</div><div class="kpi-sub">${k.sub}</div>`;
  kpiGrid.appendChild(card);
  animateCountUp(card.querySelector('.kpi-value'), k.value, !!k.isPct);
});

/* ───────────────────────── Scatter: Utilization vs Acceptance ───────────────────────── */
let scatterChartInst, mixDonutInst;
function rebuildCharts(){
  if(scatterChartInst) scatterChartInst.destroy();
  if(mixDonutInst) mixDonutInst.destroy();
  buildScatter();
  buildMixDonut();
}

function buildScatter(){
  const ctx = document.getElementById('scatterChart');
  const sample = DATA.pickers.length > 2500 ? DATA.pickers.filter((_,i)=> i % Math.ceil(DATA.pickers.length/2500) === 0) : DATA.pickers;
  const points = sample.map(p => ({x:p.ut, y:p.ar, nm:p.nm, fc:p.fc}));
  scatterChartInst = new Chart(ctx, {
    type:'scatter',
    data:{ datasets:[{
      data: points,
      backgroundColor: pt => {
        const v = pt.raw; 
        if(v.x < DATA.config.lowUtilThreshold || v.y < DATA.config.highDeclineThreshold) return '#FF5C7299';
        return '#22D3EE88';
      },
      pointRadius:3.5, pointHoverRadius:6,
    }]},
    options:{
      responsive:true, maintainAspectRatio:false,
      scales:{
        x:{ title:{display:true,text:'Utilization %', color:cssVar('--text-dim')}, min:0, max:100, grid:{color:cssVar('--border')}, ticks:{color:cssVar('--text-dim')} },
        y:{ title:{display:true,text:'Acceptance Rate %', color:cssVar('--text-dim')}, min:0, max:100, grid:{color:cssVar('--border')}, ticks:{color:cssVar('--text-dim')} },
      },
      plugins:{
        legend:{display:false},
        tooltip:{ callbacks:{ label: c => `${c.raw.nm} (${c.raw.fc}) — Util ${c.raw.x}%, Accept ${c.raw.y}%` } }
      }
    }
  });
}

function buildMixDonut(){
  const ctx = document.getElementById('mixDonut');
  const totals = DATA.pickers.reduce((acc,p) => {
    acc.p += (p.mp||0); acc.s += (p.ms||0); acc.b += (p.mb||0); acc.e += (p.me||0); return acc;
  }, {p:0,s:0,b:0,e:0});
  mixDonutInst = new Chart(ctx, {
    type:'doughnut',
    data:{
      labels:['Picking','Stacking','Bin Audit','Empty Binning'],
      datasets:[{ data:[totals.p,totals.s,totals.b,totals.e], backgroundColor:['#F5A623','#22D3EE','#A78BFA','#3DD68C'], borderWidth:0 }]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ position:'bottom', labels:{ color:cssVar('--text-dim'), padding:14, font:{size:11} } } }
    }
  });
}
buildScatter(); buildMixDonut();

/* ───────────────────────── Ranking lists ───────────────────────── */
function renderRankList(containerId, items, nameKey, valKey, isBottom, isPct, onClick){
  const el = document.getElementById(containerId);
  const maxVal = Math.max(...items.map(i=>Math.abs(i[valKey]||0)), 1);
  el.innerHTML = items.map((item,i) => {
    const v = item[valKey];
    const pctWidth = Math.min(Math.abs(v)/maxVal*100, 100);
    return `<div class="rank-row ${isBottom?'bottom':''}" data-idx="${i}">
      <span class="rank-num">${i+1}</span>
      <div class="rank-bar-wrap">
        <div class="rank-name"><span>${item[nameKey]}</span><span class="v">${isPct?fmtPct(v):fmtNum(v)}</span></div>
        <div class="rank-bar"><i style="width:${pctWidth}%"></i></div>
      </div>
    </div>`;
  }).join('');
  el.querySelectorAll('.rank-row').forEach(row => {
    row.addEventListener('click', () => onClick(items[+row.dataset.idx]));
  });
}

renderRankList('topCitiesList', DATA.topCities, 'ci', 'rk', false, false, item => openExplorerAt('city', item.ci));
renderRankList('bottomCitiesList', DATA.bottomCities, 'ci', 'rk', true, false, item => openExplorerAt('city', item.ci));
renderRankList('topFCsList', DATA.topFCs, 'fc', 'rk', false, false, item => openExplorerAt('fc', item.fc));
renderRankList('bottomFCsList', DATA.bottomFCs, 'fc', 'rk', true, false, item => openExplorerAt('fc', item.fc));
renderRankList('topPickersList', DATA.topPickers, 'nm', 'ps', false, false, item => openExplorerAt('picker', item.id));
renderRankList('bottomPickersList', DATA.bottomPickers, 'nm', 'ps', true, false, item => openExplorerAt('picker', item.id));

/* ───────────────────────── Exception Reports Table ───────────────────────── */
const excState = { tab:'lowUtil', search:'', sortKey:'ut', sortDir:1, page:1, pageSize:25 };
const excHeaders = [
  {k:'ci', label:'City'}, {k:'fc', label:'FC'}, {k:'nm', label:'Picker'},
  {k:'ut', label:'Utilization %', pct:true}, {k:'ar', label:'Acceptance %', pct:true},
  {k:'oh', label:'Orders/Hr'}, {k:'qh', label:'Qty/Hr'}, {k:'ps', label:'Score'},
];

document.querySelectorAll('.exc-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.exc-tab').forEach(t=>t.classList.remove('active'));
    tab.classList.add('active');
    excState.tab = tab.dataset.exc; excState.page = 1;
    renderExcTable();
  });
});
document.getElementById('excSearch').addEventListener('input', e => { excState.search = e.target.value.toLowerCase(); excState.page=1; renderExcTable(); });
document.getElementById('excExport').addEventListener('click', () => {
  const rows = getExcFiltered();
  exportCSV(rows, excHeaders.map(h=>h.label), excHeaders.map(h=>h.k), `${excState.tab}.csv`);
});

function getExcFiltered(){
  let rows = DATA[excState.tab] || [];
  if(excState.search){
    rows = rows.filter(r => (r.nm||'').toLowerCase().includes(excState.search) || (r.fc||'').toLowerCase().includes(excState.search) || (r.ci||'').toLowerCase().includes(excState.search));
  }
  rows = [...rows].sort((a,b) => ((a[excState.sortKey]||0) - (b[excState.sortKey]||0)) * excState.sortDir);
  return rows;
}

function renderExcTable(){
  document.getElementById('cnt-lowUtil').textContent = `(${DATA.lowUtil.length})`;
  document.getElementById('cnt-highDecline').textContent = `(${DATA.highDecline.length})`;
  document.getElementById('cnt-zeroProd').textContent = `(${DATA.zeroProd.length})`;

  const rows = getExcFiltered();
  const totalPages = Math.max(1, Math.ceil(rows.length / excState.pageSize));
  excState.page = Math.min(excState.page, totalPages);
  const pageRows = rows.slice((excState.page-1)*excState.pageSize, excState.page*excState.pageSize);

  const thead = document.querySelector('#excTable thead');
  thead.innerHTML = '<tr>' + excHeaders.map(h => {
    const arrow = excState.sortKey===h.k ? (excState.sortDir>0?'▲':'▼') : '';
    return `<th data-k="${h.k}">${h.label} <span class="arrow">${arrow}</span></th>`;
  }).join('') + '</tr>';
  thead.querySelectorAll('th').forEach(th => {
    th.addEventListener('click', () => {
      const k = th.dataset.k;
      if(excState.sortKey===k) excState.sortDir *= -1; else { excState.sortKey=k; excState.sortDir=-1; }
      renderExcTable();
    });
  });

  const tbody = document.querySelector('#excTable tbody');
  tbody.innerHTML = pageRows.map(r => `<tr data-id="${r.id||''}">
    <td>${r.ci||''}</td><td>${r.fc||''}</td><td class="name-cell">${r.nm||''}</td>
    <td>${badge(r.ut, 70, DATA.config.lowUtilThreshold)}</td>
    <td>${badge(r.ar, 95, DATA.config.highDeclineThreshold)}</td>
    <td>${fmtNum(r.oh)}</td><td>${fmtNum(r.qh)}</td><td>${fmtNum(r.ps)}</td>
  </tr>`).join('');
  tbody.querySelectorAll('tr').forEach(tr => {
    tr.addEventListener('click', () => openExplorerAt('picker', tr.dataset.id));
  });

  const pager = document.getElementById('excPager');
  pager.innerHTML = `<span>${rows.length} rows — page ${excState.page} of ${totalPages}</span>
    <div class="pager-btns">
      <button id="excPrev" ${excState.page<=1?'disabled':''}>‹</button>
      <button id="excNext" ${excState.page>=totalPages?'disabled':''}>›</button>
    </div>`;
  document.getElementById('excPrev')?.addEventListener('click', () => { excState.page--; renderExcTable(); });
  document.getElementById('excNext')?.addEventListener('click', () => { excState.page++; renderExcTable(); });
}
renderExcTable();

/* ───────────────────────── Drilldown Explorer ───────────────────────── */
const explorerState = { level:'india', city:null, fc:null,
  search:'', sortKey:'ps', sortDir:-1, page:1, pageSize:30 };

function setCrumb(){
  const crumb = document.getElementById('explorerCrumb');
  let html = `<span class="crumb ${explorerState.level==='india'?'current':''}" data-go="india">India</span>`;
  if(explorerState.city){
    html += `<span class="sep">/</span><span class="crumb ${explorerState.level==='city'?'current':''}" data-go="city">${explorerState.city}</span>`;
  }
  if(explorerState.fc){
    html += `<span class="sep">/</span><span class="crumb current" data-go="fc">${explorerState.fc}</span>`;
  }
  crumb.innerHTML = html;
  crumb.querySelectorAll('.crumb').forEach(c => c.addEventListener('click', () => {
    const go = c.dataset.go;
    if(go==='india'){ explorerState.level='india'; explorerState.city=null; explorerState.fc=null; }
    if(go==='city'){ explorerState.level='city'; explorerState.fc=null; }
    if(go==='fc'){ explorerState.level='fc'; }
    renderExplorer();
  }));
}

function openExplorerAt(type, value){
  if(type==='city'){ explorerState.level='city'; explorerState.city=value; explorerState.fc=null; }
  if(type==='fc'){
    const fcRec = DATA.fcs.find(f=>f.fc===value);
    explorerState.level='fc'; explorerState.fc=value; explorerState.city = fcRec ? fcRec.ci : explorerState.city;
  }
  if(type==='picker'){
    const pRec = DATA.pickers.find(p=>p.id===value);
    if(pRec){ explorerState.level='fc'; explorerState.city=pRec.ci; explorerState.fc=pRec.fc; }
  }
  renderExplorer();
  document.getElementById('sec-explorer').scrollIntoView({behavior:'smooth', block:'start'});
}

function renderExplorer(){
  setCrumb();
  const body = document.getElementById('explorerBody');

  if(explorerState.level==='india'){
    body.innerHTML = `<div class="explorer-cards" id="cityCards"></div>`;
    const cards = document.getElementById('cityCards');
    const sorted = [...DATA.cities].sort((a,b)=>b.rk-a.rk);
    cards.innerHTML = sorted.map(c => `<div class="explorer-card" data-ci="${c.ci}">
      <div class="ec-title">${c.ci}</div>
      <div class="ec-stats">
        <div><b>${fmtNum(c.pc)}</b>pickers</div>
        <div><b>${fmtPct(c.ut,0)}</b>util</div>
        <div><b>${fmtNum(c.rk,0)}</b>score</div>
      </div></div>`).join('');
    cards.querySelectorAll('.explorer-card').forEach(card => card.addEventListener('click', () => openExplorerAt('city', card.dataset.ci)));
    return;
  }

  if(explorerState.level==='city'){
    const fcs = DATA.fcs.filter(f => f.ci === explorerState.city);
    body.innerHTML = `<div class="explorer-cards" id="fcCards"></div>`;
    const cards = document.getElementById('fcCards');
    const sorted = [...fcs].sort((a,b)=>b.rk-a.rk);
    cards.innerHTML = sorted.map(f => `<div class="explorer-card" data-fc="${f.fc}">
      <div class="ec-title">${f.fc}</div>
      <div class="ec-stats">
        <div><b>${fmtNum(f.pc)}</b>pickers</div>
        <div><b>${fmtPct(f.ut,0)}</b>util</div>
        <div><b>${fmtNum(f.rk,0)}</b>score</div>
      </div></div>`).join('');
    cards.querySelectorAll('.explorer-card').forEach(card => card.addEventListener('click', () => openExplorerAt('fc', card.dataset.fc)));
    return;
  }

  if(explorerState.level==='fc'){
    renderPickerTable();
  }
}

function getExplorerPickers(){
  let rows = DATA.pickers.filter(p => p.fc === explorerState.fc);
  if(explorerState.search){
    const s = explorerState.search.toLowerCase();
    rows = rows.filter(r => (r.nm||'').toLowerCase().includes(s) || (r.id||'').toLowerCase().includes(s));
  }
  rows = [...rows].sort((a,b)=> ((a[explorerState.sortKey]||0)-(b[explorerState.sortKey]||0)) * explorerState.sortDir);
  return rows;
}

const explorerHeaders = [
  {k:'id', label:'ID'}, {k:'nm', label:'Picker'}, {k:'lg', label:'Login (min)'},
  {k:'ut', label:'Util %', pct:true}, {k:'ar', label:'Accept %', pct:true},
  {k:'oh', label:'Orders/Hr'}, {k:'qh', label:'Qty/Hr'}, {k:'ps', label:'Score'},
];

function renderPickerTable(){
  const body = document.getElementById('explorerBody');
  body.innerHTML = `
    <div class="table-toolbar">
      <input class="search-input" id="explSearch" placeholder="Search picker name or ID…" value="${explorerState.search}">
      <button class="btn" id="explExport">⬇ Export CSV</button>
    </div>
    <div class="table-scroll"><table id="explTable"><thead></thead><tbody></tbody></table></div>
    <div class="pager" id="explPager"></div>`;

  document.getElementById('explSearch').addEventListener('input', e => { explorerState.search = e.target.value; explorerState.page=1; renderPickerTableBody(); });
  document.getElementById('explExport').addEventListener('click', () => {
    const rows = getExplorerPickers();
    exportCSV(rows, explorerHeaders.map(h=>h.label), explorerHeaders.map(h=>h.k), `${explorerState.fc}_pickers.csv`);
  });
  renderPickerTableBody();
}

function renderPickerTableBody(){
  const rows = getExplorerPickers();
  const totalPages = Math.max(1, Math.ceil(rows.length/explorerState.pageSize));
  explorerState.page = Math.min(explorerState.page, totalPages);
  const pageRows = rows.slice((explorerState.page-1)*explorerState.pageSize, explorerState.page*explorerState.pageSize);

  const thead = document.querySelector('#explTable thead');
  thead.innerHTML = '<tr>' + explorerHeaders.map(h => {
    const arrow = explorerState.sortKey===h.k ? (explorerState.sortDir>0?'▲':'▼') : '';
    return `<th data-k="${h.k}">${h.label} <span class="arrow">${arrow}</span></th>`;
  }).join('') + '</tr>';
  thead.querySelectorAll('th').forEach(th => {
    th.addEventListener('click', () => {
      const k = th.dataset.k;
      if(explorerState.sortKey===k) explorerState.sortDir *= -1; else { explorerState.sortKey=k; explorerState.sortDir=-1; }
      renderPickerTableBody();
    });
  });

  const tbody = document.querySelector('#explTable tbody');
  tbody.innerHTML = pageRows.map(r => `<tr>
    <td>${r.id||''}</td><td class="name-cell">${r.nm||''}</td><td>${fmtNum(r.lg,0)}</td>
    <td>${badge(r.ut,70,DATA.config.lowUtilThreshold)}</td><td>${badge(r.ar,95,DATA.config.highDeclineThreshold)}</td>
    <td>${fmtNum(r.oh)}</td><td>${fmtNum(r.qh)}</td><td>${fmtNum(r.ps)}</td>
  </tr>`).join('');

  const pager = document.getElementById('explPager');
  pager.innerHTML = `<span>${rows.length} pickers — page ${explorerState.page} of ${totalPages}</span>
    <div class="pager-btns">
      <button id="explPrev" ${explorerState.page<=1?'disabled':''}>‹</button>
      <button id="explNext" ${explorerState.page>=totalPages?'disabled':''}>›</button>
    </div>`;
  document.getElementById('explPrev')?.addEventListener('click', () => { explorerState.page--; renderPickerTableBody(); });
  document.getElementById('explNext')?.addEventListener('click', () => { explorerState.page++; renderPickerTableBody(); });
}

renderExplorer();
</script>
</body>
</html>
"""


# ──────────────────────────────────────────────────────────────────────────
# Config
# ──────────────────────────────────────────────────────────────────────────

DEFAULT_CONFIG = {
    "input_folder": "input_files",
    "output_folder": "output",
    "output_file": "picker_efficiency_report.xlsx",
    "dashboard_file": "index.html",
    "low_utilization_threshold": 50,
    "high_decline_threshold": 90,
    "top_n": 10,
    "bottom_n": 10,
    "cities": [],
    "push_to_google_sheets": False,
    "deploy_to_github_pages": False,
}


def load_config():
    """Load config.json next to this script, falling back to defaults for any
    missing keys so the script always runs even with a partial/missing file."""
    cfg_path = Path(__file__).resolve().parent / "config.json"
    cfg = dict(DEFAULT_CONFIG)
    if cfg_path.exists():
        try:
            user_cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
            cfg.update({k: v for k, v in user_cfg.items() if v is not None})
        except Exception as e:
            print(f"  [Config] Warning: could not parse config.json ({e}), using defaults.")
    return cfg


# ──────────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────────

# Columns that are null when a picker did zero of that activity that day —
# fill with 0 rather than dropping (genuine zero activity, not missing data).
PICKING_NULL_FILL_COLS = [
    "total_order_picked", "total_sku_picked", "picked_quantity", "avg_picking_time_min",
]
STACKING_NULL_FILL_COLS = [
    "Total_Qty_Stacked", "Total_sku_stacked", "Avg_stacking_time_per_SKU_min",
]
BINAUDIT_NULL_FILL_COLS = [
    "Bin_Audit_Scheduled_sku_count", "Avg_time_per_bin",
    "Avg_Scheduled_Bin_Audit_time_Per_Sku", "Bin_Audit_Triggered_sku_count",
]

# 100% null in the source file profiled for this project — zero information,
# dropped entirely rather than carried through as a dead column.
COLUMN_TO_DROP = "Avg_Triggered_Bin_Audit_time_Per_Sku"

JOB_TYPES = ["Picking", "Stacking", "BinAudit", "EmptyBinning"]


# ──────────────────────────────────────────────────────────────────────────
# Data loading & cleaning
# ──────────────────────────────────────────────────────────────────────────

def load_input_files(input_folder: Path) -> pd.DataFrame:
    """Concatenate every CSV in input_files/ into one DataFrame. Supports a
    single file (current dataset) or multiple day-files dropped in later
    without any code change."""
    csv_files = sorted(input_folder.glob("*.csv"))
    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in {input_folder}")

    frames = []
    for f in csv_files:
        df = pd.read_csv(f)
        frames.append(df)
        print(f"  [Load] {f.name}: {len(df):,} rows")

    combined = pd.concat(frames, ignore_index=True)
    print(f"  [Load] Total combined: {len(combined):,} rows from {len(csv_files)} file(s)")
    return combined


def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Clean the raw picker performance data. Returns (cleaned_df, validation_info)
    where validation_info captures every cleaning action taken, for the
    Validation_Log sheet.

    Cleaning steps (in order):
      1. Drop fully-null column (Avg_Triggered_Bin_Audit_time_Per_Sku)
      2. Fill nulls in picking/stacking/bin-audit output columns with 0
         (null means zero activity of that type that day, not missing data)
      3. Strip whitespace from text identifier columns
      4. Cast date column to a real date type
      5. Filter to role == 'G2' only (G1 out of scope for this automation)
      6. De-duplicate on (user_id, fc_id, date) defensively — sum numeric
         columns if a true duplicate combo is ever found
    """
    info = {
        "rows_in": len(df),
        "columns_in": len(df.columns),
        "actions": [],
    }

    df = df.copy()

    # 1. Drop the fully-null column
    if COLUMN_TO_DROP in df.columns:
        null_pct = df[COLUMN_TO_DROP].isna().mean() * 100
        df = df.drop(columns=[COLUMN_TO_DROP])
        info["actions"].append(
            f"Dropped column '{COLUMN_TO_DROP}' ({null_pct:.0f}% null, no usable data)"
        )

    # 2. Fill nulls — genuine zero activity, not missing data
    for col_group, label in [
        (PICKING_NULL_FILL_COLS, "picking"),
        (STACKING_NULL_FILL_COLS, "stacking"),
        (BINAUDIT_NULL_FILL_COLS, "bin-audit"),
    ]:
        for col in col_group:
            if col in df.columns:
                n_filled = df[col].isna().sum()
                if n_filled > 0:
                    df[col] = df[col].fillna(0)
                    info["actions"].append(
                        f"Filled {n_filled:,} null(s) in '{col}' ({label}) with 0 "
                        f"(no {label} activity that day)"
                    )

    # 3. Strip whitespace from identifier/text columns
    text_cols = ["fc_name", "city", "Picker_Name", "user_id", "fc_id", "role"]
    for col in text_cols:
        if col in df.columns and df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()

    # 4. Cast date to real date type
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"]).dt.date

    # 5. Filter to G2 only — G1 out of scope for this automation
    role_counts = df["role"].value_counts().to_dict()
    g1_count = role_counts.get("G1", 0)
    g2_count = role_counts.get("G2", 0)
    other_count = len(df) - g1_count - g2_count

    df = df[df["role"] == "G2"].copy()
    info["actions"].append(
        f"Filtered to role == 'G2' only: kept {g2_count:,} rows, "
        f"excluded {g1_count:,} G1 rows"
        + (f" and {other_count:,} rows with other/unknown role" if other_count else "")
    )
    info["g1_excluded"] = g1_count
    info["g2_kept"] = g2_count
    info["other_excluded"] = other_count

    # 6. Defensive de-dup on (user_id, fc_id, date) — sum numeric cols if found
    dup_key = ["user_id", "fc_id", "date"]
    if all(c in df.columns for c in dup_key):
        dup_mask = df.duplicated(subset=dup_key, keep=False)
        n_dup_rows = dup_mask.sum()
        if n_dup_rows > 0:
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            non_numeric_cols = [c for c in df.columns if c not in numeric_cols and c not in dup_key]
            agg_dict = {c: "sum" for c in numeric_cols}
            agg_dict.update({c: "first" for c in non_numeric_cols})
            df = df.groupby(dup_key, as_index=False).agg(agg_dict)
            info["actions"].append(
                f"Found {n_dup_rows:,} duplicate (user_id, fc_id, date) rows — "
                f"summed numeric columns into {n_dup_rows - len(df[dup_mask.index]):,} combined rows"
            )
        else:
            info["actions"].append("No duplicate (user_id, fc_id, date) combinations found")

    info["rows_out"] = len(df)
    return df.reset_index(drop=True), info


# ──────────────────────────────────────────────────────────────────────────
# KPI calculations (picker-level, row-by-row)
# ──────────────────────────────────────────────────────────────────────────

def _safe_div(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    """Element-wise division that returns NaN where denominator is 0,
    instead of inf/error. Used everywhere a rate is 'undefined' rather
    than 0 when nothing was assigned (e.g. Acceptance Rate with 0 assigned
    jobs is not a 0% acceptance rate — it's simply not applicable)."""
    num = numerator.astype(float)
    den = denominator.astype(float)
    return np.where(den > 0, num / den, np.nan)


def compute_kpis(df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute every picker-level KPI as new columns on top of the cleaned
    G2 dataframe. One row in -> one row out (same grain: user_id x fc_id x date).

    Formulas (see project Notes sheet for the same explanations in plain
    English):

      Login Hours              = logged_in_min / 60
      Utilization %            = Busy_time_min / logged_in_min * 100        (0 if logged_in_min = 0)
      Idle Time (min)          = logged_in_min - Busy_time_min
      Idle %                   = Idle Time / logged_in_min * 100            (0 if logged_in_min = 0)

      Picking Acceptance Rate %  = (Assigned_Picking_jobs - Declined_Picking_jobs) / Assigned_Picking_jobs * 100   (NaN if Assigned_Picking_jobs = 0)
      Picking Decline Rate %     = Declined_Picking_jobs / Assigned_Picking_jobs * 100                              (NaN if Assigned_Picking_jobs = 0)
      Orders Per Hour            = total_order_picked / Login Hours          (0 if Login Hours = 0)
      SKU Per Hour (Picking)     = total_sku_picked / Login Hours
      Qty Per Hour (Picking)     = picked_quantity / Login Hours

      Stacking Acceptance Rate % = (Assigned_Stacking_jobs - Declined_Stacking_jobs) / Assigned_Stacking_jobs * 100  (NaN if Assigned_Stacking_jobs = 0)
      Stacking SKU Per Hour      = Total_sku_stacked / Login Hours
      Stacking Qty Per Hour      = Total_Qty_Stacked / Login Hours

      Bin Audit Acceptance Rate % = (Assigned_BinAudit_jobs - Declined_BinAudit_jobs) / Assigned_BinAudit_jobs * 100 (NaN if Assigned_BinAudit_jobs = 0)
      Bin Audit SKU Per Hour      = Bin_Audit_Scheduled_sku_count / Login Hours

      Empty Binning Acceptance Rate %  = (Assigned_EmptyBinning_jobs - Declined_EmptyBinning_jobs) / Assigned_EmptyBinning_jobs * 100  (NaN if Assigned_EmptyBinning_jobs = 0)
      Empty Binning Completion Rate %  = Completed_EmptyBinning_jobs / Assigned_EmptyBinning_jobs * 100                                 (NaN if Assigned_EmptyBinning_jobs = 0)

      Total Jobs Assigned (Overall) = total_jobs_assigned   (source column, used as-is)
      Total Jobs Declined (Overall) = total_jobs_declined
      Overall Acceptance Rate %     = (total_jobs_assigned - total_jobs_declined) / total_jobs_assigned * 100  (NaN if total_jobs_assigned = 0)
      Total Activity Time (min)     = Picking_Time + Stacking_time + Bin_Audit_time + Empty_Binning_time
      Activity Mix % (Picking/Stacking/BinAudit/EmptyBinning)
                                     = each *_Time / Total Activity Time * 100  (0 if Total Activity Time = 0)

    Note: total_jobs_assigned/total_jobs_declined do NOT always equal the sum
    of the 4 per-activity Assigned_*/Declined_* columns in this dataset (only
    ~42% of rows match exactly) — both are reported as-is, never reconciled
    against each other, per the confirmed data profiling.
    """
    df = df.copy()

    # Login hours & utilization
    df["Login Hours"] = df["logged_in_min"] / 60.0
    df["Utilization %"] = np.where(
        df["logged_in_min"] > 0,
        df["Busy_time_min"] / df["logged_in_min"] * 100,
        0.0,
    )
    df["Idle Time (min)"] = df["logged_in_min"] - df["Busy_time_min"]
    df["Idle %"] = np.where(
        df["logged_in_min"] > 0,
        df["Idle Time (min)"] / df["logged_in_min"] * 100,
        0.0,
    )

    login_hours_safe = df["Login Hours"].replace(0, np.nan)

    # Picking
    df["Picking Acceptance Rate %"] = _safe_div(
        df["Assigned_Picking_jobs"] - df["Declined_Picking_jobs"], df["Assigned_Picking_jobs"]
    ) * 100
    df["Picking Decline Rate %"] = _safe_div(
        df["Declined_Picking_jobs"], df["Assigned_Picking_jobs"]
    ) * 100
    df["Orders Per Hour"] = (df["total_order_picked"] / login_hours_safe).fillna(0)
    df["SKU Per Hour (Picking)"] = (df["total_sku_picked"] / login_hours_safe).fillna(0)
    df["Qty Per Hour (Picking)"] = (df["picked_quantity"] / login_hours_safe).fillna(0)

    # Stacking
    df["Stacking Acceptance Rate %"] = _safe_div(
        df["Assigned_Stacking_jobs"] - df["Declined_Stacking_jobs"], df["Assigned_Stacking_jobs"]
    ) * 100
    df["Stacking SKU Per Hour"] = (df["Total_sku_stacked"] / login_hours_safe).fillna(0)
    df["Stacking Qty Per Hour"] = (df["Total_Qty_Stacked"] / login_hours_safe).fillna(0)

    # Bin Audit
    df["Bin Audit Acceptance Rate %"] = _safe_div(
        df["Assigned_BinAudit_jobs"] - df["Declined_BinAudit_jobs"], df["Assigned_BinAudit_jobs"]
    ) * 100
    df["Bin Audit SKU Per Hour"] = (df["Bin_Audit_Scheduled_sku_count"] / login_hours_safe).fillna(0)

    # Empty Binning
    df["Empty Binning Acceptance Rate %"] = _safe_div(
        df["Assigned_EmptyBinning_jobs"] - df["Declined_EmptyBinning_jobs"], df["Assigned_EmptyBinning_jobs"]
    ) * 100
    df["Empty Binning Completion Rate %"] = _safe_div(
        df["Completed_EmptyBinning_jobs"], df["Assigned_EmptyBinning_jobs"]
    ) * 100

    # Overall (combined across all 4 job types)
    df["Overall Acceptance Rate %"] = _safe_div(
        df["total_jobs_assigned"] - df["total_jobs_declined"], df["total_jobs_assigned"]
    ) * 100
    df["Total Activity Time (min)"] = (
        df["Picking_Time"] + df["Stacking_time"] + df["Bin_Audit_time"] + df["Empty_Binning_time"]
    )
    total_activity_safe = df["Total Activity Time (min)"].replace(0, np.nan)
    df["Activity Mix % - Picking"] = (df["Picking_Time"] / total_activity_safe * 100).fillna(0)
    df["Activity Mix % - Stacking"] = (df["Stacking_time"] / total_activity_safe * 100).fillna(0)
    df["Activity Mix % - BinAudit"] = (df["Bin_Audit_time"] / total_activity_safe * 100).fillna(0)
    df["Activity Mix % - EmptyBinning"] = (df["Empty_Binning_time"] / total_activity_safe * 100).fillna(0)

    # Combined throughput (used by Performance Score) — sum of picking + stacking
    # quantity per hour, since "qty handled" is meaningful across both activities
    df["Combined Qty Per Hour"] = (
        (df["picked_quantity"] + df["Total_Qty_Stacked"]) / login_hours_safe
    ).fillna(0)

    return df


def compute_performance_score(df: pd.DataFrame) -> pd.DataFrame:
    """
    Performance Score (0-100), G2 only. Each component is converted to a
    percentile rank (0-100) within the G2 population before weighting —
    more robust to outliers than min-max scaling given the wide spread in
    this data (e.g. Busy_time_min ranges from 0 to 750+ minutes).

    Weights:
      Utilization %            -> 30%
      Overall Acceptance Rate % -> 20%
      Combined Qty Per Hour    -> 30%
      Orders Per Hour          -> 20%   (picking-only metric — see Notes:
                                          "orders" has no equivalent concept
                                          in stacking/bin-audit/empty-binning)

    Rows with NaN in Overall Acceptance Rate % (total_jobs_assigned = 0)
    contribute 0 to that component rather than being dropped from the score
    entirely, so a picker isn't penalized into having no score at all just
    because the job-assignment pipeline didn't log anything for them that day.
    """
    df = df.copy()

    def pct_rank(series: pd.Series) -> pd.Series:
        # percentile rank 0-100, NaNs excluded from ranking and filled with 0 after
        ranked = series.rank(pct=True, na_option="keep") * 100
        return ranked.fillna(0)

    util_score = pct_rank(df["Utilization %"])
    accept_score = pct_rank(df["Overall Acceptance Rate %"])
    qty_score = pct_rank(df["Combined Qty Per Hour"])
    orders_score = pct_rank(df["Orders Per Hour"])

    df["Performance Score"] = (
        util_score * 0.30
        + accept_score * 0.20
        + qty_score * 0.30
        + orders_score * 0.20
    ).round(1)

    return df


# ──────────────────────────────────────────────────────────────────────────
# Rollups (FC level, City level)
# ──────────────────────────────────────────────────────────────────────────

def rollup_by(df: pd.DataFrame, group_cols: list) -> pd.DataFrame:
    """
    Aggregate picker-level KPIs up to FC or City level. Rollups recompute
    rates from the SUMMED raw counts (not an average-of-percentages), which
    is the statistically correct way to combine rate KPIs across pickers —
    e.g. FC Overall Acceptance Rate % = (sum of assigned - sum of declined) /
    sum of assigned, not the mean of each picker's individual rate.
    """
    agg = df.groupby(group_cols).agg(
        Picker_Count=("user_id", "nunique"),
        Total_Login_Hours=("Login Hours", "sum"),
        Total_Busy_Min=("Busy_time_min", "sum"),
        Total_Login_Min=("logged_in_min", "sum"),
        Total_Jobs_Assigned=("total_jobs_assigned", "sum"),
        Total_Jobs_Declined=("total_jobs_declined", "sum"),
        Total_Orders_Picked=("total_order_picked", "sum"),
        Total_SKU_Picked=("total_sku_picked", "sum"),
        Total_Qty_Picked=("picked_quantity", "sum"),
        Total_Qty_Stacked=("Total_Qty_Stacked", "sum"),
        Total_SKU_Stacked=("Total_sku_stacked", "sum"),
        Assigned_Picking_Jobs=("Assigned_Picking_jobs", "sum"),
        Declined_Picking_Jobs=("Declined_Picking_jobs", "sum"),
        Assigned_Stacking_Jobs=("Assigned_Stacking_jobs", "sum"),
        Declined_Stacking_Jobs=("Declined_Stacking_jobs", "sum"),
        Assigned_BinAudit_Jobs=("Assigned_BinAudit_jobs", "sum"),
        Declined_BinAudit_Jobs=("Declined_BinAudit_jobs", "sum"),
        Avg_Performance_Score=("Performance Score", "mean"),
    ).reset_index()

    agg["Utilization %"] = np.where(
        agg["Total_Login_Min"] > 0, agg["Total_Busy_Min"] / agg["Total_Login_Min"] * 100, 0.0
    )
    agg["Overall Acceptance Rate %"] = _safe_div(
        agg["Total_Jobs_Assigned"] - agg["Total_Jobs_Declined"], agg["Total_Jobs_Assigned"]
    ) * 100
    agg["Picking Acceptance Rate %"] = _safe_div(
        agg["Assigned_Picking_Jobs"] - agg["Declined_Picking_Jobs"], agg["Assigned_Picking_Jobs"]
    ) * 100
    agg["Stacking Acceptance Rate %"] = _safe_div(
        agg["Assigned_Stacking_Jobs"] - agg["Declined_Stacking_Jobs"], agg["Assigned_Stacking_Jobs"]
    ) * 100
    agg["Bin Audit Acceptance Rate %"] = _safe_div(
        agg["Assigned_BinAudit_Jobs"] - agg["Declined_BinAudit_Jobs"], agg["Assigned_BinAudit_Jobs"]
    ) * 100
    login_hours_safe = agg["Total_Login_Hours"].replace(0, np.nan)
    agg["Orders Per Hour"] = (agg["Total_Orders_Picked"] / login_hours_safe).fillna(0)
    agg["Qty Per Hour (Combined)"] = (
        (agg["Total_Qty_Picked"] + agg["Total_Qty_Stacked"]) / login_hours_safe
    ).fillna(0)
    agg["Avg_Performance_Score"] = agg["Avg_Performance_Score"].round(1)

    agg["Rank Score"] = (
        agg["Utilization %"].rank(pct=True) * 30
        + agg["Overall Acceptance Rate %"].fillna(0).rank(pct=True) * 20
        + agg["Qty Per Hour (Combined)"].rank(pct=True) * 30
        + agg["Orders Per Hour"].rank(pct=True) * 20
    ).round(1)

    return agg.sort_values("Rank Score", ascending=False).reset_index(drop=True)


# ──────────────────────────────────────────────────────────────────────────
# Rankings & exception reports
# ──────────────────────────────────────────────────────────────────────────

def top_bottom(df: pd.DataFrame, score_col: str, n: int, label_cols: list) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (top_n, bottom_n) sorted by score_col descending/ascending."""
    ranked = df.sort_values(score_col, ascending=False).reset_index(drop=True)
    top = ranked.head(n).copy()
    bottom = ranked.tail(n).sort_values(score_col, ascending=True).reset_index(drop=True).copy()
    return top, bottom


def exception_reports(df: pd.DataFrame, low_util_threshold: float, high_decline_threshold: float) -> dict:
    """
    Build the three exception reports:

      Low Utilization Pickers  : Utilization % < low_util_threshold (default 50)
      High Decline Pickers     : Overall Acceptance Rate % < high_decline_threshold (default 90)
                                  — excludes rows where total_jobs_assigned = 0
                                  (no jobs assigned means no acceptance signal,
                                  not a decline problem)
      Zero Productivity Pickers: logged_in_min > 0 AND zero jobs assigned across
                                  ALL FOUR job types (picking, stacking, bin-audit,
                                  empty-binning) — corrected definition that does
                                  not penalize pickers who did stacking/bin-audit
                                  work even with zero picking jobs
    """
    low_util = df[df["Utilization %"] < low_util_threshold].sort_values("Utilization %").reset_index(drop=True)

    high_decline = df[
        (df["total_jobs_assigned"] > 0) & (df["Overall Acceptance Rate %"] < high_decline_threshold)
    ].sort_values("Overall Acceptance Rate %").reset_index(drop=True)

    zero_jobs_mask = (
        (df["Assigned_Picking_jobs"] == 0)
        & (df["Assigned_Stacking_jobs"] == 0)
        & (df["Assigned_BinAudit_jobs"] == 0)
        & (df["Assigned_EmptyBinning_jobs"] == 0)
    )
    zero_productivity = df[
        (df["logged_in_min"] > 0) & zero_jobs_mask
    ].sort_values("logged_in_min", ascending=False).reset_index(drop=True)

    return {
        "low_utilization": low_util,
        "high_decline": high_decline,
        "zero_productivity": zero_productivity,
    }


# ──────────────────────────────────────────────────────────────────────────
# Excel output — single workbook, multiple sheets
# ──────────────────────────────────────────────────────────────────────────

HEADER_FMT_PROPS = {
    "bold": True, "font_color": "white", "bg_color": "#1F4E78",
    "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
}

PICKER_DETAIL_COLS = [
    "date", "city", "fc_name", "user_id", "Picker_Name",
    "logged_in_min", "Busy_time_min", "available_time_min", "Login Hours",
    "Utilization %", "Idle Time (min)", "Idle %",
    "Assigned_Picking_jobs", "Declined_Picking_jobs", "Picking Acceptance Rate %", "Picking Decline Rate %",
    "total_order_picked", "total_sku_picked", "picked_quantity", "Orders Per Hour", "SKU Per Hour (Picking)", "Qty Per Hour (Picking)",
    "Assigned_Stacking_jobs", "Declined_Stacking_jobs", "Stacking Acceptance Rate %",
    "Total_sku_stacked", "Total_Qty_Stacked", "Stacking SKU Per Hour", "Stacking Qty Per Hour",
    "Assigned_BinAudit_jobs", "Declined_BinAudit_jobs", "Bin Audit Acceptance Rate %",
    "Bin_Audit_Scheduled_sku_count", "Bin Audit SKU Per Hour",
    "Assigned_EmptyBinning_jobs", "Declined_EmptyBinning_jobs", "Completed_EmptyBinning_jobs",
    "Empty Binning Acceptance Rate %", "Empty Binning Completion Rate %",
    "total_jobs_assigned", "total_jobs_declined", "Overall Acceptance Rate %",
    "Picking_Time", "Stacking_time", "Bin_Audit_time", "Empty_Binning_time", "Total Activity Time (min)",
    "Activity Mix % - Picking", "Activity Mix % - Stacking", "Activity Mix % - BinAudit", "Activity Mix % - EmptyBinning",
    "Combined Qty Per Hour", "Performance Score",
]

PCT_COLS = {
    "Utilization %", "Idle %", "Picking Acceptance Rate %", "Picking Decline Rate %",
    "Stacking Acceptance Rate %", "Bin Audit Acceptance Rate %",
    "Empty Binning Acceptance Rate %", "Empty Binning Completion Rate %",
    "Overall Acceptance Rate %", "Activity Mix % - Picking", "Activity Mix % - Stacking",
    "Activity Mix % - BinAudit", "Activity Mix % - EmptyBinning",
}


def _write_df_sheet(workbook, df: pd.DataFrame, sheet_name: str, fmts: dict, freeze_row=1, autofilter=True):
    """Write a DataFrame to a worksheet with standard formatting: frozen
    header row, autofilter, navy header band, percent format on % columns,
    centered numeric columns, sensible column widths."""
    ws = workbook.add_worksheet(sheet_name[:31])

    if df.empty:
        ws.write(0, 0, f"No data for {sheet_name}")
        return ws

    cols = list(df.columns)
    for c_idx, col_name in enumerate(cols):
        ws.write(0, c_idx, col_name, fmts["header"])

    for r_idx, (_, row) in enumerate(df.iterrows(), start=1):
        for c_idx, col_name in enumerate(cols):
            val = row[col_name]
            is_pct = col_name in PCT_COLS
            if pd.isna(val):
                ws.write_blank(r_idx, c_idx, None, fmts["pct"] if is_pct else fmts["default"])
            elif is_pct:
                ws.write_number(r_idx, c_idx, float(val) / 100.0, fmts["pct"])
            elif isinstance(val, (int, np.integer)):
                ws.write_number(r_idx, c_idx, int(val), fmts["int"])
            elif isinstance(val, (float, np.floating)):
                ws.write_number(r_idx, c_idx, float(val), fmts["float"])
            else:
                ws.write(r_idx, c_idx, str(val), fmts["default"])

    # Column widths — based on header length, capped
    for c_idx, col_name in enumerate(cols):
        width = min(max(len(str(col_name)) + 2, 10), 28)
        ws.set_column(c_idx, c_idx, width)

    ws.freeze_panes(freeze_row, 1)
    if autofilter:
        ws.autofilter(0, 0, len(df), len(cols) - 1)

    # Conditional formatting (3-color scale) on Utilization % / Acceptance Rate % cols
    for c_idx, col_name in enumerate(cols):
        if "Utilization %" in col_name or "Acceptance Rate %" in col_name:
            ws.conditional_format(1, c_idx, len(df), c_idx, {
                "type": "3_color_scale",
                "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B",
            })

    return ws


def write_excel_report(df, fc_scorecard, city_scorecard, top_pickers, bottom_pickers,
                        top_fcs, bottom_fcs, top_cities, bottom_cities,
                        exc_reports, validation_info, output_path: Path, cfg: dict):
    """Write the single picker_efficiency_report.xlsx with all required sheets."""
    workbook = xlsxwriter.Workbook(str(output_path))

    fmts = {
        "header": workbook.add_format(HEADER_FMT_PROPS),
        "default": workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"}),
        "int": workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "#,##0"}),
        "float": workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "#,##0.00"}),
        "pct": workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "0.0%"}),
        "title": workbook.add_format({"bold": True, "font_size": 14, "font_color": "#1F4E78"}),
        "label": workbook.add_format({"bold": True, "border": 1, "bg_color": "#D9EAF7"}),
        "value": workbook.add_format({"border": 1, "align": "center"}),
    }

    # 1. Executive_Summary
    ws = workbook.add_worksheet("Executive_Summary")
    ws.write(0, 0, "Picker Efficiency — Executive Summary (G2 Only)", fmts["title"])
    g2_total = len(df)
    summary_rows = [
        ("Date Range", f"{df['date'].min()} to {df['date'].max()}"),
        ("Total G2 Picker-Days", f"{g2_total:,}"),
        ("Unique G2 Pickers", f"{df['user_id'].nunique():,}"),
        ("Unique FCs", f"{df['fc_name'].nunique():,}"),
        ("Unique Cities", f"{df['city'].nunique():,}"),
        ("Avg Utilization %", f"{df['Utilization %'].mean():.1f}%"),
        ("Avg Overall Acceptance Rate %", f"{df['Overall Acceptance Rate %'].mean():.1f}%"),
        ("Avg Performance Score", f"{df['Performance Score'].mean():.1f}"),
        ("Total Orders Picked", f"{df['total_order_picked'].sum():,.0f}"),
        ("Total Qty Picked", f"{df['picked_quantity'].sum():,.0f}"),
        ("Total Qty Stacked", f"{df['Total_Qty_Stacked'].sum():,.0f}"),
        ("Total Jobs Assigned (Overall)", f"{df['total_jobs_assigned'].sum():,.0f}"),
        ("Total Jobs Declined (Overall)", f"{df['total_jobs_declined'].sum():,.0f}"),
        ("Low Utilization Pickers", f"{len(exc_reports['low_utilization']):,}"),
        ("High Decline Pickers", f"{len(exc_reports['high_decline']):,}"),
        ("Zero Productivity Pickers", f"{len(exc_reports['zero_productivity']):,}"),
    ]
    r = 2
    for label, value in summary_rows:
        ws.write(r, 0, label, fmts["label"])
        ws.write(r, 1, value, fmts["value"])
        r += 1
    ws.set_column(0, 0, 32)
    ws.set_column(1, 1, 24)

    # 2. Picker_Detail
    detail_cols = [c for c in PICKER_DETAIL_COLS if c in df.columns]
    _write_df_sheet(workbook, df[detail_cols], "Picker_Detail", fmts)

    # 3. City_Scorecard
    _write_df_sheet(workbook, city_scorecard, "City_Scorecard", fmts)

    # 4. FC_Scorecard
    _write_df_sheet(workbook, fc_scorecard, "FC_Scorecard", fmts)

    # 5-10. Top/Bottom N
    _write_df_sheet(workbook, top_pickers[detail_cols], "Top10_Pickers", fmts)
    _write_df_sheet(workbook, bottom_pickers[detail_cols], "Bottom10_Pickers", fmts)
    _write_df_sheet(workbook, top_fcs, "Top10_FCs", fmts)
    _write_df_sheet(workbook, bottom_fcs, "Bottom10_FCs", fmts)
    _write_df_sheet(workbook, top_cities, "Top10_Cities", fmts)
    _write_df_sheet(workbook, bottom_cities, "Bottom10_Cities", fmts)

    # 11-13. Exception reports
    _write_df_sheet(workbook, exc_reports["low_utilization"][detail_cols], "Low_Utilization_Pickers", fmts)
    _write_df_sheet(workbook, exc_reports["high_decline"][detail_cols], "High_Decline_Pickers", fmts)
    _write_df_sheet(workbook, exc_reports["zero_productivity"][detail_cols], "Zero_Productivity_Pickers", fmts)

    # 14. Validation_Log
    ws = workbook.add_worksheet("Validation_Log")
    ws.write(0, 0, "Data Validation Log", fmts["title"])
    ws.write(2, 0, "Rows In (raw)", fmts["label"])
    ws.write(2, 1, validation_info["rows_in"], fmts["value"])
    ws.write(3, 0, "Rows Out (cleaned, G2 only)", fmts["label"])
    ws.write(3, 1, validation_info["rows_out"], fmts["value"])
    ws.write(4, 0, "G1 Rows Excluded", fmts["label"])
    ws.write(4, 1, validation_info["g1_excluded"], fmts["value"])
    ws.write(5, 0, "G2 Rows Kept", fmts["label"])
    ws.write(5, 1, validation_info["g2_kept"], fmts["value"])
    ws.write(6, 0, "Report Generated", fmts["label"])
    ws.write(6, 1, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), fmts["value"])
    r = 9
    ws.write(r, 0, "Cleaning Actions Performed:", fmts["title"])
    r += 1
    for action in validation_info["actions"]:
        ws.write(r, 0, f"• {action}")
        r += 1
    ws.set_column(0, 0, 90)
    ws.set_column(1, 1, 20)

    # 15. Notes
    ws = workbook.add_worksheet("Notes")
    ws.write(0, 0, "Picker Efficiency Report — Formulas & Cleaning Notes", fmts["title"])
    notes = [
        "SCOPE: This report covers G2 staff only. G1 is excluded — G1's job-assignment "
        "tracking columns (Assigned_*/Declined_* for all 4 activity types) are 0 for "
        "every row even when G1 shows real picking/stacking/bin-audit output, so G1 has "
        "no usable Acceptance/Decline Rate signal and is out of scope for this automation.",
        "",
        "KPI FORMULAS",
        "Login Hours = logged_in_min / 60",
        "Utilization % = Busy_time_min / logged_in_min  (0% if not logged in)",
        "Idle Time (min) = logged_in_min - Busy_time_min",
        "Picking Acceptance Rate % = (Assigned_Picking_jobs - Declined_Picking_jobs) / Assigned_Picking_jobs"
        "  (blank if no picking jobs were assigned that day — not 0%, genuinely not applicable)",
        "Orders/SKU/Qty Per Hour = respective total / Login Hours",
        "Same Acceptance Rate pattern applies to Stacking, Bin Audit, and Empty Binning.",
        "Overall Acceptance Rate % uses the source total_jobs_assigned/total_jobs_declined "
        "columns directly — these do NOT always equal the sum of the 4 per-activity columns "
        "in the source data, so both are reported independently rather than reconciled.",
        "Total Activity Time = Picking_Time + Stacking_time + Bin_Audit_time + Empty_Binning_time",
        "Activity Mix % = each activity's time / Total Activity Time — shows how a picker's "
        "day split across the 4 task types.",
        "",
        "PERFORMANCE SCORE (0-100)",
        "Each picker is percentile-ranked (0-100) within the G2 population on 4 components, "
        "then weighted: Utilization % (30%), Overall Acceptance Rate % (20%), Combined Qty "
        "Per Hour = picked_quantity + Total_Qty_Stacked, summed, over Login Hours (30%), "
        "Orders Per Hour (20%). Percentile ranking is used instead of min-max scaling because "
        "it is far less sensitive to extreme outliers in this dataset.",
        "LIMITATION: Orders Per Hour only reflects picking activity — stacking, bin-audit, "
        "and empty-binning do not have an equivalent 'order' concept in this data.",
        "",
        "DATA CLEANING",
        "Nulls in picking/stacking/bin-audit output columns were filled with 0 — a null in "
        "this dataset means the picker did zero of that activity that day, not missing data.",
        "Avg_Triggered_Bin_Audit_time_Per_Sku was dropped entirely (100% null in source file).",
        "Picker identity is keyed by user_id, not Picker_Name — the same name can appear "
        "under multiple different user_ids in this dataset.",
        "",
        "EXCEPTION REPORT DEFINITIONS",
        f"Low Utilization Pickers: Utilization % < {cfg['low_utilization_threshold']}% (configurable in config.json)",
        f"High Decline Pickers: Overall Acceptance Rate % < {cfg['high_decline_threshold']}% "
        "AND total_jobs_assigned > 0 (pickers with zero jobs assigned are excluded — no "
        "acceptance signal to flag)",
        "Zero Productivity Pickers: logged in (logged_in_min > 0) AND zero jobs assigned "
        "across ALL FOUR activity types (picking, stacking, bin-audit, empty-binning) — a "
        "picker who did stacking or bin-audit work is NOT flagged even with zero picking jobs.",
    ]
    r = 2
    for line in notes:
        ws.write(r, 0, line)
        r += 1
    ws.set_column(0, 0, 110)

    workbook.close()
    print(f"  [Excel] Written: {output_path}")


# ──────────────────────────────────────────────────────────────────────────
# HTML dashboard
# ──────────────────────────────────────────────────────────────────────────

def _r(series, ndigits=1):
    """Round a pandas Series to ndigits, replacing NaN with 0, return list."""
    return series.fillna(0).round(ndigits).tolist()


def _build_dashboard_data(df, fc_scorecard, city_scorecard,
                           top_pickers, bottom_pickers,
                           top_fcs, bottom_fcs, top_cities, bottom_cities,
                           exc_reports, cfg) -> dict:
    """
    Assemble the compact JSON payload embedded into the HTML dashboard.
    Keys are deliberately short (id/nm/ut/ar/...) to keep the single-file
    HTML small with 10k+ picker rows embedded client-side.
    """
    # ---- pickers (full grain, compact keys) ----
    pickers = pd.DataFrame({
        "id": df["user_id"].astype(str),
        "nm": df["Picker_Name"].astype(str),
        "ci": df["city"].astype(str),
        "fc": df["fc_name"].astype(str),
        "dt": df["date"].astype(str),
        "lg": _r(df["logged_in_min"], 0),
        "ut": _r(df["Utilization %"]),
        "ar": _r(df["Overall Acceptance Rate %"]),
        "oh": _r(df["Orders Per Hour"]),
        "qh": _r(df["Combined Qty Per Hour"]),
        "ps": _r(df["Performance Score"]),
        "mp": _r(df["Activity Mix % - Picking"]),
        "ms": _r(df["Activity Mix % - Stacking"]),
        "mb": _r(df["Activity Mix % - BinAudit"]),
        "me": _r(df["Activity Mix % - EmptyBinning"]),
    }).to_dict(orient="records")

    # ---- full FC list (for explorer cards under a city) ----
    fcs = pd.DataFrame({
        "fc": fc_scorecard["fc_name"].astype(str),
        "ci": df.groupby("fc_name")["city"].first().reindex(fc_scorecard["fc_name"]).astype(str).values,
        "pc": fc_scorecard["Picker_Count"],
        "ut": _r(fc_scorecard["Utilization %"]),
        "rk": _r(fc_scorecard["Rank Score"]),
    }).to_dict(orient="records")

    # ---- full city list (for treemap / explorer top level) ----
    cities = pd.DataFrame({
        "ci": city_scorecard["city"].astype(str),
        "pc": city_scorecard["Picker_Count"],
        "ut": _r(city_scorecard["Utilization %"]),
        "rk": _r(city_scorecard["Rank Score"]),
    }).to_dict(orient="records")

    def _rank_rows(rdf, key_col, key_short):
        return pd.DataFrame({
            key_short: rdf[key_col].astype(str),
            "ci": rdf["city"].astype(str) if "city" in rdf.columns else "",
            "pc": rdf["Picker_Count"] if "Picker_Count" in rdf.columns else None,
            "ut": _r(rdf["Utilization %"]),
            "rk": _r(rdf["Rank Score"]),
        }).to_dict(orient="records")

    def _picker_rank_rows(rdf):
        return pd.DataFrame({
            "id": rdf["user_id"].astype(str),
            "nm": rdf["Picker_Name"].astype(str),
            "ci": rdf["city"].astype(str),
            "fc": rdf["fc_name"].astype(str),
            "ut": _r(rdf["Utilization %"]),
            "ar": _r(rdf["Overall Acceptance Rate %"]),
            "ps": _r(rdf["Performance Score"]),
        }).to_dict(orient="records")

    top_fcs_j = _rank_rows(top_fcs, "fc_name", "fc")
    bottom_fcs_j = _rank_rows(bottom_fcs, "fc_name", "fc")
    top_cities_j = _rank_rows(top_cities, "city", "ci")
    bottom_cities_j = _rank_rows(bottom_cities, "city", "ci")

    def _exc_rows(rdf):
        return pd.DataFrame({
            "id": rdf["user_id"].astype(str),
            "nm": rdf["Picker_Name"].astype(str),
            "ci": rdf["city"].astype(str),
            "fc": rdf["fc_name"].astype(str),
            "ut": _r(rdf["Utilization %"]),
            "ar": _r(rdf["Overall Acceptance Rate %"]),
            "lg": _r(rdf["logged_in_min"], 0),
        }).to_dict(orient="records")

    kpis = {
        "totalPickers": int(df["user_id"].nunique()),
        "totalFCs": int(df["fc_name"].nunique()),
        "totalCities": int(df["city"].nunique()),
        "avgUtilization": round(float(df["Utilization %"].mean()), 1),
        "avgAcceptance": round(float(df["Overall Acceptance Rate %"].mean(skipna=True)), 1),
        "avgScore": round(float(df["Performance Score"].mean()), 1),
        "totalOrders": int(df["total_order_picked"].sum()),
        "totalJobsAssigned": int(df["total_jobs_assigned"].sum()),
        "totalJobsDeclined": int(df["total_jobs_declined"].sum()),
        "lowUtilCount": int(len(exc_reports["low_utilization"])),
        "highDeclineCount": int(len(exc_reports["high_decline"])),
        "zeroProdCount": int(len(exc_reports["zero_productivity"])),
        "dateMin": str(df["date"].min()),
        "dateMax": str(df["date"].max()),
    }

    return {
        "kpis": kpis,
        "config": {
            "lowUtilThreshold": cfg["low_utilization_threshold"],
            "highDeclineThreshold": cfg["high_decline_threshold"],
        },
        "pickers": pickers,
        "fcs": fcs,
        "cities": cities,
        "topPickers": _picker_rank_rows(top_pickers),
        "bottomPickers": _picker_rank_rows(bottom_pickers),
        "topFCs": top_fcs_j,
        "bottomFCs": bottom_fcs_j,
        "topCities": top_cities_j,
        "bottomCities": bottom_cities_j,
        "lowUtil": _exc_rows(exc_reports["low_utilization"]),
        "highDecline": _exc_rows(exc_reports["high_decline"]),
        "zeroProd": _exc_rows(exc_reports["zero_productivity"]),
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def write_dashboard(data: dict, output_path: Path):
    """Render the self-contained HTML dashboard by injecting the JSON
    payload into the _DASHBOARD_TEMPLATE string in place of the
    __DASHBOARD_DATA__ marker."""
    payload = json.dumps(data, separators=(",", ":"), default=str)
    # Defend against a stray "</script>" inside any embedded string (e.g. a
    # picker/FC/city name) prematurely closing the script block and breaking
    # the page. Safe to do after json.dumps since '<' is never meaningful
    # JSON syntax.
    payload = payload.replace("</script", "<\\/script")
    html = _DASHBOARD_TEMPLATE.replace("__DASHBOARD_DATA__", payload)
    output_path.write_text(html, encoding="utf-8")
    size_mb = output_path.stat().st_size / (1024 * 1024)
    print(f"  [Dashboard] Written: {output_path} ({size_mb:.1f} MB)")


# ──────────────────────────────────────────────────────────────────────────
# Google Sheets export (optional — opt in via config.json: "push_to_google_sheets": true)
# ──────────────────────────────────────────────────────────────────────────

def push_to_google_sheets(fc_scorecard, city_scorecard, top_pickers, bottom_pickers, base_dir: Path):
    """
    Push the FC Scorecard, City Scorecard, and Top/Bottom Pickers sheets to a
    Google Sheet. Requires:
      - google_credentials.json (OAuth client secret from Google Cloud Console)
      - google_sheet_id.txt (the target spreadsheet's ID, from its URL)
    token.json is created automatically on first run after you complete the
    browser OAuth consent flow, and reused (refreshed) on every run after.

    Silently skipped (with a clear message) if gspread isn't installed, or
    credentials/sheet ID aren't configured yet — this never blocks the
    Excel/dashboard outputs.
    """
    if not GSPREAD_AVAILABLE:
        print("  [Sheets] Skipped: gspread / google-auth-oauthlib not installed "
              "(see requirements.txt — pip install -r requirements.txt).")
        return

    creds_path = base_dir / "google_credentials.json"
    sheet_id_path = base_dir / "google_sheet_id.txt"
    token_path = base_dir / "token.json"

    if not creds_path.exists():
        print("  [Sheets] Skipped: google_credentials.json not found. "
              "See README.txt for setup steps.")
        return
    sheet_id = sheet_id_path.read_text(encoding="utf-8").strip() if sheet_id_path.exists() else ""
    if not sheet_id or sheet_id.startswith("REPLACE"):
        print("  [Sheets] Skipped: google_sheet_id.txt not set. "
              "Paste your spreadsheet ID (from its URL) into that file.")
        return

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = None
    if token_path.exists():
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        creds = Credentials.from_authorized_user_file(str(token_path), scopes)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            token_path.write_text(creds.to_json(), encoding="utf-8")

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), scopes)
        creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
        print("  [Sheets] First-time login complete — token.json saved for future runs.")

    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)

    def _push(df, tab_name):
        try:
            ws = sh.worksheet(tab_name)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=tab_name, rows=str(len(df) + 10), cols=str(len(df.columns) + 2))
        ws.update([df.columns.tolist()] + df.astype(object).where(df.notna(), "").values.tolist())
        print(f"  [Sheets] Pushed '{tab_name}' ({len(df):,} rows)")

    _push(fc_scorecard, "FC Scorecard")
    _push(city_scorecard, "City Scorecard")
    _push(top_pickers, "Top Pickers")
    _push(bottom_pickers, "Bottom Pickers")
    print(f"  [Sheets] Done: https://docs.google.com/spreadsheets/d/{sheet_id}")


# ──────────────────────────────────────────────────────────────────────────
# GitHub Pages deployment (optional — opt in via config.json: "deploy_to_github_pages": true)
# ──────────────────────────────────────────────────────────────────────────

def deploy_to_github_pages(dashboard_path: Path, base_dir: Path):
    """
    Copy the generated dashboard into a local GitHub Pages repo checkout and
    git add/commit/push it, so the live dashboard URL updates automatically.

    Requires github_pages_path.txt to contain the absolute path to a local
    clone of the gh-pages (or main, for a *.github.io repo) branch that's
    already configured with a remote and credentials (e.g. via `gh auth
    login` or an SSH key) — this script does not configure git auth itself.

    Silently skipped (with a clear message) if that file is missing/empty,
    or the path doesn't exist, or any git command fails — this never blocks
    the Excel/dashboard outputs.
    """
    import shutil
    import subprocess

    path_file = base_dir / "github_pages_path.txt"
    if not path_file.exists():
        print("  [GitHub Pages] Skipped: github_pages_path.txt not found.")
        return
    repo_path_str = path_file.read_text(encoding="utf-8").strip()
    if not repo_path_str or repo_path_str.startswith("REPLACE"):
        print("  [GitHub Pages] Skipped: github_pages_path.txt not set. "
              "Paste the absolute path to your local Pages repo checkout into that file.")
        return

    repo_path = Path(repo_path_str)
    if not repo_path.exists() or not (repo_path / ".git").exists():
        print(f"  [GitHub Pages] Skipped: '{repo_path}' doesn't exist or isn't a git repo.")
        return

    try:
        dest = repo_path / "index.html"
        shutil.copy2(dashboard_path, dest)

        def _run(cmd):
            return subprocess.run(cmd, cwd=repo_path, capture_output=True, text=True, check=True)

        _run(["git", "add", "index.html"])
        status = subprocess.run(["git", "status", "--porcelain"], cwd=repo_path,
                                 capture_output=True, text=True).stdout.strip()
        if not status:
            print("  [GitHub Pages] Skipped commit: dashboard unchanged since last deploy.")
            return
        commit_msg = f"Update picker efficiency dashboard ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
        _run(["git", "commit", "-m", commit_msg])
        _run(["git", "push"])
        print(f"  [GitHub Pages] Deployed: copied, committed, and pushed from {repo_path}")
    except subprocess.CalledProcessError as e:
        print(f"  [GitHub Pages] Skipped: git command failed ({e.cmd}): {e.stderr.strip()[:200]}")
    except Exception as e:
        print(f"  [GitHub Pages] Skipped: {e}")


# ──────────────────────────────────────────────────────────────────────────
# Orchestration
# ──────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("Picker Efficiency Automation — G2 Only")
    print("=" * 70)

    cfg = load_config()
    base_dir = Path(__file__).resolve().parent
    input_folder = base_dir / cfg["input_folder"]
    output_folder = base_dir / cfg["output_folder"]
    output_folder.mkdir(parents=True, exist_ok=True)

    print("\n[1/7] Loading input files...")
    df_raw = load_input_files(input_folder)

    print("\n[2/7] Cleaning data...")
    df, validation_info = clean_data(df_raw)

    print("\n[3/7] Computing KPIs & Performance Score...")
    df = compute_kpis(df)
    df = compute_performance_score(df)

    if cfg.get("cities"):
        df = df[df["city"].isin(cfg["cities"])].reset_index(drop=True)
        print(f"  [Filter] Restricted to cities: {cfg['cities']} -> {len(df):,} rows")

    print("\n[4/7] Building rollups, rankings, and exception reports...")
    fc_scorecard = rollup_by(df, ["fc_name"])
    city_scorecard = rollup_by(df, ["city"])

    detail_cols = [c for c in PICKER_DETAIL_COLS if c in df.columns]
    top_pickers, bottom_pickers = top_bottom(df, "Performance Score", cfg["top_n"], detail_cols)
    top_fcs, bottom_fcs = top_bottom(fc_scorecard, "Rank Score", cfg["top_n"], [])
    top_cities, bottom_cities = top_bottom(city_scorecard, "Rank Score", cfg["top_n"], [])

    exc_reports = exception_reports(df, cfg["low_utilization_threshold"], cfg["high_decline_threshold"])

    print(f"  FC Scorecard: {len(fc_scorecard)} FCs")
    print(f"  City Scorecard: {len(city_scorecard)} cities")
    print(f"  Low Utilization Pickers: {len(exc_reports['low_utilization']):,}")
    print(f"  High Decline Pickers: {len(exc_reports['high_decline']):,}")
    print(f"  Zero Productivity Pickers: {len(exc_reports['zero_productivity']):,}")

    print("\n[5/7] Writing Excel report...")
    excel_path = output_folder / cfg["output_file"]
    write_excel_report(
        df, fc_scorecard, city_scorecard, top_pickers, bottom_pickers,
        top_fcs, bottom_fcs, top_cities, bottom_cities,
        exc_reports, validation_info, excel_path, cfg,
    )

    print("\n[6/7] Building HTML dashboard...")
    dash_data = _build_dashboard_data(
        df, fc_scorecard, city_scorecard, top_pickers, bottom_pickers,
        top_fcs, bottom_fcs, top_cities, bottom_cities, exc_reports, cfg,
    )
    dashboard_path = output_folder / cfg["dashboard_file"]
    write_dashboard(dash_data, dashboard_path)

    print("\n[7/7] Optional integrations...")
    if cfg.get("push_to_google_sheets"):
        push_to_google_sheets(fc_scorecard, city_scorecard, top_pickers, bottom_pickers, base_dir)
    else:
        print("  [Sheets] Skipped: push_to_google_sheets is false in config.json.")

    if cfg.get("deploy_to_github_pages"):
        deploy_to_github_pages(dashboard_path, base_dir)
    else:
        print("  [GitHub Pages] Skipped: deploy_to_github_pages is false in config.json.")

    print("\n" + "=" * 70)
    print("Done.")
    print(f"  Excel:     {excel_path}")
    print(f"  Dashboard: {dashboard_path}")
    print("=" * 70)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print("\nERROR:", exc)
        sys.exit(1)