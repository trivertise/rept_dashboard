[sales_dashboard (5).html](https://github.com/user-attachments/files/28343902/sales_dashboard.5.html)
# rept_dashboard
Reporting dashboard for trivertise media 
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Campaign Performance Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap" rel="stylesheet"/>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f0f4f8;
  --white:#ffffff;
  --border:#e2e8f0;
  --radius:16px;
  --shadow:0 2px 12px rgba(0,0,0,.07),0 1px 3px rgba(0,0,0,.05);
  --shadow-hover:0 8px 32px rgba(0,0,0,.13),0 2px 8px rgba(0,0,0,.09);
  --blue:#2563eb;--blue-tint:#eff6ff;
  --teal:#0891b2;--teal-tint:#ecfeff;
  --green:#059669;--green-tint:#ecfdf5;
  --amber:#d97706;--amber-tint:#fffbeb;
  --violet:#7c3aed;--violet-tint:#f5f3ff;
  --rose:#e11d48;--rose-tint:#fff1f2;
  --text:#0f172a;--text-2:#475569;--text-3:#94a3b8;
}
html{font-family:'Poppins',sans-serif;background:var(--bg);color:var(--text);font-size:15px}
body{min-height:100vh;padding:0 0 72px}

/* HEADER */
.header{text-align:center;padding:52px 24px 0}
.eyebrow{display:inline-flex;align-items:center;gap:8px;background:var(--blue-tint);border:1px solid #bfdbfe;border-radius:999px;padding:6px 16px;font-size:.72rem;font-weight:600;color:var(--blue);letter-spacing:.07em;text-transform:uppercase;margin-bottom:18px}
.eyebrow-dot{width:7px;height:7px;border-radius:50%;background:var(--blue);animation:blink 2s infinite}
@keyframes blink{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.4;transform:scale(.75)}}
.header h1{font-size:2.5rem;font-weight:800;line-height:1.15;letter-spacing:-.025em;margin-bottom:10px}
.header h1 .accent{color:var(--blue)}
.header p{color:var(--text-2);font-size:.95rem;max-width:500px;margin:0 auto 36px;line-height:1.65}
.divider{width:60px;height:4px;margin:0 auto;border-radius:999px;background:linear-gradient(90deg,var(--blue),var(--teal))}

/* UPLOAD */
.upload-wrap{max-width:540px;margin:48px auto 0;padding:0 24px}
.upload-card{
  background:var(--white);border:2px dashed var(--border);border-radius:var(--radius);
  padding:52px 36px;text-align:center;cursor:pointer;
  transition:border-color .2s,box-shadow .2s,transform .2s,background .2s;
}
.upload-card:hover,.upload-card.drag{border-color:var(--blue);background:var(--blue-tint);box-shadow:0 0 0 4px #bfdbfe38;transform:translateY(-2px)}
.upload-icon{width:60px;height:60px;margin:0 auto 18px;background:var(--blue-tint);border-radius:50%;display:flex;align-items:center;justify-content:center}
.upload-icon svg{width:26px;height:26px;stroke:var(--blue)}
.upload-card h3{font-size:1.15rem;font-weight:700;margin-bottom:8px}
.upload-card p{color:var(--text-2);font-size:.875rem;margin-bottom:22px;line-height:1.65}
.btn-primary{display:inline-flex;align-items:center;gap:8px;background:var(--blue);color:#fff;border:none;border-radius:999px;padding:11px 24px;font-family:'Poppins',sans-serif;font-size:.85rem;font-weight:600;cursor:pointer;transition:background .2s,transform .15s,box-shadow .2s;box-shadow:0 4px 14px #2563eb38}
.btn-primary:hover{background:#1d4ed8;transform:translateY(-1px);box-shadow:0 6px 20px #2563eb48}
#file-input{display:none}
.format-hint{margin-top:14px;font-size:.75rem;color:var(--text-3)}

/* DASHBOARD */
.dashboard{display:none;max-width:1320px;margin:0 auto;padding:0 24px}

/* STATUS BAR */
.status-bar{display:flex;align-items:center;justify-content:space-between;margin:38px 0 28px;flex-wrap:wrap;gap:12px}
.record-pill{display:inline-flex;align-items:center;gap:8px;background:var(--white);border:1px solid var(--border);border-radius:999px;padding:7px 16px;font-size:.78rem;font-weight:600;color:var(--text-2);box-shadow:var(--shadow)}
.live-dot{width:8px;height:8px;border-radius:50%;background:var(--green);animation:livepulse 2s infinite}
@keyframes livepulse{0%{box-shadow:0 0 0 0 #05966940}70%{box-shadow:0 0 0 8px transparent}100%{box-shadow:0 0 0 0 transparent}}
.record-count{color:var(--blue);font-weight:800}
.btn-reset{display:inline-flex;align-items:center;gap:6px;background:var(--white);color:var(--text-2);border:1px solid var(--border);border-radius:999px;padding:7px 16px;font-family:'Poppins',sans-serif;font-size:.78rem;font-weight:600;cursor:pointer;transition:all .2s;box-shadow:var(--shadow)}
.btn-reset:hover{border-color:var(--rose);color:var(--rose);background:var(--rose-tint)}

/* KPI GRID */
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:22px}
@media(max-width:860px){.kpi-grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:480px){.kpi-grid{grid-template-columns:1fr}}
.kpi-card{
  background:var(--white);border-radius:var(--radius);padding:22px 20px;
  border-left:4px solid transparent;box-shadow:var(--shadow);
  transition:transform .2s,box-shadow .2s;
}
.kpi-card:hover{transform:translateY(-4px);box-shadow:var(--shadow-hover)}
.kpi-card[data-c="blue"]{border-left-color:var(--blue);background:linear-gradient(135deg,#fff 55%,var(--blue-tint))}
.kpi-card[data-c="teal"]{border-left-color:var(--teal);background:linear-gradient(135deg,#fff 55%,var(--teal-tint))}
.kpi-card[data-c="green"]{border-left-color:var(--green);background:linear-gradient(135deg,#fff 55%,var(--green-tint))}
.kpi-card[data-c="amber"]{border-left-color:var(--amber);background:linear-gradient(135deg,#fff 55%,var(--amber-tint))}
.kpi-card[data-c="violet"]{border-left-color:var(--violet);background:linear-gradient(135deg,#fff 55%,var(--violet-tint))}
.kpi-card[data-c="rose"]{border-left-color:var(--rose);background:linear-gradient(135deg,#fff 55%,var(--rose-tint))}
.kpi-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
.kpi-label{font-size:.7rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--text-2)}
.kpi-icon{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center}
.kpi-icon svg{width:17px;height:17px;stroke-width:2;fill:none;stroke-linecap:round;stroke-linejoin:round}
.kpi-card[data-c="blue"] .kpi-icon{background:var(--blue-tint)} .kpi-card[data-c="blue"] .kpi-icon svg{stroke:var(--blue)}
.kpi-card[data-c="teal"] .kpi-icon{background:var(--teal-tint)} .kpi-card[data-c="teal"] .kpi-icon svg{stroke:var(--teal)}
.kpi-card[data-c="green"] .kpi-icon{background:var(--green-tint)} .kpi-card[data-c="green"] .kpi-icon svg{stroke:var(--green)}
.kpi-card[data-c="amber"] .kpi-icon{background:var(--amber-tint)} .kpi-card[data-c="amber"] .kpi-icon svg{stroke:var(--amber)}
.kpi-card[data-c="violet"] .kpi-icon{background:var(--violet-tint)} .kpi-card[data-c="violet"] .kpi-icon svg{stroke:var(--violet)}
.kpi-card[data-c="rose"] .kpi-icon{background:var(--rose-tint)} .kpi-card[data-c="rose"] .kpi-icon svg{stroke:var(--rose)}
.kpi-value{font-size:1.8rem;font-weight:800;letter-spacing:-.025em;line-height:1}
.kpi-card[data-c="blue"] .kpi-value{color:var(--blue)}
.kpi-card[data-c="teal"] .kpi-value{color:var(--teal)}
.kpi-card[data-c="green"] .kpi-value{color:var(--green)}
.kpi-card[data-c="amber"] .kpi-value{color:var(--amber)}
.kpi-card[data-c="violet"] .kpi-value{color:var(--violet)}
.kpi-card[data-c="rose"] .kpi-value{color:var(--rose)}
.kpi-sub{font-size:.73rem;color:var(--text-3);margin-top:5px}

/* CHART ROWS */
.row-2{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}
.row-3{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-bottom:20px}
@media(max-width:860px){.row-2,.row-3{grid-template-columns:1fr}}
.chart-card{background:var(--white);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;transition:transform .2s,box-shadow .2s}
.chart-card:hover{transform:translateY(-3px);box-shadow:var(--shadow-hover)}
.chart-header{padding:20px 22px 0}
.chart-title{font-size:.875rem;font-weight:700;color:var(--text)}
.chart-sub{font-size:.72rem;color:var(--text-3);margin-top:2px}
.chart-body{padding:0 4px 8px}

/* PROJECT BADGE (multi-project) */
.project-tabs{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px}
.proj-badge{display:inline-flex;align-items:center;gap:6px;border-radius:999px;padding:5px 14px;font-size:.75rem;font-weight:600;border:1.5px solid transparent;cursor:pointer;transition:all .15s}
.proj-badge.active,.proj-badge:hover{border-color:currentColor;opacity:1}
.proj-badge{opacity:.6}
.proj-dot{width:8px;height:8px;border-radius:50%}

/* ANIMATIONS */
@keyframes fadeUp{from{opacity:0;transform:translateY(22px)}to{opacity:1;transform:translateY(0)}}
.fu{opacity:0;animation:fadeUp .5s ease forwards}

/* ── TAB NAV ── */
.tab-nav{display:none;max-width:1320px;margin:32px auto 0;padding:0 24px}
.tab-nav-inner{display:flex;gap:4px;background:var(--white);border:1px solid var(--border);border-radius:12px;padding:4px;box-shadow:var(--shadow);width:fit-content}
.tab-btn{display:inline-flex;align-items:center;gap:8px;padding:9px 20px;border-radius:9px;border:none;background:transparent;font-family:'Poppins',sans-serif;font-size:.82rem;font-weight:600;color:var(--text-2);cursor:pointer;transition:all .18s}
.tab-btn:hover{background:var(--bg);color:var(--text)}
.tab-btn.active{background:var(--blue);color:#fff;box-shadow:0 2px 8px #2563eb40}
.tab-btn svg{width:15px;height:15px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}

/* ── DOWNLOAD TAB ── */
.dl-tab{display:none;max-width:1320px;margin:0 auto;padding:0 24px 72px}
.dl-hero{background:var(--white);border-radius:var(--radius);box-shadow:var(--shadow);padding:40px;margin-bottom:20px;display:flex;align-items:center;gap:32px;flex-wrap:wrap}
.dl-hero-icon{width:72px;height:72px;min-width:72px;border-radius:18px;background:linear-gradient(135deg,var(--green-tint),#d1fae5);display:flex;align-items:center;justify-content:center;box-shadow:0 4px 16px rgba(5,150,105,.15)}
.dl-hero-icon svg{width:32px;height:32px;stroke:var(--green);fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
.dl-hero-text h2{font-size:1.35rem;font-weight:800;margin-bottom:6px}
.dl-hero-text p{color:var(--text-2);font-size:.9rem;line-height:1.6;max-width:540px}
.dl-btn{display:inline-flex;align-items:center;gap:10px;background:var(--green);color:#fff;border:none;border-radius:999px;padding:13px 28px;font-family:'Poppins',sans-serif;font-size:.9rem;font-weight:700;cursor:pointer;transition:all .2s;box-shadow:0 4px 16px rgba(5,150,105,.3);white-space:nowrap;margin-top:4px}
.dl-btn:hover{background:#047857;transform:translateY(-2px);box-shadow:0 8px 24px rgba(5,150,105,.35)}
.dl-btn:active{transform:translateY(0)}
.dl-btn svg{width:17px;height:17px;stroke:currentColor;fill:none;stroke-width:2.5;stroke-linecap:round;stroke-linejoin:round}
.dl-btn.loading{opacity:.7;pointer-events:none}

.dl-sheets{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:20px}
@media(max-width:720px){.dl-sheets{grid-template-columns:1fr}}
.sheet-card{background:var(--white);border-radius:var(--radius);box-shadow:var(--shadow);padding:22px 20px;border-top:3px solid transparent;transition:transform .2s,box-shadow .2s}
.sheet-card:hover{transform:translateY(-3px);box-shadow:var(--shadow-hover)}
.sheet-card[data-sc="blue"]{border-top-color:var(--blue)}
.sheet-card[data-sc="teal"]{border-top-color:var(--teal)}
.sheet-card[data-sc="violet"]{border-top-color:var(--violet)}
.sheet-icon{width:38px;height:38px;border-radius:10px;display:flex;align-items:center;justify-content:center;margin-bottom:12px}
.sheet-card[data-sc="blue"] .sheet-icon{background:var(--blue-tint)}
.sheet-card[data-sc="teal"] .sheet-icon{background:var(--teal-tint)}
.sheet-card[data-sc="violet"] .sheet-icon{background:var(--violet-tint)}
.sheet-icon svg{width:18px;height:18px;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
.sheet-card[data-sc="blue"] .sheet-icon svg{stroke:var(--blue)}
.sheet-card[data-sc="teal"] .sheet-icon svg{stroke:var(--teal)}
.sheet-card[data-sc="violet"] .sheet-icon svg{stroke:var(--violet)}
.sheet-name{font-size:.85rem;font-weight:700;margin-bottom:4px}
.sheet-desc{font-size:.78rem;color:var(--text-2);line-height:1.55}
.sheet-cols{margin-top:10px;display:flex;flex-wrap:wrap;gap:5px}
.col-tag{background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:2px 8px;font-size:.68rem;font-weight:600;color:var(--text-2)}

.dl-note{background:var(--amber-tint);border:1px solid #fde68a;border-radius:12px;padding:14px 18px;font-size:.8rem;color:#92400e;display:flex;gap:10px;align-items:flex-start}
.dl-note svg{min-width:16px;width:16px;height:16px;stroke:#d97706;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;margin-top:1px}

</style>
</head>
<body>

<div class="header">
  <div class="eyebrow"><span class="eyebrow-dot"></span>Analytics Suite</div>
  <h1>Campaign <span class="accent">Performance</span> Dashboard</h1>
  <p>Upload your ad campaign JSON to explore impressions, clicks, spends, CTR, and lead delivery over time.</p>
  <div class="divider"></div>
</div>


<div class="tab-nav" id="tab-nav">
  <div class="tab-nav-inner">
    <button class="tab-btn active" id="tab-dashboard" onclick="switchTab('dashboard')">
      <svg viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>
      Dashboard
    </button>
    <button class="tab-btn" id="tab-download" onclick="switchTab('download')">
      <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="12" x2="12" y2="18"/><polyline points="9 15 12 18 15 15"/></svg>
      Download Report
    </button>
  </div>
</div>

<div class="upload-wrap" id="upload-section">
  <div class="upload-card" id="drop-zone">
    <div class="upload-icon">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
        <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
      </svg>
    </div>
    <h3>Drop your JSON file here</h3>
    <p>Supports the campaign export format with Date, Impressions, Clicks, CTR%, Leads Delivered, and Spends columns.</p>
    <button class="btn-primary" onclick="document.getElementById('file-input').click()">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
      Choose JSON File
    </button>
    <input type="file" id="file-input" accept=".json"/>
    <p class="format-hint">JSON array export — campaign rows with Date · Impressions · Clicks · CTR · Leads · Spends</p>
  </div>
</div>

<div class="dashboard" id="dashboard">
  <div class="status-bar">
    <div class="record-pill">
      <span class="live-dot"></span>
      <span><span class="record-count" id="rec-count">0</span> data rows loaded</span>
    </div>
    <button class="btn-reset" onclick="resetDashboard()">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3.51"/></svg>
      Upload New File
    </button>
  </div>

  <!-- project filter badges -->
  <div class="project-tabs" id="proj-tabs"></div>

  <!-- KPIs -->
  <div class="kpi-grid" id="kpi-grid">
    <div class="kpi-card fu" data-c="blue" style="animation-delay:.05s">
      <div class="kpi-header"><span class="kpi-label">Total Impressions</span>
        <div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg></div></div>
      <div class="kpi-value" id="kpi-imp">—</div><div class="kpi-sub">Total ad views delivered</div>
    </div>
    <div class="kpi-card fu" data-c="teal" style="animation-delay:.1s">
      <div class="kpi-header"><span class="kpi-label">Total Clicks</span>
        <div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4M10 17l5-5-5-5M13 12H3"/></svg></div></div>
      <div class="kpi-value" id="kpi-clk">—</div><div class="kpi-sub">Clicks across all campaigns</div>
    </div>
    <div class="kpi-card fu" data-c="green" style="animation-delay:.15s">
      <div class="kpi-header"><span class="kpi-label">Leads Delivered</span>
        <div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg></div></div>
      <div class="kpi-value" id="kpi-leads">—</div><div class="kpi-sub">Total qualified leads</div>
    </div>
    <div class="kpi-card fu" data-c="amber" style="animation-delay:.2s">
      <div class="kpi-header"><span class="kpi-label">Total Spends</span>
        <div class="kpi-icon"><svg viewBox="0 0 24 24"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg></div></div>
      <div class="kpi-value" id="kpi-spend">—</div><div class="kpi-sub">Total ad budget consumed</div>
    </div>
    <div class="kpi-card fu" data-c="violet" style="animation-delay:.22s">
      <div class="kpi-header"><span class="kpi-label">Avg. CTR</span>
        <div class="kpi-icon"><svg viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg></div></div>
      <div class="kpi-value" id="kpi-ctr">—</div><div class="kpi-sub">Click-through rate avg</div>
    </div>
    <div class="kpi-card fu" data-c="rose" style="animation-delay:.24s">
      <div class="kpi-header"><span class="kpi-label">Cost Per Lead</span>
        <div class="kpi-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg></div></div>
      <div class="kpi-value" id="kpi-cpl">—</div><div class="kpi-sub">Avg spend per lead</div>
    </div>
  </div>

  <!-- ROW 1: Impressions+Clicks timeline | Spends+Leads timeline -->
  <div class="row-2">
    <div class="chart-card fu" style="animation-delay:.28s">
      <div class="chart-header"><div class="chart-title">Impressions & Clicks Over Time</div><div class="chart-sub">Daily trend</div></div>
      <div class="chart-body"><div id="ch-impclik" style="height:300px"></div></div>
    </div>
    <div class="chart-card fu" style="animation-delay:.33s">
      <div class="chart-header"><div class="chart-title">Spends & Leads Over Time</div><div class="chart-sub">Daily budget vs lead delivery</div></div>
      <div class="chart-body"><div id="ch-spendlead" style="height:300px"></div></div>
    </div>
  </div>

  <!-- ROW 2: CTR trend | Project breakdown donut -->
  <div class="row-2">
    <div class="chart-card fu" style="animation-delay:.38s">
      <div class="chart-header"><div class="chart-title">CTR % Daily Trend</div><div class="chart-sub">Click-through rate over time</div></div>
      <div class="chart-body"><div id="ch-ctr" style="height:300px"></div></div>
    </div>
    <div class="chart-card fu" style="animation-delay:.43s">
      <div class="chart-header"><div class="chart-title">Spend by Project</div><div class="chart-sub">Budget allocation across campaigns</div></div>
      <div class="chart-body"><div id="ch-projdonut" style="height:300px"></div></div>
    </div>
  </div>

  <!-- ROW 3: CPL trend | Leads donut | Impressions bar -->
  <div class="row-3">
    <div class="chart-card fu" style="animation-delay:.48s">
      <div class="chart-header"><div class="chart-title">Cost Per Lead</div><div class="chart-sub">Daily CPL (₹ spend ÷ leads)</div></div>
      <div class="chart-body"><div id="ch-cpl" style="height:280px"></div></div>
    </div>
    <div class="chart-card fu" style="animation-delay:.52s">
      <div class="chart-header"><div class="chart-title">Leads by Project</div><div class="chart-sub">Share of total leads</div></div>
      <div class="chart-body"><div id="ch-leadsdonut" style="height:280px"></div></div>
    </div>
    <div class="chart-card fu" style="animation-delay:.56s">
      <div class="chart-header"><div class="chart-title">Daily Impressions</div><div class="chart-sub">Volume per day</div></div>
      <div class="chart-body"><div id="ch-impbar" style="height:280px"></div></div>
    </div>
  </div>
</div>

<script>
// ── UPLOAD ──────────────────────────────────────────────────────────
const dropZone=document.getElementById('drop-zone');
const fileInput=document.getElementById('file-input');
dropZone.addEventListener('dragover',e=>{e.preventDefault();dropZone.classList.add('drag')});
dropZone.addEventListener('dragleave',()=>dropZone.classList.remove('drag'));
dropZone.addEventListener('drop',e=>{e.preventDefault();dropZone.classList.remove('drag');const f=e.dataTransfer.files[0];if(f)readFile(f)});
fileInput.addEventListener('change',e=>{if(e.target.files[0])readFile(e.target.files[0])});

function readFile(f){
  const r=new FileReader();
  r.onload=e=>{
    try{
      let raw=JSON.parse(e.target.result);
      if(!Array.isArray(raw))throw new Error('Expected a JSON array');
      const data=parseRows(raw);
      if(!data.length)throw new Error('No valid data rows found. Check the file format.');
      buildDashboard(data);
    }catch(err){alert('Could not parse file:\n'+err.message)}
  };
  r.readAsText(f);
}

// ── SMART PARSER ─────────────────────────────────────────────────────
// Handles two formats:
// 1. Keyed JSON  {"Order ID":…,"Sales":…}
// 2. Positional JSON from CSV export {"":"","__1":"","__2":"PROJECT","__3":"01-05-2026",…}
//    Row 1 (index 1) is the header row; data starts at index 2.
function parseRows(raw){
  // Detect positional format: check if any row has "__3" key and no "Date" key
  const isPositional = raw.length>0 && ('__3' in raw[0] || '__3' in (raw[1]||{}));

  if(isPositional){
    // Find the header row (the one where __4 looks like "Impressions" or a number header)
    // Data rows: __3 looks like a date string (contains "-" or "/")
    let currentProject='Unknown';
    const rows=[];
    raw.forEach(r=>{
      // detect project name row: __2 is non-empty, __3 looks like a date or __4 is a number
      if(r['__2'] && String(r['__2']).trim() && !String(r['__2']).toLowerCase().includes('date')){
        currentProject=String(r['__2']).trim();
      }
      const dateVal=String(r['__3']||'').trim();
      const impVal=r['__4'];
      // Valid data row: __3 is a date-like string and __4 is numeric
      if(dateVal && /\d{2}[-\/]\d{2}[-\/]\d{4}/.test(dateVal) && impVal!==undefined && impVal!==''){
        const ctrRaw=String(r['__6']||'0').replace('%','').trim();
        rows.push({
          project: currentProject,
          date: parseDate(dateVal),
          dateStr: dateVal,
          impressions: +impVal||0,
          clicks: +r['__5']||0,
          ctr: parseFloat(ctrRaw)||0,
          leads: +r['__7']||0,
          spends: +r['__8']||0,
        });
      }
    });
    return rows;
  }

  // Keyed format (standard)
  return raw.filter(r=>r['Date']||r['Order Date']||r['Sales']!=null).map(r=>({
    project: r['Campaign']||r['Project']||'Campaign',
    date: parseDate(r['Date']||r['Order Date']||''),
    dateStr: r['Date']||r['Order Date']||'',
    impressions: +r['Impressions']||0,
    clicks: +r['Clicks']||0,
    ctr: parseFloat(String(r['CTR%']||r['CTR']||'0').replace('%',''))||0,
    leads: +r['LEADS DELIVERD']||+r['Leads']||0,
    spends: +r['SPENDS']||+r['Spends']||+r['Sales']||0,
  }));
}

function parseDate(s){
  if(!s)return null;
  s=String(s).trim();
  // DD-MM-YYYY or DD/MM/YYYY
  const m=s.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})$/);
  if(m)return new Date(+m[3],+m[2]-1,+m[1]);
  return new Date(s)||null;
}

function resetDashboard(){
  document.getElementById('dashboard').style.display='none';
  document.getElementById('dl-tab').style.display='none';
  document.getElementById('tab-nav').style.display='none';
  document.getElementById('upload-section').style.display='block';
  document.getElementById('tab-dashboard').classList.add('active');
  document.getElementById('tab-download').classList.remove('active');
  fileInput.value='';
}

// ── PLOTLY HELPERS ───────────────────────────────────────────────────
const PF={family:'Poppins,sans-serif',color:'#475569'};
const GR={gridcolor:'#e2e8f0',zerolinecolor:'#e2e8f0'};
const CFG={displayModeBar:false,responsive:true};

function BL(extra){
  return Object.assign({
    font:PF,paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',
    margin:{t:24,r:20,b:48,l:56},
    legend:{font:PF,bgcolor:'rgba(0,0,0,0)',borderwidth:0},
    xaxis:Object.assign({showgrid:false,tickfont:PF},GR),
    yaxis:Object.assign({tickfont:PF},GR),
  },extra||{});
}

const PROJ_COLORS=['#2563eb','#0891b2','#059669','#d97706','#7c3aed','#e11d48','#0284c7','#dc2626'];

// ── BUILD DASHBOARD ───────────────────────────────────────────────────
let _allData=[];
let _activeProjects=new Set();

function buildDashboard(data){
  _allData=data;
  document.getElementById('upload-section').style.display='none';
  document.getElementById('tab-nav').style.display='block';
  const dash=document.getElementById('dashboard');
  dash.style.display='block';
  document.getElementById('rec-count').textContent=data.length.toLocaleString();

  // Project tabs
  const projects=[...new Set(data.map(d=>d.project))];
  _activeProjects=new Set(projects);
  renderProjectTabs(projects);
  renderCharts(data,projects);
}

function renderProjectTabs(projects){
  const tabs=document.getElementById('proj-tabs');
  if(projects.length<=1){tabs.style.display='none';return;}
  tabs.style.display='flex';
  tabs.innerHTML='';
  projects.forEach((p,i)=>{
    const c=PROJ_COLORS[i%PROJ_COLORS.length];
    const btn=document.createElement('button');
    btn.className='proj-badge active';
    btn.style.color=c;btn.style.backgroundColor=c+'18';
    btn.innerHTML=`<span class="proj-dot" style="background:${c}"></span>${p}`;
    btn.onclick=()=>{
      if(_activeProjects.has(p)&&_activeProjects.size===1)return;
      if(_activeProjects.has(p))_activeProjects.delete(p); else _activeProjects.add(p);
      btn.classList.toggle('active');
      const filtered=_allData.filter(d=>_activeProjects.has(d.project));
      renderCharts(filtered,[...new Set(filtered.map(d=>d.project))]);
    };
    tabs.appendChild(btn);
  });
}

function renderCharts(data,projects){
  const n=v=>+v||0;
  const fmtK=v=>'₹'+Math.round(v).toLocaleString('en-IN');
  const fmtN=v=>Math.round(v).toLocaleString('en-IN');

  const totalImp=data.reduce((s,d)=>s+n(d.impressions),0);
  const totalClk=data.reduce((s,d)=>s+n(d.clicks),0);
  const totalLeads=data.reduce((s,d)=>s+n(d.leads),0);
  const totalSpend=data.reduce((s,d)=>s+n(d.spends),0);
  const avgCTR=data.length?data.reduce((s,d)=>s+n(d.ctr),0)/data.length:0;
  const cpl=totalLeads?totalSpend/totalLeads:0;

  document.getElementById('kpi-imp').textContent=fmtN(totalImp);
  document.getElementById('kpi-clk').textContent=fmtN(totalClk);
  document.getElementById('kpi-leads').textContent=totalLeads.toLocaleString('en-IN');
  document.getElementById('kpi-spend').textContent=fmtK(totalSpend);
  document.getElementById('kpi-ctr').textContent=avgCTR.toFixed(2)+'%';
  document.getElementById('kpi-cpl').textContent=fmtK(cpl);

  // Sort by date
  const sorted=[...data].sort((a,b)=>(a.date||0)-(b.date||0));
  const dates=sorted.map(d=>d.dateStr);
  const imps=sorted.map(d=>d.impressions);
  const clks=sorted.map(d=>d.clicks);
  const spends=sorted.map(d=>d.spends);
  const leads=sorted.map(d=>d.leads);
  const ctrs=sorted.map(d=>d.ctr);
  const cpls=sorted.map(d=>d.leads?d.spends/d.leads:null);

  // ── Chart 1: Impressions & Clicks ──
  Plotly.react('ch-impclik',[
    {x:dates,y:imps,name:'Impressions',type:'scatter',mode:'lines',yaxis:'y',
     line:{color:'#2563eb',width:2.5,shape:'spline'},fill:'tozeroy',fillcolor:'rgba(37,99,235,.07)'},
    {x:dates,y:clks,name:'Clicks',type:'scatter',mode:'lines',yaxis:'y2',
     line:{color:'#0891b2',width:2,shape:'spline',dash:'dot'}},
  ],BL({
    yaxis:{...GR,tickfont:PF,title:{text:'Impressions',font:{...PF,size:11}}},
    yaxis2:{overlaying:'y',side:'right',tickfont:PF,title:{text:'Clicks',font:{...PF,size:11}},showgrid:false,zeroline:false},
    legend:{orientation:'h',y:1.08,x:.5,xanchor:'center',font:PF,bgcolor:'rgba(0,0,0,0)'},
    margin:{t:30,r:56,b:48,l:56}
  }),CFG);

  // ── Chart 2: Spends & Leads ──
  Plotly.react('ch-spendlead',[
    {x:dates,y:spends,name:'Spends (₹)',type:'bar',yaxis:'y',
     marker:{color:'rgba(124,58,237,.15)',line:{color:'#7c3aed',width:1.5}}},
    {x:dates,y:leads,name:'Leads',type:'scatter',mode:'lines+markers',yaxis:'y2',
     line:{color:'#059669',width:2.5,shape:'spline'},marker:{size:6,color:'#059669'}},
  ],BL({
    yaxis:{...GR,tickfont:PF,title:{text:'Spends (₹)',font:{...PF,size:11}}},
    yaxis2:{overlaying:'y',side:'right',tickfont:PF,title:{text:'Leads',font:{...PF,size:11}},showgrid:false,zeroline:false},
    legend:{orientation:'h',y:1.08,x:.5,xanchor:'center',font:PF,bgcolor:'rgba(0,0,0,0)'},
    margin:{t:30,r:56,b:48,l:56},barmode:'group'
  }),CFG);

  // ── Chart 3: CTR Trend ──
  Plotly.react('ch-ctr',[
    {x:dates,y:ctrs,name:'CTR%',type:'scatter',mode:'lines+markers',
     line:{color:'#d97706',width:2.5,shape:'spline'},
     marker:{size:7,color:'#d97706',line:{color:'#fff',width:1.5}},
     fill:'tozeroy',fillcolor:'rgba(217,119,6,.07)',
     hovertemplate:'%{x}<br>CTR: <b>%{y:.2f}%</b><extra></extra>'}
  ],BL({
    yaxis:{...GR,tickfont:PF,ticksuffix:'%'},
    margin:{t:20,r:20,b:48,l:52}
  }),CFG);

  // ── Chart 4: Spend donut by project ──
  const projSpend={};
  data.forEach(d=>{projSpend[d.project]=(projSpend[d.project]||0)+d.spends});
  const pKeys=Object.keys(projSpend);
  Plotly.react('ch-projdonut',[{
    labels:pKeys,values:pKeys.map(k=>projSpend[k]),type:'pie',hole:.56,
    marker:{colors:PROJ_COLORS.slice(0,pKeys.length)},
    textinfo:'label+percent',textfont:PF,
    hovertemplate:'<b>%{label}</b><br>₹%{value:,.0f}<extra></extra>'
  }],Object.assign(BL(),{showlegend:false,margin:{t:16,r:16,b:16,l:16}}),CFG);

  // ── Chart 5: CPL trend ──
  Plotly.react('ch-cpl',[
    {x:dates,y:cpls,name:'CPL',type:'scatter',mode:'lines+markers',
     line:{color:'#e11d48',width:2.5,shape:'spline'},
     marker:{size:7,color:'#e11d48',line:{color:'#fff',width:1.5}},
     fill:'tozeroy',fillcolor:'rgba(225,29,72,.07)',
     connectgaps:false,
     hovertemplate:'%{x}<br>CPL: <b>₹%{y:,.0f}</b><extra></extra>'}
  ],BL({
    yaxis:{...GR,tickfont:PF,tickprefix:'₹'},
    margin:{t:20,r:16,b:48,l:60}
  }),CFG);

  // ── Chart 6: Leads donut by project ──
  const projLeads={};
  data.forEach(d=>{projLeads[d.project]=(projLeads[d.project]||0)+d.leads});
  const lKeys=Object.keys(projLeads);
  Plotly.react('ch-leadsdonut',[{
    labels:lKeys,values:lKeys.map(k=>projLeads[k]),type:'pie',hole:.56,
    marker:{colors:PROJ_COLORS.slice(0,lKeys.length)},
    textinfo:'label+percent',textfont:PF,
    hovertemplate:'<b>%{label}</b><br>%{value} leads<extra></extra>'
  }],Object.assign(BL(),{showlegend:false,margin:{t:16,r:16,b:16,l:16}}),CFG);

  // ── Chart 7: Impressions bar ──
  Plotly.react('ch-impbar',[
    {x:dates,y:imps,type:'bar',
     marker:{color:imps.map((_,i)=>{
       const t=i/(imps.length-1||1);
       const r=Math.round(37+(8-37)*t),g=Math.round(99+(145-99)*t),b=Math.round(235+(178-235)*t);
       return`rgba(${r},${g},${b},.85)`;
     }),line:{width:0}},
     hovertemplate:'%{x}<br><b>%{y:,.0f}</b> impressions<extra></extra>'}
  ],BL({
    yaxis:{...GR,tickfont:PF},
    margin:{t:16,r:16,b:48,l:60}
  }),CFG);
}

// ── TAB SWITCHING ──────────────────────────────────────────────────────
function switchTab(tab){
  document.getElementById('tab-dashboard').classList.toggle('active', tab==='dashboard');
  document.getElementById('tab-download').classList.toggle('active', tab==='download');
  document.getElementById('dashboard').style.display = tab==='dashboard' ? 'block' : 'none';
  document.getElementById('dl-tab').style.display = tab==='download' ? 'block' : 'none';
}

// ── EXCEL EXPORT ───────────────────────────────────────────────────────
function downloadExcel(){
  if(!_allData || !_allData.length){alert('No data loaded. Please upload a file first.');return;}
  var btn=document.getElementById('dl-btn-main');
  btn.classList.add('loading');
  btn.innerHTML='<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="2" x2="12" y2="6"/><line x1="12" y1="18" x2="12" y2="22"/><line x1="4.93" y1="4.93" x2="7.76" y2="7.76"/><line x1="16.24" y1="16.24" x2="19.07" y2="19.07"/><line x1="2" y1="12" x2="6" y2="12"/><line x1="18" y1="12" x2="22" y2="12"/><line x1="4.93" y1="19.07" x2="7.76" y2="16.24"/><line x1="16.24" y1="7.76" x2="19.07" y2="4.93"/></svg> Generating…';

  setTimeout(function(){
    try{
      var filtered = _allData.filter(function(d){ return _activeProjects.has(d.project); });
      var sorted = filtered.slice().sort(function(a,b){ return (a.date||0)-(b.date||0); });
      function n(v){ return +v||0; }

      var totalImp   = sorted.reduce(function(s,d){return s+n(d.impressions);},0);
      var totalClk   = sorted.reduce(function(s,d){return s+n(d.clicks);},0);
      var totalLeads = sorted.reduce(function(s,d){return s+n(d.leads);},0);
      var totalSpend = sorted.reduce(function(s,d){return s+n(d.spends);},0);
      var avgCTR     = sorted.length ? sorted.reduce(function(s,d){return s+n(d.ctr);},0)/sorted.length : 0;
      var cpl        = totalLeads ? totalSpend/totalLeads : 0;
      var projects   = [...new Set(sorted.map(function(d){return d.project;}))];
      var dateRange  = sorted.length ? sorted[0].dateStr+' to '+sorted[sorted.length-1].dateStr : '—';

      var wb = XLSX.utils.book_new();

      // ── Sheet 1: Summary KPIs ──
      var s1 = [
        ['CAMPAIGN PERFORMANCE SUMMARY REPORT'],
        [''],
        ['Report Generated', new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'short',year:'numeric'})],
        ['Date Range', dateRange],
        ['Projects', projects.join(', ')],
        ['Total Days', sorted.length],
        [''],
        ['Metric', 'Value'],
        ['Total Impressions', totalImp],
        ['Total Clicks', totalClk],
        ['Total Leads Delivered', totalLeads],
        ['Total Spends (Rs)', totalSpend],
        ['Average CTR (%)', parseFloat(avgCTR.toFixed(2))],
        ['Cost Per Lead (Rs)', parseFloat(cpl.toFixed(2))],
        ['Click-to-Lead Rate (%)', totalClk ? parseFloat((totalLeads/totalClk*100).toFixed(2)) : 0],
      ];
      var ws1 = XLSX.utils.aoa_to_sheet(s1);
      ws1['!cols'] = [{wch:32},{wch:28}];
      XLSX.utils.book_append_sheet(wb, ws1, 'Summary KPIs');

      // ── Sheet 2: Daily Data ──
      var s2 = [['Project','Date','Impressions','Clicks','CTR (%)','Leads Delivered','Spends (Rs)','Cost Per Lead (Rs)']];
      sorted.forEach(function(d){
        s2.push([
          d.project, d.dateStr,
          n(d.impressions), n(d.clicks),
          parseFloat(n(d.ctr).toFixed(2)),
          n(d.leads), n(d.spends),
          d.leads ? parseFloat((n(d.spends)/n(d.leads)).toFixed(2)) : ''
        ]);
      });
      s2.push(['TOTAL','',totalImp,totalClk,parseFloat(avgCTR.toFixed(2)),totalLeads,totalSpend,parseFloat(cpl.toFixed(2))]);
      var ws2 = XLSX.utils.aoa_to_sheet(s2);
      ws2['!cols'] = [{wch:18},{wch:14},{wch:16},{wch:10},{wch:10},{wch:18},{wch:14},{wch:20}];
      XLSX.utils.book_append_sheet(wb, ws2, 'Daily Data');

      // ── Sheet 3: Project Breakdown ──
      var projMap = {};
      sorted.forEach(function(d){
        if(!projMap[d.project]) projMap[d.project]={imp:0,clk:0,leads:0,spend:0,ctrSum:0,days:0};
        var p=projMap[d.project];
        p.imp+=n(d.impressions); p.clk+=n(d.clicks);
        p.leads+=n(d.leads); p.spend+=n(d.spends);
        p.ctrSum+=n(d.ctr); p.days+=1;
      });
      var s3 = [['Project','Days Active','Total Impressions','Total Clicks','Total Leads','Total Spends (Rs)','Avg CTR (%)','Cost Per Lead (Rs)','Impressions/Day']];
      Object.keys(projMap).forEach(function(proj){
        var p=projMap[proj];
        s3.push([
          proj, p.days, p.imp, p.clk, p.leads, p.spend,
          parseFloat((p.ctrSum/p.days).toFixed(2)),
          p.leads ? parseFloat((p.spend/p.leads).toFixed(2)) : '',
          parseFloat((p.imp/p.days).toFixed(0))
        ]);
      });
      var ws3 = XLSX.utils.aoa_to_sheet(s3);
      ws3['!cols'] = [{wch:18},{wch:13},{wch:20},{wch:14},{wch:14},{wch:18},{wch:12},{wch:20},{wch:16}];
      XLSX.utils.book_append_sheet(wb, ws3, 'Project Breakdown');

      var fname = 'Campaign_Report_' + new Date().toISOString().slice(0,10) + '.xlsx';
      XLSX.writeFile(wb, fname);

    } catch(e){
      alert('Export failed: ' + e.message + '\n' + e.stack);
    }

    btn.classList.remove('loading');
    btn.innerHTML = '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download Excel Report';
  }, 80);
}

</script>

<!-- DOWNLOAD TAB -->
<div class="dl-tab" id="dl-tab">
  <div style="margin-top:40px">
    <div class="dl-hero">
      <div class="dl-hero-icon">
        <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="13" x2="12" y2="19"/><polyline points="9 16 12 19 15 16"/></svg>
      </div>
      <div class="dl-hero-text">
        <h2>Download Campaign Report</h2>
        <p>Export your full campaign data as a formatted Excel workbook with 3 sheets — Summary KPIs, Daily Data, and Project Breakdown. Ready to share with clients.</p>
      </div>
      <button class="dl-btn" id="dl-btn-main" onclick="downloadExcel()">
        <svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Download Excel Report
      </button>
    </div>

    <div class="dl-sheets">
      <div class="sheet-card" data-sc="blue">
        <div class="sheet-icon"><svg viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg></div>
        <div class="sheet-name">📊 Sheet 1 — Summary KPIs</div>
        <div class="sheet-desc">High-level campaign totals and averages across the full date range.</div>
        <div class="sheet-cols">
          <span class="col-tag">Total Impressions</span><span class="col-tag">Total Clicks</span>
          <span class="col-tag">Total Leads</span><span class="col-tag">Total Spends</span>
          <span class="col-tag">Avg CTR%</span><span class="col-tag">Cost Per Lead</span>
        </div>
      </div>
      <div class="sheet-card" data-sc="teal">
        <div class="sheet-icon"><svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg></div>
        <div class="sheet-name">📅 Sheet 2 — Daily Data</div>
        <div class="sheet-desc">Full row-by-row daily breakdown with all metrics and computed CPL.</div>
        <div class="sheet-cols">
          <span class="col-tag">Project</span><span class="col-tag">Date</span>
          <span class="col-tag">Impressions</span><span class="col-tag">Clicks</span>
          <span class="col-tag">CTR%</span><span class="col-tag">Leads</span>
          <span class="col-tag">Spends</span><span class="col-tag">CPL</span>
        </div>
      </div>
      <div class="sheet-card" data-sc="violet">
        <div class="sheet-icon"><svg viewBox="0 0 24 24"><path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/></svg></div>
        <div class="sheet-name">🏗 Sheet 3 — Project Breakdown</div>
        <div class="sheet-desc">Aggregated totals per project/campaign with performance ratios.</div>
        <div class="sheet-cols">
          <span class="col-tag">Project</span><span class="col-tag">Days Active</span>
          <span class="col-tag">Impressions</span><span class="col-tag">Clicks</span>
          <span class="col-tag">Leads</span><span class="col-tag">Spends</span>
          <span class="col-tag">Avg CTR%</span><span class="col-tag">CPL</span>
        </div>
      </div>
    </div>

    <div class="dl-note">
      <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
      <span>The Excel file is generated entirely in your browser — no data is uploaded to any server. Each download reflects the currently filtered data.</span>
    </div>
  </div>
</div>

</body>
</html>
