#!/usr/bin/env python3
"""
Rebuild "Cut off / Open gate dashboard.html" from a daily CUT_OFF-style .xls file
and a Wharf.xls code-to-name lookup.

Usage:
    python3 build_dashboard.py CUT_OFF.xls Wharf.xls -o "Cut off - Open Gate Dashboard.html"

Requirements:
    pip install xlrd

Rules applied (matches the standing spec):
    - Cut off (Dry)    = ETA - 24 hours
    - Cut off (Reefer) = ETA - 1 hour
    - Open gate        = ETD - 6 days
    - Rows where Skip == "Y" are excluded
    - Wharf codes are resolved to full names using Wharf.xls (code shown underneath)
"""
import argparse
import json
import sys
from datetime import datetime, timedelta

try:
    import xlrd
except ImportError:
    sys.exit("Missing dependency. Install it first:  pip install xlrd")


def parse_dt(v):
    if isinstance(v, str):
        return datetime.strptime(v.strip(), "%Y-%m-%d %H:%M")
    return v


def load_cutoff_records(xls_path):
    wb = xlrd.open_workbook(xls_path)
    sh = wb.sheet_by_index(0)
    records = []
    skipped = 0
    for r in range(1, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        (service, vessel_code, vessel_name, op_liner, seq, vyg, bound,
         vyg_bound, wharf, pol, pod, skip, no_d, no_l, used,
         eta, etb, etd) = row

        if skip == "Y":
            skipped += 1
            continue

        eta_dt = parse_dt(eta)
        etb_dt = parse_dt(etb)
        etd_dt = parse_dt(etd)

        cutoff_dry = eta_dt - timedelta(hours=24)
        cutoff_reefer = eta_dt - timedelta(hours=1)
        opengate = etd_dt - timedelta(days=6)

        records.append({
            "service": service, "vessel_code": vessel_code, "vessel": vessel_name,
            "op_liner": op_liner, "seq": seq, "vyg": vyg, "bound": bound,
            "vyg_bound": vyg_bound, "wharf": wharf, "pol": pol, "pod": pod,
            "no_d": no_d, "no_l": no_l, "used": used,
            "eta": eta_dt.strftime("%Y-%m-%d %H:%M"),
            "etb": etb_dt.strftime("%Y-%m-%d %H:%M"),
            "etd": etd_dt.strftime("%Y-%m-%d %H:%M"),
            "cutoff_dry": cutoff_dry.strftime("%Y-%m-%d %H:%M"),
            "cutoff_reefer": cutoff_reefer.strftime("%Y-%m-%d %H:%M"),
            "opengate": opengate.strftime("%Y-%m-%d %H:%M"),
        })

    records.sort(key=lambda x: x["eta"])
    print(f"Loaded {len(records)} voyages ({skipped} Skip=Y rows excluded) from {xls_path}")
    return records


def load_wharf_map(xls_path):
    wb = xlrd.open_workbook(xls_path)
    sh = wb.sheet_by_index(0)
    m = {}
    for r in range(1, sh.nrows):
        code = str(sh.cell_value(r, 0)).strip()
        name = str(sh.cell_value(r, 1)).strip()
        if code:
            m[code] = name
    print(f"Loaded {len(m)} wharf codes from {xls_path}")
    return m


def check_wharf_coverage(records, wharf_map):
    used = {r["wharf"] for r in records}
    missing = used - set(wharf_map.keys())
    if missing:
        print(f"Warning: {len(missing)} wharf code(s) not found in Wharf.xls: {sorted(missing)}")


TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>HEUNG A LINE — Cut Off / Open Gate Dashboard - THBKK & THLCH</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root{
    --navy:#0b2540; --navy2:#123a5e; --navy3:#1a4a72;
    --canvas:#eef1f2; --card:#ffffff; --line:#dde3e7; --line-strong:#c3ccd2;
    --text:#132230; --text2:#586975; --muted:#93a2ac;
    --bkk:#2563a6; --bkk-bg:#e7f0fa;
    --lch:#1f8a72; --lch-bg:#e4f4ef;
    --coral:#c1483f; --coral-bg:#fbebe9;
    --amber:#c1791f; --amber-bg:#faf0dd; --amber-strong:#dd9a35;
    --teal:#1a7862; --teal-bg:#e4f4ef;
  }
  *{box-sizing:border-box;}
  body{margin:0;font-family:'Space Grotesk',-apple-system,'Segoe UI',sans-serif;background:var(--canvas);color:var(--text);}
  .mono{font-family:'JetBrains Mono',ui-monospace,monospace;}
  .wrap{max-width:none;margin:0;padding:0 16px 60px;}

  /* Hero */
  .hero{background:var(--navy);background-image:linear-gradient(160deg,var(--navy) 0%,#0e2f52 60%,var(--navy2) 100%);color:#fff;padding:28px 20px 24px;margin-bottom:22px;}
  .hero-inner{max-width:none;margin:0;display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap;}
  .hero h1{font-size:24px;font-weight:700;margin:0 0 6px;letter-spacing:-.01em;}
  .hero p{font-size:13px;color:#a9bccd;margin:0;max-width:520px;line-height:1.5;}
  .hero-mark{display:flex;align-items:center;gap:8px;color:#7fa8cb;font-size:11px;text-transform:uppercase;letter-spacing:.08em;font-weight:600;}
  .hero-logo{background:#fff;border-radius:10px;padding:9px 18px;flex:none;box-shadow:0 2px 10px rgba(0,0,0,.18);}
  .hero-logo svg{height:42px;width:auto;display:block;}

  .rule-strip{display:flex;gap:0;flex-wrap:wrap;margin:16px 0 20px;border:1px solid var(--line);border-radius:12px;overflow:hidden;background:var(--card);}
  .rule{flex:1;min-width:200px;padding:12px 16px;font-size:12.5px;color:var(--text2);display:flex;align-items:center;gap:9px;border-right:1px solid var(--line);}
  .rule:last-child{border-right:none;}
  .rule .tick{width:3px;align-self:stretch;border-radius:2px;flex:none;}
  .rule b{color:var(--text);font-weight:600;font-family:'JetBrains Mono';}

  /* Board summary strip */
  .board{background:var(--navy);border-radius:14px;padding:4px 4px;margin-bottom:20px;display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));overflow:hidden;}
  .board .seg-stat{padding:14px 18px;border-right:1px dashed rgba(255,255,255,.16);}
  .board .seg-stat:last-child{border-right:none;}
  .board .seg-stat .lbl{font-size:10.5px;color:#8ba6bf;text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;}
  .board .seg-stat .val{font-family:'JetBrains Mono';font-size:19px;font-weight:600;color:#fff;line-height:1.25;}
  .board .seg-stat .val.small{font-size:12.5px;font-weight:500;}
  .board .seg-stat .val .amber{color:var(--amber-strong);}

  .controls{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:12px;}
  .controls input[type=text]{flex:1;min-width:180px;padding:10px 13px;border:1px solid var(--line-strong);border-radius:9px;font-size:13.5px;background:var(--card);font-family:inherit;}
  .controls select{padding:10px 11px;border:1px solid var(--line-strong);border-radius:9px;font-size:13px;background:var(--card);color:var(--text);font-family:inherit;}
  .seg{display:flex;border:1px solid var(--line-strong);border-radius:9px;overflow:hidden;background:var(--card);}
  .seg button{border:none;background:transparent;padding:9px 14px;font-size:13px;cursor:pointer;color:var(--text2);font-family:inherit;}
  .seg button.active{background:var(--navy);color:#fff;}
  .seg button:not(:last-child){border-right:1px solid var(--line);}
  .clearbtn{padding:9px 13px;border:1px solid var(--line-strong);border-radius:9px;background:var(--card);font-size:13px;cursor:pointer;color:var(--text2);font-family:inherit;}
  .clearbtn:hover{background:#f3f5f6;}

  #tableHolder{overflow-x:auto;}
  table{width:auto;max-width:100%;border-collapse:collapse;background:var(--card);border-radius:12px;overflow:hidden;border:1px solid var(--line);font-size:12.8px;}
  thead th{background:var(--navy);text-align:left;padding:11px 12px;font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;color:#9fb7cc;font-weight:600;cursor:pointer;white-space:nowrap;}
  thead th:hover{color:#fff;}
  thead th.sorted{color:var(--amber-strong);}
  tbody td{padding:10px 12px;border-bottom:1px solid var(--line);white-space:nowrap;}
  tbody tr{position:relative;}
  tbody tr td:first-child{box-shadow:inset 4px 0 0 0 var(--rowaccent, transparent);}
  tbody tr:last-child td{border-bottom:none;}
  tbody tr:hover{background:#f6f8f9;}
  .pol-badge{font-size:11px;font-weight:600;padding:2px 9px;border-radius:5px;font-family:'JetBrains Mono';}
  .pol-bkk{background:var(--bkk-bg);color:var(--bkk);}
  .pol-lch{background:var(--lch-bg);color:var(--lch);}
  .vessel{font-weight:600;}
  .sub{color:var(--muted);font-size:11.5px;margin-top:1px;}
  .dtcell{display:flex;flex-direction:column;line-height:1.4;font-family:'JetBrains Mono';}
  .dtcell .d{font-weight:500;}
  .dtcell .t{color:var(--muted);font-size:11.3px;}
  .flag{display:inline-block;font-size:10px;font-weight:700;padding:1px 6px;border-radius:5px;margin-left:5px;vertical-align:1px;text-transform:uppercase;letter-spacing:.03em;font-family:'Space Grotesk';}
  .flag-soon{background:var(--amber-bg);color:var(--amber);}
  .flag-past{background:var(--coral-bg);color:var(--coral);}
  .flag-open{background:var(--teal-bg);color:var(--teal);}
  .noload{background:#eef0e9;color:#65695a;font-size:10px;padding:1px 6px;border-radius:5px;font-family:'Space Grotesk';}
  .empty{padding:50px 20px;text-align:center;color:var(--muted);font-size:14px;}
  .count{font-size:12px;color:var(--muted);margin:10px 2px 0;font-family:'JetBrains Mono';}

  /* Mobile ticket-stub card view */
  .mcards{display:flex;flex-direction:column;gap:12px;}
  .mcard{display:flex;background:var(--card);border:1px solid var(--line);border-radius:14px;overflow:hidden;position:relative;}
  .mcard-main{flex:1;padding:14px 15px;min-width:0;}
  .mcard-top{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;}
  .mcard-vessel{font-weight:600;font-size:14px;}
  .mcard-sub{color:var(--muted);font-size:11.3px;margin-top:2px;}
  .mcard-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px 14px;margin-top:12px;font-family:'JetBrains Mono';}
  .mcard-item .lbl{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:3px;font-family:'Space Grotesk';}
  .mcard-item .val{font-size:12.6px;font-weight:500;}
  .mcard-item .val .t{color:var(--muted);font-size:11px;font-weight:400;margin-left:4px;}
  .mcard-item.full{grid-column:1 / -1;}
  .mcard-stub{width:46px;flex:none;background:var(--navy);position:relative;display:flex;align-items:center;justify-content:center;background-image:linear-gradient(180deg,var(--navy),#0e2f52);}
  .mcard-stub::before{content:'';position:absolute;left:0;top:0;bottom:0;width:0;border-left:2px dashed rgba(255,255,255,.3);}
  .mcard-stub::after{content:'';position:absolute;top:-7px;left:-7px;width:14px;height:14px;border-radius:50%;background:var(--canvas);box-shadow:0 calc(100% - 14px) 0 0 var(--canvas);}
  .mcard-stub-pol{writing-mode:vertical-rl;transform:rotate(180deg);color:#fff;font-family:'JetBrains Mono';font-weight:600;font-size:12.5px;letter-spacing:.08em;}

  @media (max-width:760px){
    table{display:block;overflow-x:auto;}
    .board{grid-template-columns:repeat(2,1fr);}
    .hero h1{font-size:20px;}
  }
</style>
</head>
<body>
<div class="hero">
  <div class="hero-inner">
    <div>
      <div class="hero-mark">⚓ Terminal Gate Operations</div>
      <h1>Cut off / Open gate dashboard</h1>
      <p>POL THBKK &amp; THLCH — calculated from the vessel schedule (ETA / ETD), latest update from __SOURCE_FILENAME__</p>
    </div>
    <div class="hero-logo">
      <svg viewBox="0 0 340 60" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="HEUNG A LINE 興亞 LINE 株式會社">
        <ellipse cx="31" cy="30" rx="30" ry="25" fill="#e9f2fc"/>
        <ellipse cx="31" cy="30" rx="30" ry="25" fill="none" stroke="#0d5aa7" stroke-width="2.5"/>
        <ellipse cx="31" cy="30" rx="21" ry="16" fill="#c8102e"/>
        <path d="M8 33 Q19 24 31 33 T54 33" fill="none" stroke="#ffffff" stroke-width="2.4" stroke-linecap="round"/>
        <text x="31" y="34.5" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="13" font-weight="700" font-style="italic" fill="#ffffff">HAL</text>
        <text x="74" y="28" font-family="Arial, Helvetica, sans-serif" font-size="23" font-weight="700" fill="#0d5aa7" letter-spacing=".5">HEUNG A LINE</text>
        <text x="75" y="50" font-family="Arial, Helvetica, sans-serif" font-size="14" font-weight="700" fill="#111111" letter-spacing="2">興亞 LINE 株式會社</text>
      </svg>
    </div>
  </div>
</div>
<div class="wrap">

<div class="rule-strip">
  <div class="rule"><span class="tick" style="background:#2563a6"></span>Dry cut off = <b>ETA − 24h</b></div>
  <div class="rule"><span class="tick" style="background:#1a7862"></span>Reefer cut off = <b>ETA − 1h</b></div>
  <div class="rule"><span class="tick" style="background:var(--amber-strong)"></span>Open gate (empty return) = <b>ETD − 6d</b></div>
</div>

<div class="board" id="summary"></div>

<div class="controls">
  <div class="seg" id="viewSeg">
    <button data-v="desktop" class="active">🖥️ Computer</button>
    <button data-v="mobile">📱 Mobile</button>
  </div>
  <div class="seg" id="polSeg">
    <button data-v="ALL" class="active">All</button>
    <button data-v="THBKK">THBKK</button>
    <button data-v="THLCH">THLCH</button>
  </div>
  <input type="text" id="search" placeholder="Search vessel, voyage, service, wharf...">
  <select id="serviceFilter"><option value="ALL">All services</option></select>
  <select id="timeFilter">
    <option value="ALL">All time</option>
    <option value="7">Next 7 days</option>
    <option value="14">Next 14 days</option>
    <option value="30">Next 30 days</option>
  </select>
  <button class="clearbtn" id="clearBtn">Clear filters</button>
</div>

<div id="tableHolder"></div>
<div class="count" id="rowCount"></div>
</div>

<script>
const RAW = __RAW_JSON__;
const WHARF_MAP = __WHARF_JSON__;

function parseDT(s){ return new Date(s.replace(' ','T')+':00'); }
const NOW = new Date();

const data = RAW.map(r=>({
  ...r,
  etaDT: parseDT(r.eta),
  etdDT: parseDT(r.etd),
  cutoffDryDT: parseDT(r.cutoff_dry),
  cutoffReeferDT: parseDT(r.cutoff_reefer),
  opengateDT: parseDT(r.opengate),
}));

const state = { pol:'ALL', q:'', service:'ALL', time:'ALL', sortKey:'eta', sortDir:1, view: (window.innerWidth <= 760 ? 'mobile' : 'desktop') };

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function fmtDT(dt){
  const d = String(dt.getDate()).padStart(2,'0') + ' ' + MONTHS[dt.getMonth()];
  const t = dt.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',hour12:false});
  return {d,t};
}

function hoursUntil(dt){ return (dt - NOW)/36e5; }

function flagFor(dt, urgentHours, pastLabel, pastClass){
  const h = hoursUntil(dt);
  if(h < 0) return `<span class="flag ${pastClass || 'flag-past'}">${pastLabel || 'Closed'}</span>`;
  if(h <= urgentHours) return '<span class="flag flag-soon">Due soon</span>';
  return '';
}

// populate service filter
const services = [...new Set(data.map(r=>r.service))].sort();
const svcSel = document.getElementById('serviceFilter');
services.forEach(s=>{
  const o = document.createElement('option');
  o.value = s; o.textContent = s;
  svcSel.appendChild(o);
});

document.querySelectorAll('#viewSeg button').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('#viewSeg button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    state.view = btn.dataset.v;
    render();
  });
});
document.querySelectorAll('#polSeg button').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('#polSeg button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    state.pol = btn.dataset.v;
    render();
  });
});
document.getElementById('search').addEventListener('input', e=>{ state.q = e.target.value.trim().toLowerCase(); render(); });
svcSel.addEventListener('change', e=>{ state.service = e.target.value; render(); });
document.getElementById('timeFilter').addEventListener('change', e=>{ state.time = e.target.value; render(); });
document.getElementById('clearBtn').addEventListener('click', ()=>{
  state.pol='ALL'; state.q=''; state.service='ALL'; state.time='ALL';
  document.getElementById('search').value='';
  svcSel.value='ALL';
  document.getElementById('timeFilter').value='ALL';
  document.querySelectorAll('#polSeg button').forEach(b=>b.classList.remove('active'));
  document.querySelector('#polSeg button[data-v="ALL"]').classList.add('active');
  render();
});

const COLS = [
  {key:'pol', label:'POL'},
  {key:'service', label:'Service'},
  {key:'vessel', label:'Vessel / Voyage'},
  {key:'wharf', label:'Wharf'},
  {key:'pod', label:'Next Port'},
  {key:'eta', label:'ETA'},
  {key:'etd', label:'ETD'},
  {key:'opengate', label:'Open gate'},
  {key:'cutoff_dry', label:'Cut off (Dry)'},
  {key:'cutoff_reefer', label:'Cut off (Reefer)'},
];

function sortData(rows){
  const {sortKey, sortDir} = state;
  const dtKeyMap = {eta:'etaDT', etd:'etdDT', cutoff_dry:'cutoffDryDT', cutoff_reefer:'cutoffReeferDT', opengate:'opengateDT'};
  return rows.slice().sort((a,b)=>{
    let av, bv;
    if(dtKeyMap[sortKey]){ av=a[dtKeyMap[sortKey]]; bv=b[dtKeyMap[sortKey]]; return sortDir*(av-bv); }
    av = (a[sortKey]||'').toString().toLowerCase();
    bv = (b[sortKey]||'').toString().toLowerCase();
    if(av<bv) return -1*sortDir;
    if(av>bv) return 1*sortDir;
    return 0;
  });
}

function filterData(){
  let rows = data;
  if(state.pol !== 'ALL') rows = rows.filter(r=>r.pol === state.pol);
  if(state.service !== 'ALL') rows = rows.filter(r=>r.service === state.service);
  if(state.time !== 'ALL'){
    const days = parseInt(state.time,10);
    const limit = new Date(NOW.getTime() + days*24*3600*1000);
    rows = rows.filter(r=> r.etaDT <= limit && r.etaDT >= new Date(NOW.getTime()-24*3600*1000));
  }
  if(state.q){
    const q = state.q;
    rows = rows.filter(r=>
      r.vessel.toLowerCase().includes(q) ||
      r.vyg_bound.toLowerCase().includes(q) ||
      r.service.toLowerCase().includes(q) ||
      r.wharf.toLowerCase().includes(q) ||
      (WHARF_MAP[r.wharf]||'').toLowerCase().includes(q) ||
      r.pod.toLowerCase().includes(q) ||
      r.pol.toLowerCase().includes(q)
    );
  }
  return rows;
}

function render(){
  document.querySelectorAll('#viewSeg button').forEach(b=>{
    b.classList.toggle('active', b.dataset.v === state.view);
  });
  const filtered = sortData(filterData());
  renderSummary(filtered);
  if(state.view === 'mobile'){
    renderCards(filtered);
  } else {
    renderTable(filtered);
  }
  document.getElementById('rowCount').textContent = `Showing ${filtered.length} of ${data.length} voyages`;
}

function rowAccent(r){
  const dryPast = hoursUntil(r.cutoffDryDT) < 0;
  const reeferPast = hoursUntil(r.cutoffReeferDT) < 0;
  const dryDue = hoursUntil(r.cutoffDryDT) >= 0 && hoursUntil(r.cutoffDryDT) <= 24;
  const reeferDue = hoursUntil(r.cutoffReeferDT) >= 0 && hoursUntil(r.cutoffReeferDT) <= 6;
  if(dryPast || reeferPast) return 'var(--coral)';
  if(dryDue || reeferDue) return 'var(--amber-strong)';
  const gateDue = hoursUntil(r.opengateDT) >= 0 && hoursUntil(r.opengateDT) <= 24;
  if(gateDue) return 'var(--amber-strong)';
  return 'transparent';
}

function renderSummary(rows){
  const upcoming7 = rows.filter(r=> hoursUntil(r.etaDT) >=0 && hoursUntil(r.etaDT) <= 24*7).length;
  const nextCutoff = rows.filter(r=>hoursUntil(r.cutoffDryDT)>=0).sort((a,b)=>a.cutoffDryDT-b.cutoffDryDT)[0];
  const nextGate = rows.filter(r=>hoursUntil(r.opengateDT)>=0).sort((a,b)=>a.opengateDT-b.opengateDT)[0];
  const bkkCount = rows.filter(r=>r.pol==='THBKK').length;
  const lchCount = rows.filter(r=>r.pol==='THLCH').length;

  const cards = [
    {lbl:'Voyages shown', val: rows.length, small:false},
    {lbl:'THBKK / THLCH', val: `${bkkCount} / ${lchCount}`, small:false},
    {lbl:'Arriving next 7 days', val: upcoming7, small:false},
    {lbl:'Nearest cut off (dry)', val: nextCutoff ? `${nextCutoff.vessel} — ${fmtDT(nextCutoff.cutoffDryDT).d} ${fmtDT(nextCutoff.cutoffDryDT).t}` : '—', small:true},
    {lbl:'Nearest open gate', val: nextGate ? `${nextGate.vessel} — ${fmtDT(nextGate.opengateDT).d} ${fmtDT(nextGate.opengateDT).t}` : '—', small:true},
  ];
  document.getElementById('summary').innerHTML = cards.map(c=>`
    <div class="seg-stat"><div class="lbl">${c.lbl}</div><div class="val ${c.small?'small':''}">${c.val}</div></div>
  `).join('');
}

function renderTable(rows){
  const holder = document.getElementById('tableHolder');
  if(rows.length===0){
    holder.innerHTML = '<div class="empty">No results match the current filters</div>';
    return;
  }
  let thead = '<thead><tr>' + COLS.map(c=>{
    const sorted = state.sortKey===c.key;
    const arrow = sorted ? (state.sortDir===1?' ▲':' ▼') : '';
    return `<th data-key="${c.key}" class="${sorted?'sorted':''}">${c.label}${arrow}</th>`;
  }).join('') + '</tr></thead>';

  let tbody = '<tbody>' + rows.map(r=>{
    const eta = fmtDT(r.etaDT), etd = fmtDT(r.etdDT);
    const cd = fmtDT(r.cutoffDryDT), cr = fmtDT(r.cutoffReeferDT), og = fmtDT(r.opengateDT);
    const polClass = r.pol==='THBKK' ? 'pol-bkk' : 'pol-lch';
    const noLoad = r.no_l === 'Y' ? '<span class="noload">No load</span>' : '';
    return `<tr style="--rowaccent:${rowAccent(r)}">
      <td><span class="pol-badge ${polClass}">${r.pol}</span></td>
      <td>${r.service}</td>
      <td><div class="vessel">${r.vessel}</div><div class="sub">${r.vyg_bound} · ${r.op_liner}</div></td>
      <td><div class="vessel" style="font-weight:500">${WHARF_MAP[r.wharf]||r.wharf}</div><div class="sub">${r.wharf}</div></td>
      <td>${r.pod}</td>
      <td><div class="dtcell"><span class="d">${eta.d}</span><span class="t">${eta.t}</span></div></td>
      <td><div class="dtcell"><span class="d">${etd.d}</span><span class="t">${etd.t}</span></div></td>
      <td><div class="dtcell"><span class="d">${og.d}${flagFor(r.opengateDT,24,'Open','flag-open')}</span><span class="t">${og.t}</span></div></td>
      <td><div class="dtcell"><span class="d">${cd.d}${flagFor(r.cutoffDryDT,24,'Closed')}</span><span class="t">${cd.t}</span></div></td>
      <td><div class="dtcell"><span class="d">${cr.d}${flagFor(r.cutoffReeferDT,6,'Closed')}${noLoad}</span><span class="t">${cr.t}</span></div></td>
    </tr>`;
  }).join('') + '</tbody>';

  holder.innerHTML = `<table>${thead}${tbody}</table>`;

  holder.querySelectorAll('th').forEach(th=>{
    th.addEventListener('click', ()=>{
      const key = th.dataset.key;
      if(state.sortKey === key){ state.sortDir *= -1; }
      else { state.sortKey = key; state.sortDir = 1; }
      render();
    });
  });
}

function renderCards(rows){
  const holder = document.getElementById('tableHolder');
  if(rows.length===0){
    holder.innerHTML = '<div class="empty">No results match the current filters</div>';
    return;
  }
  holder.innerHTML = '<div class="mcards">' + rows.map(r=>{
    const eta = fmtDT(r.etaDT), etd = fmtDT(r.etdDT);
    const cd = fmtDT(r.cutoffDryDT), cr = fmtDT(r.cutoffReeferDT), og = fmtDT(r.opengateDT);
    const noLoad = r.no_l === 'Y' ? '<span class="noload">No load</span>' : '';
    return `<div class="mcard">
      <div class="mcard-main">
        <div class="mcard-top">
          <div>
            <div class="mcard-vessel">${r.vessel}</div>
            <div class="mcard-sub">${r.vyg_bound} · ${r.op_liner} · ${r.service}</div>
          </div>
        </div>
        <div class="mcard-grid">
          <div class="mcard-item"><div class="lbl">Wharf</div><div class="val">${WHARF_MAP[r.wharf]||r.wharf}</div></div>
          <div class="mcard-item"><div class="lbl">Next Port</div><div class="val">${r.pod}</div></div>
          <div class="mcard-item"><div class="lbl">ETA</div><div class="val">${eta.d}<span class="t">${eta.t}</span></div></div>
          <div class="mcard-item"><div class="lbl">ETD</div><div class="val">${etd.d}<span class="t">${etd.t}</span></div></div>
          <div class="mcard-item full"><div class="lbl">Open gate</div><div class="val">${og.d}<span class="t">${og.t}</span>${flagFor(r.opengateDT,24,'Open','flag-open')}</div></div>
          <div class="mcard-item"><div class="lbl">Cut off (Dry)</div><div class="val">${cd.d}<span class="t">${cd.t}</span>${flagFor(r.cutoffDryDT,24,'Closed')}</div></div>
          <div class="mcard-item"><div class="lbl">Cut off (Reefer)</div><div class="val">${cr.d}<span class="t">${cr.t}</span>${flagFor(r.cutoffReeferDT,6,'Closed')}${noLoad}</div></div>
        </div>
      </div>
      <div class="mcard-stub"><div class="mcard-stub-pol">${r.pol}</div></div>
    </div>`;
  }).join('') + '</div>';
}

render();
</script>
</body>
</html>
"""


def build_html(records, wharf_map, source_filename):
    html = TEMPLATE.replace("__RAW_JSON__", json.dumps(records, ensure_ascii=False))
    html = html.replace("__WHARF_JSON__", json.dumps(wharf_map, ensure_ascii=False))
    html = html.replace("__SOURCE_FILENAME__", source_filename)
    return html


def main():
    ap = argparse.ArgumentParser(description="Rebuild the Cut off / Open gate dashboard.")
    ap.add_argument("cutoff_xls", help="Path to the daily CUT_OFF-style .xls file")
    ap.add_argument("wharf_xls", help="Path to Wharf.xls (Wharf code -> Wharf Name)")
    ap.add_argument("-o", "--output", default="Cut off - Open Gate Dashboard.html",
                     help="Output HTML path (default: %(default)s)")
    args = ap.parse_args()

    records = load_cutoff_records(args.cutoff_xls)
    wharf_map = load_wharf_map(args.wharf_xls)
    check_wharf_coverage(records, wharf_map)

    html = build_html(records, wharf_map, args.cutoff_xls.split("/")[-1].split("\\\\")[-1])

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Done -> {args.output}")


if __name__ == "__main__":
    main()
