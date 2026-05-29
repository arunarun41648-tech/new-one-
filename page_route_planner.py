# page_route_planner.py  — Route Planner Admin Page
# Drop this file next to app.py and call page_route_planner() from your admin tabs
#
# DEPENDENCIES already in requirements.txt: streamlit, pandas, gspread, google-auth
# This page uses an embedded HTML/JS component (streamlit.components.v1) so no extra installs.

import streamlit as st
import streamlit.components.v1 as components
import json, uuid
from datetime import datetime, date

# ── import helpers from app.py ──────────────────────────────────────────────
# These are expected to be importable from app.py in the same directory.
# If you're calling this from inside app.py instead, just reference them directly.

ROUTE_PLANNER_HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
:root{
  --bg:#08101a;--surf:#0d1a28;--card:#0f1f30;--card2:#122437;
  --border:#1a2e42;--border2:#1e3550;
  --cyan:#00d4ff;--orange:#ff6b2b;--green:#2ecc71;
  --red:#e74c3c;--yellow:#f39c12;--purple:#9b59b6;
  --text:#d4e8f5;--muted:#4a6a85;--dim:#2a4560;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}
body::before{content:'';position:fixed;inset:0;
  background-image:linear-gradient(rgba(0,212,255,.02) 1px,transparent 1px),linear-gradient(90deg,rgba(0,212,255,.02) 1px,transparent 1px);
  background-size:50px 50px;pointer-events:none;z-index:0}
.wrap{position:relative;z-index:1;padding:14px 18px}
.hdr{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;
  background:var(--surf);border:1px solid var(--border);border-radius:12px;
  padding:16px 22px;margin-bottom:14px;position:relative;overflow:hidden}
.hdr::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--cyan),var(--orange),var(--green),var(--cyan))}
.hdr-brand h1{font-family:'Space Mono',monospace;font-size:1rem;color:var(--cyan);letter-spacing:2px;text-transform:uppercase}
.hdr-brand p{font-size:.68rem;color:var(--muted);font-family:'Space Mono',monospace;margin-top:2px}
.hdr-r{display:flex;align-items:center;gap:8px;flex-wrap:wrap}

/* BUTTONS */
.btn{display:inline-flex;align-items:center;gap:5px;padding:7px 14px;border:none;border-radius:7px;
  cursor:pointer;font-family:'Space Mono',monospace;font-size:.63rem;letter-spacing:.8px;
  text-transform:uppercase;font-weight:700;transition:all .15s;white-space:nowrap}
.bc{background:var(--cyan);color:#08101a}.bc:hover{background:#33ddff;transform:translateY(-1px)}
.bo{background:linear-gradient(135deg,var(--orange),#e74c3c);color:#fff}.bo:hover{filter:brightness(1.1)}
.bg{background:rgba(46,204,113,.12);color:var(--green);border:1px solid rgba(46,204,113,.3)}.bg:hover{background:rgba(46,204,113,.2)}
.bgh{background:rgba(255,255,255,.04);color:var(--muted);border:1px solid var(--border)}.bgh:hover{border-color:var(--cyan);color:var(--cyan)}
.bsm{padding:4px 9px;font-size:.58rem}
.bred{background:rgba(231,76,60,.12);color:var(--red);border:1px solid rgba(231,76,60,.25)}.bred:hover{background:rgba(231,76,60,.22)}
.bpur{background:rgba(155,89,182,.12);color:var(--purple);border:1px solid rgba(155,89,182,.3)}.bpur:hover{background:rgba(155,89,182,.22)}
.bsubmit{background:linear-gradient(135deg,var(--green),#27ae60);color:#fff;font-size:.75rem;padding:10px 22px;border-radius:9px;box-shadow:0 4px 15px rgba(46,204,113,.3)}
.bsubmit:hover{filter:brightness(1.1);transform:translateY(-1px);box-shadow:0 6px 20px rgba(46,204,113,.4)}
.bsubmit:disabled{opacity:.4;cursor:not-allowed;transform:none}

/* STATS */
.stats-bar{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px}
.sc{flex:1;min-width:110px;background:var(--surf);border:1px solid var(--border);border-radius:10px;padding:11px 14px;position:relative;overflow:hidden}
.sc::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--cyan)}
.sc.o::before{background:var(--orange)}.sc.g::before{background:var(--green)}
.sc.y::before{background:var(--yellow)}.sc.r::before{background:var(--red)}
.sl2{font-family:'Space Mono',monospace;font-size:.52rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-bottom:3px}
.sv{font-family:'Space Mono',monospace;font-size:1.15rem;color:var(--text);font-weight:700}
.ss{font-size:.58rem;color:var(--muted);margin-top:1px}

/* TABS */
.tabs{display:flex;gap:3px;background:var(--surf);border:1px solid var(--border);border-radius:10px;padding:3px;margin-bottom:14px;flex-wrap:wrap}
.tab{flex:1;min-width:90px;text-align:center;padding:7px 10px;border-radius:7px;cursor:pointer;
  font-family:'Space Mono',monospace;font-size:.58rem;letter-spacing:.8px;text-transform:uppercase;
  color:var(--muted);transition:all .18s}
.tab.active{background:var(--card2);color:var(--cyan);border:1px solid rgba(0,212,255,.2)}
.tab-content{display:none}.tab-content.active{display:block}

/* PANELS */
.panel{background:var(--surf);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:14px}
.ph{display:flex;align-items:center;justify-content:space-between;padding:11px 16px;
  border-bottom:1px solid var(--border);background:rgba(0,0,0,.2);flex-wrap:wrap;gap:8px}
.pt{font-family:'Space Mono',monospace;font-size:.62rem;color:var(--muted);text-transform:uppercase;letter-spacing:1.5px}
.pb{padding:16px}

/* TABLE */
.rt{width:100%;border-collapse:collapse;font-size:.74rem}
.rt th{font-family:'Space Mono',monospace;font-size:.54rem;color:var(--muted);text-transform:uppercase;
  letter-spacing:1px;padding:7px 9px;border-bottom:1px solid var(--border);text-align:left;white-space:nowrap}
.rt td{padding:7px 9px;border-bottom:1px solid rgba(26,46,66,.4);vertical-align:middle}
.rt tr:hover td{background:rgba(0,212,255,.02)}
.rt tr:last-child td{border:none}

/* TRIP COLORS */
.tc0{background:linear-gradient(135deg,#c0392b,#e74c3c)}
.tc1{background:linear-gradient(135deg,#d35400,#e67e22)}
.tc2{background:linear-gradient(135deg,#b7950b,#f39c12)}
.tc3{background:linear-gradient(135deg,#1a8c4e,#27ae60)}
.tc4{background:linear-gradient(135deg,#148f77,#1abc9c)}
.tc5{background:linear-gradient(135deg,#1f618d,#2980b9)}
.tc6{background:linear-gradient(135deg,#7d3c98,#9b59b6)}
.tc7{background:linear-gradient(135deg,#a93226,#e91e63)}
.tc8{background:linear-gradient(135deg,#bf360c,#ff5722)}
.tc9{background:linear-gradient(135deg,#4e342e,#795548)}
.tc10{background:linear-gradient(135deg,#00695c,#00897b)}
.tc11{background:linear-gradient(135deg,#283593,#3949ab)}
.trip-tag{display:inline-block;padding:2px 7px;border-radius:4px;font-family:'Space Mono',monospace;font-size:.58rem;font-weight:700;color:#fff}
const TCLS=['tc0','tc1','tc2','tc3','tc4','tc5','tc6','tc7','tc8','tc9','tc10','tc11'];

/* ROUTEMAP CARDS */
.rm-grid{display:flex;gap:8px;overflow-x:auto;padding-bottom:8px;flex-wrap:wrap}
.trip-col{min-width:160px;max-width:175px;flex-shrink:0;border-radius:9px;overflow:hidden;
  border:1px solid var(--border);background:var(--card);animation:fup .3s ease}
@keyframes fup{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
.trip-hdr{padding:6px 8px 4px;font-size:.6rem;font-weight:700;color:#fff;line-height:1.3}
.trip-title{font-family:'Space Mono',monospace;letter-spacing:.4px;font-size:.62rem}
.trip-cc{font-size:.54rem;opacity:.8;margin-top:1px}
.trip-ton{background:rgba(0,0,0,.35);padding:3px 8px;font-family:'Space Mono',monospace;font-size:.63rem;
  color:#fff;border-bottom:1px solid rgba(255,255,255,.08);display:flex;justify-content:space-between;align-items:center}
.trip-customers{padding:4px 6px;display:flex;flex-direction:column;gap:3px}
.cust-card{background:rgba(0,0,0,.18);border-radius:5px;padding:5px 7px;border:1px solid rgba(255,255,255,.04)}
.cust-name{font-size:.67rem;font-weight:600;color:var(--text);line-height:1.2;margin-bottom:2px}
.cust-meta{font-family:'Space Mono',monospace;font-size:.55rem;color:var(--muted);line-height:1.7}
.cust-meta .cr{color:var(--cyan)}.cust-meta .ti{color:var(--yellow)}.cust-meta .tn{color:var(--orange)}
.snum{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;
  border-radius:50%;font-family:'Space Mono',monospace;font-size:.5rem;font-weight:700;
  background:var(--cyan);color:#08101a;margin-right:3px;vertical-align:middle;flex-shrink:0}

/* MAP */
#leafMap{height:480px;width:100%;border-radius:0}

/* LOG */
.logbox{font-family:'Space Mono',monospace;font-size:.6rem;max-height:120px;overflow-y:auto;display:flex;flex-direction:column;gap:2px}
.le{display:flex;gap:8px;line-height:1.6;animation:fi .2s ease}
@keyframes fi{from{opacity:0}to{opacity:1}}
.lt2{color:var(--dim);flex-shrink:0}.lm{color:var(--text)}
.lok .lm{color:var(--green)}.lwarn .lm{color:var(--yellow)}.lerr .lm{color:var(--red)}.linfo .lm{color:var(--cyan)}

/* EMPTY */
.empty{text-align:center;padding:32px;color:var(--muted);font-family:'Space Mono',monospace;font-size:.68rem}
.eico{font-size:1.8rem;margin-bottom:8px;opacity:.3}

/* SUBMIT PANEL */
.submit-panel{background:linear-gradient(135deg,rgba(46,204,113,.08),rgba(0,212,255,.05));
  border:1.5px solid rgba(46,204,113,.25);border-radius:14px;padding:20px 24px;margin:14px 0}
.submit-title{font-family:'Space Mono',monospace;font-size:.75rem;color:var(--green);text-transform:uppercase;letter-spacing:1.5px;margin-bottom:12px}

/* DRIVER ASSIGN */
.driver-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px;margin-top:10px}
.driver-card{background:var(--card2);border:1.5px solid var(--border);border-radius:10px;padding:12px 14px;
  cursor:pointer;transition:all .15s;position:relative}
.driver-card:hover{border-color:rgba(0,212,255,.4)}
.driver-card.selected{border-color:var(--green)!important;background:rgba(46,204,113,.08)!important}
.driver-card.selected::after{content:'✓';position:absolute;top:8px;right:10px;
  color:var(--green);font-weight:700;font-size:.85rem}
.driver-name{font-weight:600;font-size:.82rem;margin-bottom:3px}
.driver-meta{font-family:'Space Mono',monospace;font-size:.56rem;color:var(--muted);line-height:1.7}
.driver-status{display:inline-block;font-size:.58rem;padding:2px 8px;border-radius:10px;font-weight:600;margin-top:4px}
.ds-active{background:rgba(46,204,113,.15);color:var(--green)}
.ds-offline{background:rgba(74,106,133,.15);color:var(--muted)}

/* TRIP-DRIVER ASSIGN TABLE */
.assign-row{display:grid;grid-template-columns:120px 1fr 220px;gap:10px;align-items:center;
  padding:8px 12px;border-bottom:1px solid rgba(26,46,66,.5);font-size:.78rem}
.assign-row:last-child{border:none}
.assign-hdr{font-family:'Space Mono',monospace;font-size:.55rem;color:var(--muted);text-transform:uppercase;
  background:rgba(0,0,0,.2);padding:6px 12px;border-bottom:1px solid var(--border)}

/* FILTER BAR */
.filter-bar{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:12px 16px;
  background:rgba(0,0,0,.15);border-bottom:1px solid var(--border)}
.filter-bar label{font-family:'Space Mono',monospace;font-size:.58rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px}
.filter-bar select,.filter-bar input{background:var(--card2);border:1px solid var(--border);
  color:var(--text);padding:6px 10px;border-radius:7px;font-family:'DM Sans',sans-serif;font-size:.8rem;outline:none}
.filter-bar select:focus,.filter-bar input:focus{border-color:var(--cyan)}

/* CONFIRM TOAST */
.toast{position:fixed;bottom:24px;right:24px;background:var(--green);color:#fff;
  padding:12px 20px;border-radius:10px;font-family:'Space Mono',monospace;font-size:.7rem;
  font-weight:700;z-index:9999;animation:slideUp .3s ease;display:none;box-shadow:0 4px 20px rgba(46,204,113,.4)}
.toast.show{display:block}
@keyframes slideUp{from{transform:translateY(20px);opacity:0}to{transform:translateY(0);opacity:1}}

/* PROGRESS */
.prog-bar{height:4px;background:var(--border);border-radius:2px;overflow:hidden;margin-top:6px}
.prog-fill{height:100%;background:linear-gradient(90deg,var(--cyan),var(--green));border-radius:2px;transition:width .3s}

select option{background:var(--card2)}
</style>
</head>
<body>
<div class="wrap">

<div class="hdr">
  <div class="hdr-brand">
    <h1>🧄 Route Planner</h1>
    <p>Load orders → Optimize routes → Assign drivers → Create trips</p>
  </div>
  <div class="hdr-r">
    <span id="dateLabel" style="font-family:'Space Mono',monospace;font-size:.65rem;color:var(--yellow)">No date selected</span>
    <button class="btn bc" onclick="runOptimize()">▶ OPTIMIZE ROUTES</button>
  </div>
</div>

<!-- STATS -->
<div class="stats-bar" id="statsBar">
  <div class="sc"><div class="sl2">Orders</div><div class="sv" id="s-ord">0</div><div class="ss">from sheet</div></div>
  <div class="sc o"><div class="sl2">Trips</div><div class="sv" id="s-trips">0</div><div class="ss">auto-grouped</div></div>
  <div class="sc g"><div class="sl2">Total Crates</div><div class="sv" id="s-crates">0</div><div class="ss">units</div></div>
  <div class="sc y"><div class="sl2">Total kg</div><div class="sv" id="s-ton">0</div><div class="ss">tonnage</div></div>
  <div class="sc r"><div class="sl2">Est. Distance</div><div class="sv" id="s-dist">—</div><div class="ss">all trips</div></div>
</div>

<!-- TABS -->
<div class="tabs">
  <div class="tab active" onclick="showTab('orders')">📋 Orders</div>
  <div class="tab" onclick="showTab('routemap')">🗺 Route Cards</div>
  <div class="tab" onclick="showTab('map')">📍 Live Map</div>
  <div class="tab" onclick="showTab('table')">📊 Route Table</div>
  <div class="tab" onclick="showTab('submit')">✅ Submit Trips</div>
</div>

<!-- TAB: ORDERS -->
<div id="tab-orders" class="tab-content active">
  <div class="panel">
    <div class="ph">
      <span class="pt">Orders for selected date</span>
      <div style="display:flex;gap:6px;align-items:center">
        <span id="ordCountLbl" style="font-family:'Space Mono',monospace;font-size:.6rem;color:var(--muted)">0 orders</span>
        <button class="btn bc bsm" onclick="runOptimize()">▶ RUN OPTIMIZER</button>
        <button class="btn bred bsm" onclick="clearOrders()">✕ CLEAR</button>
      </div>
    </div>
    <div style="overflow-x:auto">
      <table class="rt">
        <thead><tr>
          <th>#</th><th>Order ID</th><th>Customer</th><th>Address</th>
          <th>Crates</th><th>Tonnage (kg)</th><th>Slot</th><th>Lat</th><th>Lng</th>
        </tr></thead>
        <tbody id="ordBody"><tr><td colspan="9"><div class="empty"><div class="eico">📋</div>Select a date above to load orders</div></td></tr></tbody>
      </table>
    </div>
  </div>
</div>

<!-- TAB: ROUTE MAP -->
<div id="tab-routemap" class="tab-content">
  <div class="panel">
    <div class="ph">
      <span class="pt">Optimized Route Cards</span>
      <span id="rmSub" style="font-family:'Space Mono',monospace;font-size:.58rem;color:var(--muted)">Run optimizer to populate</span>
    </div>
    <div class="pb">
      <div class="rm-grid" id="rmGrid">
        <div class="empty"><div class="eico">🗺️</div>Load orders and run optimizer</div>
      </div>
    </div>
  </div>
</div>

<!-- TAB: LIVE MAP -->
<div id="tab-map" class="tab-content">
  <div class="panel" style="overflow:hidden">
    <div class="ph">
      <span class="pt">Live Map — OpenStreetMap</span>
      <div style="display:flex;gap:8px;align-items:center">
        <select id="mapTripFilter" onchange="filterMapTrip(this.value)"
          style="background:var(--card2);border:1px solid var(--border);color:var(--text);padding:5px 9px;border-radius:6px;font-size:.72rem;outline:none">
          <option value="all">All Trips</option>
        </select>
        <button class="btn bgh bsm" onclick="mapFitAll()">⊞ FIT ALL</button>
      </div>
    </div>
    <div id="leafMap"></div>
  </div>
</div>

<!-- TAB: ROUTE TABLE -->
<div id="tab-table" class="tab-content">
  <div class="panel">
    <div class="ph">
      <span class="pt">Optimized Route Sequence</span>
      <span style="font-family:'Space Mono',monospace;font-size:.55rem;color:var(--dim);background:rgba(0,212,255,.06);border:1px solid rgba(0,212,255,.15);padding:2px 8px;border-radius:4px">Greedy Seed + 2-OPT + OR-OPT</span>
    </div>
    <div style="overflow-x:auto">
      <table class="rt">
        <thead><tr>
          <th>Stop</th><th>Trip</th><th>Order ID</th><th>Customer</th><th>Address</th>
          <th>Crates</th><th>kg</th><th>Slot</th><th>Leg km</th><th>Cum km</th><th>ETA</th>
        </tr></thead>
        <tbody id="routeBody"><tr><td colspan="11"><div class="empty"><div class="eico">📍</div>Run optimizer</div></td></tr></tbody>
      </table>
    </div>
  </div>
</div>

<!-- TAB: SUBMIT TRIPS -->
<div id="tab-submit" class="tab-content">
  <div class="submit-panel" id="submitPanel">
    <div class="submit-title">📋 Trip Summary — Ready to Submit</div>
    <div id="submitSummary">
      <div class="empty"><div class="eico">⚡</div>Run optimizer first to see trips here</div>
    </div>
  </div>

  <!-- DRIVER ASSIGNMENT -->
  <div class="panel" id="assignPanel" style="display:none">
    <div class="ph">
      <span class="pt">🚚 Assign Drivers to Trips</span>
      <span style="font-family:'Space Mono',monospace;font-size:.58rem;color:var(--muted)">Select a driver for each trip (optional)</span>
    </div>
    <div class="pb">
      <div id="assignTable">
        <div class="empty"><div class="eico">🚚</div>Loading drivers...</div>
      </div>
    </div>
  </div>

  <!-- SUBMIT BUTTON -->
  <div id="submitBtnArea" style="display:none;text-align:center;padding:20px 0 10px">
    <div style="margin-bottom:12px;font-family:'Space Mono',monospace;font-size:.65rem;color:var(--muted)">
      Review the trip cards above, assign drivers, then submit to create trips in the system.
    </div>
    <button class="btn bsubmit" id="submitBtn" onclick="submitTrips()">
      ✅ &nbsp;CREATE TRIPS IN SYSTEM
    </button>
    <div id="submitProgress" style="display:none;margin-top:14px">
      <div style="font-family:'Space Mono',monospace;font-size:.62rem;color:var(--cyan);margin-bottom:6px" id="submitMsg">Submitting...</div>
      <div class="prog-bar"><div class="prog-fill" id="submitFill" style="width:0%"></div></div>
    </div>
  </div>
</div>

<!-- LOG -->
<div class="panel">
  <div class="ph">
    <span class="pt">System Log</span>
    <button class="btn bgh bsm" onclick="clearLog()">CLEAR</button>
  </div>
  <div class="pb" style="padding:10px 14px">
    <div class="logbox" id="cronLog"></div>
  </div>
</div>

</div><!-- /wrap -->

<div id="toast" class="toast"></div>

<script>
// ══════════════════════════════════════════════════════════════════
// STATE
// ══════════════════════════════════════════════════════════════════
let orders = [];
let optimizedTrips = {};
let leafMap = null;
let mapLayers = [];
let selectedMapTrip = null;
let drivers = [];
let tripDriverMap = {};   // { tripKey: {uid, name} }

const TCLS = ['tc0','tc1','tc2','tc3','tc4','tc5','tc6','tc7','tc8','tc9','tc10','tc11'];
const DEPOT = {lat: 12.9716, lng: 77.5946, name: 'Warehouse'};

// ══════════════════════════════════════════════════════════════════
// INIT — receive data from Streamlit via window.initData
// ══════════════════════════════════════════════════════════════════
window.addEventListener('message', e => {
  if (e.data && e.data.type === 'INIT_DATA') {
    const d = e.data;
    if (d.orders)  loadOrders(d.orders);
    if (d.drivers) loadDrivers(d.drivers);
    if (d.date)    document.getElementById('dateLabel').textContent = '📅 ' + d.date;
  }
  if (e.data && e.data.type === 'LOAD_ORDERS') {
    loadOrders(e.data.orders);
    document.getElementById('dateLabel').textContent = '📅 ' + (e.data.date || '');
  }
});

// Also accept direct JS call from Streamlit component iframe bridge
function receiveData(json) {
  const d = JSON.parse(json);
  if (d.orders)  loadOrders(d.orders);
  if (d.drivers) loadDrivers(d.drivers);
  if (d.date)    document.getElementById('dateLabel').textContent = '📅 ' + d.date;
}

// ══════════════════════════════════════════════════════════════════
// LOAD ORDERS
// ══════════════════════════════════════════════════════════════════
function loadOrders(raw) {
  orders = raw.map((r, i) => ({
    id:       r.id        || r['SaleOrderId']   || r['Order ID']    || `ORD-${String(i+1).padStart(3,'0')}`,
    customer: r.customer  || r['Customer']      || r['Customer shop name'] || `Customer ${i+1}`,
    address:  r.address   || r['Shop Location'] || r['address']     || '—',
    lat:      parseFloat(r.lat || r['Latitude'] || 0),
    lng:      parseFloat(r.lng || r['Longitude']|| 0),
    crates:   parseFloat(r.crates || r['TotalCrates'] || r['OrderedQty'] || 0),
    tonnage:  parseFloat(r.tonnage|| r['OrderKg']     || r['OrderTotal']  || 0),
    window:   r.window    || r['Slot']          || r['DeliverySlot'] || '07:00-08:00',
    priority: r.priority  || 'med',
    custId:   r.custId    || r['CustomerId']    || r['CustomerId']   || '',
    trip:     parseInt(r.trip || r['Tripid'] || 1) || 1,
  }));
  // Re-number trips 1,2,3...
  const raw_keys = [...new Set(orders.map(o => o.trip))].sort((a,b) => a-b);
  const remap = {}; raw_keys.forEach((k,i) => remap[k] = i+1);
  orders.forEach(o => { o.trip = remap[o.trip] || o.trip; });
  renderOrdersTable();
  updateStats();
  log(`✓ Loaded ${orders.length} orders across ${[...new Set(orders.map(o=>o.trip))].length} trips`, 'lok');
}

function loadDrivers(raw) {
  drivers = raw.map(d => ({
    uid:     d.uid    || d['Driver ID'] || '',
    name:    d.name   || d['Full Name'] || '',
    vehicle: d.vehicle|| d['Vehicle Type'] || '',
    vnum:    d.vnum   || d['Vehicle Number'] || '',
    status:  (d.status|| d['Active Status'] || 'Offline').toLowerCase(),
  }));
  log(`✓ Loaded ${drivers.length} drivers`, 'lok');
  renderAssignTable();
}

// ══════════════════════════════════════════════════════════════════
// RENDER ORDERS TABLE
// ══════════════════════════════════════════════════════════════════
function renderOrdersTable() {
  const tb = document.getElementById('ordBody');
  document.getElementById('ordCountLbl').textContent = orders.length + ' orders';
  if (!orders.length) {
    tb.innerHTML = '<tr><td colspan="9"><div class="empty"><div class="eico">📋</div>No orders for this date</div></td></tr>';
    return;
  }
  const ti = {}; [...new Set(orders.map(o=>o.trip))].sort((a,b)=>a-b).forEach((t,i) => ti[t]=i);
  tb.innerHTML = orders.map((o,i) => `<tr>
    <td style="font-family:'Space Mono',monospace;font-size:.6rem;color:var(--muted)">${i+1}</td>
    <td style="font-family:'Space Mono',monospace;font-size:.62rem;color:var(--cyan)">${o.id}</td>
    <td style="font-weight:600;max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${o.customer}</td>
    <td style="font-size:.68rem;color:var(--muted);max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${o.address}</td>
    <td style="font-family:'Space Mono',monospace;color:var(--cyan)">${o.crates}</td>
    <td style="font-family:'Space Mono',monospace;color:var(--orange)">${o.tonnage}</td>
    <td style="font-family:'Space Mono',monospace;font-size:.62rem;color:var(--yellow)">${o.window}</td>
    <td style="font-family:'Space Mono',monospace;font-size:.58rem;color:var(--muted)">${o.lat ? o.lat.toFixed(4) : '—'}</td>
    <td style="font-family:'Space Mono',monospace;font-size:.58rem;color:var(--muted)">${o.lng ? o.lng.toFixed(4) : '—'}</td>
  </tr>`).join('');
}

// ══════════════════════════════════════════════════════════════════
// STATS
// ══════════════════════════════════════════════════════════════════
function updateStats() {
  document.getElementById('s-ord').textContent   = orders.length;
  document.getElementById('s-trips').textContent = Object.keys(optimizedTrips).length || [...new Set(orders.map(o=>o.trip))].length;
  document.getElementById('s-crates').textContent= orders.reduce((s,o) => s+o.crates, 0).toFixed(0);
  document.getElementById('s-ton').textContent   = orders.reduce((s,o) => s+o.tonnage, 0).toFixed(1);
}

// ══════════════════════════════════════════════════════════════════
// OPTIMIZATION ENGINE  (Greedy Seed + 2-opt + Or-opt)
// ══════════════════════════════════════════════════════════════════
function hav(a, b) {
  if (!a.lat || !b.lat || !a.lng || !b.lng) return 5;
  const R=6371, dLa=(b.lat-a.lat)*Math.PI/180, dLn=(b.lng-a.lng)*Math.PI/180;
  const s = Math.sin(dLa/2)**2 + Math.cos(a.lat*Math.PI/180)*Math.cos(b.lat*Math.PI/180)*Math.sin(dLn/2)**2;
  return R*2*Math.atan2(Math.sqrt(s), Math.sqrt(1-s));
}
function parseWindowStart(win) {
  if (!win || typeof win !== 'string') return 480;
  const m = win.match(/(\d{1,2}):(\d{2})/);
  return m ? +m[1]*60+(+m[2]) : 480;
}
const PW = {high:0, med:2, low:4};
function routeDist(stops, depot) {
  if (!stops.length) return 0;
  let d = hav(depot, stops[0]);
  for (let i=1; i<stops.length; i++) d += hav(stops[i-1], stops[i]);
  return d + hav(stops[stops.length-1], depot);
}
function greedySeed(stops, depot) {
  if (!stops.length) return [];
  let unvis=[...stops], route=[], cur=depot;
  while (unvis.length) {
    let best=null, bs=Infinity;
    unvis.forEach(o => {
      const d = hav(cur, o);
      const score = d + (PW[o.priority]||2)*1.5 + (parseWindowStart(o.window)/60)*0.3;
      if (score < bs) { bs=score; best=o; }
    });
    unvis = unvis.filter(o => o !== best);
    route.push(best); cur=best;
  }
  return route;
}
function twoOpt(route, depot) {
  if (route.length < 4) return route;
  let best=[...route], bestD=routeDist(best,depot), improved=true, iters=0;
  while (improved && iters<200) {
    improved=false; iters++;
    for (let i=0; i<best.length-1; i++) {
      for (let j=i+2; j<best.length; j++) {
        const c=[...best.slice(0,i+1),...best.slice(i+1,j+1).reverse(),...best.slice(j+1)];
        const d=routeDist(c,depot);
        if (d < bestD-0.0001) { bestD=d; best=c; improved=true; }
      }
    }
  }
  return best;
}
function orOpt1(route, depot) {
  if (route.length < 3) return route;
  let best=[...route], bestD=routeDist(best,depot), improved=true, iters=0;
  while (improved && iters<150) {
    improved=false; iters++;
    for (let i=0; i<best.length; i++) {
      const node=best[i], without=best.filter((_,idx)=>idx!==i);
      for (let j=0; j<=without.length; j++) {
        const c=[...without.slice(0,j),node,...without.slice(j)];
        const d=routeDist(c,depot);
        if (d < bestD-0.0001) { bestD=d; best=c; improved=true; break; }
      }
      if (improved) break;
    }
  }
  return best;
}
function optTrip(stops) {
  if (!stops.length) return [];
  const withC  = stops.filter(o => o.lat && o.lng && Math.abs(o.lat) > 0.001);
  const noC    = stops.filter(o => !o.lat || !o.lng || Math.abs(o.lat) <= 0.001);
  let route    = greedySeed(withC, DEPOT);
  if (route.length >= 4) route = twoOpt(route, DEPOT);
  if (route.length >= 3) route = orOpt1(route, DEPOT);
  const before = routeDist(greedySeed(withC, DEPOT), DEPOT).toFixed(2);
  const after  = routeDist(route.filter(o=>o.lat&&o.lng), DEPOT).toFixed(2);
  log(`Trip opt: seed ${before}km → optimized ${after}km (${((before-after)/before*100).toFixed(1)}% saved)`, 'lok');
  return [...route, ...noC];
}
function buildRouteMeta(route) {
  let cum=0, prev=DEPOT, elapsed=0;
  return route.map((o,i) => {
    const d=hav(prev,o); cum+=d;
    elapsed += Math.round(d/35*60) + 5;
    const hh=7+Math.floor(elapsed/60), mm=elapsed%60;
    prev=o;
    return {...o, stop:i+1, legDist:d.toFixed(2), cumDist:cum.toFixed(2),
            eta:`${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`,_cumKm:cum};
  });
}
function runOptimize() {
  if (!orders.length) { log('⚠ No orders loaded — select a date first', 'lwarn'); return; }
  log(`▶ Optimizing ${orders.length} orders...`, 'linfo');
  const tripKeys = [...new Set(orders.map(o => o.trip))].sort((a,b) => a-b);
  let grand=0; optimizedTrips={};
  tripKeys.forEach(tk => {
    const sorted = optTrip(orders.filter(o => o.trip==tk));
    const meta   = buildRouteMeta(sorted);
    optimizedTrips[tk] = meta;
    const last   = meta[meta.length-1];
    grand += (last ? +last.cumDist + hav(last, DEPOT) : 0);
    log(`✓ Trip ${tk}: ${meta.length} stops`, 'lok');
  });
  document.getElementById('s-dist').textContent = grand.toFixed(1)+'km';
  document.getElementById('s-trips').textContent = tripKeys.length;
  document.getElementById('rmSub').textContent   = `${tripKeys.length} trips · ${orders.length} stops · ${grand.toFixed(1)}km · 2-OPT+OR-OPT`;
  renderRoutemapCards();
  renderRouteTable();
  renderSubmitSummary();
  renderAssignTable();
  updateMapTripFilter();
  if (leafMap) { clearMapLayers(); drawAllTrips(); }
  showTab('routemap');
  log(`✓ Done: ${tripKeys.length} trips, ${grand.toFixed(1)}km`, 'lok');
}

// ══════════════════════════════════════════════════════════════════
// ROUTEMAP CARDS
// ══════════════════════════════════════════════════════════════════
function renderRoutemapCards() {
  const g    = document.getElementById('rmGrid');
  const keys = Object.keys(optimizedTrips).sort((a,b) => +a-+b);
  if (!keys.length) { g.innerHTML='<div class="empty"><div class="eico">🗺️</div>Run optimizer</div>'; return; }
  g.innerHTML = keys.map((tk,ci) => {
    const custs = optimizedTrips[tk];
    const ton   = custs.reduce((s,c) => s+(+c.tonnage||0), 0).toFixed(1);
    const km    = custs.length ? custs[custs.length-1].cumDist : '—';
    return `<div class="trip-col">
      <div class="trip-hdr ${TCLS[ci%12]}">
        <div class="trip-title">Trip ${tk}</div>
        <div class="trip-cc">${custs.length} stops · ${km} km</div>
      </div>
      <div class="trip-ton"><span>Tonnage: ${ton} kg</span><span style="opacity:.6;font-size:.55rem">${custs.length} stops</span></div>
      <div class="trip-customers">
        ${custs.map((c,i) => `<div class="cust-card">
          <div class="cust-name"><span class="snum">${c.stop}</span>${c.customer}</div>
          <div class="cust-meta">
            <span class="cr">Crates: ${c.crates}</span>&nbsp;
            <span class="ti">${c.window}</span><br>
            <span class="tn">ETA: ${c.eta}</span>
          </div>
        </div>`).join('')}
      </div>
    </div>`;
  }).join('');
}

// ══════════════════════════════════════════════════════════════════
// ROUTE TABLE
// ══════════════════════════════════════════════════════════════════
function renderRouteTable() {
  const tb   = document.getElementById('routeBody');
  const rows = Object.values(optimizedTrips).flat();
  if (!rows.length) { tb.innerHTML='<tr><td colspan="11"><div class="empty"><div class="eico">📍</div>Run optimizer</div></td></tr>'; return; }
  const ti = {}; Object.keys(optimizedTrips).sort((a,b)=>+a-+b).forEach((t,i) => ti[t]=i);
  tb.innerHTML = rows.map(r => `<tr>
    <td><div style="display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;background:var(--cyan);color:#08101a;font-family:'Space Mono',monospace;font-size:.58rem;font-weight:700">${r.stop}</div></td>
    <td><span class="trip-tag ${TCLS[ti[r.trip]%12]}">${r.trip}</span></td>
    <td style="font-family:'Space Mono',monospace;font-size:.6rem;color:var(--cyan)">${r.id}</td>
    <td style="font-weight:600">${r.customer}</td>
    <td style="font-size:.68rem;color:var(--muted)">${r.address}</td>
    <td style="font-family:'Space Mono',monospace;color:var(--cyan)">${r.crates}</td>
    <td style="font-family:'Space Mono',monospace;color:var(--orange)">${r.tonnage}</td>
    <td style="font-family:'Space Mono',monospace;font-size:.62rem;color:var(--yellow)">${r.window}</td>
    <td style="font-family:'Space Mono',monospace;font-size:.65rem">${r.legDist} km</td>
    <td style="font-family:'Space Mono',monospace;font-size:.65rem;color:var(--muted)">${r.cumDist} km</td>
    <td style="font-family:'Space Mono',monospace;color:var(--green)">${r.eta}</td>
  </tr>`).join('');
}

// ══════════════════════════════════════════════════════════════════
// MAP
// ══════════════════════════════════════════════════════════════════
function initMap() {
  if (leafMap) return;
  leafMap = L.map('leafMap').setView([12.9716, 77.5946], 11);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution:'© OpenStreetMap', maxZoom:18
  }).addTo(leafMap);
  log('Map initialized', 'lok');
}
function clearMapLayers() {
  mapLayers.forEach(l => { try { leafMap.removeLayer(l); } catch(e){} });
  mapLayers = [];
}
const MCOLORS = ['#e74c3c','#e67e22','#f39c12','#27ae60','#1abc9c','#2980b9','#9b59b6','#e91e63','#ff5722','#795548','#00897b','#3949ab'];
function drawAllTrips() {
  if (!leafMap) return;
  clearMapLayers();
  const keys = Object.keys(optimizedTrips).sort((a,b)=>+a-+b);
  let allBounds = [];
  keys.forEach((tk,ci) => {
    const route  = optimizedTrips[tk];
    const color  = MCOLORS[ci % MCOLORS.length];
    const coords = route.filter(o=>o.lat&&o.lng).map(o=>[o.lat,o.lng]);
    if (coords.length) {
      const pl = L.polyline([[DEPOT.lat,DEPOT.lng],...coords,[DEPOT.lat,DEPOT.lng]], {color,weight:2.5,opacity:.7});
      pl.addTo(leafMap); mapLayers.push(pl);
      coords.forEach(c => allBounds.push(c));
    }
    route.forEach((o,i) => {
      if (!o.lat||!o.lng) return;
      const icon = L.divIcon({className:'',html:`<div style="background:${color};color:#fff;border-radius:50%;width:22px;height:22px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;font-family:monospace;border:2px solid #fff;box-shadow:0 2px 4px rgba(0,0,0,.4)">${o.stop}</div>`,iconSize:[22,22],iconAnchor:[11,11]});
      const mk = L.marker([o.lat,o.lng],{icon});
      mk.bindPopup(`<div style="font-family:'DM Sans',sans-serif;font-size:12px;background:#0d1a28;color:#d4e8f5;padding:6px;min-width:150px"><b style="color:${color}">Stop ${o.stop} · Trip ${tk}</b><br>${o.customer}<br><span style="color:#4a6a85;font-size:11px">${o.address}</span><br><span style="font-family:monospace;font-size:10px;color:#f39c12">ETA: ${o.eta}</span></div>`);
      mk.addTo(leafMap); mapLayers.push(mk);
    });
  });
  // Depot marker
  const depotIcon = L.divIcon({className:'',html:`<div style="background:#ff6b2b;color:#fff;border-radius:8px;padding:3px 7px;font-size:9px;font-weight:700;font-family:monospace;border:2px solid #fff;box-shadow:0 2px 4px rgba(0,0,0,.5)">📦 DEPOT</div>`,iconAnchor:[30,12]});
  const depMk = L.marker([DEPOT.lat,DEPOT.lng],{icon:depotIcon}).addTo(leafMap);
  mapLayers.push(depMk);
  if (allBounds.length) {
    try { leafMap.fitBounds([[...allBounds,{lat:DEPOT.lat,lng:DEPOT.lng}].map(c=>Array.isArray(c)?c:[c.lat,c.lng])],{padding:[30,30]}); } catch(e){}
  }
}
function updateMapTripFilter() {
  const sel = document.getElementById('mapTripFilter');
  const keys= Object.keys(optimizedTrips).sort((a,b)=>+a-+b);
  sel.innerHTML = '<option value="all">All Trips</option>' +
    keys.map((tk,i) => `<option value="${tk}" style="color:${MCOLORS[i%MCOLORS.length]}">Trip ${tk} (${optimizedTrips[tk].length} stops)</option>`).join('');
}
function filterMapTrip(val) {
  if (!leafMap) return;
  if (val === 'all') { drawAllTrips(); mapFitAll(); return; }
  clearMapLayers();
  const route = optimizedTrips[val]; if (!route) return;
  const ci    = Object.keys(optimizedTrips).sort((a,b)=>+a-+b).indexOf(val);
  const color = MCOLORS[ci % MCOLORS.length];
  const coords= route.filter(o=>o.lat&&o.lng).map(o=>[o.lat,o.lng]);
  if (coords.length) {
    L.polyline([[DEPOT.lat,DEPOT.lng],...coords,[DEPOT.lat,DEPOT.lng]],{color,weight:3,opacity:.8}).addTo(leafMap);
    route.forEach(o => {
      if (!o.lat||!o.lng) return;
      const icon=L.divIcon({className:'',html:`<div style="background:${color};color:#fff;border-radius:50%;width:24px;height:24px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;font-family:monospace;border:2px solid #fff;box-shadow:0 2px 5px rgba(0,0,0,.5)">${o.stop}</div>`,iconSize:[24,24],iconAnchor:[12,12]});
      L.marker([o.lat,o.lng],{icon}).bindPopup(`<div style="font-family:'DM Sans',sans-serif;font-size:12px;background:#0d1a28;color:#d4e8f5;padding:6px"><b style="color:${color}">Stop ${o.stop}</b><br>${o.customer}<br><span style="color:#4a6a85;font-size:11px">${o.address}</span><br>ETA: ${o.eta}</div>`).addTo(leafMap);
    });
    try { leafMap.fitBounds(coords.map(c=>({lat:c[0],lng:c[1]})).concat([DEPOT]).map(c=>[c.lat||c[0],c.lng||c[1]]),{padding:[40,40]}); } catch(e){}
  }
}
function mapFitAll() {
  selectedMapTrip = null;
  if (leafMap && Object.keys(optimizedTrips).length) drawAllTrips();
}

// ══════════════════════════════════════════════════════════════════
// SUBMIT SUMMARY
// ══════════════════════════════════════════════════════════════════
function renderSubmitSummary() {
  const keys = Object.keys(optimizedTrips).sort((a,b) => +a-+b);
  if (!keys.length) return;
  const ss = document.getElementById('submitSummary');
  ss.innerHTML = `
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px;margin-bottom:16px">
      ${keys.map((tk,ci) => {
        const route = optimizedTrips[tk];
        const km    = route.length ? route[route.length-1].cumDist : 0;
        const ton   = route.reduce((s,o) => s+(+o.tonnage||0), 0).toFixed(1);
        const crates= route.reduce((s,o) => s+(+o.crates||0), 0).toFixed(0);
        return `<div style="background:var(--card2);border:1.5px solid var(--border);border-radius:10px;padding:12px;border-top:3px solid ${MCOLORS[ci%MCOLORS.length]}">
          <div style="font-family:'Space Mono',monospace;font-size:.7rem;font-weight:700;color:${MCOLORS[ci%MCOLORS.length]};margin-bottom:6px">TRIP ${tk}</div>
          <div style="font-size:.75rem;margin-bottom:4px"><b>${route.length}</b> stops</div>
          <div style="font-family:'Space Mono',monospace;font-size:.62rem;color:var(--muted);line-height:1.8">
            <span style="color:var(--cyan)">${km} km</span> &nbsp;·&nbsp;
            <span style="color:var(--orange)">${ton} kg</span><br>
            <span style="color:var(--yellow)">${crates} crates</span>
          </div>
        </div>`;
      }).join('')}
    </div>
    <div style="font-family:'Space Mono',monospace;font-size:.62rem;color:var(--muted);text-align:right">
      ${keys.length} trip(s) · ${orders.length} total orders · ${orders.reduce((s,o)=>s+o.crates,0).toFixed(0)} crates · ${orders.reduce((s,o)=>s+o.tonnage,0).toFixed(1)} kg
    </div>`;
  document.getElementById('assignPanel').style.display = 'block';
  document.getElementById('submitBtnArea').style.display = 'block';
}

// ══════════════════════════════════════════════════════════════════
// DRIVER ASSIGN TABLE
// ══════════════════════════════════════════════════════════════════
function renderAssignTable() {
  const keys = Object.keys(optimizedTrips).sort((a,b) => +a-+b);
  const at   = document.getElementById('assignTable');
  if (!keys.length) { at.innerHTML='<div class="empty"><div class="eico">🚚</div>Run optimizer first</div>'; return; }
  if (!drivers.length) { at.innerHTML='<div class="empty"><div class="eico">🚚</div>No drivers loaded</div>'; return; }
  at.innerHTML = `
    <div class="assign-hdr" style="display:grid;grid-template-columns:110px 1fr 240px;gap:10px">
      <div>Trip</div><div>Stops / Route</div><div>Assign Driver</div>
    </div>
    ${keys.map((tk,ci) => {
      const route = optimizedTrips[tk];
      const km    = route.length ? route[route.length-1].cumDist : 0;
      const drvSel = drivers.map(d => `<option value="${d.uid}">${d.status==='active'?'🟢':'⚫'} ${d.name} | ${d.uid} | ${d.vehicle}</option>`).join('');
      const cur   = tripDriverMap[tk];
      return `<div class="assign-row">
        <div>
          <span class="trip-tag ${TCLS[ci%12]}">Trip ${tk}</span>
          <div style="font-family:'Space Mono',monospace;font-size:.55rem;color:var(--muted);margin-top:4px">${km} km · ${route.length} stops</div>
        </div>
        <div style="font-size:.72rem;color:var(--muted)">
          ${route.slice(0,3).map(o=>`<div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:200px">${o.stop}. ${o.customer}</div>`).join('')}
          ${route.length>3 ? `<div style="color:var(--dim);font-size:.62rem">+${route.length-3} more</div>` : ''}
        </div>
        <div>
          <select id="drv-sel-${tk}" onchange="selectDriver('${tk}',this.value,this.options[this.selectedIndex].text)"
            style="width:100%;background:var(--card2);border:1px solid var(--border);color:var(--text);padding:6px 8px;border-radius:7px;font-size:.72rem;outline:none">
            <option value="">⬜ Assign later</option>
            ${drvSel}
          </select>
          ${cur ? `<div style="font-family:'Space Mono',monospace;font-size:.56rem;color:var(--green);margin-top:3px">✓ ${cur.name}</div>` : ''}
        </div>
      </div>`;
    }).join('')}`;
}

function selectDriver(tk, uid, label) {
  if (!uid) { delete tripDriverMap[tk]; return; }
  // Parse name from label format "🟢 Name | UID | Vehicle"
  const parts = label.replace(/[🟢⚫]/g,'').trim().split('|');
  tripDriverMap[tk] = { uid: uid.trim(), name: (parts[0]||'').trim() };
  log(`Trip ${tk} → Driver: ${tripDriverMap[tk].name}`, 'lok');
}

// ══════════════════════════════════════════════════════════════════
// SUBMIT TRIPS — sends data to Streamlit parent
// ══════════════════════════════════════════════════════════════════
function submitTrips() {
  const keys = Object.keys(optimizedTrips).sort((a,b) => +a-+b);
  if (!keys.length) { log('⚠ No trips to submit', 'lwarn'); return; }
  document.getElementById('submitBtn').disabled = true;
  document.getElementById('submitProgress').style.display = 'block';
  let prog = 0;
  const tick = setInterval(() => {
    prog = Math.min(90, prog + 10);
    document.getElementById('submitFill').style.width = prog + '%';
  }, 120);
  const payload = keys.map(tk => ({
    tripKey:   tk,
    stops:     optimizedTrips[tk].map(o => ({
      custId:   o.custId || o.id,
      customer: o.customer,
      address:  o.address,
      lat:      o.lat,
      lng:      o.lng,
      stop:     o.stop,
      orderId:  o.id,
      crates:   o.crates,
      tonnage:  o.tonnage,
      window:   o.window,
      eta:      o.eta,
      legKm:    o.legDist,
      cumKm:    o.cumDist,
    })),
    driverUid:  (tripDriverMap[tk]||{}).uid  || '',
    driverName: (tripDriverMap[tk]||{}).name || '',
    totalKm:    optimizedTrips[tk].length ? optimizedTrips[tk][optimizedTrips[tk].length-1].cumDist : 0,
    totalStops: optimizedTrips[tk].length,
  }));
  setTimeout(() => {
    clearInterval(tick);
    document.getElementById('submitFill').style.width = '100%';
    document.getElementById('submitMsg').textContent  = '✅ Trips submitted successfully!';
    log(`✅ Submitted ${payload.length} trip(s) to system`, 'lok');
    // Send to Streamlit parent via postMessage
    window.parent.postMessage({ type: 'TRIPS_SUBMITTED', trips: payload }, '*');
    showToast(`✅ ${payload.length} trip(s) created successfully!`);
    setTimeout(() => { document.getElementById('submitBtn').disabled = false; }, 3000);
  }, 1400);
}

// ══════════════════════════════════════════════════════════════════
// TABS / UI HELPERS
// ══════════════════════════════════════════════════════════════════
function showTab(id) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-'+id).classList.add('active');
  const idx = ['orders','routemap','map','table','submit'].indexOf(id);
  if (idx >= 0) document.querySelectorAll('.tab')[idx].classList.add('active');
  if (id === 'map') {
    if (!leafMap) initMap();
    setTimeout(() => {
      if (leafMap) { leafMap.invalidateSize(); if (Object.keys(optimizedTrips).length) drawAllTrips(); }
    }, 200);
  }
}

function clearOrders() {
  orders=[]; optimizedTrips={}; tripDriverMap={};
  renderOrdersTable(); updateStats();
  document.getElementById('rmGrid').innerHTML = '<div class="empty"><div class="eico">🗺️</div>Load orders and run optimizer</div>';
  document.getElementById('routeBody').innerHTML = '<tr><td colspan="11"><div class="empty"><div class="eico">📍</div>Run optimizer</div></td></tr>';
  document.getElementById('submitSummary').innerHTML = '<div class="empty"><div class="eico">⚡</div>Run optimizer first</div>';
  document.getElementById('assignPanel').style.display = 'none';
  document.getElementById('submitBtnArea').style.display = 'none';
  document.getElementById('s-dist').textContent = '—';
  if (leafMap) clearMapLayers();
  log('Cleared all orders', 'lwarn');
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3500);
}

function log(msg, cls='') {
  const el = document.getElementById('cronLog');
  const d  = document.createElement('div');
  d.className = 'le ' + cls;
  d.innerHTML = `<span class="lt2">[${new Date().toLocaleTimeString('en-US',{hour12:false})}]</span><span class="lm">${msg}</span>`;
  el.prepend(d);
  while (el.children.length > 80) el.removeChild(el.lastChild);
}
function clearLog() { document.getElementById('cronLog').innerHTML=''; log('Log cleared','linfo'); }

// ══════════════════════════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════════════════════════
window.addEventListener('load', () => {
  log('✓ Route Planner ready — waiting for order data from Streamlit', 'lok');
});
</script>
</body>
</html>
"""

