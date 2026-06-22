#!/usr/bin/env python3
"""
build_dashboard.py  —  Generate a single static dashboard at site/index.html.

Reads the product catalog (products/) and the status layer (status/), embeds the data
as JSON inside one self-contained HTML file, and writes site/index.html. No server, no
build step, no dependencies: double-click it, or host site/ on GitHub Pages / Cloudflare.

Three views in one page:
  - Catalog  : product data (image, name, vendor, pricing, images, docs)
  - Status   : a grid of colored channel chips per item (green=live, red=error, ...)
  - Detail   : click any row for full product data + per-channel status

Usage:
  python scripts/build_dashboard.py
"""

import glob
import json
import os
import sys
from datetime import date

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from status_lib import CHANNELS

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRODUCTS_DIR = os.path.join(REPO_ROOT, "products")
STATUS_DIR = os.path.join(REPO_ROOT, "status")
OUT = os.path.join(REPO_ROOT, "site", "index.html")


def build_records():
    status = {}
    for p in glob.glob(os.path.join(STATUS_DIR, "*.json")):
        with open(p, encoding="utf-8") as f:
            s = json.load(f)
        status[str(s.get("ds_number"))] = s

    records = []
    for path in sorted(glob.glob(os.path.join(PRODUCTS_DIR, "**", "*.json"), recursive=True)):
        with open(path, encoding="utf-8") as f:
            p = json.load(f)
        ds = str(p.get("ds_number") or "")
        pricing = p.get("pricing") or {}
        ident = p.get("identity") or {}
        mc = (p.get("dimensions") or {}).get("master_case", {}) or {}
        content = p.get("content") or {}
        comp = p.get("compliance") or {}
        imgs = (p.get("images") or {}).get("urls") or []
        docs = comp.get("documents") or []
        st = status.get(ds, {})
        chans = st.get("channels") or {}
        records.append({
            "ds": ds,
            "vendor": p.get("vendor") or "",
            "brand": p.get("brand") or "",
            "name": p.get("product_name") or "",
            "upc": ident.get("upc") or "",
            "vin": p.get("vendor_item_number") or "",
            "cost": pricing.get("wholesale_cost"),
            "msrp": pricing.get("msrp"),
            "dropship": pricing.get("drop_ship_cost"),
            "caseqty": mc.get("case_qty"),
            "supplier": (p.get("_meta") or {}).get("supplier_id"),
            "imgs": imgs,
            "docs": docs,
            "desc": content.get("description") or "",
            "bullets": [b for b in (content.get("bullets") or []) if b],
            "lifecycle": st.get("lifecycle") or "active",
            "channels": {c: {
                "state": (chans.get(c) or {}).get("state", "not_listed"),
                "id": (chans.get(c) or {}).get("id"),
                "url": (chans.get(c) or {}).get("url"),
                "case": (chans.get(c) or {}).get("case_number"),
                "issue": (chans.get(c) or {}).get("issue"),
                "notes": (chans.get(c) or {}).get("notes"),
            } for c in CHANNELS},
        })
    return records


def main():
    records = build_records()
    html = (TEMPLATE
            .replace("__DATA__", json.dumps(records, ensure_ascii=False, default=str))
            .replace("__CHANNELS__", json.dumps(CHANNELS, ensure_ascii=False))
            .replace("__GENERATED__", date.today().isoformat())
            .replace("__COUNT__", str(len(records))))
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote dashboard ({len(records)} items) -> {os.path.relpath(OUT, REPO_ROOT)}")
    print("Open it: double-click site/index.html (or host the site/ folder).")


TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DS Catalog &amp; Channel Status</title>
<style>
  :root{
    --bg:#0f1419; --panel:#171e26; --panel2:#1e2832; --line:#2a3642; --txt:#e6edf3;
    --muted:#8b98a5; --accent:#3b82f6;
    --live:#22c55e; --error:#ef4444; --suppressed:#f97316; --pending:#38bdf8;
    --setup_in_progress:#eab308; --planned:#64748b; --not_listed:#374151; --discontinued:#4b5563;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--txt);font:14px/1.4 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif}
  header{padding:14px 18px;border-bottom:1px solid var(--line);display:flex;gap:18px;align-items:center;flex-wrap:wrap;position:sticky;top:0;background:var(--bg);z-index:10}
  h1{font-size:16px;margin:0;font-weight:600}
  .stats{color:var(--muted);font-size:12px;display:flex;gap:14px;flex-wrap:wrap}
  .stats b{color:var(--txt)}
  .controls{padding:10px 18px;border-bottom:1px solid var(--line);display:flex;gap:10px;align-items:center;flex-wrap:wrap;background:var(--panel)}
  input,select{background:var(--panel2);color:var(--txt);border:1px solid var(--line);border-radius:6px;padding:6px 9px;font-size:13px}
  input#q{min-width:240px}
  .seg{display:inline-flex;border:1px solid var(--line);border-radius:6px;overflow:hidden}
  .seg button{background:var(--panel2);color:var(--muted);border:0;padding:6px 12px;cursor:pointer;font-size:13px}
  .seg button.on{background:var(--accent);color:#fff}
  table{width:100%;border-collapse:collapse}
  th,td{text-align:left;padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:middle}
  th{position:sticky;top:53px;background:var(--panel);color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.04em;cursor:pointer;white-space:nowrap}
  tbody tr{cursor:pointer}
  tbody tr:hover{background:var(--panel)}
  .thumb{width:40px;height:40px;object-fit:contain;background:#fff;border-radius:4px}
  .noimg{width:40px;height:40px;border-radius:4px;background:var(--panel2);display:flex;align-items:center;justify-content:center;color:var(--muted);font-size:9px}
  .nm{max-width:380px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .muted{color:var(--muted)}
  .badge{display:inline-block;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:600;text-transform:capitalize}
  .lc-active{background:rgba(34,197,94,.18);color:var(--live)}
  .lc-discontinued{background:rgba(75,85,99,.3);color:#cbd5e1}
  .lc-on_hold{background:rgba(234,179,8,.18);color:var(--setup_in_progress)}
  .lc-planned{background:rgba(56,189,248,.16);color:var(--pending)}
  .lc-dropped{background:rgba(239,68,68,.16);color:var(--error)}
  .chips{display:flex;gap:3px;flex-wrap:nowrap}
  .chip{width:15px;height:15px;border-radius:3px;display:inline-block;cursor:help}
  .chip.na{background:transparent!important;border:1.5px dashed var(--muted)}
  .chiprow{display:flex;gap:6px;flex-wrap:wrap;margin-top:4px}
  .chiplbl{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--muted);background:var(--panel2);border:1px solid var(--line);border-radius:5px;padding:3px 7px}
  .legend{display:flex;gap:12px;flex-wrap:wrap;color:var(--muted);font-size:11px;align-items:center}
  .legend .chip{cursor:default}
  /* drawer */
  #drawer{position:fixed;top:0;right:0;height:100%;width:min(560px,92vw);background:var(--panel);border-left:1px solid var(--line);transform:translateX(100%);transition:transform .18s;overflow-y:auto;z-index:20;box-shadow:-10px 0 30px rgba(0,0,0,.4)}
  #drawer.open{transform:translateX(0)}
  .dh{padding:16px 18px;border-bottom:1px solid var(--line);position:sticky;top:0;background:var(--panel)}
  .dh h2{margin:0 0 4px;font-size:16px}
  .db{padding:16px 18px}
  .close{position:absolute;top:14px;right:16px;cursor:pointer;color:var(--muted);font-size:20px;background:0;border:0}
  .kv{display:grid;grid-template-columns:130px 1fr;gap:4px 10px;font-size:13px;margin:10px 0}
  .kv div:nth-child(odd){color:var(--muted)}
  .gallery{display:flex;gap:6px;flex-wrap:wrap;margin:8px 0}
  .gallery img{width:64px;height:64px;object-fit:contain;background:#fff;border-radius:4px}
  .sect{margin-top:18px;font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);border-bottom:1px solid var(--line);padding-bottom:5px}
  a{color:var(--accent)}
  .pill{font-size:10px;padding:1px 6px;border-radius:99px;color:#fff;text-transform:capitalize}
  #backdrop{position:fixed;inset:0;background:rgba(0,0,0,.45);opacity:0;pointer-events:none;transition:opacity .18s;z-index:15}
  #backdrop.open{opacity:1;pointer-events:auto}
  .empty{padding:40px;text-align:center;color:var(--muted)}
</style>
</head>
<body>
<header>
  <h1>DS Catalog &amp; Channel Status</h1>
  <div class="stats" id="stats"></div>
</header>
<div class="controls">
  <input id="q" placeholder="Search name, DS#, UPC, vendor, brand...">
  <select id="vendor"></select>
  <select id="lifecycle"><option value="">All lifecycle</option></select>
  <select id="channel"><option value="">Any channel</option></select>
  <select id="cstate"><option value="">Any state</option></select>
  <span class="seg"><button data-view="catalog" class="on">Catalog</button><button data-view="status">Status</button></span>
  <span class="legend" id="legend"></span>
</div>
<table>
  <thead id="thead"></thead>
  <tbody id="tbody"></tbody>
</table>
<div class="empty" id="empty" style="display:none">No items match.</div>

<div id="backdrop"></div>
<div id="drawer"><button class="close" id="dclose">&times;</button><div id="dcontent"></div></div>

<script>
const DATA = __DATA__;
const CHANNELS = __CHANNELS__;
const STATES = ["live","error","suppressed","pending","setup_in_progress","planned","discontinued","not_listed"];
const cssVar = s => getComputedStyle(document.body).getPropertyValue('--'+s) || '#888';
let view = "catalog";

function money(v){ return (v===null||v===undefined||v==="")?'<span class="muted">—</span>':('$'+Number(v).toFixed(2)); }
function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function chip(state){ return `<span class="chip" style="background:${cssVar(state)}"></span>`; }

// populate filters
const vendors = [...new Set(DATA.map(d=>d.vendor))].sort();
vendor.innerHTML = '<option value="">All vendors</option>'+vendors.map(v=>`<option>${esc(v)}</option>`).join('');
["active","planned","on_hold","discontinued","dropped"].forEach(l=>lifecycle.innerHTML+=`<option value="${l}">${l}</option>`);
Object.entries(CHANNELS).forEach(([k,v])=>channel.innerHTML+=`<option value="${k}">${esc(v)}</option>`);
STATES.forEach(s=>cstate.innerHTML+=`<option value="${s}">${s.replace(/_/g,' ')}</option>`);
legend.innerHTML = STATES.map(s=>`${chip(s)}<span>${s.replace(/_/g,' ')}</span>`).join(' &nbsp; ');

function filtered(){
  const q=document.getElementById('q').value.toLowerCase().trim();
  const v=vendor.value, lc=lifecycle.value, ch=channel.value, cs=cstate.value;
  return DATA.filter(d=>{
    if(v && d.vendor!==v) return false;
    if(lc && d.lifecycle!==lc) return false;
    if(ch && cs && d.channels[ch].state!==cs) return false;
    if(ch && !cs && d.channels[ch].state==='not_listed') return false;
    if(!ch && cs && !Object.values(d.channels).some(c=>c.state===cs)) return false;
    if(q){ const hay=(d.ds+' '+d.name+' '+d.upc+' '+d.vendor+' '+d.brand+' '+d.vin).toLowerCase();
           if(!hay.includes(q)) return false; }
    return true;
  });
}

function render(){
  const rows=filtered();
  document.getElementById('empty').style.display = rows.length?'none':'block';
  const liveTot = DATA.reduce((a,d)=>a+Object.values(d.channels).filter(c=>c.state==='live').length,0);
  const errTot = DATA.reduce((a,d)=>a+Object.values(d.channels).filter(c=>c.state==='error').length,0);
  stats.innerHTML = `<span><b>${DATA.length}</b> items</span><span><b>${vendors.length}</b> vendors</span>`
    +`<span><b>${liveTot}</b> live listings</span><span><b style="color:var(--error)">${errTot}</b> errors</span>`
    +`<span>showing <b>${rows.length}</b></span>`;
  if(view==='catalog'){
    thead.innerHTML='<tr><th></th><th>DS#</th><th>Product</th><th>Vendor</th><th>Brand</th><th>Cost</th><th>MSRP</th><th>Case</th><th>Imgs</th><th>Life</th></tr>';
    tbody.innerHTML=rows.map((d,i)=>`<tr data-i="${DATA.indexOf(d)}">
      <td>${d.imgs.length?`<img class="thumb" loading="lazy" src="${esc(d.imgs[0])}">`:'<div class="noimg">none</div>'}</td>
      <td class="muted">${esc(d.ds)}</td>
      <td class="nm" title="${esc(d.name)}">${esc(d.name)}</td>
      <td>${esc(d.vendor)}</td><td class="muted">${esc(d.brand)}</td>
      <td>${money(d.cost)}</td><td>${money(d.msrp)}</td>
      <td class="muted">${d.caseqty??'—'}</td><td class="muted">${d.imgs.length}</td>
      <td><span class="badge lc-${d.lifecycle}">${d.lifecycle}</span></td></tr>`).join('');
  } else {
    const heads = Object.values(CHANNELS).map(l=>`<th title="${esc(l)}">${esc(l.split(' ')[0])}</th>`).join('');
    thead.innerHTML=`<tr><th>DS#</th><th>Product</th><th>Vendor</th><th>Life</th>${heads}</tr>`;
    tbody.innerHTML=rows.map(d=>`<tr data-i="${DATA.indexOf(d)}">
      <td class="muted">${esc(d.ds)}</td><td class="nm" title="${esc(d.name)}">${esc(d.name)}</td>
      <td>${esc(d.vendor)}</td><td><span class="badge lc-${d.lifecycle}">${d.lifecycle}</span></td>
      ${Object.keys(CHANNELS).map(k=>{const c=d.channels[k];
        const cov=(k==='walmart_3p' && d.channels.walmart_1p && d.channels.walmart_1p.state==='live' && c.state==='not_listed');
        const tip=cov?'Walmart 3P: not needed (listed on Walmart 1P)'
          :`${CHANNELS[k]}: ${c.state}`+(c.id?` · ${c.id}`:'')+(c.case?` · case ${c.case}`:'')+(c.issue?` · ${c.issue}`:'');
        return `<td><span class="chip${cov?' na':''}" style="background:${cov?'transparent':cssVar(c.state)}" title="${esc(tip)}"></span></td>`;}).join('')}
    </tr>`).join('');
  }
}

function openDrawer(d){
  const ch = Object.keys(CHANNELS).map(k=>{const c=d.channels[k];
    const cov=(k==='walmart_3p' && d.channels.walmart_1p && d.channels.walmart_1p.state==='live' && c.state==='not_listed');
    if(cov) return `<div class="chiplbl"><span class="chip na"></span>${esc(CHANNELS[k])}: <span class="muted">not needed (on Walmart 1P)</span></div>`;
    return `<div class="chiplbl"><span class="chip" style="background:${cssVar(c.state)}"></span>${esc(CHANNELS[k])}: ${c.state.replace(/_/g,' ')}`
      +(c.id?(c.url?` · <a href="${esc(c.url)}" target="_blank">${esc(c.id)}</a>`:` · <span class="muted">${esc(c.id)}</span>`):'')
      +(c.case?` · case ${esc(c.case)}`:'')
      +(c.issue?` · <span style="color:var(--error)">${esc(c.issue)}</span>`:'')+`</div>`;}).join('');
  dcontent.innerHTML=`<div class="dh"><h2>${esc(d.name)}</h2>
    <div class="muted">DS${esc(d.ds)} · ${esc(d.vendor)} · ${esc(d.brand)} <span class="badge lc-${d.lifecycle}">${d.lifecycle}</span></div></div>
  <div class="db">
    ${d.imgs.length?`<div class="gallery">${d.imgs.map(u=>`<a href="${esc(u)}" target="_blank"><img loading="lazy" src="${esc(u)}"></a>`).join('')}</div>`:'<p class="muted">No images yet.</p>'}
    <div class="sect">Pricing &amp; identity</div>
    <div class="kv"><div>UPC</div><div>${esc(d.upc)||'—'}</div>
      <div>Vendor item #</div><div>${esc(d.vin)||'—'}</div>
      <div>Wholesale</div><div>${money(d.cost)}</div>
      <div>Drop-ship</div><div>${money(d.dropship)}</div>
      <div>MSRP</div><div>${money(d.msrp)}</div>
      <div>Case qty</div><div>${d.caseqty??'—'}</div>
      <div>Supplier ID</div><div>${d.supplier??'—'}</div></div>
    ${d.bullets.length?`<div class="sect">Bullets</div><ul>${d.bullets.map(b=>`<li>${esc(b)}</li>`).join('')}</ul>`:''}
    ${d.desc?`<div class="sect">Description</div><p>${esc(d.desc)}</p>`:''}
    <div class="sect">Channel status</div><div class="chiprow">${ch}</div>
    ${d.docs.length?`<div class="sect">Compliance docs</div>${d.docs.map(x=>`<div>· ${esc(x.type||'doc')}: <a href="${esc(x.url)}" target="_blank">link</a></div>`).join('')}`:''}
  </div>`;
  drawer.classList.add('open'); backdrop.classList.add('open');
}
function closeDrawer(){ drawer.classList.remove('open'); backdrop.classList.remove('open'); }

document.querySelectorAll('.seg button').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('.seg button').forEach(x=>x.classList.remove('on'));
  b.classList.add('on'); view=b.dataset.view; render();
});
['q','vendor','lifecycle','channel','cstate'].forEach(id=>document.getElementById(id).addEventListener('input',render));
tbody.addEventListener('click',e=>{const tr=e.target.closest('tr'); if(tr) openDrawer(DATA[+tr.dataset.i]);});
dclose.onclick=closeDrawer; backdrop.onclick=closeDrawer;
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDrawer()});
render();
</script>
<footer style="padding:12px 18px;color:var(--muted);font-size:11px;border-top:1px solid var(--line)">
  Generated __GENERATED__ from the DS Setup Repo · __COUNT__ items · rebuild with <code>python scripts/build_dashboard.py</code>
</footer>
</body>
</html>"""


if __name__ == "__main__":
    main()
