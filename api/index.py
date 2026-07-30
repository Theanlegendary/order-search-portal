"""
api/index.py  —  Order Search Portal (Vercel Serverless)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Credentials come from Vercel Environment Variables (never hardcoded)
• In-memory module-level cache (Vercel keeps warm instances alive, so
  the same instance handles requests for several minutes, giving us a
  free 5-min cache window without needing Redis / KV)
• Full-text search across every column (order ID, phone, name, note …)
• Category filters: Delayed, Missing, Issue statuses
• Download filtered results as Excel (.xlsx)

Vercel ENV variables required:
  API_URL          https://gw-express.metfone.com.kh/tms-report/api/v1/reports/stages/export-detail
  API_BEARER       eyJhbGci...
  API_BRANCH       MEGA,PRE,PNP,SVA,KAN,KAM,...
  API_CLIENT_ID    TMS_ANDROID
  API_REFERER      https://opsexpress.metfone.com.kh/
  SEARCH_PASSWORD  (optional) simple password to protect the portal
"""

import io
import json
import os
import time
import threading
from datetime import datetime, timedelta

import pandas as pd
import requests
from flask import Flask, request, Response, send_file

app = Flask(__name__)

# ── Env vars (set in Vercel dashboard) ─────────────────────────────────────────
API_URL       = os.environ.get("API_URL", "")
API_BEARER    = os.environ.get("API_BEARER", "")
API_BRANCH    = os.environ.get("API_BRANCH", "MEGA,MEGA1,PRE,PNP,SVA,KAN")
API_CLIENT_ID = os.environ.get("API_CLIENT_ID", "TMS_ANDROID")
API_REFERER   = os.environ.get("API_REFERER", "https://opsexpress.metfone.com.kh/")
PORTAL_PASS   = os.environ.get("SEARCH_PASSWORD", "")   # leave blank = no auth
CACHE_TTL     = int(os.environ.get("CACHE_TTL_SEC", "300"))  # 5 minutes default

# ── Status code labels ─────────────────────────────────────────────────────────
STATUS_LABELS = {
    "110": "Pickup Pending",    "120": "Pickup Failed",     "200": "At Origin Store",
    "210": "In Transit",        "230": "In Transit",        "300": "At Hub",
    "302": "Received Hub",      "306": "At Branch",         "309": "At Agent Store",
    "310": "At Agent",          "311": "Dispatched",        "400": "Delivery Assigned",
    "401": "Out for Delivery",  "402": "Re-Delivery",       "420": "Notify Customer",
    "430": "Contact Receiver",  "460": "Return Initiated",  "470": "Returning",
    "471": "Return Verify",     "472": "Return Issue",      "480": "Rerouting",
    "500": "Returning",         "510": "Return Transit",    "511": "Return Branch",
    "512": "Return Store",      "201": "Delivered",         "410": "Completed",
    "520": "Return Completed",
}
DONE_CODES  = {"201", "410", "520"}
ISSUE_CODES = {"420", "460", "472", "480", "500"}

# ── In-memory cache ────────────────────────────────────────────────────────────
_cache: dict = {
    "df":    pd.DataFrame(),
    "mtime": 0.0,
    "lock":  threading.Lock(),
}


def _download_df() -> pd.DataFrame:
    """Download 14-day order data from API and return as DataFrame."""
    today     = datetime.utcnow() + timedelta(hours=7)   # Phnom Penh time
    from_date = (today - timedelta(days=14)).strftime("%Y%m%d")
    to_date   = today.strftime("%Y%m%d")

    headers = {
        "Authorization": f"Bearer {API_BEARER}",
        "Referer":        API_REFERER,
        "Accept-Language": "vi-VN",
        "Accept":         "application/json, text/plain, */*",
        "Content-Type":   "application/json",
        "x-client-id":    API_CLIENT_ID,
        "User-Agent":     "Mozilla/5.0 (compatible; OrderPortal/1.0)",
    }
    payload = {
        "from_date":   from_date,
        "to_date":     to_date,
        "branch_code": API_BRANCH,
    }

    resp = requests.post(API_URL, headers=headers, json=payload, timeout=120)
    resp.raise_for_status()

    ctype = resp.headers.get("Content-Type", "")

    # Direct Excel binary
    if any(k in ctype for k in ("spreadsheet", "octet-stream", "excel")) or API_URL.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(resp.content))

    # JSON wrapper with a download URL
    try:
        data = resp.json()
    except Exception:
        return pd.read_excel(io.BytesIO(resp.content))

    file_url = (data.get("data") or {}).get("url") or data.get("url")
    if file_url:
        r2 = requests.get(file_url, headers={"Authorization": headers["Authorization"]}, timeout=120)
        r2.raise_for_status()
        return pd.read_excel(io.BytesIO(r2.content))

    raise RuntimeError("Cannot find Excel data in API response")


def _enrich(df: pd.DataFrame) -> pd.DataFrame:
    """Add computed columns: age, status code, status label."""
    now = datetime.utcnow() + timedelta(hours=7)
    date_col = "CREATED DATE" if "CREATED DATE" in df.columns else "CURRENT TIME"
    if date_col in df.columns:
        parsed = pd.to_datetime(df[date_col], dayfirst=True, format="mixed", errors="coerce")
        df = df.copy()
        df["_age_days"] = ((now - parsed).dt.total_seconds() / 86400).fillna(0).clip(lower=0)
    else:
        df = df.copy()
        df["_age_days"] = 0.0

    # Extract 3-digit status code
    if "CURRENT STATUS" in df.columns:
        sc = df["CURRENT STATUS"].astype(str).str.extract(r"(\d{3})")[0]
        df["_sc"]    = sc.fillna("")
        df["_slabel"] = sc.map(STATUS_LABELS).fillna(sc)
    else:
        df["_sc"]    = ""
        df["_slabel"] = ""

    return df


def get_data() -> pd.DataFrame:
    """Return cached DataFrame, refreshing if stale."""
    with _cache["lock"]:
        age = time.time() - _cache["mtime"]
        if age > CACHE_TTL or _cache["df"].empty:
            try:
                raw = _download_df()
                _cache["df"]    = _enrich(raw)
                _cache["mtime"] = time.time()
            except Exception as exc:
                app.logger.error("Data refresh failed: %s", exc)
                if _cache["df"].empty:
                    raise
        return _cache["df"]


def do_search(q: str, cat: str) -> pd.DataFrame:
    df = get_data()
    if df.empty:
        return df

    # Exclude done orders unless explicitly requested
    if cat != "done":
        df = df[~df["_sc"].isin(DONE_CODES)]

    # Keyword filter (case-insensitive across every column)
    q = q.strip()
    if q:
        mask = df.astype(str).apply(
            lambda row: row.str.contains(q, case=False, na=False, regex=False).any(), axis=1
        )
        df = df[mask]

    # Category filter
    if   cat == "missing":  df = df[df["_age_days"] > 7]
    elif cat == "delayed":  df = df[df["_age_days"] > 1]
    elif cat == "issue":    df = df[df["_sc"].isin(ISSUE_CODES)]
    elif cat == "done":
        full = get_data()
        df = full[full["_sc"].isin(DONE_CODES)]
        if q:
            mask = df.astype(str).apply(
                lambda row: row.str.contains(q, case=False, na=False, regex=False).any(), axis=1
            )
            df = df[mask]

    return df


def make_excel_bytes(df: pd.DataFrame) -> bytes:
    drop = {"_age_days", "_sc", "_slabel"}
    out  = df.drop(columns=[c for c in drop if c in df.columns], errors="ignore")
    buf  = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="Orders")
        ws = w.sheets["Orders"]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                max(len(str(col[0].value or "")), 10) + 4, 50
            )
    return buf.getvalue()


# ── Helpers ────────────────────────────────────────────────────────────────────
def _badge(sc: str) -> tuple[str, str]:
    """Return (css_class, emoji) for a status code."""
    if sc in DONE_CODES:          return "done",     "✅"
    if sc in ISSUE_CODES:         return "issue",    "⚠️"
    if sc in {"470","471","472","480","500","510","511","512"}: return "return", "↩️"
    if sc in {"401","402","420","430"}:  return "delivery", "🚚"
    if sc in {"110","120","200"}:        return "pickup",   "📦"
    if sc in {"210","230","300","302","306","309","310","311"}: return "transit", "🔄"
    return "default", ""


def _cache_info() -> tuple[str, str]:
    age = int(time.time() - _cache["mtime"])
    if age < 60:   return f"🟢 Fresh ({age}s ago)", "fresh"
    if age < 300:  return f"🟡 {age//60}m {age%60}s ago", "fresh"
    return f"🔴 Stale ({age//60}m ago)", "stale"


PAGE_SIZE = 250

def _build_table(df: pd.DataFrame, page: int) -> tuple[str, str]:
    if df.empty:
        return (
            '<div class="empty"><div class="empty-icon">🔍</div>'
            '<p>No orders found matching your search.</p>'
            '<p class="sub">Try a different keyword or category.</p></div>',
            ""
        )

    total       = len(df)
    total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
    page        = max(1, min(page, total_pages))
    start       = (page - 1) * PAGE_SIZE
    chunk       = df.iloc[start : start + PAGE_SIZE]

    SHOW_COLS = [
        ("ORDER ID",                "Order ID"),
        ("CREATED DATE",            "Created"),
        ("_slabel",                 "Status"),
        ("SENDER",                  "Sender"),
        ("SENDER PHONE",            "Sender Phone"),
        ("RECEIVER",                "Receiver"),
        ("RECEIVER PHONE",          "Receiver Phone"),
        ("RECEIVE POST OFFICE",     "Origin PO"),
        ("DELIVERY POST OFFICE",    "Dest PO"),
        ("CURRENT POST OFFICE",     "Current PO"),
        ("CURRENT TIME",            "Last Update"),
        ("_age_days",               "Age"),
    ]
    avail = [(c, lbl) for c, lbl in SHOW_COLS if c in chunk.columns]

    ths = "".join(f"<th>{lbl}</th>" for _, lbl in avail)

    rows = []
    for _, row in chunk.iterrows():
        age   = float(row.get("_age_days", 0) or 0)
        sc    = str(row.get("_sc", ""))
        bcls, bico = _badge(sc)
        age_cls = "age-ok" if age <= 1 else ("age-warn" if age <= 7 else "age-danger")

        tds = []
        for col, _ in avail:
            val = row.get(col, "")
            safe = "" if str(val) in ("nan", "None", "") else str(val)
            if col == "_slabel":
                label = STATUS_LABELS.get(sc, safe)
                tds.append(f'<td><span class="badge b-{bcls}">{bico} {label}</span></td>')
            elif col == "_age_days":
                days = int(age)
                hrs  = int((age - days) * 24)
                txt  = f"{days}d" if days else f"{hrs}h"
                tds.append(f'<td class="{age_cls}">{txt}</td>')
            elif col == "ORDER ID":
                tds.append(f'<td class="mono">{safe}</td>')
            else:
                tds.append(f'<td title="{safe}">{safe[:55]}</td>')

        rows.append(f"<tr>{''.join(tds)}</tr>")

    table = (
        f'<div class="tbl-wrap"><table>'
        f'<thead><tr>{ths}</tr></thead>'
        f'<tbody>{"".join(rows)}</tbody>'
        f'</table></div>'
    )

    # Pager
    pager_parts = []
    q_param = request.args.get("q", "")
    cat_param = request.args.get("cat", "all")
    for p in range(1, total_pages + 1):
        active = " active" if p == page else ""
        pager_parts.append(
            f'<a class="pg-btn{active}" href="/?q={q_param}&cat={cat_param}&page={p}">{p}</a>'
        )
    pager_parts.append(f'<span class="pg-info">Rows {start+1}–{min(start+PAGE_SIZE, total):,} of {total:,}</span>')
    return table, "".join(pager_parts)


# ── HTML template ──────────────────────────────────────────────────────────────
HTML_TMPL = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Order Search Portal</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,300;0,14..32,400;0,14..32,500;0,14..32,600;0,14..32,700&display=swap" rel="stylesheet">
<style>
:root{--bg:#070d1a;--surface:#0f1829;--s2:#162035;--s3:#1c2a45;--border:#1e2d45;
  --accent:#4f8ef7;--a2:#38bdf8;--green:#22c55e;--amber:#f59e0b;--red:#ef4444;
  --purple:#a78bfa;--orange:#fb923c;--text:#e2e8f0;--muted:#64748b;--r:10px;--r2:6px}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',system-ui,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px}

/* ── Header ── */
.hdr{background:linear-gradient(135deg,#0a1428 0%,#12204a 50%,#0a1428 100%);
  border-bottom:1px solid var(--border);padding:18px 28px;
  display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:100;
  box-shadow:0 4px 32px rgba(0,0,0,.5)}
.hdr-logo{font-size:24px}
.hdr h1{font-size:18px;font-weight:700;
  background:linear-gradient(120deg,#60a5fa,#a78bfa,#38bdf8);
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.hdr-sub{font-size:11px;color:var(--muted);margin-top:1px}
.hdr-right{margin-left:auto;display:flex;gap:10px;align-items:center}
.cache-pill{font-size:11px;padding:3px 10px;border-radius:20px;
  border:1px solid var(--border);background:var(--s2);color:var(--muted)}
.cache-pill.fresh{color:var(--green)}
.cache-pill.stale{color:var(--red)}
.refresh-btn{font-size:11px;padding:4px 12px;border-radius:20px;border:1px solid var(--border);
  background:var(--s2);color:var(--muted);cursor:pointer;transition:.15s}
.refresh-btn:hover{color:var(--text);border-color:var(--accent)}

/* ── Layout ── */
.page{max-width:1700px;margin:0 auto;padding:22px 28px}

/* ── Stat cards ── */
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:22px}
.sc{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);
  padding:15px 18px;cursor:pointer;transition:.2s;user-select:none}
.sc:hover{border-color:var(--accent);transform:translateY(-1px);box-shadow:0 4px 20px rgba(0,0,0,.3)}
.sc.active{border-color:var(--accent);background:rgba(79,142,247,.1)}
.sc .n{font-size:30px;font-weight:700;line-height:1;margin-bottom:5px}
.sc .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.sc-all    .n{color:var(--accent)}
.sc-delayed .n{color:var(--amber)}
.sc-missing .n{color:var(--red)}
.sc-issue   .n{color:var(--purple)}
.sc-done    .n{color:var(--green)}

/* ── Search row ── */
.srow{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:center}
.sinput-wrap{flex:1;min-width:220px;position:relative}
.sinput-wrap input{width:100%;background:var(--surface);border:1px solid var(--border);
  border-radius:var(--r2);padding:11px 14px 11px 40px;color:var(--text);font-size:13px;
  font-family:inherit;outline:none;transition:.15s}
.sinput-wrap input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(79,142,247,.15)}
.sinput-wrap input::placeholder{color:var(--muted)}
.sinput-ico{position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:15px}
.btn{padding:10px 18px;border:none;border-radius:var(--r2);cursor:pointer;
  font-size:12px;font-weight:600;font-family:inherit;transition:.15s;white-space:nowrap}
.btn-search{background:linear-gradient(135deg,var(--accent),var(--a2));color:#fff}
.btn-search:hover{opacity:.9;transform:translateY(-1px)}
.btn-xl{background:rgba(34,197,94,.12);color:var(--green);border:1px solid rgba(34,197,94,.3)}
.btn-xl:hover{background:rgba(34,197,94,.2)}
.btn-clear{background:var(--s2);color:var(--muted);border:1px solid var(--border)}
.btn-clear:hover{color:var(--text)}

/* ── Result info ── */
.rinfo{font-size:12px;color:var(--muted);margin-bottom:10px}
.rinfo strong{color:var(--text)}

/* ── Table ── */
.tbl-wrap{overflow-x:auto;border-radius:var(--r);border:1px solid var(--border);
  background:var(--surface);box-shadow:0 2px 16px rgba(0,0,0,.2)}
table{width:100%;border-collapse:collapse;font-size:12.5px}
thead tr{background:var(--s2)}
th{padding:10px 12px;text-align:left;font-weight:600;color:var(--muted);
  font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;white-space:nowrap;
  border-bottom:1px solid var(--border)}
td{padding:9px 12px;border-bottom:1px solid rgba(30,45,69,.5);white-space:nowrap;
  max-width:200px;overflow:hidden;text-overflow:ellipsis}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(255,255,255,.02)}
.mono{font-family:ui-monospace,'JetBrains Mono',monospace;font-size:11.5px}

/* ── Badges ── */
.badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10.5px;font-weight:600;white-space:nowrap}
.b-transit  {background:rgba(56,189,248,.12);color:#38bdf8}
.b-delivery {background:rgba(79,142,247,.12);color:#7aa8f9}
.b-pickup   {background:rgba(167,139,250,.12);color:#c084fc}
.b-done     {background:rgba(34,197,94,.12);color:#4ade80}
.b-issue    {background:rgba(239,68,68,.12);color:#f87171}
.b-return   {background:rgba(245,158,11,.12);color:#fbbf24}
.b-default  {background:rgba(100,116,139,.12);color:var(--muted)}

/* ── Age colours ── */
.age-ok    {color:var(--green)}
.age-warn  {color:var(--amber)}
.age-danger{color:var(--red);font-weight:600}

/* ── Empty ── */
.empty{padding:80px 20px;text-align:center;color:var(--muted)}
.empty-icon{font-size:48px;margin-bottom:14px}
.empty p{font-size:14px;margin-bottom:6px}
.empty .sub{font-size:12px;color:var(--border)}

/* ── Pagination ── */
.pager{display:flex;gap:6px;align-items:center;margin-top:14px;flex-wrap:wrap}
.pg-btn{background:var(--surface);border:1px solid var(--border);color:var(--text);
  padding:5px 11px;border-radius:var(--r2);cursor:pointer;font-size:11px;
  text-decoration:none;transition:.15s}
.pg-btn:hover,.pg-btn.active{background:var(--accent);border-color:var(--accent);color:#fff}
.pg-info{font-size:11px;color:var(--muted);margin-left:6px}

/* ── Loading overlay ── */
#ld{display:none;position:fixed;inset:0;background:rgba(7,13,26,.75);z-index:999;
  align-items:center;justify-content:center;flex-direction:column;gap:12px}
#ld.on{display:flex}
.spin{width:36px;height:36px;border:3px solid var(--border);border-top-color:var(--accent);
  border-radius:50%;animation:sp .7s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
#ld p{font-size:13px;color:var(--muted)}

/* ── Auth wall ── */
.auth-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;background:var(--bg)}
.auth-box{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);
  padding:40px;width:340px;text-align:center}
.auth-box h2{font-size:18px;margin-bottom:6px;
  background:linear-gradient(120deg,#60a5fa,#a78bfa);-webkit-background-clip:text;
  -webkit-text-fill-color:transparent;background-clip:text}
.auth-box p{font-size:12px;color:var(--muted);margin-bottom:22px}
.auth-box input{width:100%;background:var(--s2);border:1px solid var(--border);
  border-radius:var(--r2);padding:10px 14px;color:var(--text);font-size:13px;
  font-family:inherit;outline:none;margin-bottom:12px;text-align:center}
.auth-box input:focus{border-color:var(--accent)}
.auth-err{color:var(--red);font-size:12px;margin-top:-6px;margin-bottom:8px}

@media(max-width:700px){.page{padding:14px}.hdr{padding:12px 14px}
  .stats{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>

<div id="ld"><div class="spin"></div><p>Loading…</p></div>

<header class="hdr">
  <div class="hdr-logo">🚚</div>
  <div>
    <h1>Order Search Portal</h1>
    <div class="hdr-sub">14-Day Tracker · Live Bill Lookup</div>
  </div>
  <div class="hdr-right">
    <span class="cache-pill {cc}" id="cpill">{cl}</span>
    <button class="refresh-btn" onclick="forceRefresh()">↺ Refresh</button>
  </div>
</header>

<div class="page">

  <!-- Stats -->
  <div class="stats">
    <div class="sc sc-all {a_all}" onclick="gocat('all')">
      <div class="n">{cnt_all}</div>
      <div class="l">📦 Active</div>
    </div>
    <div class="sc sc-delayed {a_delayed}" onclick="gocat('delayed')">
      <div class="n">{cnt_delayed}</div>
      <div class="l">⏰ Delayed &gt;1d</div>
    </div>
    <div class="sc sc-missing {a_missing}" onclick="gocat('missing')">
      <div class="n">{cnt_missing}</div>
      <div class="l">🚨 Missing &gt;7d</div>
    </div>
    <div class="sc sc-issue {a_issue}" onclick="gocat('issue')">
      <div class="n">{cnt_issue}</div>
      <div class="l">⚠️ Issues</div>
    </div>
    <div class="sc sc-done {a_done}" onclick="gocat('done')">
      <div class="n">{cnt_done}</div>
      <div class="l">✅ Completed</div>
    </div>
  </div>

  <!-- Search -->
  <form method="get" action="/" onsubmit="showLoading()">
    <input type="hidden" name="cat" id="catIn" value="{cat}">
    <div class="srow">
      <div class="sinput-wrap">
        <span class="sinput-ico">🔍</span>
        <input type="text" name="q" id="qIn" value="{q}"
          placeholder="Search by order ID · phone · sender · receiver · post office · note …"
          autocomplete="off" autofocus>
      </div>
      <button type="submit" class="btn btn-search">Search</button>
      <button type="button" class="btn btn-clear" onclick="clearAll()">✕ Clear</button>
      <button type="button" class="btn btn-xl" onclick="dlExcel()">⬇ Export Excel</button>
    </div>
  </form>

  <div class="rinfo">{rinfo}</div>

  {table}

  <div class="pager">{pager}</div>

</div>

<script>
var _cat = "{cat}", _q = "{q}";

function gocat(c) {{
  _cat = c;
  document.getElementById('catIn').value = c;
  showLoading();
  window.location = '/?cat=' + c + (_q ? '&q=' + encodeURIComponent(_q) : '');
}}
function showLoading() {{ document.getElementById('ld').classList.add('on'); }}
function clearAll() {{ window.location = '/?cat=' + _cat; }}
function forceRefresh() {{
  showLoading();
  fetch('/api/refresh', {{method:'POST'}}).then(()=>window.location.reload());
}}
function dlExcel() {{
  showLoading();
  var url = '/api/export?cat=' + _cat + (_q?'&q='+encodeURIComponent(_q):'');
  fetch(url).then(r=>r.blob()).then(b=>{{
    var a=document.createElement('a');
    a.href=URL.createObjectURL(b);
    a.download='orders_{cat}_{stamp}.xlsx';
    a.click();
    document.getElementById('ld').classList.remove('on');
  }});
}}
window.addEventListener('pageshow', ()=>document.getElementById('ld').classList.remove('on'));
// Auto refresh cache pill every 30s
setInterval(()=>{{
  fetch('/api/cache_status').then(r=>r.json()).then(d=>{{
    var p=document.getElementById('cpill');
    p.textContent=d.label; p.className='cache-pill '+d.cls;
  }});
}},30000);
</script>
</body>
</html>"""

AUTH_TMPL = """<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>Login — Order Portal</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
<style>*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',sans-serif;background:#070d1a;color:#e2e8f0;min-height:100vh;
display:flex;align-items:center;justify-content:center}}
.box{{background:#0f1829;border:1px solid #1e2d45;border-radius:10px;padding:40px;width:320px;text-align:center}}
h2{{font-size:20px;margin-bottom:6px;background:linear-gradient(120deg,#60a5fa,#a78bfa);
-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}}
p{{font-size:12px;color:#64748b;margin-bottom:22px}}
input{{width:100%;background:#162035;border:1px solid #1e2d45;border-radius:6px;
padding:11px;color:#e2e8f0;font-size:14px;font-family:inherit;outline:none;
margin-bottom:12px;text-align:center;letter-spacing:2px}}
input:focus{{border-color:#4f8ef7}}
button{{width:100%;padding:11px;background:linear-gradient(135deg,#4f8ef7,#38bdf8);
color:#fff;border:none;border-radius:6px;font-size:13px;font-weight:600;cursor:pointer}}
button:hover{{opacity:.9}}
.err{{color:#ef4444;font-size:12px;margin-bottom:10px}}
</style></head>
<body><div class="box">
<h2>🔐 Order Portal</h2>
<p>Enter your access password</p>
{err}
<form method="post" action="/auth">
<input type="password" name="pw" placeholder="Password" autofocus>
<button type="submit">Enter</button>
</form>
</div></body></html>"""


# ── Auth helpers ───────────────────────────────────────────────────────────────
_sessions: set = set()

def _check_auth() -> bool:
    if not PORTAL_PASS:
        return True
    token = request.cookies.get("session")
    return token in _sessions


# ── Routes ─────────────────────────────────────────────────────────────────────
@app.get("/auth")
def auth_page():
    return AUTH_TMPL.replace("{err}", ""), 200


@app.post("/auth")
def auth_submit():
    from flask import make_response, redirect
    pw = request.form.get("pw", "")
    if pw == PORTAL_PASS:
        import secrets
        tok = secrets.token_hex(16)
        _sessions.add(tok)
        resp = make_response(redirect("/"))
        resp.set_cookie("session", tok, max_age=86400 * 7, httponly=True, samesite="Lax")
        return resp
    err = '<p class="err">Wrong password. Try again.</p>'
    return AUTH_TMPL.replace("{err}", err), 401


@app.get("/")
def index():
    from flask import redirect
    if not _check_auth():
        return redirect("/auth")

    q    = request.args.get("q", "").strip()
    cat  = request.args.get("cat", "all")
    page = max(1, int(request.args.get("page", 1)))

    try:
        df = do_search(q, cat)
    except Exception as exc:
        return f"<pre>Error loading data:\n{exc}</pre>", 500

    # Counts (always on full active dataset)
    full = get_data()
    active = full[~full["_sc"].isin(DONE_CODES)]
    counts = {
        "all":     len(active),
        "delayed": int((active["_age_days"] > 1).sum()),
        "missing": int((active["_age_days"] > 7).sum()),
        "issue":   int(active["_sc"].isin(ISSUE_CODES).sum()),
        "done":    int(full["_sc"].isin(DONE_CODES).sum()),
    }

    cl, cc   = _cache_info()
    table, pager = _build_table(df, page)

    cat_names = {"all":"All Active","delayed":"Delayed >1d","missing":"Missing >7d","issue":"Issues","done":"Completed"}
    if q or cat != "all":
        rinfo = (
            f'Found <strong>{len(df):,}</strong> orders'
            + (f' matching "<strong>{q}</strong>"' if q else "")
            + f' › <strong>{cat_names.get(cat, cat)}</strong>'
        )
    else:
        rinfo = f'Showing <strong>{len(df):,}</strong> active orders in the last 14 days'

    def ac(c): return "active" if cat == c else ""

    stamp = datetime.utcnow().strftime("%d%m%Y")
    html  = HTML_TMPL.format(
        q=q, cat=cat, page=page,
        cl=cl, cc=cc,
        cnt_all=f"{counts['all']:,}",      cnt_delayed=f"{counts['delayed']:,}",
        cnt_missing=f"{counts['missing']:,}", cnt_issue=f"{counts['issue']:,}",
        cnt_done=f"{counts['done']:,}",
        a_all=ac("all"), a_delayed=ac("delayed"), a_missing=ac("missing"),
        a_issue=ac("issue"), a_done=ac("done"),
        rinfo=rinfo, table=table, pager=pager,
        stamp=stamp,
    )
    return html, 200


@app.get("/api/export")
def api_export():
    if not _check_auth():
        return "Unauthorized", 401
    q   = request.args.get("q", "").strip()
    cat = request.args.get("cat", "all")
    try:
        df = do_search(q, cat)
    except Exception as exc:
        return str(exc), 500

    xlsx = make_excel_bytes(df)
    stamp = datetime.utcnow().strftime("%d%m%Y_%H%M")
    fname = f"orders_{cat}_{stamp}.xlsx"
    return Response(
        xlsx,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.post("/api/refresh")
def api_refresh():
    if not _check_auth():
        return "Unauthorized", 401
    with _cache["lock"]:
        _cache["mtime"] = 0.0   # force re-download on next request
    return "ok"


@app.get("/api/cache_status")
def api_cache_status():
    label, cls = _cache_info()
    return {"label": label, "cls": cls}


@app.get("/health")
def health():
    return {"status": "ok", "rows": len(_cache["df"])}


# Local dev entry-point
if __name__ == "__main__":
    # For local dev, load config.json from parent directory
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
    parent_cfg = os.path.join(os.path.dirname(__file__), "..", "config.json")
    if os.path.exists(parent_cfg):
        with open(parent_cfg, encoding="utf-8") as f:
            cfg = json.load(f)["api"]
        os.environ.setdefault("API_URL",       cfg.get("url", ""))
        os.environ.setdefault("API_BEARER",    cfg.get("bearer_token", ""))
        os.environ.setdefault("API_BRANCH",    cfg.get("branch_code", ""))
        os.environ.setdefault("API_CLIENT_ID", cfg.get("x_client_id", "TMS_ANDROID"))
        # Re-read env vars
        globals()["API_URL"]       = os.environ["API_URL"]
        globals()["API_BEARER"]    = os.environ["API_BEARER"]
        globals()["API_BRANCH"]    = os.environ["API_BRANCH"]
        globals()["API_CLIENT_ID"] = os.environ["API_CLIENT_ID"]

    app.run(host="0.0.0.0", port=5055, debug=True)
