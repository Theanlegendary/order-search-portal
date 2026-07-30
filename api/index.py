"""
api/index.py  —  Order Search Portal (Vercel Serverless)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Vercel ENV variables required:
  API_URL          https://gw-express.metfone.com.kh/tms-report/...
  API_BEARER       eyJhbGci...
  API_BRANCH       MEGA,PRE,PNP,SVA,KAN,KAM,...
  API_CLIENT_ID    TMS_ANDROID
  API_REFERER      https://opsexpress.metfone.com.kh/
  SEARCH_PASSWORD  (optional) simple password to protect the portal
  CACHE_TTL_SEC    300 (5 minutes default)
"""

import io
import json
import os
import re
import time
import threading
from datetime import datetime, timedelta

import pandas as pd
import requests
from flask import Flask, request, Response

app = Flask(__name__)

# ── Env vars ────────────────────────────────────────────────────────────────────
API_URL       = os.environ.get("API_URL", "")
API_BEARER    = os.environ.get("API_BEARER", "")
API_BRANCH    = os.environ.get("API_BRANCH", "MEGA,MEGA1,PRE,PNP,SVA,KAN")
API_CLIENT_ID = os.environ.get("API_CLIENT_ID", "TMS_ANDROID")
API_REFERER   = os.environ.get("API_REFERER", "https://opsexpress.metfone.com.kh/")
PORTAL_PASS   = os.environ.get("SEARCH_PASSWORD", "")
CACHE_TTL     = int(os.environ.get("CACHE_TTL_SEC", "300"))

# ── Status labels ────────────────────────────────────────────────────────────────
STATUS_LABELS = {
    "110":"Pickup Pending",   "120":"Pickup Failed",    "200":"At Origin Store",
    "210":"In Transit",       "230":"In Transit",       "300":"At Hub",
    "302":"Received Hub",     "306":"At Branch",        "309":"At Agent Store",
    "310":"At Agent",         "311":"Dispatched",       "400":"Delivery Assigned",
    "401":"Out for Delivery", "402":"Re-Delivery",      "420":"Notify Customer",
    "430":"Contact Receiver", "460":"Return Initiated", "470":"Returning",
    "471":"Return Verify",    "472":"Return Issue",     "480":"Rerouting",
    "500":"Returning",        "510":"Return Transit",   "511":"Return Branch",
    "512":"Return Store",     "201":"Delivered",        "410":"Completed",
    "520":"Return Completed",
}
DONE_CODES   = {"201", "410"}
CANCEL_CODES = {"520", "470", "471", "472", "480", "500", "510", "511", "512", "120", "460"}
ISSUE_CODES  = {"420", "460", "472", "480", "500"}

BRANCH_NAME_MAP = {
    "PNP": "Phnom Penh", "PRE": "Prey Veng", "SVA": "Svay Rieng", "KAN": "Kandal",
    "KAM": "Kampot", "KOH": "Koh Kong", "SIH": "Preah Sihanouk", "SPE": "Kampong Speu",
    "TAK": "Takeo", "BAN": "Banteay Meanchey", "BAT": "Battambang", "CHH": "Kampong Chhnang",
    "PUR": "Pursat", "SIE": "Siem Reap", "PRH": "Preah Vihear", "ODD": "Oddar Meanchey",
    "THO": "Kampong Thom", "CHA": "Kampong Cham", "KRA": "Kratie", "TBK": "Tboung Khmum",
    "ROT": "Ratanak Kiri", "MON": "Mondul Kiri", "STU": "Stung Treng", "KEP": "Kep", "PAI": "Pailin"
}

# ── In-memory cache ──────────────────────────────────────────────────────────────
_cache = {"df": pd.DataFrame(), "mtime": 0.0, "lock": threading.Lock()}


def _download_df():
    today     = datetime.utcnow() + timedelta(hours=7)
    from_date = (today - timedelta(days=14)).strftime("%Y%m%d")
    to_date   = today.strftime("%Y%m%d")
    hdrs = {
        "Authorization": f"Bearer {API_BEARER}",
        "Referer": API_REFERER,
        "Accept-Language": "vi-VN",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "x-client-id": API_CLIENT_ID,
        "User-Agent": "Mozilla/5.0 (compatible; OrderPortal/1.0)",
    }
    resp = requests.post(API_URL, headers=hdrs,
                         json={"from_date": from_date, "to_date": to_date, "branch_code": API_BRANCH},
                         timeout=120)
    resp.raise_for_status()
    ctype = resp.headers.get("Content-Type", "")
    if any(k in ctype for k in ("spreadsheet", "octet-stream", "excel")):
        return pd.read_excel(io.BytesIO(resp.content))
    try:
        data = resp.json()
    except Exception:
        return pd.read_excel(io.BytesIO(resp.content))
    file_url = (data.get("data") or {}).get("url") or data.get("url")
    if file_url:
        r2 = requests.get(file_url, headers={"Authorization": hdrs["Authorization"]}, timeout=120)
        r2.raise_for_status()
        return pd.read_excel(io.BytesIO(r2.content))
    raise RuntimeError("No Excel data in API response")


def _get_facility_type(code):
    code = str(code or "").strip().upper()
    m = re.search(r"^[A-Z]{3}([PSA])\d+", code)
    if m:
        letter = m.group(1)
        if letter == "P": return "Post Office"
        if letter == "S": return "Showroom"
        if letter == "A": return "Agent"
    return "Other"


def _get_branch_code(code):
    code = str(code or "").strip().upper()
    m = re.match(r"^([A-Z]{3})", code)
    return m.group(1) if m else "OTHER"


def _enrich(df):
    now = datetime.utcnow() + timedelta(hours=7)
    date_col = "CREATED DATE" if "CREATED DATE" in df.columns else "CURRENT TIME"
    df = df.copy()
    if date_col in df.columns:
        parsed = pd.to_datetime(df[date_col], dayfirst=True, format="mixed", errors="coerce")
        df["_parsed_date"] = parsed
        df["_date_only"]   = parsed.dt.date
        df["_age_days"]    = ((now - parsed).dt.total_seconds() / 86400).fillna(0).clip(lower=0)
    else:
        df["_parsed_date"] = pd.NaT
        df["_date_only"]   = None
        df["_age_days"]    = 0.0

    if "CURRENT STATUS" in df.columns:
        sc = df["CURRENT STATUS"].astype(str).str.extract(r"(\d{3})")[0]
        df["STATUS CODE"] = sc.fillna("")
        df["_sc"]         = sc.fillna("")
        df["_slabel"]     = sc.map(STATUS_LABELS).fillna(sc)
    else:
        df["STATUS CODE"] = ""
        df["_sc"]         = ""
        df["_slabel"]     = ""

    po_col = "CURRENT POST OFFICE" if "CURRENT POST OFFICE" in df.columns else ("DELIVERY POST OFFICE" if "DELIVERY POST OFFICE" in df.columns else "RECEIVE POST OFFICE")
    if po_col in df.columns:
        df["_facility"] = df[po_col].apply(_get_facility_type)
        df["_branch"]   = df[po_col].apply(_get_branch_code)
    else:
        df["_facility"] = "Other"
        df["_branch"]   = "OTHER"

    return df


def get_data():
    with _cache["lock"]:
        if time.time() - _cache["mtime"] > CACHE_TTL or _cache["df"].empty:
            try:
                raw = _download_df()
                _cache["df"]    = _enrich(raw)
                _cache["mtime"] = time.time()
            except Exception as exc:
                app.logger.error("Refresh failed: %s", exc)
                if _cache["df"].empty:
                    raise
        return _cache["df"]


def do_search(q, cat, branch="", date_filter="all", sort_order="desc"):
    df = get_data()
    if df.empty:
        return df

    # Branch filter
    if branch and branch != "ALL":
        df = df[df["_branch"] == branch.upper()]

    # Date range filter
    today = (datetime.utcnow() + timedelta(hours=7)).date()
    if date_filter == "today":
        df = df[df["_date_only"] == today]
    elif date_filter == "yesterday":
        df = df[df["_date_only"] == (today - timedelta(days=1))]
    elif date_filter == "last3":
        df = df[df["_date_only"] >= (today - timedelta(days=2))]
    elif date_filter == "last7":
        df = df[df["_date_only"] >= (today - timedelta(days=6))]
    elif re.match(r"^\d{4}-\d{2}-\d{2}$", date_filter):
        target_d = pd.to_datetime(date_filter).date()
        df = df[df["_date_only"] == target_d]

    # Category filters
    if cat == "active":
        df = df[~df["_sc"].isin(DONE_CODES | CANCEL_CODES)]
    elif cat == "po_only":
        df = df[df["_facility"] == "Post Office"]
    elif cat == "missing":
        df = df[(df["_age_days"] > 7) & (~df["_sc"].isin(DONE_CODES))]
    elif cat == "delayed":
        df = df[(df["_age_days"] > 1) & (~df["_sc"].isin(DONE_CODES))]
    elif cat == "issue":
        df = df[df["_sc"].isin(ISSUE_CODES)]
    elif cat == "done":
        df = df[df["_sc"].isin(DONE_CODES)]
    elif cat == "cancel":
        df = df[df["_sc"].isin(CANCEL_CODES)]
    elif cat == "all":
        pass

    # Keyword search across all columns
    q = q.strip()
    if q:
        mask = df.astype(str).apply(
            lambda row: row.str.contains(q, case=False, na=False, regex=False).any(), axis=1
        )
        df = df[mask]

    # Date sorting
    ascending = (sort_order == "asc")
    if "_parsed_date" in df.columns:
        df = df.sort_values(by="_parsed_date", ascending=ascending, na_position="last")

    return df


def make_excel_bytes(df):
    drop = {"_age_days", "_sc", "_slabel", "_facility", "_branch", "_parsed_date", "_date_only"}
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


def _badge(sc):
    if sc in DONE_CODES:                                          return "done",     "&#10003;"
    if sc in ISSUE_CODES:                                         return "issue",    "&#9888;"
    if sc in CANCEL_CODES:                                        return "return_",  "&#8617;"
    if sc in {"401","402","420","430"}:                           return "delivery", "&#128666;"
    if sc in {"110","120","200"}:                                 return "pickup",   "&#128230;"
    if sc in {"210","230","300","302","306","309","310","311"}:   return "transit",  "&#8635;"
    return "default_", ""


def _cache_info():
    age = int(time.time() - _cache["mtime"])
    if age <  60:  return f"Fresh ({age}s ago)", "fresh"
    if age < 300:  return f"{age//60}m {age%60}s ago", "fresh"
    return f"Stale ({age//60}m ago)", "stale"


def _render(tmpl, **kw):
    for k, v in kw.items():
        tmpl = tmpl.replace(f"~~{k}~~", str(v))
    return tmpl


PAGE_SIZE = 250

def _build_table(df, page=1, cat="all", branch="", date_filter="all", sort_order="desc"):
    if df.empty:
        return (
            '<div class="empty">'
            '<div style="font-size:48px;margin-bottom:14px">&#128269;</div>'
            '<p>No orders found.</p>'
            '<p style="font-size:12px;color:#475569;margin-top:6px">Try a different keyword, date range, branch, or category.</p>'
            '</div>'
        ), ""

    total       = len(df)
    total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
    page        = max(1, min(page, total_pages))
    start       = (page - 1) * PAGE_SIZE
    chunk       = df.iloc[start: start + PAGE_SIZE]

    COLS = [
        ("ORDER ID",             "Order ID"),
        ("CREATED DATE",         "Created Date 📅"),
        ("STATUS CODE",          "Status Code"),
        ("_slabel",              "Status Description"),
        ("SENDER",               "Sender"),
        ("SENDER PHONE",         "Sender Phone"),
        ("RECEIVER",             "Receiver"),
        ("RECEIVER PHONE",       "Receiver Phone"),
        ("RECEIVE POST OFFICE",  "Origin PO"),
        ("DELIVERY POST OFFICE", "Dest PO"),
        ("CURRENT POST OFFICE",  "Current PO"),
        ("_facility",            "Facility"),
        ("CURRENT TIME",         "Last Update"),
        ("_age_days",            "Age"),
    ]
    avail = [(c, l) for c, l in COLS if c in chunk.columns]
    ths   = "".join(f"<th>{l}</th>" for _, l in avail)

    rows = []
    for _, row in chunk.iterrows():
        age  = float(row.get("_age_days", 0) or 0)
        sc   = str(row.get("_sc", ""))
        bcls, bico = _badge(sc)
        age_cls = "age-ok" if age <= 1 else ("age-warn" if age <= 7 else "age-danger")
        tds = []
        for col, _ in avail:
            val  = row.get(col, "")
            safe = "" if str(val) in ("nan", "None", "") else str(val)
            if col == "STATUS CODE":
                tds.append(f'<td><span class="badge b-{bcls}" style="font-family:monospace;font-size:12px;font-weight:700">{sc or safe or "N/A"}</span></td>')
            elif col == "_slabel":
                label = STATUS_LABELS.get(sc, safe)
                tds.append(f'<td><span class="badge b-{bcls}">{bico} {label}</span></td>')
            elif col == "_facility":
                f_color = "#38bdf8" if safe == "Post Office" else ("#f59e0b" if safe == "Agent" else "#a78bfa")
                tds.append(f'<td><span style="color:{f_color};font-size:11px;font-weight:600">{safe}</span></td>')
            elif col == "_age_days":
                days = int(age); hrs = int((age - days) * 24)
                txt  = f"{days}d" if days else f"{hrs}h"
                tds.append(f'<td class="{age_cls}">{txt}</td>')
            elif col == "CREATED DATE":
                tds.append(f'<td style="color:#60a5fa;font-weight:500">{safe}</td>')
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

    q_p   = request.args.get("q", "")
    pager = []
    for p in range(1, total_pages + 1):
        cls = ' class="active"' if p == page else ""
        pager.append(f'<a{cls} href="/?q={q_p}&cat={cat}&branch={branch}&date={date_filter}&sort={sort_order}&page={p}">{p}</a>')
    pager.append(f'<span class="pg-info">Rows {start+1}&#8211;{min(start+PAGE_SIZE,total):,} of {total:,}</span>')
    return table, "".join(pager)


# ── HTML template ─────────────────────────────────────────────────────────────────
HTML_TMPL = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Order Search Portal</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {
  --bg:#070d1a; --surface:#0f1829; --s2:#162035; --s3:#1c2a45;
  --border:#1e2d45; --accent:#4f8ef7; --a2:#38bdf8;
  --green:#22c55e; --amber:#f59e0b; --red:#ef4444;
  --purple:#a78bfa; --text:#e2e8f0; --muted:#64748b;
}
* { box-sizing:border-box; margin:0; padding:0; }
body { font-family:'Inter',system-ui,sans-serif; background:var(--bg); color:var(--text); min-height:100vh; font-size:14px; }

.hdr { background:linear-gradient(135deg,#0a1428 0%,#12204a 50%,#0a1428 100%);
  border-bottom:1px solid var(--border); padding:18px 28px;
  display:flex; align-items:center; gap:14px;
  position:sticky; top:0; z-index:100; box-shadow:0 4px 32px rgba(0,0,0,.5); }
.hdr h1 { font-size:18px; font-weight:700;
  background:linear-gradient(120deg,#60a5fa,#a78bfa,#38bdf8);
  -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; }
.hdr-sub { font-size:11px; color:var(--muted); margin-top:1px; }
.hdr-right { margin-left:auto; display:flex; gap:10px; align-items:center; }
.cpill { font-size:11px; padding:3px 10px; border-radius:20px;
  border:1px solid var(--border); background:var(--s2); color:var(--muted); }
.cpill.fresh { color:var(--green); }
.cpill.stale { color:var(--red); }
.rbtn { font-size:11px; padding:4px 12px; border-radius:20px;
  border:1px solid var(--border); background:var(--s2); color:var(--muted);
  cursor:pointer; transition:.15s; }
.rbtn:hover { color:var(--text); border-color:var(--accent); }

.page { max-width:1700px; margin:0 auto; padding:22px 28px; }

/* Guide Box */
.guide-box { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:14px 18px; margin-bottom:18px; font-size:12.5px; }
.guide-title { font-weight:700; color:#60a5fa; margin-bottom:8px; display:flex; align-items:center; gap:8px; cursor:pointer; }
.guide-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:10px; margin-top:8px; }
.g-item { background:var(--s2); border:1px solid var(--border); border-radius:6px; padding:8px 12px; }
.g-code { font-family:monospace; font-weight:700; color:#38bdf8; font-size:13px; margin-right:6px; }

.stats { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:10px; margin-bottom:22px; }
.sc { background:var(--surface); border:1px solid var(--border); border-radius:10px;
  padding:14px 16px; cursor:pointer; transition:.2s; user-select:none; }
.sc:hover { border-color:var(--accent); transform:translateY(-1px); box-shadow:0 4px 20px rgba(0,0,0,.3); }
.sc.active { border-color:var(--accent); background:rgba(79,142,247,.1); }
.sc .n { font-size:26px; font-weight:700; line-height:1; margin-bottom:5px; }
.sc .l { font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.5px; }
.sc-all .n     { color:var(--text); }
.sc-active .n  { color:var(--accent); }
.sc-po .n      { color:var(--a2); }
.sc-delayed .n { color:var(--amber); }
.sc-missing .n { color:var(--red); }
.sc-issue .n   { color:var(--purple); }
.sc-done .n    { color:var(--green); }
.sc-cancel .n  { color:#f87171; }

.srow { display:flex; gap:10px; margin-bottom:16px; flex-wrap:wrap; align-items:center; }
.sinput-wrap { flex:1; min-width:220px; position:relative; }
.sinput-wrap input { width:100%; background:var(--surface); border:1px solid var(--border);
  border-radius:6px; padding:11px 14px 11px 40px; color:var(--text);
  font-size:13px; font-family:inherit; outline:none; transition:.15s; }
.sinput-wrap input:focus { border-color:var(--accent); box-shadow:0 0 0 3px rgba(79,142,247,.15); }
.sinput-wrap input::placeholder { color:var(--muted); }
.ico { position:absolute; left:12px; top:50%; transform:translateY(-50%); font-size:15px; }

.bselect { background:var(--surface); border:1px solid var(--border); border-radius:6px;
  padding:10px 14px; color:var(--text); font-size:13px; font-family:inherit; outline:none;
  cursor:pointer; transition:.15s; }
.bselect:focus { border-color:var(--accent); }

.btn { padding:10px 18px; border:none; border-radius:6px; cursor:pointer;
  font-size:12px; font-weight:600; font-family:inherit; transition:.15s; white-space:nowrap; }
.btn-s { background:linear-gradient(135deg,var(--accent),var(--a2)); color:#fff; }
.btn-s:hover { opacity:.9; transform:translateY(-1px); }
.btn-xl { background:rgba(34,197,94,.12); color:var(--green); border:1px solid rgba(34,197,94,.3); }
.btn-xl:hover { background:rgba(34,197,94,.2); }
.btn-c { background:var(--s2); color:var(--muted); border:1px solid var(--border); }
.btn-c:hover { color:var(--text); }

.rinfo { font-size:12px; color:var(--muted); margin-bottom:10px; }
.rinfo strong { color:var(--text); }

.tbl-wrap { overflow-x:auto; border-radius:10px; border:1px solid var(--border);
  background:var(--surface); box-shadow:0 2px 16px rgba(0,0,0,.2); }
table { width:100%; border-collapse:collapse; font-size:12.5px; }
thead tr { background:var(--s2); }
th { padding:10px 12px; text-align:left; font-weight:600; color:var(--muted);
  font-size:10.5px; text-transform:uppercase; letter-spacing:.4px;
  white-space:nowrap; border-bottom:1px solid var(--border); }
td { padding:9px 12px; border-bottom:1px solid rgba(30,45,69,.5);
  white-space:nowrap; max-width:200px; overflow:hidden; text-overflow:ellipsis; }
tr:last-child td { border-bottom:none; }
tr:hover td { background:rgba(255,255,255,.02); }
.mono { font-family:ui-monospace,monospace; font-size:11.5px; }

.badge { display:inline-block; padding:2px 8px; border-radius:20px; font-size:10.5px; font-weight:600; white-space:nowrap; }
.b-transit   { background:rgba(56,189,248,.12); color:#38bdf8; }
.b-delivery  { background:rgba(79,142,247,.12); color:#7aa8f9; }
.b-pickup    { background:rgba(167,139,250,.12); color:#c084fc; }
.b-done      { background:rgba(34,197,94,.12); color:#4ade80; }
.b-issue     { background:rgba(239,68,68,.12); color:#f87171; }
.b-return_   { background:rgba(245,158,11,.12); color:#fbbf24; }
.b-default_  { background:rgba(100,116,139,.12); color:var(--muted); }

.age-ok     { color:var(--green); }
.age-warn   { color:var(--amber); }
.age-danger { color:var(--red); font-weight:600; }

.empty { padding:80px 20px; text-align:center; color:var(--muted); }

.pager { display:flex; gap:6px; align-items:center; margin-top:14px; flex-wrap:wrap; }
.pager a { background:var(--surface); border:1px solid var(--border); color:var(--text);
  padding:5px 11px; border-radius:6px; cursor:pointer; font-size:11px;
  text-decoration:none; transition:.15s; }
.pager a:hover, .pager a.active { background:var(--accent); border-color:var(--accent); color:#fff; }
.pg-info { font-size:11px; color:var(--muted); margin-left:6px; }

#ld { display:none; position:fixed; inset:0; background:rgba(7,13,26,.75);
  z-index:999; align-items:center; justify-content:center; flex-direction:column; gap:12px; }
#ld.on { display:flex; }
.spin { width:36px; height:36px; border:3px solid var(--border);
  border-top-color:var(--accent); border-radius:50%; animation:sp .7s linear infinite; }
@keyframes sp { to { transform:rotate(360deg); } }

@media(max-width:700px) {
  .page { padding:14px; }
  .hdr { padding:12px 14px; }
  .stats { grid-template-columns:repeat(2,1fr); }
}
</style>
</head>
<body>

<div id="ld"><div class="spin"></div><p style="font-size:13px;color:#64748b">Loading&hellip;</p></div>

<header class="hdr">
  <div style="font-size:24px">&#128666;</div>
  <div>
    <h1>Order Search Portal</h1>
    <div class="hdr-sub">14-Day Tracker &bull; Live Bill Lookup</div>
  </div>
  <div class="hdr-right">
    <span class="cpill ~~cc~~" id="cpill">~~cl~~</span>
    <button class="rbtn" onclick="forceRefresh()">&#8635; Refresh</button>
  </div>
</header>

<div class="page">

  <div class="stats">
    <div class="sc sc-active ~~a_active~~" onclick="gocat('active')">
      <div class="n">~~cnt_active~~</div><div class="l">&#128230; Active Pending</div>
    </div>
    <div class="sc sc-po ~~a_po_only~~" onclick="gocat('po_only')">
      <div class="n">~~cnt_po_only~~</div><div class="l">&#127963; PO Only</div>
    </div>
    <div class="sc sc-delayed ~~a_delayed~~" onclick="gocat('delayed')">
      <div class="n">~~cnt_delayed~~</div><div class="l">&#9200; Delayed &gt;1d</div>
    </div>
    <div class="sc sc-missing ~~a_missing~~" onclick="gocat('missing')">
      <div class="n">~~cnt_missing~~</div><div class="l">&#128680; Missing &gt;7d</div>
    </div>
    <div class="sc sc-issue ~~a_issue~~" onclick="gocat('issue')">
      <div class="n">~~cnt_issue~~</div><div class="l">&#9888; Issues</div>
    </div>
    <div class="sc sc-done ~~a_done~~" onclick="gocat('done')">
      <div class="n">~~cnt_done~~</div><div class="l">&#10003; 410 Complete</div>
    </div>
    <div class="sc sc-cancel ~~a_cancel~~" onclick="gocat('cancel')">
      <div class="n">~~cnt_cancel~~</div><div class="l">&#8617; Cancel/Return</div>
    </div>
    <div class="sc sc-all ~~a_all~~" onclick="gocat('all')">
      <div class="n">~~cnt_all~~</div><div class="l">&#128203; All Bills</div>
    </div>
  </div>

  <form method="get" action="/" onsubmit="showLoading()">
    <input type="hidden" name="cat" id="catIn" value="~~cat~~">
    <div class="srow">
      <div class="sinput-wrap">
        <span class="ico">&#128269;</span>
        <input type="text" name="q" id="qIn" value="~~q~~"
          placeholder="Search order ID &bull; phone &bull; status code (410, 306, 401) &bull; sender &bull; receiver &bull; post office &bull; note &hellip;"
          autocomplete="off" autofocus>
      </div>

      <select name="date" id="dateSelect" class="bselect" onchange="this.form.submit(); showLoading();">
        ~~date_options~~
      </select>

      <select name="branch" id="branchSelect" class="bselect" onchange="this.form.submit(); showLoading();">
        ~~branch_options~~
      </select>

      <select name="sort" id="sortSelect" class="bselect" onchange="this.form.submit(); showLoading();">
        <option value="desc" ~~sort_desc_sel~~>📅 Newest First</option>
        <option value="asc" ~~sort_asc_sel~~>📅 Oldest First</option>
      </select>

      <button type="submit" class="btn btn-s">Search</button>
      <button type="button" class="btn btn-c" onclick="clearAll()">&times; Clear</button>
      <button type="button" class="btn btn-xl" onclick="dlExcel()">&#8659; Export Excel</button>
    </div>
  </form>

  <div class="rinfo">~~rinfo~~</div>

  ~~table~~

  <div class="pager">~~pager~~</div>

</div>

<script>
var _cat="~~cat~~", _q="~~q~~", _branch="~~branch~~", _date="~~date~~", _sort="~~sort~~";
function toggleGuide() {
  var g = document.getElementById('guideGrid');
  g.style.display = (g.style.display === 'none') ? 'grid' : 'none';
}
function gocat(c) {
  _cat=c; document.getElementById('catIn').value=c;
  showLoading();
  window.location='/?cat='+c+(_branch?'&branch='+encodeURIComponent(_branch):'')+(_date?'&date='+encodeURIComponent(_date):'')+(_sort?'&sort='+encodeURIComponent(_sort):'')+(_q?'&q='+encodeURIComponent(_q):'');
}
function showLoading() { document.getElementById('ld').classList.add('on'); }
function clearAll() { window.location='/?cat=active'; }
function forceRefresh() {
  showLoading();
  fetch('/api/refresh',{method:'POST'}).then(function(){ window.location.reload(); });
}
function dlExcel() {
  showLoading();
  var url='/api/export?cat='+_cat+(_branch?'&branch='+encodeURIComponent(_branch):'')+(_date?'&date='+encodeURIComponent(_date):'')+(_sort?'&sort='+encodeURIComponent(_sort):'')+(_q?'&q='+encodeURIComponent(_q):'');
  fetch(url).then(function(r){ return r.blob(); }).then(function(b){
    var a=document.createElement('a');
    a.href=URL.createObjectURL(b);
    a.download='orders_~~cat~~_~~stamp~~.xlsx';
    a.click();
    document.getElementById('ld').classList.remove('on');
  });
}
window.addEventListener('pageshow',function(){ document.getElementById('ld').classList.remove('on'); });
setInterval(function(){
  fetch('/api/cache_status').then(function(r){ return r.json(); }).then(function(d){
    var p=document.getElementById('cpill');
    p.textContent=d.label; p.className='cpill '+d.cls;
  });
},30000);
</script>
</body>
</html>"""

AUTH_TMPL = """<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>Login</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
<style>
* { box-sizing:border-box; margin:0; padding:0; }
body { font-family:'Inter',sans-serif; background:#070d1a; color:#e2e8f0;
  min-height:100vh; display:flex; align-items:center; justify-content:center; }
.box { background:#0f1829; border:1px solid #1e2d45; border-radius:10px; padding:40px; width:320px; text-align:center; }
h2 { font-size:20px; margin-bottom:6px;
  background:linear-gradient(120deg,#60a5fa,#a78bfa);
  -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; }
p { font-size:12px; color:#64748b; margin-bottom:22px; }
input { width:100%; background:#162035; border:1px solid #1e2d45; border-radius:6px;
  padding:11px; color:#e2e8f0; font-size:14px; font-family:inherit;
  outline:none; margin-bottom:12px; text-align:center; letter-spacing:2px; }
input:focus { border-color:#4f8ef7; }
button { width:100%; padding:11px; background:linear-gradient(135deg,#4f8ef7,#38bdf8);
  color:#fff; border:none; border-radius:6px; font-size:13px; font-weight:600; cursor:pointer; }
.err { color:#ef4444; font-size:12px; margin-bottom:10px; }
</style></head>
<body><div class="box">
<h2>&#128274; Order Portal</h2>
<p>Enter your access password</p>
~~err~~
<form method="post" action="/auth">
<input type="password" name="pw" placeholder="Password" autofocus>
<button type="submit">Enter</button>
</form>
</div></body></html>"""


# ── Auth ─────────────────────────────────────────────────────────────────────────
_sessions: set = set()

def _check_auth():
    if not PORTAL_PASS:
        return True
    return request.cookies.get("session") in _sessions


# ── Routes ────────────────────────────────────────────────────────────────────────
@app.get("/auth")
def auth_page():
    return _render(AUTH_TMPL, err=""), 200


@app.post("/auth")
def auth_submit():
    from flask import make_response, redirect
    import secrets
    if request.form.get("pw", "") == PORTAL_PASS:
        tok = secrets.token_hex(16)
        _sessions.add(tok)
        resp = make_response(redirect("/"))
        resp.set_cookie("session", tok, max_age=86400*7, httponly=True, samesite="Lax")
        return resp
    return _render(AUTH_TMPL, err='<p class="err">Wrong password.</p>'), 401


@app.get("/")
def index():
    from flask import redirect
    if not _check_auth():
        return redirect("/auth")

    q           = request.args.get("q", "").strip()
    cat         = request.args.get("cat", "active")
    branch      = request.args.get("branch", "ALL").strip().upper()
    date_filter = request.args.get("date", "all").strip()
    sort_order  = request.args.get("sort", "desc").strip().lower()
    page        = max(1, int(request.args.get("page", 1)))

    try:
        df = do_search(q, cat, branch, date_filter, sort_order)
    except Exception as exc:
        return f"<pre style='color:red;padding:20px'>Error loading data:\n{exc}</pre>", 500

    full = get_data()

    # Apply branch filter to stat counts if branch selected
    stat_df = full if (not branch or branch == "ALL") else full[full["_branch"] == branch]

    # Apply date filter to stat counts if date selected
    today = (datetime.utcnow() + timedelta(hours=7)).date()
    if date_filter == "today":
        stat_df = stat_df[stat_df["_date_only"] == today]
    elif date_filter == "yesterday":
        stat_df = stat_df[stat_df["_date_only"] == (today - timedelta(days=1))]
    elif date_filter == "last3":
        stat_df = stat_df[stat_df["_date_only"] >= (today - timedelta(days=2))]
    elif date_filter == "last7":
        stat_df = stat_df[stat_df["_date_only"] >= (today - timedelta(days=6))]
    elif re.match(r"^\d{4}-\d{2}-\d{2}$", date_filter):
        stat_df = stat_df[stat_df["_date_only"] == pd.to_datetime(date_filter).date()]

    active_df = stat_df[~stat_df["_sc"].isin(DONE_CODES | CANCEL_CODES)]

    counts = {
        "active":  len(active_df),
        "po_only": int((stat_df["_facility"] == "Post Office").sum()),
        "delayed": int(((active_df["_age_days"] > 1)).sum()),
        "missing": int(((active_df["_age_days"] > 7)).sum()),
        "issue":   int(active_df["_sc"].isin(ISSUE_CODES).sum()),
        "done":    int(stat_df["_sc"].isin(DONE_CODES).sum()),
        "cancel":  int(stat_df["_sc"].isin(CANCEL_CODES).sum()),
        "all":     len(stat_df),
    }

    # Build branch options
    branches_in_data = sorted(full["_branch"].unique())
    b_opts = ['<option value="ALL">🏢 All Branches</option>']
    for b in branches_in_data:
        if not b or b == "OTHER": continue
        b_name = BRANCH_NAME_MAP.get(b, b)
        sel = ' selected' if branch == b else ''
        b_opts.append(f'<option value="{b}"{sel}>{b} - {b_name}</option>')
    branch_options = "".join(b_opts)

    # Build date options
    d_opts = [
        ('<option value="all"' + (' selected' if date_filter == 'all' else '') + '>📅 All 14 Days</option>'),
        ('<option value="today"' + (' selected' if date_filter == 'today' else '') + '>📅 Today (' + today.strftime("%d/%m") + ')</option>'),
        ('<option value="yesterday"' + (' selected' if date_filter == 'yesterday' else '') + '>📅 Yesterday (' + (today - timedelta(days=1)).strftime("%d/%m") + ')</option>'),
        ('<option value="last3"' + (' selected' if date_filter == 'last3' else '') + '>📅 Last 3 Days</option>'),
        ('<option value="last7"' + (' selected' if date_filter == 'last7' else '') + '>📅 Last 7 Days</option>'),
    ]

    # Add available dates dynamically
    unique_dates = sorted([d for d in full["_date_only"].dropna().unique() if d], reverse=True)
    d_opts.append('<optgroup label="Specific Dates">')
    for ud in unique_dates:
        ud_str = ud.strftime("%Y-%m-%d")
        sel = ' selected' if date_filter == ud_str else ''
        d_opts.append(f'<option value="{ud_str}"{sel}>{ud.strftime("%d/%m/%Y")}</option>')
    d_opts.append('</optgroup>')
    date_options = "".join(d_opts)

    cl, cc        = _cache_info()
    table, pager  = _build_table(df, page, cat, branch, date_filter, sort_order)

    cat_names = {
        "active": "Active Packages", "po_only": "Post Office Only (Excl. Agent/Showroom)",
        "delayed": "Delayed >1d", "missing": "Missing >7d", "issue": "Issues",
        "done": "Status 410 / Completed", "cancel": "Cancel / Returned", "all": "All Orders"
    }

    rinfo_parts = [f'Found <strong>{len(df):,}</strong> orders']
    if q:
        rinfo_parts.append(f'matching &ldquo;<strong>{q}</strong>&rdquo;')
    if date_filter != "all":
        rinfo_parts.append(f'on <strong>{date_filter}</strong>')
    if branch and branch != "ALL":
        rinfo_parts.append(f'in Branch <strong>{branch} ({BRANCH_NAME_MAP.get(branch, branch)})</strong>')
    rinfo_parts.append(f'&rsaquo; <strong>{cat_names.get(cat, cat)}</strong>')
    rinfo_parts.append(f'({ "Newest First" if sort_order == "desc" else "Oldest First" })')

    rinfo = " ".join(rinfo_parts)

    def ac(c): return "active" if cat == c else ""
    stamp = datetime.utcnow().strftime("%d%m%Y")

    html = _render(
        HTML_TMPL,
        q=q, cat=cat, branch=branch, date=date_filter, sort=sort_order, page=page, stamp=stamp,
        cl=cl, cc=cc,
        cnt_active=f"{counts['active']:,}",
        cnt_po_only=f"{counts['po_only']:,}",
        cnt_delayed=f"{counts['delayed']:,}",
        cnt_missing=f"{counts['missing']:,}",
        cnt_issue=f"{counts['issue']:,}",
        cnt_done=f"{counts['done']:,}",
        cnt_cancel=f"{counts['cancel']:,}",
        cnt_all=f"{counts['all']:,}",
        a_active=ac("active"), a_po_only=ac("po_only"),
        a_delayed=ac("delayed"), a_missing=ac("missing"),
        a_issue=ac("issue"), a_done=ac("done"), a_cancel=ac("cancel"), a_all=ac("all"),
        branch_options=branch_options,
        date_options=date_options,
        sort_desc_sel=('selected' if sort_order == 'desc' else ''),
        sort_asc_sel=('selected' if sort_order == 'asc' else ''),
        rinfo=rinfo, table=table, pager=pager,
    )
    return html, 200


@app.get("/api/export")
def api_export():
    if not _check_auth():
        return "Unauthorized", 401
    q           = request.args.get("q", "").strip()
    cat         = request.args.get("cat", "active")
    branch      = request.args.get("branch", "ALL").strip().upper()
    date_filter = request.args.get("date", "all").strip()
    sort_order  = request.args.get("sort", "desc").strip().lower()
    try:
        df = do_search(q, cat, branch, date_filter, sort_order)
    except Exception as exc:
        return str(exc), 500
    xlsx  = make_excel_bytes(df)
    stamp = datetime.utcnow().strftime("%d%m%Y_%H%M")
    fname = f"orders_{cat}_{branch}_{stamp}.xlsx"
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
        _cache["mtime"] = 0.0
    return "ok"


@app.get("/api/cache_status")
def api_cache_status():
    label, cls = _cache_info()
    return {"label": label, "cls": cls}


@app.get("/health")
def health():
    return {"status": "ok", "rows": len(_cache["df"])}


# ── Local dev ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
    parent_cfg = os.path.join(os.path.dirname(__file__), "..", "config.json")
    if os.path.exists(parent_cfg):
        with open(parent_cfg, encoding="utf-8") as f:
            cfg = json.load(f)["api"]
        for k, v in [("API_URL", cfg.get("url","")),
                     ("API_BEARER", cfg.get("bearer_token","")),
                     ("API_BRANCH", cfg.get("branch_code","")),
                     ("API_CLIENT_ID", cfg.get("x_client_id","TMS_ANDROID")),
                     ("API_REFERER", cfg.get("referer","https://opsexpress.metfone.com.kh/"))]:
            os.environ.setdefault(k, v)
            globals()[k] = os.environ[k]
    app.run(host="0.0.0.0", port=5055, debug=True)
