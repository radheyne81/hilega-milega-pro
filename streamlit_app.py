#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HILEGA-MILEGA v4.0 PRO — STREAMLIT CLOUD & MOBILE/WEB TERMINAL
==============================================================
Ready for 1-click deployment on https://share.streamlit.io/
Fixes line 498 SyntaxError (unterminated string literal) and includes:
  1. Full Interactive 8-Tab PRO Web Terminal (embedded HTML/JS + Chart.js + 360 X-Ray)
  2. Native Streamlit Mobile & Desktop Tables + Fool-Proof 6-Gate Intraday Validator
  3. Live NSE Quote Fetcher (NIFTY 50, Sector Indices, and Nifty-500 Stocks)
  4. Built-in Auto-Refresh Timer (OFF / 15s / 30s / 60s / 5m)
  5. Drag-and-Drop EOD .xlsx Uploader + Automatic hm_data.json / HTML Loader
"""

from __future__ import annotations

import datetime as dt
import io
import json
import math
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Hilega-Milega v4.0 PRO Terminal",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_DIR = Path(__file__).resolve().parent
HTML_PATH = APP_DIR / "HilegaMilega_PRO_Terminal.html"
JSON_PATH = APP_DIR / "hm_data.json"

# Custom dark styling for Streamlit chrome
st.markdown(
    """
    <style>
      .stApp { background-color: #070d19; color: #e2e8f0; }
      .block-container { padding-top: 1.0rem !important; padding-bottom: 1.5rem !important; max-width: 1650px !important; }
      div[data-testid="stMetric"] {
          background: rgba(15, 23, 42, 0.85);
          border: 1px solid rgba(51, 65, 85, 0.75);
          border-radius: 12px;
          padding: 10px 14px;
      }
      .hm-banner {
          background: linear-gradient(90deg, rgba(15,23,42,0.95) 0%, rgba(30,41,59,0.90) 100%);
          border: 1px solid rgba(59,130,246,0.35);
          border-left: 4px solid #3b82f6;
          border-radius: 14px;
          padding: 12px 18px;
          margin-bottom: 12px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# DATA EXTRACTION HELPERS (.xlsx / hm_data.json / HTML fallback)
# -----------------------------------------------------------------------------
def _clean_val(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, (int, float, bool)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return None
        return round(v, 4) if isinstance(v, float) else v
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def extract_workbook_bytes_to_json(xlsx_bytes: bytes) -> Dict[str, Any]:
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    compact: Dict[str, Any] = {}

    for name in wb.sheetnames:
        if name == "Dashboard":
            continue
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = 2 if name == "Guide" else 3
        if len(rows) <= hdr_idx:
            continue

        if name == "Guide":
            cols = ["Topic", "Explanation"]
        else:
            cols = [
                str(c).strip() if c is not None else f"Col_{i}"
                for i, c in enumerate(rows[hdr_idx])
            ]

        data_rows = []
        for r in rows[hdr_idx + 1 :]:
            if r is None or all(x is None or str(x).strip() == "" for x in r):
                continue
            data_rows.append(
                [_clean_val(r[i]) if i < len(r) else None for i in range(len(cols))]
            )

        compact[name] = {"c": cols, "r": data_rows}

    if "Dashboard" in wb.sheetnames:
        ws_d = wb["Dashboard"]
        d_rows = list(ws_d.iter_rows(values_only=True))

        def cell_str(r_idx: int, c_idx: int) -> str:
            if r_idx < len(d_rows) and c_idx < len(d_rows[r_idx]):
                val = d_rows[r_idx][c_idx]
                return str(val).strip() if val is not None else ""
            return ""

        subtitle = cell_str(1, 0)
        regime_banner = cell_str(2, 0)
        regime_label = "NEUTRAL"
        regime_score = 0
        m_reg = re.search(
            r"MARKET REGIME:\s*([A-Z]+)\s*\(score\s*([+-]?\d+)\)", regime_banner
        )
        if m_reg:
            regime_label = m_reg.group(1)
            regime_score = int(m_reg.group(2))

        stance = (
            regime_banner.split(">>")[-1].strip()
            if ">>" in regime_banner
            else regime_banner
        )

        nifty_map: Dict[str, Any] = {}
        if len(d_rows) > 10:
            hdrs = d_rows[9] if len(d_rows) > 9 else []
            vals = d_rows[10] if len(d_rows) > 10 else []
            for i, h in enumerate(hdrs):
                if h is not None and i < len(vals) and vals[i] is not None:
                    key = (
                        str(h)
                        .replace("NIFTY ", "")
                        .replace("CPR ", "")
                        .replace(" (14)", "")
                        .strip()
                    )
                    nifty_map[key] = _clean_val(vals[i])
        nifty_map["read"] = cell_str(11, 0)

        narratives = []
        for r_idx in range(14, min(23, len(d_rows))):
            t = cell_str(r_idx, 0)
            txt = cell_str(r_idx, 2)
            if t and txt:
                narratives.append({"title": t, "text": txt})

        compact["DashboardMeta"] = {
            "subtitle": subtitle,
            "regimeBanner": regime_banner,
            "regimeLabel": regime_label,
            "regimeScore": regime_score,
            "regimeStance": stance,
            "kpis": [
                {
                    "label": cell_str(4, 0) or "NIFTY 50",
                    "value": cell_str(5, 0),
                    "sub": cell_str(6, 0),
                    "tone": "emerald",
                },
                {
                    "label": cell_str(4, 2) or "BREADTH > 50-DMA",
                    "value": cell_str(5, 2),
                    "sub": cell_str(6, 2),
                    "tone": "amber",
                },
                {
                    "label": cell_str(4, 4) or "ADVANCE / DECLINE",
                    "value": cell_str(5, 4),
                    "sub": cell_str(6, 4),
                    "tone": "emerald",
                },
                {
                    "label": cell_str(4, 6) or "HM DAILY BULL vs BEAR",
                    "value": cell_str(5, 6),
                    "sub": cell_str(6, 6),
                    "tone": "rose",
                },
                {
                    "label": cell_str(4, 8) or "20D HIGHS / LOWS",
                    "value": cell_str(5, 8),
                    "sub": cell_str(6, 8),
                    "tone": "sky",
                },
                {
                    "label": cell_str(4, 10) or "AVG DELIVERY",
                    "value": cell_str(5, 10),
                    "sub": cell_str(6, 10),
                    "tone": "violet",
                },
            ],
            "niftyMap": nifty_map,
            "narratives": narratives,
        }

    return compact


@st.cache_data(show_spinner=False)
def load_default_hm_data() -> Dict[str, Any]:
    # 1. Try hm_data.json
    if JSON_PATH.exists():
        try:
            return json.loads(JSON_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass

    # 2. Try newest HilegaMilega_PRO_*.xlsx in folder
    xlsx_files = sorted(
        APP_DIR.glob("HilegaMilega_PRO_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if xlsx_files:
        try:
            return extract_workbook_bytes_to_json(xlsx_files[0].read_bytes())
        except Exception:
            pass

    # 3. Extract embedded RAW_HM_DATA from HilegaMilega_PRO_Terminal.html
    if HTML_PATH.exists():
        try:
            html_text = HTML_PATH.read_text(encoding="utf-8")
            m = re.search(
                r"const\s+RAW_HM_DATA\s*=\s*(\{.*?\});\s*\n", html_text, re.DOTALL
            )
            if m:
                return json.loads(m.group(1))
        except Exception:
            pass

    return {}


def sheet_to_df(hm_data: Dict[str, Any], sheet_name: str) -> pd.DataFrame:
    payload = hm_data.get(sheet_name)
    if not payload or "c" not in payload or "r" not in payload:
        return pd.DataFrame()
    return pd.DataFrame(payload["r"], columns=payload["c"])


# -----------------------------------------------------------------------------
# LIVE NSE QUOTE FETCHER (Cached 15s)
# -----------------------------------------------------------------------------
@st.cache_data(ttl=15, show_spinner=False)
def fetch_live_nse_snapshot() -> Dict[str, Any]:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "en-US,en;q=0.9",
            "Referer": "https://www.nseindia.com/market-data/live-equity-market",
        }
    )
    quotes: Dict[str, Dict[str, float]] = {}
    indices: Dict[str, Dict[str, float]] = {}

    try:
        session.get("https://www.nseindia.com", timeout=6)
        resp = session.get(
            "https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20500",
            timeout=8,
        )
        if resp.ok:
            for item in resp.json().get("data", []):
                sym = str(item.get("symbol", "")).strip().upper()
                if not sym:
                    continue
                quotes[sym] = {
                    "open": float(item.get("open") or 0),
                    "high": float(item.get("dayHigh") or 0),
                    "low": float(item.get("dayLow") or 0),
                    "ltp": float(item.get("lastPrice") or 0),
                    "prevClose": float(item.get("previousClose") or 0),
                    "pChange": float(item.get("pChange") or 0),
                }
    except Exception:
        pass

    try:
        resp_i = session.get("https://www.nseindia.com/api/allIndices", timeout=8)
        if resp_i.ok:
            for item in resp_i.json().get("data", []):
                idx_name = str(item.get("index", "")).strip().upper()
                indices[idx_name] = {
                    "open": float(item.get("open") or 0),
                    "high": float(item.get("high") or 0),
                    "low": float(item.get("low") or 0),
                    "ltp": float(item.get("last") or 0),
                    "prevClose": float(item.get("previousClose") or 0),
                    "pChange": float(item.get("percentChange") or 0),
                }
    except Exception:
        pass

    ist_now = dt.datetime.utcnow() + dt.timedelta(hours=5, minutes=30)
    hhmm = ist_now.strftime("%H:%M")
    if hhmm < "09:30":
        time_bucket = "09:20"
    elif hhmm <= "11:30":
        time_bucket = "09:35"
    elif hhmm <= "13:30":
        time_bucket = "12:15"
    elif hhmm <= "14:30":
        time_bucket = "13:45"
    else:
        time_bucket = "14:45"

    return {
        "status": "ok" if quotes else "offline",
        "timestamp": ist_now.strftime("%d-%b-%Y %H:%M:%S IST"),
        "hhmm": hhmm,
        "timeBucket": time_bucket,
        "nifty50": indices.get("NIFTY 50") or quotes.get("NIFTY 50") or {},
        "indices": indices,
        "quotes": quotes,
    }


# -----------------------------------------------------------------------------
# POSITION SIZING & 9-FACTOR SUPER-CONFLUENCE ENGINE
# -----------------------------------------------------------------------------
def calc_size(
    entry: float,
    sl: float,
    capital: float,
    risk_pct: float,
    max_pos_pct: float,
    leverage: float = 1.0,
) -> Tuple[int, int, int]:
    risk_per_share = abs(float(entry or 0) - float(sl or 0))
    if entry <= 0 or risk_per_share <= 0:
        return 0, 0, 0
    q_risk = int((capital * (risk_pct / 100.0)) // risk_per_share)
    q_cap = int((capital * leverage * (max_pos_pct / 100.0)) // entry)
    qty = max(0, min(q_risk, q_cap))
    return qty, int(round(qty * entry)), int(round(qty * risk_per_share))


def compute_super_confluence_df(
    hm_data: Dict[str, Any],
    capital: float,
    intra_risk_pct: float,
    max_pos_pct: float,
    intra_lev: float,
) -> pd.DataFrame:
    df_intra = sheet_to_df(hm_data, "Intraday Plan")
    if df_intra.empty:
        return pd.DataFrame()

    df_swing = sheet_to_df(hm_data, "Swing Plan")
    df_ind = sheet_to_df(hm_data, "Industry Heat")
    df_del = sheet_to_df(hm_data, "Delivery Analysis")
    df_wbo = sheet_to_df(hm_data, "Weekly Breakouts")
    df_ib = sheet_to_df(hm_data, "Inside Bar Setups")
    df_bt = sheet_to_df(hm_data, "Backtest Stats")
    df_sig = sheet_to_df(hm_data, "Signals - Stocks")

    swing_map = (
        {str(r["Symbol"]).upper(): r for _, r in df_swing.iterrows()}
        if not df_swing.empty
        else {}
    )
    ind_map = (
        {str(r["Industry"]): r for _, r in df_ind.iterrows()}
        if not df_ind.empty
        else {}
    )
    del_map = (
        {str(r["Symbol"]).upper(): r for _, r in df_del.iterrows()}
        if not df_del.empty
        else {}
    )
    wbo_map = (
        {str(r["Symbol"]).upper(): r for _, r in df_wbo.iterrows()}
        if not df_wbo.empty
        else {}
    )
    bt_map = (
        {str(r["Symbol"]).upper(): r for _, r in df_bt.iterrows()}
        if not df_bt.empty
        else {}
    )

    ib_daily_set = set()
    if not df_ib.empty:
        for _, r in df_ib.iterrows():
            if str(r.get("TF", "")) == "Daily":
                ib_daily_set.add(str(r.get("Symbol", "")).upper())

    sig_daily_map = {}
    if not df_sig.empty:
        for _, r in df_sig.iterrows():
            if str(r.get("TF", "")) == "Daily":
                sig_daily_map[str(r.get("Symbol", "")).upper()] = r

    out_rows = []
    for _, r in df_intra.iterrows():
        sym = str(r["Symbol"]).upper()
        side = str(r["Side"]).upper()
        is_long = side == "BUY"

        sw = swing_map.get(sym, {})
        sw_score = float(sw.get("Score") or 0) if len(sw) else 0.0
        ind_obj = ind_map.get(str(r.get("Industry", "")), {})
        iq = str(ind_obj.get("Quadrant") or "N/A")

        del_obj = del_map.get(sym, {})
        del_pct = float(r.get("DEL %") or del_obj.get("DEL %") or 0)
        del_trend = float(del_obj.get("DEL Trend pp") or 0)

        wbo_obj = wbo_map.get(sym, {})
        wbo_type = str(wbo_obj.get("Breakout Type") or "")

        sig_d = sig_daily_map.get(sym, {})
        hm5_tag = str(sig_d.get("HM 5EMA Tag") or "")
        rsi9 = float(sig_d.get("RSI9") or sw.get("RSI9") or 50.0)

        bt_obj = bt_map.get(sym, {})
        edge5 = float(bt_obj.get("Edge pp") or 0)
        brk_pct = float(r.get("Hist PDH/PDL Break %") or 0)

        pp = float(r.get("Pivot PP") or 0)
        bc = float(r.get("CPR BC") or 0)
        tc = float(r.get("CPR TC") or 0)
        cpr_w = (abs(tc - bc) / pp * 100.0) if pp > 0 else 1.0

        pts = float(r.get("Score") or 0) * 0.45
        pillars: List[str] = []

        is_dual = sw_score >= 55 and (
            (is_long and str(sw.get("Side")) == "BUY")
            or (not is_long and str(sw.get("Side")) == "SELL")
        )
        if is_dual:
            pts += sw_score * 0.20
            pillars.append(f"Dual I+S ({sw_score:.0f})")
        else:
            pts += 8.0

        sector_aligned = (is_long and iq in ("LEADING", "IMPROVING")) or (
            not is_long and iq in ("LAGGING", "WEAKENING")
        )
        if sector_aligned:
            pts += 8.0
            pillars.append(f"Sector {iq}")
        else:
            pts -= 4.0

        if (
            is_long
            and r.get("HM Daily") == "BULL"
            and r.get("HM Weekly") == "BULL"
        ) or (
            not is_long
            and r.get("HM Daily") == "BEAR"
            and r.get("HM Weekly") == "BEAR"
        ):
            pts += 6.0
            pillars.append(f"MTF D+W {r.get('HM Daily')}")

        if del_pct >= 60 and del_trend >= 5:
            pts += 8.0
            pillars.append(f"Inst. DEL {del_pct:.0f}% (+{del_trend:.0f}pp)")
        elif del_pct >= 50 and del_trend >= 0:
            pts += 5.0
            pillars.append(f"DEL {del_pct:.0f}%")
        elif del_pct < 35 or del_trend <= -10:
            pts -= 5.0

        if cpr_w <= 0.30:
            pts += 4.0
            pillars.append(f"Narrow CPR ({cpr_w:.2f}%)")
        if sym in ib_daily_set or "Compression" in str(r.get("Setup", "")):
            pts += 3.5
            pillars.append("NR7/Inside Coil")
        if (is_long and wbo_type == "5dayweekBullBO") or (
            not is_long and wbo_type == "5dayweekBearBO"
        ):
            pts += 3.5
            pillars.append("5D Weekly BO")
        if hm5_tag in ("Hilega5EMABull", "Milega5EMABear"):
            pts += 3.0
            pillars.append("5-EMA Touch")
        if brk_pct >= 55:
            pts += 3.0
            pillars.append(f"Hist Brk {brk_pct:.0f}%")
        if edge5 >= 3.0:
            pts += 3.0
            pillars.append(f"5D Edge +{edge5:.1f}pp")
        if (is_long and rsi9 > 80) or (not is_long and rsi9 < 20):
            pts -= 6.0
            pillars.append(f"RSI9 {rsi9:.0f} Warn")

        expert_score = round(min(99.0, max(35.0, pts)), 1)
        qty, cap_rs, risk_rs = calc_size(
            float(r.get("Entry") or 0),
            float(r.get("SL") or 0),
            capital,
            intra_risk_pct,
            max_pos_pct,
            intra_lev,
        )

        row_dict = dict(r)
        row_dict["Expert Score"] = expert_score
        row_dict["Swing Score"] = round(sw_score, 1) if sw_score else None
        row_dict["Sector Quadrant"] = iq
        row_dict["CPR Width %"] = round(cpr_w, 2)
        row_dict["DEL Trend pp"] = round(del_trend, 2)
        row_dict["RSI9"] = round(rsi9, 1)
        row_dict["Dyn Qty"] = qty
        row_dict["Dyn Capital (Rs)"] = cap_rs
        row_dict["Dyn Risk (Rs)"] = risk_rs
        row_dict["Confluence Pillars"] = " · ".join(pillars)
        out_rows.append(row_dict)

    df_out = pd.DataFrame(out_rows)
    df_out.sort_values(by="Expert Score", ascending=False, inplace=True)
    df_out.reset_index(drop=True, inplace=True)
    df_out.index = df_out.index + 1
    df_out.index.name = "Super Rank"
    return df_out


# -----------------------------------------------------------------------------
# TOP CONTROL BAR (Capital Sizer, Upload .xlsx, Live Feed & Auto-Refresh)
# -----------------------------------------------------------------------------
col_brand, col_cap, col_risk, col_upload, col_live = st.columns([2.6, 1.3, 1.5, 1.6, 1.6])

with col_brand:
    st.markdown(
        """
        <div class="hm-banner">
          <div style="font-size:1.05rem;font-weight:800;color:#ffffff;">
            📈 HILEGA-MILEGA <span style="color:#60a5fa;">v4.0 PRO</span> — Cloud & Mobile Decision Terminal
          </div>
          <div style="font-size:0.76rem;color:#94a3b8;margin-top:2px;">
            9-Factor Cross-Sheet Super-Confluence · Fool-Proof 9:30 AM 6-Gate Validator · All 14 Sheets
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_cap:
    capital_rs = st.number_input(
        "Trading Capital (₹)",
        min_value=50000,
        max_value=100000000,
        value=1000000,
        step=50000,
    )
    max_pos_pct = st.number_input(
        "Max Pos % of Cap", min_value=5.0, max_value=100.0, value=20.0, step=5.0
    )

with col_risk:
    c_r1, c_r2 = st.columns(2)
    with c_r1:
        intra_risk_pct = st.number_input(
            "Intra Risk %", min_value=0.1, max_value=5.0, value=0.5, step=0.1
        )
    with c_r2:
        swing_risk_pct = st.number_input(
            "Swing Risk %", min_value=0.1, max_value=5.0, value=1.0, step=0.1
        )
    intra_leverage = st.number_input(
        "Intraday MIS Leverage (x)", min_value=1.0, max_value=5.0, value=1.0, step=0.5
    )

with col_upload:
    uploaded_xlsx = st.file_uploader(
        "📂 Load New EOD .xlsx Workbook",
        type=["xlsx", "xls"],
        help="Drop any HilegaMilega_PRO_YYYY-MM-DD_HHMM.xlsx file here to refresh all 14 sheets immediately.",
    )

with col_live:
    enable_live_nse = st.toggle(
        "📡 Fetch Live NSE Quotes",
        value=False,
        help="Fetches real-time NSE Open & LTP for NIFTY 50 and Nifty-500 stocks.",
    )
    auto_refresh_label = st.selectbox(
        "🔄 Auto-Refresh Interval",
        options=["OFF", "15s", "30s", "60s", "5m"],
        index=1,
        help="Automatically refreshes live quotes and the Intraday Gatekeeper.",
    )

# Resolve active dataset (Uploaded Excel overrides default)
if uploaded_xlsx is not None:
    try:
        hm_data = extract_workbook_bytes_to_json(uploaded_xlsx.getvalue())
        st.toast(f"Loaded {uploaded_xlsx.name} successfully!", icon="✅")
    except Exception as exc:
        st.error(f"Could not parse uploaded workbook: {exc}")
        hm_data = load_default_hm_data()
else:
    hm_data = load_default_hm_data()

live_snapshot: Dict[str, Any] = {}
if enable_live_nse:
    with st.spinner("Fetching live NSE quotes..."):
        live_snapshot = fetch_live_nse_snapshot()

# -----------------------------------------------------------------------------
# INTERFACE MODE SELECTOR
# -----------------------------------------------------------------------------
view_mode = st.radio(
    "Select Terminal View Mode",
    options=[
        "⚡ Full Interactive PRO Web Terminal (8-Tab HTML Canvas + 360° X-Ray Modal)",
        "📊 Native Streamlit Tables & 6-Gate Validator (Mobile & CSV Export Friendly)",
    ],
    index=0,
    horizontal=True,
    label_visibility="collapsed",
)

# =============================================================================
# MODE 1: EMBEDDED FULL INTERACTIVE 8-TAB HTML PRO TERMINAL
# =============================================================================
if view_mode.startswith("⚡"):
    if HTML_PATH.exists():
        html_str = HTML_PATH.read_text(encoding="utf-8")

        # If user uploaded a new Excel or hm_data is loaded, patch RAW_HM_DATA inside HTML
        if hm_data:
            json_str = json.dumps(hm_data, separators=(",", ":"))
            pattern = re.compile(
                r"const\s+RAW_HM_DATA\s*=\s*\{.*?\};\s*\n", re.DOTALL
            )
            replacement = f"const RAW_HM_DATA = {json_str};\n"
            if pattern.search(html_str):
                html_str = pattern.sub(lambda _: replacement, html_str, count=1)

        # Inject Streamlit live snapshot & capital settings cleanly before </head>
        inject_payload = {
            "capital": capital_rs,
            "intraRiskPct": intra_risk_pct,
            "swingRiskPct": swing_risk_pct,
            "maxPosPct": max_pos_pct,
            "intraLeverage": intra_leverage,
            "liveSnapshot": live_snapshot if enable_live_nse else None,
        }
        inject_json = json.dumps(inject_payload, separators=(",", ":"))
        live_inject_js = (
            "<script>"
            f"window.__STREAMLIT_CONFIG__ = {inject_json};"
            "window.__STREAMLIT_LIVE_SNAPSHOT__ = window.__STREAMLIT_CONFIG__.liveSnapshot;"
            "window.addEventListener('DOMContentLoaded', function() {"
            "  if (typeof STATE !== 'undefined' && window.__STREAMLIT_CONFIG__) {"
            "    STATE.capital = window.__STREAMLIT_CONFIG__.capital;"
            "    STATE.intraRiskPct = window.__STREAMLIT_CONFIG__.intraRiskPct;"
            "    STATE.swingRiskPct = window.__STREAMLIT_CONFIG__.swingRiskPct;"
            "    STATE.maxPosPct = window.__STREAMLIT_CONFIG__.maxPosPct;"
            "    STATE.intraLeverage = window.__STREAMLIT_CONFIG__.intraLeverage;"
            "    if (typeof renderAllViews === 'function') renderAllViews();"
            "  }"
            "});"
            "</script>"
        )
        html_str = html_str.replace("</head>", live_inject_js + "</head>", 1)

        components.html(html_str, height=1180, scrolling=True)
    else:
        st.warning(
            "HilegaMilega_PRO_Terminal.html not found in repository root. Switching to Native Streamlit View."
        )

# =============================================================================
# MODE 2: NATIVE STREAMLIT TABLES & FOOL-PROOF 6-GATE VALIDATOR
# =============================================================================
else:
    meta = hm_data.get("DashboardMeta", {})
    kpis = meta.get("kpis", [])
    if kpis:
        k_cols = st.columns(len(kpis))
        for idx, k in enumerate(kpis):
            with k_cols[idx]:
                st.metric(
                    label=k.get("label", ""),
                    value=k.get("value", ""),
                    delta=k.get("sub", ""),
                )

    tabs = st.tabs(
        [
            "🎯 Super-Confluence & 6-Gate Validator",
            "⚡ Intraday Plan (50)",
            "📅 Swing Plan 3-10D (60)",
            "🧭 Index, Sectors & Industry Heat",
            "🏦 Delivery (496), Weekly BO & Inside Bars",
            "🔬 MTF (418), Signals & Backtest Edge",
            "📖 Guide & Rules",
        ]
    )

    # TAB 1: SUPER-CONFLUENCE & 6-GATE VALIDATOR
    with tabs[0]:
        df_super = compute_super_confluence_df(
            hm_data, capital_rs, intra_risk_pct, max_pos_pct, intra_leverage
        )
        if df_super.empty:
            st.info("No Intraday Plan data found.")
        else:
            c_left, c_right = st.columns([1.15, 1.85])
            with c_left:
                st.subheader("🛡️ Fool-Proof 6-Gate Intraday Validator")
                symbols_list = df_super["Symbol"].tolist()
                sel_sym = st.selectbox(
                    "Select Intraday Setup",
                    options=symbols_list,
                    format_func=lambda s: f"{s} ({df_super[df_super['Symbol']==s].iloc[0]['Side']} | Expert {df_super[df_super['Symbol']==s].iloc[0]['Expert Score']})",
                )
                row_s = df_super[df_super["Symbol"] == sel_sym].iloc[0]
                is_long = str(row_s["Side"]).upper() == "BUY"
                entry_px = float(row_s["Entry"])
                sl_px = float(row_s["SL"])
                t1_px = float(row_s["T1 (1R)"])
                t2_px = float(row_s["T2 (2R)"])
                pp_px = float(row_s["Pivot PP"])
                prev_close = float(row_s["Prev Close"])

                live_q = (
                    live_snapshot.get("quotes", {}).get(sel_sym, {})
                    if live_snapshot
                    else {}
                )
                default_open = (
                    float(live_q.get("open"))
                    if live_q.get("open")
                    else round((prev_close + entry_px) / 2.0, 2)
                )
                default_ltp = (
                    float(live_q.get("ltp"))
                    if live_q.get("ltp")
                    else round(entry_px * (1.0015 if is_long else 0.9985), 2)
                )

                sc1, sc2 = st.columns(2)
                with sc1:
                    sim_open = st.number_input(
                        "9:15 Open Price (₹)", value=default_open, step=0.5
                    )
                    sim_time = st.selectbox(
                        "Session Time Window (IST)",
                        options=[
                            "09:15–09:29 (Wait 1st Bar)",
                            "09:30–11:30 (Prime Window)",
                            "11:30–13:30 (Midday Lull)",
                            "13:30–14:30 (Afternoon)",
                            "After 14:30 (No Fresh Entry)",
                        ],
                        index=1,
                    )
                with sc2:
                    sim_ltp = st.number_input(
                        "9:30 Close / Live LTP (₹)", value=default_ltp, step=0.5
                    )
                    sim_cue = st.selectbox(
                        "Live Nifty / Sector Cue",
                        options=[
                            "Aligned with Setup",
                            "Neutral / Range-Bound",
                            "Strongly Opposing",
                        ],
                        index=0,
                    )

                is_pre930 = sim_time.startswith("09:15")
                is_late = sim_time.startswith("After 14:30")
                gapped_past_t1 = (sim_open >= t1_px) if is_long else (sim_open <= t1_px)
                triggered = (sim_ltp >= entry_px) if is_long else (sim_ltp <= entry_px)
                sl_hit = (sim_ltp <= sl_px) if is_long else (sim_ltp >= sl_px)
                near_pp = (
                    not triggered
                    and not sl_hit
                    and (abs(sim_ltp - pp_px) / pp_px <= 0.0045)
                )

                if gapped_past_t1:
                    st.error(
                        f"🚫 **NO-GO / SKIP TRADE — Opening Gap Past T1 (₹{t1_px:,.2f})**: 9:15 Open (₹{sim_open:,.2f}) gapped beyond T1. Do not chase!"
                    )
                elif sl_hit:
                    st.error(
                        f"🛑 **INVALIDATED / STOP HIT**: Price (₹{sim_ltp:,.2f}) breached Stop Loss (₹{sl_px:,.2f})."
                    )
                elif is_pre930:
                    st.warning(
                        f"⏳ **WAIT FOR 9:30 AM CANDLE CLOSE**: Do not enter during 9:15–9:29 AM opening noise."
                    )
                elif is_late:
                    st.error(
                        "🚫 **NO FRESH ENTRY AFTER 14:30 IST**: Square-off window active (exit by 15:15 IST)."
                    )
                elif sim_cue == "Strongly Opposing":
                    st.warning(
                        "⚠️ **HOLD / CAUTION**: Live Nifty/Sector cue is opposing this direction."
                    )
                elif triggered:
                    st.success(
                        f"✅ **GO / EXECUTE {row_s['Side']}**: 9:30 Trigger confirmed ({'Above' if is_long else 'Below'} ₹{entry_px:,.2f}) | SL ₹{sl_px:,.2f} | T1 ₹{t1_px:,.2f} | T2 ₹{t2_px:,.2f} | Qty: **{int(row_s['Dyn Qty'])} sh**"
                    )
                elif near_pp:
                    st.info(
                        f"🎯 **ALT PULLBACK ENTRY ZONE**: Price testing Pivot PP (₹{pp_px:,.2f}) with SL ₹{sl_px:,.2f}."
                    )
                else:
                    st.info(
                        f"👀 **ARMED & WAITING**: Waiting for 9:30+ candle close {'above' if is_long else 'below'} ₹{entry_px:,.2f}."
                    )

                st.caption(row_s.get("Expert Read", ""))

            with c_right:
                st.subheader("🏆 9-Factor Super-Confluence Leaderboard (50 Setups)")
                f_side = st.radio(
                    "Filter Side", ["ALL", "SELL", "BUY"], horizontal=True
                )
                df_view = (
                    df_super
                    if f_side == "ALL"
                    else df_super[df_super["Side"] == f_side]
                )
                display_cols = [
                    "Symbol",
                    "Industry",
                    "Side",
                    "Expert Score",
                    "Score",
                    "Swing Score",
                    "Entry",
                    "SL",
                    "T1 (1R)",
                    "T2 (2R)",
                    "DEL %",
                    "Dyn Qty",
                    "Dyn Risk (Rs)",
                    "Confluence Pillars",
                ]
                st.dataframe(
                    df_view[[c for c in display_cols if c in df_view.columns]],
                    use_container_width=True,
                    height=580,
                )

    # TAB 2: INTRADAY PLAN
    with tabs[1]:
        df_i = sheet_to_df(hm_data, "Intraday Plan")
        st.dataframe(df_i, use_container_width=True, height=620)

    # TAB 3: SWING PLAN
    with tabs[2]:
        df_s = sheet_to_df(hm_data, "Swing Plan")
        st.dataframe(df_s, use_container_width=True, height=620)

    # TAB 4: SECTORS & INDUSTRY HEAT
    with tabs[3]:
        st.subheader("🧭 NSE Index & Sector View (17)")
        st.dataframe(
            sheet_to_df(hm_data, "Index & Sector"),
            use_container_width=True,
            height=420,
        )
        st.subheader("🔥 Nifty-500 Industry Heat (19 Industries)")
        st.dataframe(
            sheet_to_df(hm_data, "Industry Heat"),
            use_container_width=True,
            height=420,
        )

    # TAB 5: DELIVERY, WEEKLY BO & INSIDE BARS
    with tabs[4]:
        st.subheader("🏦 Institutional Delivery Analysis (496 Stocks)")
        st.dataframe(
            sheet_to_df(hm_data, "Delivery Analysis"),
            use_container_width=True,
            height=380,
        )
        c_w, c_ib = st.columns(2)
        with c_w:
            st.subheader("🚀 5-Day Weekly Breakouts (179)")
            st.dataframe(
                sheet_to_df(hm_data, "Weekly Breakouts"),
                use_container_width=True,
                height=340,
            )
        with c_ib:
            st.subheader("🪆 Inside Bar Setups (102)")
            st.dataframe(
                sheet_to_df(hm_data, "Inside Bar Setups"),
                use_container_width=True,
                height=340,
            )

    # TAB 6: MTF, SIGNALS & BACKTEST
    with tabs[5]:
        st.subheader("🔗 Multi-Timeframe Confluence (418 Symbols)")
        st.dataframe(
            sheet_to_df(hm_data, "MTF Confluence"),
            use_container_width=True,
            height=340,
        )
        st.subheader("📊 Empirical Historical HM Backtest Stats (203 Stocks)")
        st.dataframe(
            sheet_to_df(hm_data, "Backtest Stats"),
            use_container_width=True,
            height=340,
        )
        st.subheader("📡 Active HM Signals — Stocks (473)")
        st.dataframe(
            sheet_to_df(hm_data, "Signals - Stocks"),
            use_container_width=True,
            height=340,
        )

    # TAB 7: GUIDE
    with tabs[6]:
        st.dataframe(
            sheet_to_df(hm_data, "Guide"), use_container_width=True, height=520
        )

# -----------------------------------------------------------------------------
# AUTO-REFRESH ENGINE (Runs at the bottom of the script)
# -----------------------------------------------------------------------------
refresh_map = {"OFF": 0, "15s": 15, "30s": 30, "60s": 60, "5m": 300}
refresh_sec = refresh_map.get(auto_refresh_label, 0)
if refresh_sec > 0 and enable_live_nse:
    time.sleep(refresh_sec)
    st.rerun()
