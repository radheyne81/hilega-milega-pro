#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
HILEGA-MILEGA v4.0 PRO — STREAMLIT CLOUD & MOBILE AUTO-REFRESH WEB PORTAL
===============================================================================
Deployable in 1 click on https://share.streamlit.io/ via GitHub.
Combines:
  1. Server-Side Live NSE Quote Bridge (fetches real-time NIFTY 500 + 17 Indices
     directly from NSE's JSON APIs with automatic cookie session management).
  2. Auto-Refresh Portal Engine (15s / 30s / 60s / 5m hands-free live refresh).
  3. Full Access to Both Views:
     - Embedded Hilega-Milega v4.0 PRO Interactive Web & Mobile Terminal
       (with live NSE quotes injected into the 6-Gate Intraday Validator)
     - Native Streamlit Multi-Dashboard Explorer across all 14 sheets,
       9-Factor Super-Confluence Leaderboard, 6-Gate Validator, 360° Stock X-Ray,
       and Dynamic Capital & Risk Position Sizer.
===============================================================================
"""

from __future__ import annotations

import datetime as dt
import json
import math
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
import streamlit.components.v1 as components

# Import workbook extractor from HilegaMilega_Auto_Launcher
from HilegaMilega_Auto_Launcher import extract_workbook_to_json

# -----------------------------------------------------------------------------
# PAGE CONFIG & MOBILE-FRIENDLY INSTITUTIONAL DARK THEME CSS
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Hilega-Milega v4.0 PRO — Live Market Intelligence & Intraday Terminal",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.1rem !important; padding-bottom: 2rem !important; max-width: 1650px; }
      div[data-testid="stMetric"] {
          background: rgba(15, 23, 42, 0.85);
          border: 1px solid rgba(51, 65, 85, 0.75);
          border-radius: 12px;
          padding: 10px 14px;
      }
      .hm-banner {
          background: linear-gradient(90deg, rgba(136,19,55,0.75) 0%, rgba(120,53,15,0.55) 50%, rgba(15,23,42,0.9) 100%);
          border: 1px solid rgba(244,63,94,0.45);
          border-radius: 14px;
          padding: 14px 18px;
          margin-bottom: 12px;
      }
      .hm-card {
          background: rgba(15, 23, 42, 0.88);
          border: 1px solid rgba(51, 65, 85, 0.8);
          border-radius: 14px;
          padding: 14px 16px;
          margin-bottom: 10px;
      }
      .gate-pass { border-left: 4px solid #10b981; background: rgba(6, 78, 59, 0.25); padding: 8px 12px; border-radius: 8px; margin-bottom: 6px; }
      .gate-fail { border-left: 4px solid #ef4444; background: rgba(127, 29, 29, 0.25); padding: 8px 12px; border-radius: 8px; margin-bottom: 6px; }
      @media (max-width: 768px) {
          .block-container { padding-left: 0.6rem !important; padding-right: 0.6rem !important; }
      }
    </style>
    """,
    unsafe_allow_html=True,
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_JSON_PATH = APP_DIR / "hm_data.json"
DEFAULT_HTML_PATH = APP_DIR / "HilegaMilega_PRO_Terminal.html"


# -----------------------------------------------------------------------------
# STEP 1: LOAD EOD WORKBOOK / JSON DATA (ALL 14 SHEETS + DASHBOARD META)
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_default_hm_data() -> Dict[str, Any]:
    # 1. Check if a newer HilegaMilega_PRO_*.xlsx exists in the repo folder
    xlsx_files = sorted(APP_DIR.glob("HilegaMilega_PRO_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if xlsx_files:
        try:
            return extract_workbook_to_json(xlsx_files[0])
        except Exception:
            pass

    # 2. Load pre-extracted hm_data.json
    if DEFAULT_JSON_PATH.exists():
        return json.loads(DEFAULT_JSON_PATH.read_text(encoding="utf-8"))

    # 3. Fallback: extract from HilegaMilega_PRO_Terminal.html
    if DEFAULT_HTML_PATH.exists():
        content = DEFAULT_HTML_PATH.read_text(encoding="utf-8")
        m = re.search(r"const\s+RAW_HM_DATA\s*=\s*(\{.*?\});\s*\n", content, re.DOTALL)
        if m:
            return json.loads(m.group(1))
    return {}


def sheet_to_df(hm_data: Dict[str, Any], sheet_name: str) -> pd.DataFrame:
    payload = hm_data.get(sheet_name)
    if not payload or "c" not in payload or "r" not in payload:
        return pd.DataFrame()
    return pd.DataFrame(payload["r"], columns=payload["c"])


# -----------------------------------------------------------------------------
# STEP 2: LIVE NSE CLOUD BRIDGE (CACHED 15s FOR STREAMLIT CLOUD)
# -----------------------------------------------------------------------------
@st.cache_data(ttl=15, show_spinner=False)
def fetch_live_nse_snapshot() -> Dict[str, Any]:
    """
    Fetches real-time NSE quotes for NIFTY 500 stocks and all NSE Indices.
    Cached for 15 seconds to prevent rate-limiting while supporting auto-refresh.
    """
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
        resp = session.get("https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20500", timeout=8)
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
                    "totalTradedVolume": float(item.get("totalTradedVolume") or 0),
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
        "status": "ok" if quotes else "eod_fallback",
        "timestamp": ist_now.strftime("%d-%b-%Y %H:%M:%S IST"),
        "hhmm": hhmm,
        "timeBucket": time_bucket,
        "nifty50": indices.get("NIFTY 50") or quotes.get("NIFTY 50") or {},
        "indices": indices,
        "quotes": quotes,
    }


# -----------------------------------------------------------------------------
# STEP 3: 9-FACTOR SUPER-CONFLUENCE & DYNAMIC POSITION SIZER ENGINE
# -----------------------------------------------------------------------------
def calc_position_size(
    entry: float,
    sl: float,
    capital: float,
    risk_pct: float,
    max_pos_pct: float,
    leverage: float = 1.0,
) -> Dict[str, int]:
    risk_per_share = abs(float(entry) - float(sl))
    if entry <= 0 or risk_per_share <= 0:
        return {"qty": 0, "cap": 0, "risk_rs": 0}
    q_risk = math.floor((capital * (risk_pct / 100.0)) / risk_per_share)
    q_cap = math.floor((capital * leverage * (max_pos_pct / 100.0)) / float(entry))
    qty = max(0, min(q_risk, q_cap))
    return {
        "qty": qty,
        "cap": int(round(qty * float(entry))),
        "risk_rs": int(round(qty * risk_per_share)),
    }


def build_super_confluence_df(
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

    swing_map = {str(r["Symbol"]).upper(): r for _, r in df_swing.iterrows()} if not df_swing.empty else {}
    ind_map = {str(r["Industry"]): r for _, r in df_ind.iterrows()} if not df_ind.empty else {}
    del_map = {str(r["Symbol"]).upper(): r for _, r in df_del.iterrows()} if not df_del.empty else {}
    wbo_map = {str(r["Symbol"]).upper(): r for _, r in df_wbo.iterrows()} if not df_wbo.empty else {}
    bt_map = {str(r["Symbol"]).upper(): r for _, r in df_bt.iterrows()} if not df_bt.empty else {}

    ib_daily_syms = set()
    if not df_ib.empty:
        ib_daily_syms = {str(r["Symbol"]).upper() for _, r in df_ib[df_ib["TF"] == "Daily"].iterrows()}

    sig_daily_map = {}
    if not df_sig.empty:
        for _, r in df_sig[df_sig["TF"] == "Daily"].iterrows():
            sig_daily_map[str(r["Symbol"]).upper()] = r

    records = []
    for _, r in df_intra.iterrows():
        sym = str(r["Symbol"]).upper()
        side = str(r["Side"]).upper()
        is_long = side == "BUY"

        sw = swing_map.get(sym, {})
        sw_score = float(sw.get("Score") or 0.0)
        ind_row = ind_map.get(str(r["Industry"]), {})
        iq = str(ind_row.get("Quadrant") or "N/A")

        del_row = del_map.get(sym, {})
        del_pct = float(r.get("DEL %") or del_row.get("DEL %") or 0.0)
        del_trend = float(del_row.get("DEL Trend pp") or 0.0)

        wbo_row = wbo_map.get(sym, {})
        wbo_type = str(wbo_row.get("Breakout Type") or "")
        has_ib = sym in ib_daily_syms

        sig_d = sig_daily_map.get(sym, {})
        hm5_tag = str(sig_d.get("HM 5EMA Tag") or "")
        rsi9 = float(sig_d.get("RSI9") or sw.get("RSI9") or 50.0)
        adx = float(sig_d.get("ADX14") or sw.get("ADX") or 22.0)

        bt_row = bt_map.get(sym, {})
        edge5 = float(bt_row.get("Edge pp") or 0.0)
        brk_pct = float(r.get("Hist PDH/PDL Break %") or 0.0)

        pp = float(r.get("Pivot PP") or 1.0)
        bc = float(r.get("CPR BC") or pp)
        tc = float(r.get("CPR TC") or pp)
        cpr_width_pct = (abs(tc - bc) / pp) * 100.0 if pp else 1.0

        pts = float(r["Score"]) * 0.45
        pillars = []

        is_dual = sw_score >= 55.0 and ((is_long and sw.get("Side") == "BUY") or (not is_long and sw.get("Side") == "SELL"))
        if is_dual:
            pts += sw_score * 0.20
            pillars.append(f"Dual I+S ({sw_score:.0f})")
        else:
            pts += 8.0

        sector_aligned = (is_long and iq in ("LEADING", "IMPROVING")) or (not is_long and iq in ("LAGGING", "WEAKENING"))
        if sector_aligned:
            pts += 8.0
            pillars.append(f"Sector {iq}")
        else:
            pts -= 4.0

        if (is_long and r.get("HM Daily") == "BULL" and r.get("HM Weekly") == "BULL") or (
            not is_long and r.get("HM Daily") == "BEAR" and r.get("HM Weekly") == "BEAR"
        ):
            pts += 6.0
            pillars.append(f"MTF D+W {r.get('HM Daily')}")

        if del_pct >= 60.0 and del_trend >= 5.0:
            pts += 8.0
            pillars.append(f"Inst DEL {del_pct:.0f}% (+{del_trend:.0f}pp)")
        elif del_pct >= 50.0 and del_trend >= 0.0:
            pts += 5.0
            pillars.append(f"DEL {del_pct:.0f}%")
        elif del_pct < 35.0 or del_trend <= -10.0:
            pts -= 5.0

        if cpr_width_pct <= 0.30:
            pts += 4.0
            pillars.append(f"Narrow CPR ({cpr_width_pct:.2f}%)")
        if has_ib or "Compression" in str(r.get("Setup")):
            pts += 3.5
            pillars.append("NR7/Inside Coil")
        if (is_long and wbo_type == "5dayweekBullBO") or (not is_long and wbo_type == "5dayweekBearBO"):
            pts += 3.5
            pillars.append("5D Weekly BO")
        if hm5_tag in ("Hilega5EMABull", "Milega5EMABear"):
            pts += 3.0
            pillars.append("5-EMA Touch")
        if brk_pct >= 55.0:
            pts += 3.0
            pillars.append(f"Hist Brk {brk_pct:.0f}%")
        if edge5 >= 3.0:
            pts += 3.0
            pillars.append(f"5D Edge +{edge5:.1f}pp")
        if (is_long and rsi9 > 80.0) or (not is_long and rsi9 < 20.0):
            pts -= 6.0
            pillars.append(f"⚠ RSI9 {rsi9:.0f}")

        expert_score = round(min(99.0, max(35.0, pts)), 1)
        sz = calc_position_size(float(r["Entry"]), float(r["SL"]), capital, intra_risk_pct, max_pos_pct, intra_lev)

        rec = dict(r)
        rec.update(
            {
                "Expert Score": expert_score,
                "Swing Score": round(sw_score, 1) if sw_score else None,
                "Dual I+S": "YES" if is_dual else "NO",
                "Sector Quadrant": iq,
                "Sector Aligned": sector_aligned,
                "DEL Trend pp": round(del_trend, 2),
                "CPR Width %": round(cpr_width_pct, 2),
                "RSI9": round(rsi9, 1),
                "ADX": round(adx, 1),
                "5D Edge pp": round(edge5, 2),
                "Sized Qty": sz["qty"],
                "Sized Capital (Rs)": sz["cap"],
                "Sized Risk (Rs)": sz["risk_rs"],
                "Confluence Pillars": " · ".join(pillars),
            }
        )
        records.append(rec)

    df_out = pd.DataFrame(records).sort_values("Expert Score", ascending=False).reset_index(drop=True)
    df_out["Expert Rank"] = df_out.index + 1
    return df_out


# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS: MODE, LIVE NSE FEED, AUTO-REFRESH & POSITION SIZER
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🎯 HILEGA-MILEGA v4.0 PRO")
    st.caption("Full Automation Web & Mobile Portal")

    portal_mode = st.radio(
        "🖥️ Select Portal Experience",
        options=[
            "🚀 Combined Auto-Refresh Web/Mobile Terminal + Live NSE Bridge",
            "📊 Native Streamlit Multi-Dashboard Explorer (All 14 Sheets)",
        ],
        index=0,
    )

    st.divider()
    st.markdown("#### 🔄 Live NSE & Auto-Refresh Automation")
    enable_live_nse = st.toggle("📡 Fetch Live NSE Quotes (Nifty 500 + Indices)", value=True)
    auto_refresh_sec = st.selectbox(
        "⏱️ Hands-Free Auto-Refresh Interval",
        options=[0, 15, 30, 60, 300],
        index=2,
        format_func=lambda x: "OFF (Manual)" if x == 0 else f"Every {x} seconds" if x < 60 else f"Every {x // 60} minutes",
    )
    if st.button("⚡ Refresh Live Quotes Now", use_container_width=True):
        fetch_live_nse_snapshot.clear()
        st.rerun()

    st.divider()
    st.markdown("#### ⚙️ Dynamic Capital & Risk Sizer")
    cap_input = st.number_input("Trading Capital (₹)", min_value=50000, max_value=100000000, value=1000000, step=50000)
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        intra_risk = st.number_input("Intra Risk %", min_value=0.1, max_value=5.0, value=0.5, step=0.1)
    with col_r2:
        swing_risk = st.number_input("Swing Risk %", min_value=0.1, max_value=5.0, value=1.0, step=0.1)
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        max_pos_pct = st.number_input("Max Pos %", min_value=5.0, max_value=100.0, value=20.0, step=5.0)
    with col_l2:
        intra_lev = st.number_input("MIS Leverage (x)", min_value=1.0, max_value=5.0, value=1.0, step=0.5)

    st.divider()
    st.markdown("#### 📂 EOD Workbook Upload (.xlsx)")
    uploaded_xlsx = st.file_uploader("Upload HilegaMilega_PRO_*.xlsx", type=["xlsx", "xls"])

# Load EOD dataset (from uploaded .xlsx or bundled hm_data.json)
if uploaded_xlsx is not None:
    tmp_path = APP_DIR / f"_uploaded_{uploaded_xlsx.name}"
    tmp_path.write_bytes(uploaded_xlsx.getvalue())
    hm_data = extract_workbook_to_json(tmp_path)
    st.sidebar.success(f"✔ Loaded {uploaded_xlsx.name} ({len(hm_data) - 1} sheets)")
else:
    hm_data = load_default_hm_data()

# Fetch live NSE market snapshot if enabled
live_snap = fetch_live_nse_snapshot() if enable_live_nse else {"status": "disabled", "quotes": {}, "indices": {}, "nifty50": {}}
quotes_count = len(live_snap.get("quotes") or {})

with st.sidebar:
    if enable_live_nse and quotes_count > 0:
        st.success(f"🟢 Live NSE Connected: {quotes_count} stocks | {live_snap.get('timestamp')}")
    elif enable_live_nse:
        st.info(f"🕒 NSE Market Closed / EOD Mode ({live_snap.get('timestamp')})")

df_super = build_super_confluence_df(hm_data, cap_input, intra_risk, max_pos_pct, intra_lev)
meta = hm_data.get("DashboardMeta", {})

# -----------------------------------------------------------------------------
# MODE 1: EMBEDDED FULL PRO WEB & MOBILE TERMINAL WITH LIVE CLOUD BRIDGE
# -----------------------------------------------------------------------------
if portal_mode.startswith("🚀"):
    # Top status bar in Streamlit
    c1, c2, c3, c4 = st.columns([2.5, 1.5, 1.5, 1.5])
    with c1:
        st.markdown(f"**{meta.get('regimeBanner', 'MARKET REGIME: WEAK (Score -2)')}**")
    with c2:
        nf_ltp = (live_snap.get("nifty50") or {}).get("ltp") or (meta.get("niftyMap") or {}).get("Close", 23146.15)
        st.markdown(f"📈 **NIFTY 50 Spot:** `₹{float(nf_ltp):,.2f}`")
    with c3:
        st.markdown(f"📡 **Live Quotes:** `{quotes_count} NSE Stocks`")
    with c4:
        st.markdown(f"💰 **Sizer:** `₹{cap_input/100000:.1f}L ({intra_risk}%I / {swing_risk}%S)`")

    if DEFAULT_HTML_PATH.exists():
        html_str = DEFAULT_HTML_PATH.read_text(encoding="utf-8")
        # Inject latest hm_data + live NSE snapshot + user's sidebar capital settings directly into the HTML iframe!
        hm_json_str = json.dumps(hm_data, separators=(",", ":"))
        html_str = re.sub(
            r"const\s+RAW_HM_DATA\s*=\s*\{.*?\};\s*\n",
            lambda _: f"const RAW_HM_DATA = {hm_json_str};\n",
            html_str,
            count=1,
            flags=re.DOTALL,
        )
        live_inject_js = f"""
        <script>
          window.__STREAMLIT_LIVE_SNAPSHOT__ = {json.dumps(live_snap, separators=(",", ":"))};
          window.addEventListener('DOMContentLoaded', () => {{
            if (typeof STATE !== 'undefined') {{
              STATE.capital = {float(cap_input)};
              STATE.intraRiskPct = {float(intra_risk)};
              STATE.swingRiskPct = {float(swing_risk)};
              STATE.maxPosPct = {float(max_pos_pct)};
              STATE.intraLeverage = {float(intra_lev)};
              if (typeof renderAllViews === 'function') renderAllViews();
            }}
            if (typeof applyLiveSnapshotToUI === 'function' && window.__STREAMLIT_LIVE_SNAPSHOT__) {{
              applyLiveSnapshotToUI(window.__STREAMLIT_LIVE_SNAPSHOT__, "Streamlit Cloud NSE");
            }}
          }});
        </script>
        """
        html_str = html_str.replace("</head>", live_inject_js + "
</head>")
        components.html(html_str, height=1180, scrolling=True)
    else:
        st.error("HilegaMilega_PRO_Terminal.html not found in repository.")

# -----------------------------------------------------------------------------
# MODE 2: NATIVE STREAMLIT MULTI-DASHBOARD EXPLORER (ALL 14 SHEETS + X-RAY)
# -----------------------------------------------------------------------------
else:
    st.markdown(
        f"""
        <div class="hm-banner">
          <div style="font-size:11px;font-weight:700;color:#fda4af;text-transform:uppercase;letter-spacing:0.08em;">
            {meta.get('subtitle', '')}
          </div>
          <div style="font-size:20px;font-weight:800;color:#ffffff;margin-top:2px;">
            {meta.get('regimeBanner', 'MARKET REGIME: WEAK (Score -2)')}
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    tabs = st.tabs(
        [
            "🎯 Super-Confluence & 6-Gate Validator",
            "📊 Market Dashboard & Nifty Map",
            "⚡ Intraday Plan (50)",
            "📅 Swing Plan 3-10D (60)",
            "🧭 Index, Sectors & Industry Heat",
            "🏦 Delivery (496), Weekly BO & Inside Bars",
            "🔬 MTF (418), Signals & Backtest",
            "🔍 360° Stock X-Ray",
            "📖 Guide & Formulas",
        ]
    )

    # TAB 0: SUPER-CONFLUENCE & 6-GATE VALIDATOR
    with tabs[0]:
        col_gate, col_board = st.columns([5, 7], gap="medium")
        with col_gate:
            st.markdown("#### 🛡️ Fool-Proof 6-Gate Intraday Validator")
            sym_options = df_super["Symbol"].tolist()
            sel_sym = st.selectbox(
                "Select Intraday Setup Symbol",
                options=sym_options,
                format_func=lambda s: f"#{int(df_super.loc[df_super['Symbol']==s, 'Expert Rank'].values[0])} {s} ({df_super.loc[df_super['Symbol']==s, 'Side'].values[0]} | Expert {df_super.loc[df_super['Symbol']==s, 'Expert Score'].values[0]})",
            )
            row = df_super[df_super["Symbol"] == sel_sym].iloc[0]
            is_long = row["Side"] == "BUY"
            live_q = (live_snap.get("quotes") or {}).get(sel_sym, {})

            default_open = float(live_q.get("open") or ((float(row["Prev Close"]) + float(row["Entry"])) / 2.0))
            default_ltp = float(live_q.get("ltp") or (float(row["Entry"]) * (1.0015 if is_long else 0.9985)))

            g_c1, g_c2 = st.columns(2)
            with g_c1:
                sim_open = st.number_input("9:15 Open Price (₹)", value=round(default_open, 2), step=0.5)
                sim_time = st.selectbox(
                    "Session Time Window (IST)",
                    options=["09:20", "09:35", "12:15", "13:45", "14:45"],
                    index=1,
                    format_func=lambda v: {
                        "09:20": "09:15–09:29 (Wait 1st Bar)",
                        "09:35": "09:30–11:30 (Prime Window)",
                        "12:15": "11:30–13:30 (Midday Lull)",
                        "13:45": "13:30–14:30 (Afternoon)",
                        "14:45": "After 14:30 (No Fresh Entry)",
                    }[v],
                )
            with g_c2:
                sim_ltp = st.number_input("9:30 Close / Live LTP (₹)", value=round(default_ltp, 2), step=0.5)
                sim_cue = st.selectbox("Live Nifty / Sector Cue", options=["aligned", "flat", "opposed"], index=0)

            entry_px = float(row["Entry"])
            sl_px = float(row["SL"])
            t1_px = float(row["T1 (1R)"])
            t2_px = float(row["T2 (2R)"])
            t3_px = float(row["T3 (3R)"])
            pp_px = float(row["Pivot PP"])

            is_pre930 = sim_time == "09:20"
            is_late = sim_time == "14:45"
            gapped_past_t1 = (sim_open >= t1_px) if is_long else (sim_open <= t1_px)
            triggered = (sim_ltp >= entry_px) if is_long else (sim_ltp <= entry_px)
            sl_hit = (sim_ltp <= sl_px) if is_long else (sim_ltp >= sl_px)
            near_pp = (not triggered) and (not sl_hit) and (abs(sim_ltp - pp_px) / pp_px <= 0.0045)

            if gapped_past_t1:
                st.error(f"🚫 **NO-GO / SKIP TRADE — Opening Gap Past T1 (₹{t1_px:,.2f})**: 9:15 Open ₹{sim_open:,.2f} gapped beyond T1.")
            elif sl_hit:
                st.error(f"🛑 **INVALIDATED / STOP HIT**: LTP ₹{sim_ltp:,.2f} breached SL ₹{sl_px:,.2f}.")
            elif is_pre930:
                st.warning(f"⏳ **WAIT FOR 9:30 AM CANDLE CLOSE**: Do not enter during 9:15–9:30 AM opening noise.")
            elif is_late:
                st.error("🚫 **NO FRESH ENTRY AFTER 14:30 IST**: Square-off window active.")
            elif triggered and sim_cue != "opposed":
                st.success(f"✅ **GO / EXECUTE {row['Side']}**: Trigger ₹{entry_px:,.2f} confirmed! SL ₹{sl_px:,.2f} | T1 ₹{t1_px:,.2f} | T2 ₹{t2_px:,.2f}")
            elif near_pp:
                st.info(f"🎯 **ALT PULLBACK ENTRY ZONE**: Testing Central Pivot PP ₹{pp_px:,.2f} with SL ₹{sl_px:,.2f}.")
            else:
                st.warning(f"👀 **ARMED & WAITING**: Waiting for 9:30+ candle close {'above' if is_long else 'below'} ₹{entry_px:,.2f}.")

            gates = [
                ("Gate 1: 9:30 AM Candle & Time Window", not is_pre930 and not is_late, f"Window: {sim_time} IST"),
                ("Gate 2: Opening Gap-Past-T1 Guard", not gapped_past_t1, f"Open ₹{sim_open:,.2f} vs T1 ₹{t1_px:,.2f}"),
                ("Gate 3: Trigger / Pivot Pullback", triggered or near_pp, f"LTP ₹{sim_ltp:,.2f} vs Trigger ₹{entry_px:,.2f}"),
                ("Gate 4: Regime & Industry Quadrant", (not is_long) or bool(row["Sector Aligned"]), f"{row['Industry']} ({row['Sector Quadrant']})"),
                ("Gate 5: Institutional Delivery Footprint", float(row["DEL %"]) >= 45.0, f"DEL {row['DEL %']}% ({row['DEL Trend pp']:+.1f}pp vs 7D)"),
                ("Gate 6: RSI9 Overextension & Live Cue", sim_cue != "opposed" and not ((is_long and row["RSI9"] > 80) or (not is_long and row["RSI9"] < 20)), f"RSI9 {row['RSI9']} | Break {row['Hist PDH/PDL Break %']}%"),
            ]
            for g_title, g_ok, g_det in gates:
                cls = "gate-pass" if g_ok else "gate-fail"
                icon = "✔" if g_ok else "✖"
                st.markdown(f"<div class='{cls}'><strong>{icon} {g_title}</strong> — <span style='font-size:12px;opacity:0.9;'>{g_det}</span></div>", unsafe_allow_html=True)

            st.caption(f"**Position Sizing:** `{row['Sized Qty']:,} shares` | Capital `₹{row['Sized Capital (Rs)']:,}` | 1R Risk `-₹{row['Sized Risk (Rs)']:,}`")
            st.info(row["Expert Read"])

        with col_board:
            st.markdown("#### 🏆 9-Factor Expert Super-Confluence Leaderboard (All 50 Setups)")
            f1, f2, f3 = st.columns(3)
            with f1:
                side_f = st.selectbox("Filter Side", ["ALL", "SELL", "BUY"], index=0)
            with f2:
                tier_f = st.selectbox("Confluence Filter", ["ALL", "Dual I+S Only", "Sector Aligned Only", "High DEL >= 60%"], index=0)
            with f3:
                q_sup = st.text_input("Search Symbol/Industry", "")

            df_view = df_super.copy()
            if side_f != "ALL":
                df_view = df_view[df_view["Side"] == side_f]
            if tier_f == "Dual I+S Only":
                df_view = df_view[df_view["Dual I+S"] == "YES"]
            elif tier_f == "Sector Aligned Only":
                df_view = df_view[df_view["Sector Aligned"] == True]
            elif tier_f == "High DEL >= 60%":
                df_view = df_view[df_view["DEL %"] >= 60.0]
            if q_sup:
                df_view = df_view[
                    df_view["Symbol"].str.contains(q_sup, case=False, na=False)
                    | df_view["Industry"].str.contains(q_sup, case=False, na=False)
                ]

            show_cols = [
                "Expert Rank", "Symbol", "Side", "Expert Score", "Score", "Swing Score",
                "Entry", "SL", "T1 (1R)", "T2 (2R)", "Sized Qty", "Sized Risk (Rs)",
                "DEL %", "CPR Width %", "Hist PDH/PDL Break %", "Confluence Pillars", "Industry"
            ]
            st.dataframe(df_view[show_cols], use_container_width=True, height=620, hide_index=True)

    # TAB 1: MARKET DASHBOARD & NIFTY MAP
    with tabs[1]:
        kpis = meta.get("kpis", [])
        if kpis:
            k_cols = st.columns(len(kpis))
            for idx, k in enumerate(kpis):
                with k_cols[idx]:
                    st.metric(label=k.get("label", ""), value=k.get("value", ""), delta=k.get("sub", ""))

        st.markdown("#### 🗺️ NIFTY 50 Key Levels Map (Next Session)")
        nf_map = meta.get("niftyMap", {})
        nf_keys = ["S2", "S1", "PDL", "BC", "PP", "TC", "Close", "PDH", "R1", "R2", "ATR"]
        nf_cols = st.columns(len(nf_keys))
        for idx, k in enumerate(nf_keys):
            with nf_cols[idx]:
                st.metric(label=k, value=f"{float(nf_map.get(k, 0)):,.1f}")
        st.info(nf_map.get("read", ""))

        ch_c1, ch_c2 = st.columns([7, 5])
        df_idx_sec = sheet_to_df(hm_data, "Index & Sector")
        with ch_c1:
            if not df_idx_sec.empty:
                df_sec_only = df_idx_sec[df_idx_sec["Type"] == "Sector"].sort_values("Ret 20D %", ascending=False)
                fig_sec = px.bar(
                    df_sec_only,
                    x="Name",
                    y=["Ret 20D %", "Ret 5D %"],
                    barmode="group",
                    title="Sector Rotation (20D vs 5D Return %)",
                    template="plotly_dark",
                )
                fig_sec.update_layout(height=360, margin=dict(l=10, r=10, t=40, b=10))
                st.plotly_chart(fig_sec, use_container_width=True)
        with ch_c2:
            df_ind_heat = sheet_to_df(hm_data, "Industry Heat")
            if not df_ind_heat.empty:
                fig_ind = px.bar(
                    df_ind_heat.head(12),
                    x="Avg RS Rank",
                    y="Industry",
                    orientation="h",
                    color="Quadrant",
                    title="Top 12 Industries by Relative Strength (RS Rank)",
                    template="plotly_dark",
                )
                fig_ind.update_layout(height=360, margin=dict(l=10, r=10, t=40, b=10), yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_ind, use_container_width=True)

        st.markdown("#### 🧠 Automated Expert Market Narratives (8 Pillars)")
        narrs = meta.get("narratives", [])
        n_cols = st.columns(2)
        for i, n in enumerate(narrs):
            with n_cols[i % 2]:
                st.markdown(f"<div class='hm-card'><strong style='color:#60a5fa;'>{n.get('title')}</strong><br><span style='font-size:13px;'>{n.get('text')}</span></div>", unsafe_allow_html=True)

    # TAB 2: INTRADAY PLAN (50)
    with tabs[2]:
        df_intra = sheet_to_df(hm_data, "Intraday Plan")
        st.markdown("#### ⚡ Intraday Plan — All 50 High-Liquidity Setups")
        st.dataframe(df_intra, use_container_width=True, height=620, hide_index=True)

    # TAB 3: SWING PLAN (60)
    with tabs[3]:
        df_swing = sheet_to_df(hm_data, "Swing Plan")
        st.markdown("#### 📅 Swing Plan (3–10 Trading Days) — All 60 Setups")
        st.dataframe(df_swing, use_container_width=True, height=620, hide_index=True)

    # TAB 4: INDEX, SECTORS & INDUSTRY HEAT
    with tabs[4]:
        st.markdown("#### 🧭 NSE Index & Sector View (17 Indices & Sectors)")
        st.dataframe(sheet_to_df(hm_data, "Index & Sector"), use_container_width=True, height=420, hide_index=True)
        st.markdown("#### 🔥 Industry Heat (19 Nifty-500 Industries)")
        st.dataframe(sheet_to_df(hm_data, "Industry Heat"), use_container_width=True, height=420, hide_index=True)

    # TAB 5: INSTITUTIONAL FOOTPRINTS
    with tabs[5]:
        fp_choice = st.radio(
            "Select Footprint Sheet",
            ["🏦 Delivery Analysis (496)", "🚀 Weekly Breakouts (179)", "🪆 Inside Bar Setups (102)"],
            horizontal=True,
        )
        if fp_choice.startswith("🏦"):
            st.dataframe(sheet_to_df(hm_data, "Delivery Analysis"), use_container_width=True, height=600, hide_index=True)
        elif fp_choice.startswith("🚀"):
            st.dataframe(sheet_to_df(hm_data, "Weekly Breakouts"), use_container_width=True, height=600, hide_index=True)
        else:
            st.dataframe(sheet_to_df(hm_data, "Inside Bar Setups"), use_container_width=True, height=600, hide_index=True)

    # TAB 6: MTF CONFLUENCE, SIGNALS & BACKTEST
    with tabs[6]:
        mtf_choice = st.radio(
            "Select Signal / Backtest Sheet",
            [
                "🔗 MTF Confluence (418)",
                "📡 Signals - Stocks (473)",
                "📈 Signals - Index (5)",
                "🧭 Signals - Sector (13)",
                "📊 Backtest Stats (203)",
            ],
            horizontal=True,
        )
        sheet_key = mtf_choice.split(" (")[0].split(" ", 1)[1]
        st.dataframe(sheet_to_df(hm_data, sheet_key), use_container_width=True, height=600, hide_index=True)

    # TAB 7: 360° STOCK X-RAY
    with tabs[7]:
        st.markdown("#### 🔍 360° Cross-Sheet Stock X-Ray")
        df_del_all = sheet_to_df(hm_data, "Delivery Analysis")
        all_syms = sorted(df_del_all["Symbol"].dropna().unique().tolist()) if not df_del_all.empty else []
        x_sym = st.selectbox("Choose Any Nifty-500 Symbol", options=all_syms, index=all_syms.index("CUMMINSIND") if "CUMMINSIND" in all_syms else 0)
        for s_name in [
            "Intraday Plan", "Swing Plan", "MTF Confluence", "Delivery Analysis",
            "Weekly Breakouts", "Inside Bar Setups", "Backtest Stats", "Signals - Stocks"
        ]:
            df_s = sheet_to_df(hm_data, s_name)
            if not df_s.empty and "Symbol" in df_s.columns:
                m_rows = df_s[df_s["Symbol"] == x_sym]
                if not m_rows.empty:
                    st.markdown(f"**{s_name}**")
                    st.dataframe(m_rows, use_container_width=True, hide_index=True)

    # TAB 8: GUIDE
    with tabs[8]:
        st.markdown("#### 📖 Hilega-Milega v4.0 All-In-One Guide & Formula Reference")
        st.dataframe(sheet_to_df(hm_data, "Guide"), use_container_width=True, height=600, hide_index=True)

# -----------------------------------------------------------------------------
# AUTO-REFRESH TIMER TRIGGER
# -----------------------------------------------------------------------------
if auto_refresh_sec and auto_refresh_sec > 0:
    time.sleep(auto_refresh_sec)
    st.rerun()
