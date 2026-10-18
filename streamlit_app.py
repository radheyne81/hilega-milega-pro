#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HILEGA-MILEGA v4.0 PRO — CLEAN FULL-SCREEN WEB & MOBILE APP (STREAMLIT CLOUD)
=============================================================================
- Zero top-of-page input clutter: all configuration, capital sizing, EOD .xlsx
  upload, and Live NSE quote controls live inside the collapsible sidebar (`st.sidebar`).
- Full-bleed responsive Web App layout with Action Cards, 6-Gate Validator,
  360° Stock X-Ray, and sticky mobile bottom navigation.
"""

from __future__ import annotations

import datetime as dt
import io
import json
import math
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple

import openpyxl
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Hilega-Milega v4.0 PRO",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_DIR = Path(__file__).resolve().parent
HTML_PATH = APP_DIR / "HilegaMilega_PRO_Terminal.html"
JSON_PATH = APP_DIR / "hm_data.json"

# Strip default Streamlit padding so the Web App looks & feels like a standalone website on desktop & mobile
st.markdown(
    """
    <style>
      .stApp { background-color: #0b1120; color: #f1f5f9; }
      header[data-testid="stHeader"] { background: transparent !important; height: 2rem !important; }
      .block-container {
          padding-top: 0rem !important;
          padding-bottom: 0rem !important;
          padding-left: 0rem !important;
          padding-right: 0rem !important;
          max-width: 100% !important;
      }
      iframe { border: none !important; width: 100% !important; }
    </style>
    """,
    unsafe_allow_html=True,
)


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
                {"label": cell_str(4, 0) or "NIFTY 50", "value": cell_str(5, 0), "sub": cell_str(6, 0), "tone": "emerald"},
                {"label": cell_str(4, 2) or "BREADTH > 50-DMA", "value": cell_str(5, 2), "sub": cell_str(6, 2), "tone": "amber"},
                {"label": cell_str(4, 4) or "ADVANCE / DECLINE", "value": cell_str(5, 4), "sub": cell_str(6, 4), "tone": "emerald"},
                {"label": cell_str(4, 6) or "HM DAILY BULL vs BEAR", "value": cell_str(5, 6), "sub": cell_str(6, 6), "tone": "rose"},
                {"label": cell_str(4, 8) or "20D HIGHS / LOWS", "value": cell_str(5, 8), "sub": cell_str(6, 8), "tone": "sky"},
                {"label": cell_str(4, 10) or "AVG DELIVERY", "value": cell_str(5, 10), "sub": cell_str(6, 10), "tone": "violet"},
            ],
            "niftyMap": nifty_map,
            "narratives": narratives,
        }

    return compact


@st.cache_data(show_spinner=False)
def load_default_hm_data() -> Dict[str, Any]:
    if JSON_PATH.exists():
        try:
            return json.loads(JSON_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass

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
# COLLAPSIBLE SIDEBAR (All Controls & Uploaders Tucked Neatly Out of the Way)
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚙️ Terminal Controls")
    st.caption("All settings sync directly with the Web Terminal.")

    uploaded_xlsx = st.file_uploader(
        "📂 Load EOD .xlsx Workbook",
        type=["xlsx", "xls"],
        help="Upload a new HilegaMilega_PRO_YYYY-MM-DD.xlsx file to update all 14 sheets.",
    )

    st.divider()
    st.markdown("#### 📐 Position Sizing")
    capital_rs = st.number_input(
        "Trading Capital (₹)", min_value=50000, max_value=100000000, value=1000000, step=50000
    )
    c1, c2 = st.columns(2)
    with c1:
        intra_risk_pct = st.number_input("Intra Risk %", min_value=0.1, max_value=5.0, value=0.5, step=0.1)
    with c2:
        swing_risk_pct = st.number_input("Swing Risk %", min_value=0.1, max_value=5.0, value=1.0, step=0.1)

    c3, c4 = st.columns(2)
    with c3:
        max_pos_pct = st.number_input("Max Pos %", min_value=5.0, max_value=100.0, value=20.0, step=5.0)
    with c4:
        intra_leverage = st.number_input("MIS Lev (x)", min_value=1.0, max_value=5.0, value=1.0, step=0.5)

    st.divider()
    st.markdown("#### 📡 Live NSE Feed")
    enable_live_nse = st.toggle("Fetch Live NSE Quotes", value=False)
    auto_refresh_label = st.selectbox(
        "Auto-Refresh Interval", options=["OFF", "15s", "30s", "60s", "5m"], index=0
    )

if uploaded_xlsx is not None:
    try:
        hm_data = extract_workbook_bytes_to_json(uploaded_xlsx.getvalue())
        st.sidebar.success(f"Loaded {uploaded_xlsx.name}!")
    except Exception as exc:
        st.sidebar.error(f"Error parsing workbook: {exc}")
        hm_data = load_default_hm_data()
else:
    hm_data = load_default_hm_data()

live_snapshot: Dict[str, Any] = {}
if enable_live_nse:
    live_snapshot = fetch_live_nse_snapshot()

# -----------------------------------------------------------------------------
# FULL-SCREEN RESPONSIVE WEB APPLICATION RENDERER
# -----------------------------------------------------------------------------
if HTML_PATH.exists():
    html_str = HTML_PATH.read_text(encoding="utf-8")

    if hm_data:
        json_str = json.dumps(hm_data, separators=(",", ":"))
        pattern = re.compile(r"const\s+RAW_HM_DATA\s*=\s*\{.*?\};\s*\n", re.DOTALL)
        replacement = f"const RAW_HM_DATA = {json_str};\n"
        if pattern.search(html_str):
            html_str = pattern.sub(lambda _: replacement, html_str, count=1)

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

    # 920px viewport with internal flex scrolling ensures both desktop and mobile
    # navigation bars remain pinned inside the visible screen with zero double scrollbars
    components.html(html_str, height=920, scrolling=False)
else:
    st.error("HilegaMilega_PRO_Terminal.html not found in application folder.")

refresh_map = {"OFF": 0, "15s": 15, "30s": 30, "60s": 60, "5m": 300}
refresh_sec = refresh_map.get(auto_refresh_label, 0)
if refresh_sec > 0 and enable_live_nse:
    time.sleep(refresh_sec)
    st.rerun()
