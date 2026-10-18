#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
HILEGA-MILEGA v4.0 PRO — FULL AUTOMATION LAUNCHER & LIVE NSE INTRADAY BRIDGE
===============================================================================
Designed for: D:\\Daily Python Stock Analysis\\Hilega_Milega
Zero extra dependencies required beyond your existing v4.0 environment:
  pip install pandas numpy requests openpyxl

WHAT THIS AUTOMATION ENGINE DOES:
  1. EOD AUTO-PIPELINE (--mode eod):
     - Runs your Hilega-Milega v4.0 screener script (if configured) to generate
       today's HilegaMilega_PRO_YYYY-MM-DD_HHMM.xlsx.
     - Automatically locates the newest HilegaMilega_PRO_*.xlsx workbook in
       D:\\Daily Python Stock Analysis\\Hilega_Milega.
     - Extracts all 14 sheets + Dashboard KPIs, Regime Banner, Nifty 50 Map,
       and Expert Narratives into compact JSON (RAW_HM_DATA).
     - Auto-patches HilegaMilega_PRO_Terminal.html in-place so the Web Terminal
       always opens with the latest EOD data pre-loaded (zero manual upload!).

  2. LIVE 9:15 AM NSE INTRADAY GATEKEEPER SERVER (--mode live):
     - Starts a lightweight local HTTP server at http://127.0.0.1:8765
     - Serves the updated HilegaMilega_PRO_Terminal.html and opens your browser.
     - Exposes http://127.0.0.1:8765/live_quotes which fetches real-time NSE
       Open, LTP, High, Low, and PrevClose for NIFTY 50, Sector Indices, and
       all 50 Intraday Plan stocks directly from NSE's live JSON APIs.
     - Feeds the Web Terminal's Fool-Proof 6-Gate Intraday Validator every 15s
       so 9:15 Open, 9:30 Candle Close / LTP, IST Time Window, and Live Nifty
       cues update hands-free!

  3. 1-CLICK WINDOWS TASK SCHEDULER SETUP (--install-schedule):
     - Registers two automated Windows Scheduled Tasks (Mon-Fri):
         a) HM_v4_EOD_Screener_AutoBuild  -> Runs at 19:15 (7:15 PM IST)
         b) HM_v4_Live_Morning_Terminal   -> Runs at 09:14 (9:14 AM IST)
     - Also creates two 1-click desktop/folder batch launchers (.bat).
===============================================================================
"""

from __future__ import annotations

import argparse
import datetime as dt
import glob
import http.server
import json
import os
import re
import socketserver
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
import requests

# =============================================================================
# USER CONFIGURATION (Pre-configured for your folder structure)
# =============================================================================
BASE_DIR = Path(r"D:\Daily Python Stock Analysis\Hilega_Milega")

# If you run this script from another folder (or Linux/Mac), fall back to script's directory
if not BASE_DIR.exists():
    BASE_DIR = Path(__file__).resolve().parent

# Name of your v4.0 EOD Python screener script (if present in BASE_DIR, EOD mode runs it first)
SCREENER_SCRIPT_CANDIDATES = [
    "hilega_milega_v4.py",
    "HilegaMilega_PRO_v4.py",
    "HilegaMilega_Screener.py",
    "screener_v4.py",
]

HTML_FILENAME = "HilegaMilega_PRO_Terminal.html"
EXCEL_GLOB_PATTERN = "HilegaMilega_PRO_*.xlsx"

HOST = "127.0.0.1"
PORT = 8765
LIVE_CACHE_TTL_SECONDS = 12  # Refresh NSE live quotes at most every 12 seconds


# =============================================================================
# STEP 1: EXTRACT ALL 14 SHEETS & DASHBOARD META FROM LATEST EXCEL WORKBOOK
# =============================================================================
def find_latest_excel(folder: Path) -> Optional[Path]:
    files = sorted(
        folder.glob(EXCEL_GLOB_PATTERN),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return files[0] if files else None


def clean_val(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, (int, float, bool)):
        if isinstance(v, float) and (v != v):  # NaN check
            return None
        return round(v, 4) if isinstance(v, float) else v
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def extract_workbook_to_json(xlsx_path: Path) -> Dict[str, Any]:
    print(f"[AUTO-BUILD] Reading workbook: {xlsx_path.name} ...")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
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
            data_rows.append([clean_val(r[i]) if i < len(r) else None for i in range(len(cols))])

        compact[name] = {"c": cols, "r": data_rows}

    # Extract Dashboard metadata (Regime, KPIs, Nifty 50 Map, Narratives)
    if "Dashboard" in wb.sheetnames:
        ws_d = wb["Dashboard"]
        d_rows = list(ws_d.iter_rows(values_only=True))

        def cell_str(r_idx: int, c_idx: int) -> str:
            if r_idx < len(d_rows) and c_idx < len(d_rows[r_idx]):
                v = d_rows[r_idx][c_idx]
                return str(v).strip() if v is not None else ""
            return ""

        subtitle = cell_str(1, 0)
        regime_banner = cell_str(2, 0)

        regime_label = "NEUTRAL"
        regime_score = 0
        m_reg = re.search(r"MARKET REGIME:\s*([A-Z]+)\s*\(score\s*([+-]?\d+)\)", regime_banner)
        if m_reg:
            regime_label = m_reg.group(1)
            regime_score = int(m_reg.group(2))

        stance = (
            regime_banner.split(">>")[-1].strip()
            if ">>" in regime_banner
            else regime_banner
        )

        # Parse Nifty 50 Map (Row 10 headers, Row 11 values, Row 12 read)
        nifty_map: Dict[str, Any] = {}
        if len(d_rows) > 10:
            hdrs = d_rows[9] if len(d_rows) > 9 else []
            vals = d_rows[10] if len(d_rows) > 10 else []
            for i, h in enumerate(hdrs):
                if h is not None and i < len(vals) and vals[i] is not None:
                    key = str(h).replace("NIFTY ", "").replace("CPR ", "").replace(" (14)", "").strip()
                    nifty_map[key] = clean_val(vals[i])
        nifty_map["read"] = cell_str(11, 0)

        # Parse Narratives (Rows 15..22)
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


def patch_html_with_latest_data(html_path: Path, raw_hm_data: Dict[str, Any]) -> bool:
    if not html_path.exists():
        print(f"[WARN] HTML file not found at {html_path}")
        return False

    content = html_path.read_text(encoding="utf-8")
    json_str = json.dumps(raw_hm_data, separators=(",", ":"))

    pattern = re.compile(r"const\s+RAW_HM_DATA\s*=\s*\{.*?\};\s*\n", re.DOTALL)
    replacement = f"const RAW_HM_DATA = {json_str};\n"

    if pattern.search(content):
        new_content = pattern.sub(lambda _: replacement, content, count=1)
        html_path.write_text(new_content, encoding="utf-8")
        (html_path.parent / "hm_data.json").write_text(json_str, encoding="utf-8")
        print(f"[AUTO-BUILD] Injected latest EOD data into {html_path.name} ({len(json_str)//1024} KB)")
        return True
    else:
        print("[WARN] Could not locate 'const RAW_HM_DATA =' marker in HTML file.")
        return False


# =============================================================================
# STEP 2: REAL-TIME NSE QUOTE FETCHER FOR LIVE 9:15 AM – 3:30 PM GATEKEEPER
# =============================================================================
class NSELiveBridge:
    """
    Fetches live NSE market quotes (9:15 Open, LTP, High, Low, PrevClose)
    for NIFTY 50 and all Nifty-500 stocks using NSE's JSON endpoints.
    Uses a warm cookie session and caches responses for 12 seconds.
    """

    def __init__(self) -> None:
        self.session = requests.Session()
        self.session.headers.update(
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
        self._last_cookie_ts = 0.0
        self._cache_ts = 0.0
        self._cached_payload: Dict[str, Any] = {}
        self._lock = threading.Lock()

    def _ensure_cookies(self) -> None:
        now = time.time()
        if now - self._last_cookie_ts > 240:
            try:
                self.session.get("https://www.nseindia.com", timeout=8)
                self._last_cookie_ts = now
            except Exception as exc:
                print(f"[NSE-BRIDGE] Cookie warmup warning: {exc}")

    def get_live_market_snapshot(self) -> Dict[str, Any]:
        with self._lock:
            now = time.time()
            if self._cached_payload and (now - self._cache_ts < LIVE_CACHE_TTL_SECONDS):
                return self._cached_payload

            self._ensure_cookies()
            quotes: Dict[str, Dict[str, float]] = {}
            indices: Dict[str, Dict[str, float]] = {}

            # 1. Fetch NIFTY 500 live equity table (1 request gets Open, High, Low, LTP for all 500 stocks!)
            try:
                url_eq = "https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20500"
                resp = self.session.get(url_eq, timeout=10)
                if resp.status_code == 401 or resp.status_code == 403:
                    self._last_cookie_ts = 0
                    self._ensure_cookies()
                    resp = self.session.get(url_eq, timeout=10)

                if resp.ok:
                    payload = resp.json()
                    for item in payload.get("data", []):
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
            except Exception as exc:
                print(f"[NSE-BRIDGE] Equity fetch warning: {exc}")

            # 2. Fetch All NSE Indices (NIFTY 50, BANK NIFTY, NIFTY IT, etc.)
            try:
                url_idx = "https://www.nseindia.com/api/allIndices"
                resp_i = self.session.get(url_idx, timeout=10)
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
            except Exception as exc:
                print(f"[NSE-BRIDGE] Indices fetch warning: {exc}")

            ist_now = dt.datetime.now()
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

            result = {
                "status": "ok" if quotes else "offline",
                "timestamp": ist_now.strftime("%d-%b-%Y %H:%M:%S IST"),
                "hhmm": hhmm,
                "timeBucket": time_bucket,
                "nifty50": indices.get("NIFTY 50") or quotes.get("NIFTY 50") or {},
                "indices": indices,
                "quotes": quotes,
            }
            self._cached_payload = result
            self._cache_ts = now
            return result


# =============================================================================
# STEP 3: LOCAL HTTP SERVER SERVING TERMINAL + /live_quotes + /rebuild
# =============================================================================
def run_local_server(folder: Path, port: int = PORT, open_browser: bool = True) -> None:
    bridge = NSELiveBridge()

    class TerminalHandler(http.server.SimpleHTTPRequestHandler):
        def __init__(self, *args: Any, **kwargs: Any) -> None:
            super().__init__(*args, directory=str(folder), **kwargs)

        def log_message(self, fmt: str, *args: Any) -> None:
            # Keep console output clean
            pass

        def _send_json(self, data: Dict[str, Any], status: int = 200) -> None:
            raw = json.dumps(data).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)

        def do_GET(self) -> None:
            parsed = urllib.parse.urlparse(self.path)
            if parsed.path in ("/", "/index.html"):
                self.path = f"/{HTML_FILENAME}"
                return super().do_GET()

            if parsed.path == "/live_quotes":
                snap = bridge.get_live_market_snapshot()
                return self._send_json(snap)

            if parsed.path == "/rebuild":
                latest = find_latest_excel(folder)
                if latest:
                    data = extract_workbook_to_json(latest)
                    ok = patch_html_with_latest_data(folder / HTML_FILENAME, data)
                    return self._send_json({"status": "ok" if ok else "error", "excel": latest.name})
                return self._send_json({"status": "error", "message": "No Excel file found"}, 404)

            return super().do_GET()

    socketserver.TCPServer.allow_reuse_address = True
    url = f"http://{HOST}:{port}/{HTML_FILENAME}"
    print("=" * 78)
    print(f"🚀 HILEGA-MILEGA v4.0 PRO — LIVE TERMINAL SERVER RUNNING")
    print(f"   Folder       : {folder}")
    print(f"   Terminal URL : {url}")
    print(f"   Live API     : http://{HOST}:{port}/live_quotes")
    print("   Press Ctrl+C to stop the server.")
    print("=" * 78)

    if open_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    with socketserver.TCPServer((HOST, port), TerminalHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n[SERVER] Shutting down gracefully.")


# =============================================================================
# STEP 4: EOD SCREENER EXECUTION + HTML AUTO-UPDATE PIPELINE
# =============================================================================
def run_eod_pipeline(folder: Path, screener_script: Optional[str] = None, launch_browser: bool = True) -> None:
    print("=" * 78)
    print("⚡ HILEGA-MILEGA v4.0 PRO — EOD FULL AUTOMATION PIPELINE")
    print("=" * 78)

    # 1. Locate and run the Python EOD screener script if available
    script_path: Optional[Path] = None
    if screener_script:
        cand = Path(screener_script)
        script_path = cand if cand.is_absolute() else (folder / cand)
    else:
        for name in SCREENER_SCRIPT_CANDIDATES:
            if (folder / name).exists():
                script_path = folder / name
                break

    if script_path and script_path.exists():
        print(f"[1/3] Running EOD Screener Script: {script_path.name} ...")
        try:
            subprocess.run([sys.executable, str(script_path)], cwd=str(folder), check=True)
        except Exception as exc:
            print(f"[WARN] Screener script execution returned warning: {exc}")
    else:
        print("[1/3] No separate screener script specified; using newest .xlsx in folder.")

    # 2. Find newest HilegaMilega_PRO_*.xlsx and inject into HTML
    latest_xlsx = find_latest_excel(folder)
    if not latest_xlsx:
        print(f"[ERROR] No {EXCEL_GLOB_PATTERN} files found in {folder}!")
        return

    print(f"[2/3] Extracting all 14 sheets from {latest_xlsx.name} ...")
    hm_data = extract_workbook_to_json(latest_xlsx)

    html_path = folder / HTML_FILENAME
    patched = patch_html_with_latest_data(html_path, hm_data)

    # 3. Launch updated HTML terminal
    if patched and launch_browser:
        print(f"[3/3] Opening updated {HTML_FILENAME} in default browser...")
        webbrowser.open(html_path.as_uri())
    print("✔ EOD Auto-Pipeline Complete!")


# =============================================================================
# STEP 5: 1-CLICK WINDOWS TASK SCHEDULER & BATCH LAUNCHER INSTALLER
# =============================================================================
def install_windows_automation(folder: Path, screener_script: Optional[str] = None) -> None:
    launcher_py = Path(__file__).resolve()
    python_exe = sys.executable

    # Create 2 convenient 1-click .bat files in the folder
    bat_eod = folder / "1_Run_EOD_Update_and_Open_Terminal.bat"
    bat_live = folder / "2_Start_Live_915AM_Intraday_Terminal.bat"

    scr_arg = f' --screener "{screener_script}"' if screener_script else ""

    bat_eod.write_text(
        f'@echo off\r\n'
        f'title Hilega-Milega v4.0 PRO - EOD Auto-Pipeline\r\n'
        f'cd /d "{folder}"\r\n'
        f'"{python_exe}" "{launcher_py}" --mode eod{scr_arg}\r\n'
        f'pause\r\n',
        encoding="utf-8",
    )

    bat_live.write_text(
        f'@echo off\r\n'
        f'title Hilega-Milega v4.0 PRO - Live 9:15 AM Intraday Terminal\r\n'
        f'cd /d "{folder}"\r\n'
        f'"{python_exe}" "{launcher_py}" --mode live\r\n'
        f'pause\r\n',
        encoding="utf-8",
    )

    print(f"✔ Created 1-click batch launcher: {bat_eod}")
    print(f"✔ Created 1-click batch launcher: {bat_live}")

    if os.name != "nt":
        print("[INFO] Not running on Windows; skipping schtasks registration.")
        print("       On Linux/macOS, add these two lines to 'crontab -e':")
        print(f'       15 19 * * 1-5 "{python_exe}" "{launcher_py}" --mode eod')
        print(f'       14 09 * * 1-5 "{python_exe}" "{launcher_py}" --mode live')
        return

    # Register Windows Task Scheduler jobs for Mon-Fri
    tasks = [
        (
            "HM_v4_EOD_Screener_AutoBuild",
            "19:15",
            f'"{bat_eod}"',
            "Runs Hilega-Milega EOD screener at 7:15 PM IST (Mon-Fri) and updates the PRO Web Terminal",
        ),
        (
            "HM_v4_Live_Morning_Terminal",
            "09:14",
            f'"{bat_live}"',
            "Launches Hilega-Milega PRO Live Intraday Gatekeeper at 9:14 AM IST (Mon-Fri)",
        ),
    ]

    for task_name, start_time, tr_cmd, desc in tasks:
        cmd = [
            "schtasks",
            "/Create",
            "/F",
            "/SC",
            "WEEKLY",
            "/D",
            "MON,TUE,WED,THU,FRI",
            "/TN",
            task_name,
            "/TR",
            tr_cmd,
            "/ST",
            start_time,
        ]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode == 0:
                print(f"✔ Scheduled Windows Task '{task_name}' at {start_time} (Mon-Fri) — {desc}")
            else:
                print(f"[WARN] Could not register '{task_name}': {res.stderr.strip()}")
        except Exception as exc:
            print(f"[WARN] schtasks error for '{task_name}': {exc}")




def sync_to_github(folder: Path) -> None:
    """Auto-commits and pushes updated hm_data.json & HilegaMilega_PRO_Terminal.html to GitHub so Streamlit Cloud auto-redeploys."""
    try:
        stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M IST")
        subprocess.run(["git", "add", HTML_FILENAME, "hm_data.json"], cwd=str(folder), check=True)
        subprocess.run(["git", "commit", "-m", f"Auto-update EOD Hilega-Milega data ({stamp})"], cwd=str(folder), check=False)
        subprocess.run(["git", "push"], cwd=str(folder), check=True)
        print("✔ Pushed latest EOD data to GitHub -> Streamlit Cloud will auto-refresh!")
    except Exception as exc:
        print(f"[GITHUB-SYNC] Warning: {exc}")

# =============================================================================
# CLI ENTRY POINT
# =============================================================================
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Hilega-Milega v4.0 PRO — Full Automation Launcher & Live NSE Bridge"
    )
    parser.add_argument(
        "--mode",
        choices=["auto", "eod", "live"],
        default="auto",
        help="'eod' = run EOD screener & patch HTML; 'live' = start local live-quote server; 'auto' = patch latest Excel & start live server (default)",
    )
    parser.add_argument(
        "--dir",
        default=str(BASE_DIR),
        help="Path to Hilega_Milega folder containing .xlsx and .html files",
    )
    parser.add_argument(
        "--screener",
        default=None,
        help="Optional filename of your Python v4.0 screener script (e.g. hilega_milega_v4.py)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=PORT,
        help=f"Local HTTP server port (default: {PORT})",
    )
    parser.add_argument(
        "--sync-github",
        action="store_true",
        help="Auto-commit & push updated hm_data.json and HTML to GitHub after EOD build so share.streamlit.io updates",
    )
    parser.add_argument(
        "--install-schedule",
        action="store_true",
        help="Create 1-click .bat launchers and register Mon-Fri Windows Scheduled Tasks (7:15 PM EOD & 9:14 AM Live)",
    )
    args = parser.parse_args()

    target_dir = Path(args.dir).resolve()
    if not target_dir.exists():
        print(f"[ERROR] Folder does not exist: {target_dir}")
        sys.exit(1)

    if args.install_schedule:
        install_windows_automation(target_dir, args.screener)
        return

    if args.mode == "eod":
        run_eod_pipeline(target_dir, args.screener, launch_browser=True)
        if args.sync_github:
            sync_to_github(target_dir)
    elif args.mode == "live":
        # Always auto-patch from newest Excel first so the terminal is fresh
        latest = find_latest_excel(target_dir)
        if latest:
            data = extract_workbook_to_json(latest)
            patch_html_with_latest_data(target_dir / HTML_FILENAME, data)
        run_local_server(target_dir, port=args.port, open_browser=True)
    else:
        # 'auto' mode: if during market hours (09:00 - 15:45), start live server; otherwise run EOD + start server
        latest = find_latest_excel(target_dir)
        if latest:
            data = extract_workbook_to_json(latest)
            patch_html_with_latest_data(target_dir / HTML_FILENAME, data)
        run_local_server(target_dir, port=args.port, open_browser=True)


if __name__ == "__main__":
    main()
