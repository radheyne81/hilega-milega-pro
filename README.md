# 🎯 Hilega-Milega v4.0 PRO — Auto-Refresh Web & Mobile Portal (Streamlit Cloud + GitHub)

A full-stack, auto-refreshing market intelligence & intraday execution portal combining all **14 Hilega-Milega v4.0 sheets**, the **9-Factor Super-Confluence Engine**, the **Fool-Proof 6-Gate 9:30 AM Intraday Validator**, **360° Stock X-Ray**, **Dynamic Capital & Risk Position Sizer**, and **Live NSE Quote Automation**.

---

## 🚀 1-Click Deployment on GitHub & Streamlit Cloud (`https://share.streamlit.io/`)

1. **Upload to GitHub**:
   - Extract `HilegaMilega_PRO_Streamlit_GitHub_App.zip`.
   - Create a new repository on [GitHub](https://github.com/new) (e.g., `hilega-milega-pro`) and upload all files from the extracted folder (`streamlit_app.py`, `HilegaMilega_PRO_Terminal.html`, `hm_data.json`, `HilegaMilega_Auto_Launcher.py`, `requirements.txt`, `.streamlit/`, `.github/`).
2. **Deploy on Streamlit Cloud**:
   - Go to [https://share.streamlit.io/](https://share.streamlit.io/) and click **Create app**.
   - Select your GitHub repository, branch (`main`), and set **Main file path** to `streamlit_app.py`.
   - Click **Deploy!** Your live web & mobile app will be online in ~60 seconds.

---

## 📱 Install as a Mobile App (iOS & Android PWA)

- **On iPhone / iPad (Safari)**: Open your Streamlit Cloud URL (or `HilegaMilega_PRO_Terminal.html`), tap **Share** -> **Add to Home Screen**.
- **On Android (Chrome)**: Open the URL, tap the **⋮ Menu** -> **Add to Home screen / Install app**.
- The portal automatically displays a dedicated **6-tab Mobile Bottom Navigation Bar** (`Gatekeeper`, `Dashboard`, `Intra (50)`, `Swing (60)`, `Sectors`, `Delivery`) optimized for one-handed smartphone use during market hours.

---

## ⚡ Full Automation Workflow (EOD + Live 9:15 AM Market Hours)

- **Live 9:15 AM – 3:30 PM IST Intraday Auto-Refresh**:
  - Both `streamlit_app.py` (in the cloud) and `HilegaMilega_Auto_Launcher.py --mode live` (locally) fetch real-time NSE quotes (`NIFTY 500` equities + `allIndices`) and auto-refresh every **15s / 30s / 60s / 5m**.
- **Automated 7:15 PM IST EOD Sync to GitHub & Streamlit Cloud**:
  - Run locally with `--sync-github`:
    ```bash
    python HilegaMilega_Auto_Launcher.py --mode eod --sync-github
    ```
  - This extracts all 14 sheets from your newest `HilegaMilega_PRO_*.xlsx`, updates both `HilegaMilega_PRO_Terminal.html` and `hm_data.json`, and pushes to GitHub so `share.streamlit.io` updates automatically.
