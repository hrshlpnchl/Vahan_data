# ⚡ India EV & Vehicle Registration Dashboard

Maker-wise EV and vehicle registration analytics across all Indian states and UTs.  
Data scraped daily from [Vahan Portal](https://vahan.parivahan.gov.in/vahan4dashboard/) (MoRTH, Govt of India).

**Live dashboard:** `https://<your-app>.streamlit.app`

---

## How it works

```
4:00 AM IST   GitHub Actions runs scraper.py → downloads 216 xlsx files → commits to data/
5:00 AM IST   GitHub Actions runs compile_parquet.py → master.parquet committed
5:05 AM IST   Streamlit Cloud detects new commit → auto-redeploys app.py
              app.py reads master.parquet (one file, loads in <2 seconds)
```

---

## Repo structure

```
vahan-dashboard/
├── .github/workflows/
│   ├── scrape.yml          ← runs at 4am IST daily
│   └── compile.yml         ← runs at 5am IST daily (also triggers after scrape)
├── data/                   ← xlsx files (committed by scrape workflow)
├── master.parquet          ← pre-compiled data (committed by compile workflow)
├── app.py                  ← Streamlit dashboard (reads parquet only)
├── scraper.py              ← VAHAN scraper v5.5
├── compile_parquet.py      ← xlsx → parquet compiler
└── requirements.txt
```

---

## First-time setup (do this once)

### 1. Create GitHub repo

```bash
git init vahan-dashboard
cd vahan-dashboard
# copy all files here
git add .
git commit -m "initial commit"
git remote add origin https://github.com/YOUR_USERNAME/vahan-dashboard.git
git push -u origin main
```

### 2. Bootstrap data locally (first run only)

Run the scraper on your local machine to get the initial xlsx files:

```bash
pip install playwright pandas openpyxl
playwright install chromium
python scraper.py
```

This creates `vahan_downloads/` with ~216 xlsx files. Then compile:

```bash
pip install pyarrow
python compile_parquet.py --data-dir ./vahan_downloads --out master.parquet
```

Copy the xlsx files to `data/` and commit everything:

```bash
mkdir -p data
cp vahan_downloads/*.xlsx data/
git add data/ master.parquet
git commit -m "chore: initial data load"
git push
```

### 3. Deploy to Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Click **New app**
3. Select your GitHub repo → branch: `main` → Main file: `app.py`
4. Click **Deploy**

Streamlit Cloud reads `requirements.txt` automatically. No other config needed.

### 4. Enable GitHub Actions

GitHub Actions is enabled by default. The workflows will run automatically:
- `scrape.yml` at 22:30 UTC (= 4:00 AM IST)
- `compile.yml` at 23:30 UTC (= 5:00 AM IST), and also when scrape finishes

Check the **Actions** tab on GitHub to monitor runs.

---

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Dashboard shows "master.parquet not found" | Run `python compile_parquet.py` first |
| Scraper fails in GitHub Actions | Check Actions tab → scrape.yml logs; VAHAN may be down |
| No data for a state | Check `data/` for that state's xlsx; scraper may have skipped it |
| Streamlit doesn't refresh after push | Go to Streamlit Cloud → Reboot app |

---

Built by **Harshal Panchal** · Data: Vahan Portal · MoRTH, Govt of India
