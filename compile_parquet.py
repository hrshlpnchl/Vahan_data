#!/usr/bin/env python3
"""
compile_parquet.py - v2.0
Compiles cleaned Vahan xlsx files into one master.parquet for the Streamlit app.

Filename convention produced by scraper v5.5.1:
    <State>_<Combo>_<Year>_<YYYYMMDD>.xlsx
    e.g.  Gujarat_2W_PureEV_2026_20260618.xlsx

For each unique (state, combo, year) tuple, this script picks the file with
the LARGEST date suffix and ignores older duplicates. Files without a date
suffix (legacy v5.5 format) are treated as oldest.

Output: a long-format parquet with columns
    state, vehicle_category, fuel_type, year, month, maker, registrations,
    source_file, source_date
"""

import argparse, os, re, sys, glob, logging
from datetime import datetime
import pandas as pd
import openpyxl

# ---------------- CLI -----------------------------------------------
parser = argparse.ArgumentParser(description="Compile Vahan xlsx -> master.parquet")
parser.add_argument("--data-dir", default="./data", help="Folder with cleaned xlsx files")
parser.add_argument("--out",      default="master.parquet", help="Output parquet path")
args = parser.parse_args()

# ---------------- Logging -------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("compile_parquet")

# ---------------- Constants -----------------------------------------
MONTH_FULL = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
MONTH_TO_NUM = {m: i+1 for i, m in enumerate(MONTH_FULL)}

# Regex parses scraper output filenames.
# Captures:
#   state  : everything before _<combo>_
#   combo  : e.g. 2W_PureEV / 3W_AllFuel / 4W_PureEV ...
#   year   : 4-digit calendar year
#   date   : optional YYYYMMDD scrape-run date (v5.5.1+)
FNAME_RE = re.compile(
    r"^(?P<state>.+?)_(?P<combo>(?:2W|3W|4W)_(?:PureEV|AllFuel))_"
    r"(?P<year>\d{4})(?:_(?P<date>\d{8}))?\.xlsx$"
)

# ---------------- Filename parser -----------------------------------
def parse_filename(fname):
    m = FNAME_RE.match(fname)
    if not m:
        return None
    d = m.groupdict()
    cat, fuel = d["combo"].split("_")          # '2W', 'PureEV'
    return {
        "state":             d["state"].replace("_", " ").replace("and", "&"),
        "state_safe":        d["state"],
        "vehicle_category":  cat,
        "fuel_type":         fuel,
        "combo":             d["combo"],
        "year":              int(d["year"]),
        "date_suffix":       d["date"] or "00000000",  # missing date => oldest
        "filename":          fname,
    }

# ---------------- Pick latest per group ------------------------------
def pick_latest_files(data_dir):
    all_files = sorted(glob.glob(os.path.join(data_dir, "*.xlsx")))
    log.info(f"Scanning {len(all_files)} xlsx files in {data_dir}")

    parsed   = []
    skipped  = []
    for fp in all_files:
        info = parse_filename(os.path.basename(fp))
        if info is None:
            skipped.append(fp)
            continue
        info["filepath"] = fp
        parsed.append(info)

    if skipped:
        log.warning(f"Skipped {len(skipped)} files that didn't match naming pattern:")
        for s in skipped[:5]:
            log.warning(f"   - {os.path.basename(s)}")
        if len(skipped) > 5:
            log.warning(f"   ... and {len(skipped)-5} more")

    # Group by (state, combo, year) and keep max date_suffix
    df = pd.DataFrame(parsed)
    if df.empty:
        log.error("No parseable files found!")
        return [], 0

    df_sorted = df.sort_values("date_suffix", ascending=False)
    df_latest = df_sorted.drop_duplicates(subset=["state_safe", "combo", "year"], keep="first")

    dropped_count = len(df) - len(df_latest)
    log.info(f"Picked {len(df_latest)} latest files; ignored {dropped_count} older duplicates")

    # Log a few examples of what we're using
    log.info("Sample of selected latest files:")
    for _, row in df_latest.head(5).iterrows():
        log.info(f"   [{row['date_suffix']}] {row['filename']}")

    return df_latest.to_dict("records"), dropped_count

# ---------------- Read one xlsx into long-format dataframe -----------
def read_one(file_info):
    """
    Cleaned files (produced by scraper.clean_excel) have:
        Row 1 : title
        Row 2 : headers  -> S. No. | Maker | Jan YYYY | Feb YYYY | ... | Total
        Row 3+: data
    """
    fp = file_info["filepath"]
    try:
        # Try header=1 first (cleaned format)
        df = pd.read_excel(fp, header=1, engine="openpyxl")
    except Exception as e:
        log.warning(f"   read failed for {file_info['filename']}: {e}")
        return None

    if df.empty:
        return None

    # Normalise column names
    df.columns = [str(c).strip() for c in df.columns]

    # Identify maker column
    maker_col = None
    for c in df.columns:
        if c.lower() == "maker":
            maker_col = c
            break
    if maker_col is None:
        # Fallback: second column
        maker_col = df.columns[1] if len(df.columns) > 1 else None
    if maker_col is None:
        log.warning(f"   no 'Maker' column in {file_info['filename']}")
        return None

    # Identify month columns ("Jan 2026", "Feb 2026", ...)
    month_cols = []
    for c in df.columns:
        for mname in MONTH_FULL:
            if c.startswith(mname + " "):
                month_cols.append((c, mname))
                break

    if not month_cols:
        log.warning(f"   no month columns in {file_info['filename']}")
        return None

    # Melt to long
    keep = [maker_col] + [c for c, _ in month_cols]
    long_df = df[keep].melt(
        id_vars=[maker_col],
        var_name="month_col",
        value_name="registrations",
    )
    long_df = long_df.rename(columns={maker_col: "maker"})
    long_df["maker"] = long_df["maker"].astype(str).str.strip()
    long_df = long_df[~long_df["maker"].str.upper().isin(["NAN", "MAKER", "TOTAL", ""])]
    long_df = long_df[long_df["maker"].str.len() > 0]

    # Extract month name and numeric value
    long_df["month_name"] = long_df["month_col"].str.split(" ").str[0]
    long_df["month"]      = long_df["month_name"].map(MONTH_TO_NUM)
    long_df["registrations"] = (
        long_df["registrations"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    long_df["registrations"] = pd.to_numeric(long_df["registrations"], errors="coerce").fillna(0).astype("int64")

    # Add metadata columns
    long_df["state"]            = file_info["state"]
    long_df["vehicle_category"] = file_info["vehicle_category"]
    long_df["fuel_type"]        = file_info["fuel_type"]
    long_df["year"]             = file_info["year"]
    long_df["source_file"]      = file_info["filename"]
    long_df["source_date"]      = file_info["date_suffix"]

    return long_df[[
        "state", "vehicle_category", "fuel_type", "year", "month",
        "maker", "registrations", "source_file", "source_date"
    ]]

# ---------------- Main ----------------------------------------------
def main():
    if not os.path.isdir(args.data_dir):
        log.error(f"Data dir not found: {args.data_dir}")
        sys.exit(1)

    latest_files, dropped = pick_latest_files(args.data_dir)
    if not latest_files:
        log.error("Nothing to compile — exiting")
        sys.exit(1)

    frames = []
    fail_count = 0
    for fi in latest_files:
        df = read_one(fi)
        if df is None or df.empty:
            fail_count += 1
            continue
        frames.append(df)

    if not frames:
        log.error("All file reads returned empty — nothing written")
        sys.exit(1)

    master = pd.concat(frames, ignore_index=True)
    master.to_parquet(args.out, index=False, compression="snappy")

    # ---------------- Summary ---------------------------------------
    size_mb = os.path.getsize(args.out) / (1024 * 1024)
    log.info("=" * 60)
    log.info("COMPILE SUMMARY")
    log.info("=" * 60)
    log.info(f"  Output file       : {args.out}  ({size_mb:.2f} MB)")
    log.info(f"  Files used        : {len(latest_files)}")
    log.info(f"  Older duplicates  : {dropped}  (ignored)")
    log.info(f"  Files failed read : {fail_count}")
    log.info(f"  Total rows        : {len(master):,}")
    log.info(f"  Unique states     : {master['state'].nunique()}")
    log.info(f"  Unique makers     : {master['maker'].nunique()}")
    log.info(f"  Years covered     : {sorted(master['year'].unique().tolist())}")
    log.info(f"  Categories        : {sorted(master['vehicle_category'].unique().tolist())}")
    log.info(f"  Fuel types        : {sorted(master['fuel_type'].unique().tolist())}")
    log.info(f"  Source-date range : {master['source_date'].min()} -> {master['source_date'].max()}")
    log.info("=" * 60)

if __name__ == "__main__":
    main()
