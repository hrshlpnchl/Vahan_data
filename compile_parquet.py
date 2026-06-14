#!/usr/bin/env python3
"""
compile_parquet.py
Reads all xlsx files from ./data/ and writes a single master.parquet.
Run this locally once to bootstrap, then GitHub Actions runs it at 5am daily.

Usage:
    python compile_parquet.py
    python compile_parquet.py --data-dir ./vahan_downloads   # custom folder
"""

import os
import re
import argparse
import pandas as pd
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ── CLI ──────────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument("--data-dir", default="./data", help="Folder with xlsx files")
parser.add_argument("--out", default="master.parquet", help="Output parquet path")
args = parser.parse_args()

DATA_DIR = args.data_dir
OUT_PATH = args.out

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
          "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


# ── Helpers ──────────────────────────────────────────────────────────────────

def safe_int(val) -> int:
    if pd.isna(val):
        return 0
    try:
        return int(float(str(val).replace(",", "").strip()))
    except (ValueError, TypeError):
        return 0


def classify_file(filename: str):
    """
    Extract (state, vehicle_type, fuel_type) from filename.
    Pattern: {State}_{2W|3W|4W}_{PureEV|AllFuel}_{YEAR}.xlsx
    """
    base = re.sub(r"_\d{4}\.xlsx$", "", filename, flags=re.IGNORECASE)
    match = re.search(r"^(.+)_(2W|3W|4W)_(PureEV|AllFuel)$", base)
    if not match:
        return None, None, None, None

    state_raw, vtype, ftype = match.group(1), match.group(2), match.group(3)

    # Extract year from filename
    year_match = re.search(r"_(\d{4})\.xlsx$", filename, re.IGNORECASE)
    year = int(year_match.group(1)) if year_match else 2026

    state = state_raw.replace("_", " ")
    if "All Vahan4 Running States" in state or "All India" in state:
        state = "All India"

    return state, vtype, ftype, year


def parse_file(filepath: str) -> list[dict]:
    """
    Parse one cleaned Vahan xlsx (layout: Row0=title, Row1=header, Row2+=data).
    Returns list of row dicts.
    """
    try:
        df = pd.read_excel(filepath, header=None, engine="openpyxl")
    except Exception as e:
        print(f"  ❌ Cannot read {os.path.basename(filepath)}: {e}")
        return []

    # Detect active months dynamically (handles files with fewer months)
    header_row = df.iloc[1] if len(df) > 1 else pd.Series()
    active_month_cols = {}
    for col_idx, val in enumerate(header_row):
        val_str = str(val).strip().upper()
        for mi, m in enumerate(MONTHS):
            if m in val_str or val_str.startswith(m[:3]):
                active_month_cols[m] = col_idx
                break

    # Fallback: fixed layout (col2=JAN .. col7=JUN, col8=Total)
    if not active_month_cols:
        active_month_cols = {m: i + 2 for i, m in enumerate(MONTHS[:6])}

    records = []
    for idx in range(2, len(df)):
        row = df.iloc[idx]
        maker = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if not maker or maker.lower() in ("nan", "total", "maker"):
            continue

        monthly = {m: safe_int(row.iloc[col]) if col < len(row) else 0
                   for m, col in active_month_cols.items()}

        # Fill missing months with 0
        for m in MONTHS:
            monthly.setdefault(m, 0)

        total = sum(monthly.values())
        if total == 0:
            continue

        records.append({"Maker": maker, **monthly, "TOTAL": total})

    return records


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if not os.path.isdir(DATA_DIR):
        print(f"❌ Data directory not found: {DATA_DIR}")
        raise SystemExit(1)

    xlsx_files = sorted(
        f for f in os.listdir(DATA_DIR)
        if f.lower().endswith(".xlsx") and not f.startswith("~")
    )
    print(f"📂 Found {len(xlsx_files)} xlsx files in {DATA_DIR}")

    all_records = []
    loaded = skipped = 0

    for fname in xlsx_files:
        state, vtype, ftype, year = classify_file(fname)
        if state is None:
            print(f"  ⏭  SKIP (no match): {fname}")
            skipped += 1
            continue

        rows = parse_file(os.path.join(DATA_DIR, fname))
        for r in rows:
            r.update({"State": state, "VehicleType": vtype,
                       "FuelType": ftype, "Year": year})
        all_records.extend(rows)
        loaded += 1
        print(f"  ✅ {fname} → {state} | {vtype} | {ftype} | {len(rows)} rows")

    if not all_records:
        print("❌ No records — parquet not written")
        raise SystemExit(1)

    df = pd.DataFrame(all_records)
    df["Maker"] = df["Maker"].str.strip()

    # Ensure all month columns exist
    for m in MONTHS:
        if m not in df.columns:
            df[m] = 0

    # Enforce dtypes for small file size
    for m in MONTHS + ["TOTAL"]:
        df[m] = df[m].fillna(0).astype("int32")

    df["State"] = df["State"].astype("category")
    df["VehicleType"] = df["VehicleType"].astype("category")
    df["FuelType"] = df["FuelType"].astype("category")
    df["Maker"] = df["Maker"].astype("category")
    df["Year"] = df["Year"].astype("int16")

    df.to_parquet(OUT_PATH, index=False, compression="snappy")

    size_kb = os.path.getsize(OUT_PATH) / 1024
    print(f"\n✅ Written: {OUT_PATH} ({size_kb:.0f} KB, {len(df):,} rows)")
    print(f"   Loaded: {loaded} files | Skipped: {skipped} files")
    print(f"   States: {df['State'].nunique()} | Makers: {df['Maker'].nunique()}")


if __name__ == "__main__":
    main()
