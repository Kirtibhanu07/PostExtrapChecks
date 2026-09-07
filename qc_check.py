"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  QC CHECK AUTOMATION  —  Full Pipeline                                      ║
║  Replaces: MM_EATables macro + SampleTables macro (VBA)                    ║
║                                                                              ║
║  USAGE:                                                                      ║
║    python qc_check.py <adapt> <bsr> <youtube_mm> <matex> <sample>          ║
║    python qc_check.py adapt.xlsx bsr.xlsx yt.xlsx matex.xlsx sample.xlsx   ║
║    python qc_check.py ... --output MyQC_Jun26.xlsx                         ║
╚══════════════════════════════════════════════════════════════════════════════╝

WHAT THE VBA DID (manually, across multiple open workbooks):
─────────────────────────────────────────────────────────────
SUB MM_EATables  (runs on Adapt workbook):
  1. Pivot Adapt/programs → rows: country|Channel|matchday|progr.start|draID
                          → values: SUM(AVE 100%), SUM(total exposure msec)
     Copy pivot A:G → paste as values into cols I:O of RateCheck (right of itself)
     Add col P "Rate" = ROUND( AVE / (exposure/1000), 2 )
     Normalise col I (country): Macau→Macao, Pan-→Pan , Pan Global→Pan-Global
     Sort by col I (country) ASC, col P (Rate) ASC

  2. Open BSR/Database (header row 6):
     Pivot → rows: Market|Channel|Date(UTC)|Start(UTC)|Matchday
           → value: SUM(CPS Euro)
     Copy pivot A:F → paste into RateCheck cols T:Y
     
  3. Open YouTube MM/MM:
     Pivot → rows: Country|TV Channel|Date|Program Start(local)|Matchday
           → value: SUM(Rate 1 seconds)
     Paste below BSR data in RateCheck cols T onward
     Add col Z "Rate" = ROUND(Rate 1 seconds, 2)
     Normalise col T (country), Sort by col T ASC then col Z ASC
     Add col R "Check" = CONCAT(country_adapt, rate_adapt) = CONCAT(country_ref, rate_ref)
     AutoFilter on col R

  4. Ask user: "Has Sample File QC been created?"
     If NO → open Sample file → call SampleTables sub

  5. Back in Adapt/programs:
     Pivot → rows: country|Channel|progr.start|draID|matchday|BT(block)
           → value: SUM(total exposure msec)
     Add col H "BT"    = exposure / XLOOKUP(matchday, sample_col_F, sample_col_G)
     Add col I "Msec"  = exposure / XLOOKUP(matchday, sample_col_F, sample_col_H)
     Add col J "Check" = Msec% - BT%
     Open Sample file again → copy Sample!F2:H_last → paste ExtrapCheck!M1
     AutoFilter on H:J

SUB SampleTables  (runs on Sample workbook):
  1. Pivot Sample/programs → rows: matchday|brand|tool|location
                           → values: SUM(sequences), SUM(total exposure msec)
     This is the LEFT half of ExposureCheck (cols A:F)

  2. Open Matex/exposures:
     Pivot → same rows/values as above
     Copy Matex pivot A:F → paste into Sample/ExposureCheck cols J:O (RIGHT half)
     Add col H "Check" = CONCAT(left 6 cols) = CONCAT(right 6 cols)
     AutoFilter on col H

  3. Pivot Sample/programs → rows: pID(MM)|matchday|BT(block)
                           → value: SUM(total exposure msec)
     Copy col B (matchday) → paste col F → remove duplicates → unique matchdays
     Col G header "BT Block", col H header "Msec"
     G3 = SUMIF(matchday_col, F3, BT_block_col)   — total BT block per matchday
     H3 = SUMIF(matchday_col, F3, exposure_col)   — total Msec per matchday
     Fill down for all unique matchdays

OUTPUT SHEETS (all tabs coloured yellow):
  RateCheck     — Adapt pivot (cols I-P) | Check (col R) | BSR+YT pivot (cols T-Z)
  ExposureCheck — Sample pivot (cols A-F) | Check (col H) | Matex pivot (cols J-O)
  Sample        — pID/matchday/BT block pivot + matchday summary (cols F-H)
  ExtrapCheck   — Adapt pivot + BT%/Msec%/Check cols + matchday lookup (col M onward)
"""

import sys
import argparse
import traceback
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION — only edit here if column/sheet names change in source files
# ══════════════════════════════════════════════════════════════════════════════

SHEETS = {
    "adapt":  "programs",   # Adapt file
    "bsr":    "Database",   # BSR file — data starts at ROW 6 (header_row=5)
    "yt":     "MM",         # YouTube MM file
    "matex":  "exposures",  # Matex file
    "sample": "programs",   # Sample file — same schema as Adapt/programs
}

# Adapt & Sample share the same column schema
AC = {
    "country":   "country",
    "channel":   "channel",
    "matchday":  "matchday",
    "start":     "progr. start",
    "dra_id":    "draID",
    "ave":       "AVE (100%)",
    "exposure":  "total exposure (msec)",
    "bt_block":  "BT (block)",
    "pid":       "pID (MM)",
    "brand":     "brand",
    "tool":      "tool",
    "location":  "location",
    "sequences": "sequences",
}

BC = {  # BSR columns
    "market":   "Market",
    "channel":  "Channel",
    "date":     "Date(UTC)",
    "start":    "Start(UTC)",
    "matchday": "Matchday",
    "cps":      "CPS Euro",
}

YC = {  # YouTube MM columns
    "country":  "Country",
    "channel":  "TV Channel",
    "date":     "Date",
    "start":    "Program Start (local)",
    "matchday": "Matchday",
    "rate":     "Rate 1 seconds",
}

MC = {  # Matex columns
    "matchday":  "matchday",
    "brand":     "brand",
    "tool":      "tool",
    "location":  "location",
    "sequences": "sequences",
    "exposure":  "total exposure (msec)",
}

YELLOW = "FFFF00"   # Tab colour — VBA: .Color = 65535


# ══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def log(stage: str, status: str, detail: str = ""):
    marker = "✓" if status == "OK" else ("⚠" if status == "SKIPPED" else "✗")
    print(f"  [{marker}] {stage:<38} {status}   {detail}")


def load_sheet(path: str, sheet: str, header_row: int = 0) -> pd.DataFrame:
    """
    Load one worksheet from an Excel file.
    On any failure, lists the sheets that ARE available so the problem
    is immediately obvious without opening the file.
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {path}")
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=header_row, engine="openpyxl")
    except Exception as original:
        try:
            available = pd.ExcelFile(path, engine="openpyxl").sheet_names
        except Exception:
            available = ["<could not read workbook>"]
        raise ValueError(
            f"Sheet '{sheet}' not found in '{p.name}'.\n"
            f"          Available sheets: {available}"
        ) from original
    df.columns = [str(c).strip() for c in df.columns]
    return df


def assert_cols(df: pd.DataFrame, required: list, label: str):
    """Raise a clear error that lists every missing column."""
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"[{label}] Missing columns: {missing}\n"
            f"          Columns found: {list(df.columns)}"
        )


def normalise_country(series: pd.Series) -> pd.Series:
    """
    Three replacements done by VBA on country columns — order matters:
      Macau     → Macao
      Pan-      → Pan       (strip hyphen; runs BEFORE next rule)
      Pan Global → Pan-Global
    """
    s = series.astype(str).str.strip()
    s = s.str.replace("Macau",      "Macao",      regex=False)
    s = s.str.replace("Pan-",       "Pan ",        regex=False)
    s = s.str.replace("Pan Global", "Pan-Global",  regex=False)
    return s


def df_to_ws(ws, df: pd.DataFrame):
    """Write a DataFrame (with header row) into an openpyxl worksheet."""
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)


def yellow_tab(ws):
    ws.sheet_properties.tabColor = YELLOW


# ══════════════════════════════════════════════════════════════════════════════
# STAGE 1 — RATE CHECK
# ══════════════════════════════════════════════════════════════════════════════
#
# LEFT HALF  (mirrors Adapt pivot, VBA cols A-G copied to I-O):
#   Source : Adapt/programs
#   Rows   : country | Channel | matchday | progr.start | draID
#   Values : SUM(AVE 100%)  |  SUM(total exposure msec)
#   Extra  : Rate = ROUND( AVE / (exposure/1000), 2 )
#   Normalise country, sort by country ASC then Rate ASC
#
# RIGHT HALF (VBA cols T-Z):
#   Source A: BSR/Database  (header row 6)
#     Rows  : Market | Channel | Date(UTC) | Start(UTC) | Matchday
#     Value : SUM(CPS Euro)
#   Source B: YouTube MM/MM  — appended BELOW BSR
#     Rows  : Country | TV Channel | Date | Program Start(local) | Matchday
#     Value : SUM(Rate 1 seconds)
#   Extra  : Rate = ROUND(value, 2)
#   Normalise country col, sort by country ASC then Rate ASC
#
# CHECK (VBA col R):
#   =CONCAT(left_country, left_Rate) = CONCAT(right_country, right_Rate)
#   Placed BETWEEN the two halves in output for easy filtering.

def build_rate_check(adapt_df, bsr_df, yt_df) -> pd.DataFrame:

    # ── LEFT: Adapt pivot ─────────────────────────────────────────────────────
    req = [AC["country"], AC["channel"], AC["matchday"],
           AC["start"],   AC["dra_id"], AC["ave"], AC["exposure"]]
    assert_cols(adapt_df, req, "Adapt/programs → RateCheck LEFT")

    left = adapt_df.groupby(
        [AC["country"], AC["channel"], AC["matchday"], AC["start"], AC["dra_id"]],
        as_index=False, sort=False
    ).agg({AC["ave"]: "sum", AC["exposure"]: "sum"})

    left["country_norm"] = normalise_country(left[AC["country"]])
    # Rate = ROUND( AVE / (msec / 1000), 2 )
    left["Rate"] = (left[AC["ave"]] / (left[AC["exposure"]] / 1000)).round(2)
    left = left.sort_values(["country_norm", "Rate"], ascending=[True, True]).reset_index(drop=True)

    # Rename for output clarity
    left = left.rename(columns={
        AC["country"]:  "adapt_country",
        AC["channel"]:  "adapt_channel",
        AC["matchday"]: "adapt_matchday",
        AC["start"]:    "adapt_start",
        AC["dra_id"]:   "adapt_draID",
        AC["ave"]:      "adapt_AVE",
        AC["exposure"]: "adapt_exposure_msec",
        "country_norm": "adapt_country_norm",
        "Rate":         "adapt_Rate",
    })

    # ── RIGHT: BSR pivot ──────────────────────────────────────────────────────
    assert_cols(bsr_df, list(BC.values()), "BSR/Database → RateCheck RIGHT")

    bsr_pvt = bsr_df.groupby(
        [BC["market"], BC["channel"], BC["date"], BC["start"], BC["matchday"]],
        as_index=False, sort=False
    ).agg({BC["cps"]: "sum"})
    bsr_pvt["country_norm"] = normalise_country(bsr_pvt[BC["market"]])
    bsr_pvt["Rate"] = bsr_pvt[BC["cps"]].round(2)
    bsr_pvt = bsr_pvt.rename(columns={
        BC["market"]:   "ref_country_raw",
        BC["channel"]:  "ref_channel",
        BC["date"]:     "ref_date",
        BC["start"]:    "ref_start",
        BC["matchday"]: "ref_matchday",
        BC["cps"]:      "ref_value",
        "country_norm": "ref_country_norm",
        "Rate":         "ref_Rate",
    })

    # ── RIGHT: YouTube MM pivot ───────────────────────────────────────────────
    assert_cols(yt_df, list(YC.values()), "YouTube MM/MM → RateCheck RIGHT")

    yt_pvt = yt_df.groupby(
        [YC["country"], YC["channel"], YC["date"], YC["start"], YC["matchday"]],
        as_index=False, sort=False
    ).agg({YC["rate"]: "sum"})
    yt_pvt["country_norm"] = normalise_country(yt_pvt[YC["country"]])
    yt_pvt["Rate"] = yt_pvt[YC["rate"]].round(2)
    yt_pvt = yt_pvt.rename(columns={
        YC["country"]:  "ref_country_raw",
        YC["channel"]:  "ref_channel",
        YC["date"]:     "ref_date",
        YC["start"]:    "ref_start",
        YC["matchday"]: "ref_matchday",
        YC["rate"]:     "ref_value",
        "country_norm": "ref_country_norm",
        "Rate":         "ref_Rate",
    })

    # Stack BSR + YT (VBA appends YT rows below BSR in same columns)
    right = pd.concat([bsr_pvt, yt_pvt], ignore_index=True)
    right = right.sort_values(["ref_country_norm", "ref_Rate"], ascending=[True, True]).reset_index(drop=True)

    # ── CHECK column ──────────────────────────────────────────────────────────
    # VBA: =CONCAT(country_adapt, rate_adapt) = CONCAT(country_ref, rate_ref)
    # Align both halves positionally (VBA does row-by-row after pasting side-by-side)
    n = max(len(left), len(right))
    left  = left.reindex(range(n))
    right = right.reindex(range(n))

    left_key  = (left["adapt_country_norm"].astype(str).fillna("") +
                 left["adapt_Rate"].astype(str).fillna(""))
    right_key = (right["ref_country_norm"].astype(str).fillna("") +
                 right["ref_Rate"].astype(str).fillna(""))
    check_col = (left_key == right_key).rename("Check")

    return pd.concat([left, check_col, right], axis=1)


# ══════════════════════════════════════════════════════════════════════════════
# STAGE 2 — EXPOSURE CHECK
# ══════════════════════════════════════════════════════════════════════════════
#
# LEFT HALF  (VBA cols A-F, built from Sample/programs):
#   Source : Sample/programs
#   Rows   : matchday | brand | tool | location
#   Values : SUM(sequences)  |  SUM(total exposure msec)
#
# RIGHT HALF (VBA cols J-O, built from Matex/exposures):
#   Source : Matex/exposures  — exact same pivot structure
#   Pasted into cols J:O of ExposureCheck in the Sample workbook
#
# CHECK (VBA col H):
#   =CONCAT(left 6 cols) = CONCAT(right 6 cols)

def build_exposure_check(sample_df, matex_df) -> pd.DataFrame:

    grp_keys = [AC["matchday"], AC["brand"], AC["tool"], AC["location"]]
    val_keys  = [AC["sequences"], AC["exposure"]]
    assert_cols(sample_df, grp_keys + val_keys, "Sample/programs → ExposureCheck LEFT")
    assert_cols(matex_df,
                [MC["matchday"], MC["brand"], MC["tool"], MC["location"],
                 MC["sequences"], MC["exposure"]],
                "Matex/exposures → ExposureCheck RIGHT")

    # LEFT pivot — Sample
    left = sample_df.groupby(grp_keys, as_index=False, sort=False).agg({
        AC["sequences"]: "sum",
        AC["exposure"]:  "sum",
    }).rename(columns={
        AC["matchday"]:  "s_matchday",
        AC["brand"]:     "s_brand",
        AC["tool"]:      "s_tool",
        AC["location"]:  "s_location",
        AC["sequences"]: "s_sequences",
        AC["exposure"]:  "s_exposure_msec",
    })
    # Sort left side — mirrors what a pivot table does by default (row label order)
    left = left.sort_values(
        ["s_matchday", "s_brand", "s_tool", "s_location"],
        ascending=True
    ).reset_index(drop=True)

    # RIGHT pivot — Matex
    right = matex_df.groupby(
        [MC["matchday"], MC["brand"], MC["tool"], MC["location"]],
        as_index=False, sort=False
    ).agg({
        MC["sequences"]: "sum",
        MC["exposure"]:  "sum",
    }).rename(columns={
        MC["matchday"]:  "m_matchday",
        MC["brand"]:     "m_brand",
        MC["tool"]:      "m_tool",
        MC["location"]:  "m_location",
        MC["sequences"]: "m_sequences",
        MC["exposure"]:  "m_exposure_msec",
    })
    # Sort right side — same order as left so positional comparison is valid
    right = right.sort_values(
        ["m_matchday", "m_brand", "m_tool", "m_location"],
        ascending=True
    ).reset_index(drop=True)

    # Align positionally (after both sides are sorted identically)
    n = max(len(left), len(right))
    left  = left.reindex(range(n))
    right = right.reindex(range(n))

    # CHECK: CONCAT(all 6 left cols) = CONCAT(all 6 right cols)
    left_key  = (left["s_matchday"].astype(str).fillna("")  +
                 left["s_brand"].astype(str).fillna("")     +
                 left["s_tool"].astype(str).fillna("")      +
                 left["s_location"].astype(str).fillna("")  +
                 left["s_sequences"].astype(str).fillna("") +
                 left["s_exposure_msec"].astype(str).fillna(""))
    right_key = (right["m_matchday"].astype(str).fillna("")  +
                 right["m_brand"].astype(str).fillna("")     +
                 right["m_tool"].astype(str).fillna("")      +
                 right["m_location"].astype(str).fillna("")  +
                 right["m_sequences"].astype(str).fillna("") +
                 right["m_exposure_msec"].astype(str).fillna(""))
    check_col = (left_key == right_key).rename("Check")

    return pd.concat([left, check_col, right], axis=1)


# ══════════════════════════════════════════════════════════════════════════════
# STAGE 3 — SAMPLE SHEET  (built from Sample/programs)
# ══════════════════════════════════════════════════════════════════════════════
#
# PIVOT (main table, VBA cols A-D):
#   Source : Sample/programs
#   Rows   : pID(MM) | matchday | BT(block)
#   Value  : SUM(total exposure msec)
#
# SUMMARY TABLE (VBA cols F-H):
#   Col F : unique matchdays (copy of col B with duplicates removed)
#   Col G : "BT Block" = SUMIF(matchday_col, F3, BT_block_col)
#   Col H : "Msec"     = SUMIF(matchday_col, F3, exposure_col)
#   This summary is later used as the XLOOKUP source for ExtrapCheck.
#
# Returns: (sample_detail_df, sample_summary_df)

def build_sample(sample_df):

    req = [AC["pid"], AC["matchday"], AC["bt_block"], AC["exposure"]]
    assert_cols(sample_df, req, "Sample/programs → Sample sheet")

    # Main pivot: pID(MM) | matchday | BT(block) → sum exposure
    # This is the pivot output — each unique pID+matchday+BT_block is ONE row
    detail = sample_df.groupby(
        [AC["pid"], AC["matchday"], AC["bt_block"]],
        as_index=False, sort=False
    ).agg({AC["exposure"]: "sum"}).rename(columns={
        AC["pid"]:      "pID_MM",
        AC["matchday"]: "matchday",
        AC["bt_block"]: "BT_block",
        AC["exposure"]: "exposure_msec",
    })

    # VBA SUMIF runs on the PIVOT OUTPUT (col B=matchday, col C=BT_block, col D=exposure_msec)
    # NOT on the raw data — because pivot deduplicated to unique pID+matchday+BT_block rows
    # G3 = SUMIF(pivot_matchday_col, unique_matchday, pivot_BT_block_col)
    # H3 = SUMIF(pivot_matchday_col, unique_matchday, pivot_exposure_col)
    pivot_for_sumif = detail  # SUMIF source = pivot output rows

    # Unique matchdays in first-seen order (VBA: copy col B → remove duplicates)
    matchday_order = detail["matchday"].drop_duplicates().reset_index(drop=True)

    summary = (
        matchday_order.rename("matchday")
        .to_frame()
        .merge(
            pivot_for_sumif.groupby("matchday", as_index=False, sort=False).agg(
                BT_Block=("BT_block",    "sum"),   # sum of BT_block values per matchday from pivot
                Msec    =("exposure_msec","sum"),   # sum of exposure_msec per matchday from pivot
            ),
            on="matchday", how="left"
        )
    )

    # Write summary alongside detail in same sheet — VBA puts pivot (cols A-D)
    # and summary table (cols F-H) on the same Sample sheet
    max_rows = max(len(detail), len(summary))
    detail_out  = detail.reindex(range(max_rows))
    summary_out = summary.reindex(range(max_rows))
    sep = pd.Series([None] * max_rows, name="_sep")
    combined = pd.concat([detail_out, sep, summary_out], axis=1)

    return combined, summary


# ══════════════════════════════════════════════════════════════════════════════
# STAGE 4 — EXTRAP CHECK  (built from Adapt/programs + sample_summary)
# ══════════════════════════════════════════════════════════════════════════════
#
# PIVOT (VBA cols A-G):
#   Source : Adapt/programs
#   Rows   : country | Channel | progr.start | draID | matchday | BT(block)
#   Value  : SUM(total exposure msec)
#
# DERIVED COLS (VBA cols H, I, J):
#   H "BT"    = exposure / XLOOKUP(matchday, sample_F, sample_G)   → % of BT Block
#   I "Msec"  = exposure / XLOOKUP(matchday, sample_F, sample_H)   → % of total Msec
#   J "Check" = Msec% - BT%  (should be ~0 if extrapolation is correct)
#
# LOOKUP TABLE (VBA cols M onward, pasted from Sample!F2:H_last):
#   matchday | BT_Block | Msec
#   VBA hardcoded this as R1C13:R20C15 (rows 1-20, cols 13-15).
#   We replace the XLOOKUP with a pandas merge on matchday.

def build_extrap_check(adapt_df, sample_summary) -> pd.DataFrame:

    req = [AC["country"], AC["channel"], AC["start"], AC["dra_id"],
           AC["matchday"], AC["bt_block"], AC["exposure"]]
    assert_cols(adapt_df, req, "Adapt/programs → ExtrapCheck")

    for col in ["matchday", "BT_Block", "Msec"]:
        if col not in sample_summary.columns:
            raise ValueError(
                f"Sample summary missing column '{col}'. "
                f"Stage 3 must succeed before Stage 4 can run."
            )

    # Pivot Adapt — rows: country|Channel|progr.start|draID|matchday|BT(block)
    pvt = adapt_df.groupby(
        [AC["country"], AC["channel"], AC["start"], AC["dra_id"],
         AC["matchday"], AC["bt_block"]],
        as_index=False, sort=False
    ).agg({AC["exposure"]: "sum"}).rename(columns={
        AC["country"]:  "country",
        AC["channel"]:  "channel",
        AC["start"]:    "progr_start",
        AC["dra_id"]:   "draID",
        AC["matchday"]: "matchday",
        AC["bt_block"]: "BT_block",
        AC["exposure"]: "exposure_msec",
    })

    # XLOOKUP equivalent: merge sample_summary on matchday
    # sample_summary col = "matchday" | "BT_Block" | "Msec"
    # BT%   = row exposure_msec / BT_Block total for that matchday  (from Sample file)
    # Msec% = row exposure_msec / Msec total for that matchday      (from Sample file)
    # Check = Msec% - BT%   (expect 0 if extrapolation correct)
    merged = pvt.merge(
        sample_summary[["matchday", "BT_Block", "Msec"]],
        on="matchday", how="left"
    )

    merged["BT_pct"]   = (merged["BT_block"] /
                           merged["BT_Block"].replace(0, pd.NA)).round(6)
    merged["Msec_pct"] = (merged["exposure_msec"] /
                           merged["Msec"].replace(0, pd.NA)).round(6)
    merged["Check"]    = (merged["Msec_pct"] - merged["BT_pct"]).round(6)

    # Drop the lookup columns used for calculation — they will appear separately in col M onward
    # VBA: paste Sample!F2:H_last into ExtrapCheck!M1
    # We append them as separate columns after a gap column (blank separator)
    merged = merged.drop(columns=["BT_Block", "Msec"])

    # Reindex to make room for separator + lookup table
    n = len(merged)
    lookup = sample_summary[["matchday", "BT_Block", "Msec"]].reset_index(drop=True)

    # Pad shorter side with NaN rows so concat aligns
    max_rows = max(n, len(lookup))
    merged = merged.reindex(range(max_rows))
    lookup = lookup.reindex(range(max_rows))

    # Blank separator column (mirrors the gap between col J and col M in VBA)
    sep = pd.Series([None] * max_rows, name="_sep")

    return pd.concat([merged, sep, lookup], axis=1)


# ══════════════════════════════════════════════════════════════════════════════
# OUTPUT WRITER
# ══════════════════════════════════════════════════════════════════════════════

def write_output(output_path: str, sheets: dict):
    """Write all result DataFrames to one Excel workbook. All tabs coloured yellow."""
    if not sheets:
        raise RuntimeError("No sheets produced — all pipeline stages failed.")
    wb = Workbook()
    wb.remove(wb.active)  # remove default blank sheet
    for name, df in sheets.items():
        ws = wb.create_sheet(title=name)
        df_to_ws(ws, df)
        yellow_tab(ws)
    wb.save(output_path)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="QC Check Automation — replaces MM_EATables + SampleTables VBA macros",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example:
  python qc_check.py adapt.xlsx bsr.xlsx yt_mm.xlsx matex.xlsx sample.xlsx
  python qc_check.py adapt.xlsx bsr.xlsx yt_mm.xlsx matex.xlsx sample.xlsx --output QC_Jun26.xlsx
        """
    )
    parser.add_argument("adapt",      help="Adapt file (.xlsx) — sheet: 'programs'")
    parser.add_argument("bsr",        help="BSR file (.xlsx)   — sheet: 'Database' (header row 6)")
    parser.add_argument("youtube_mm", help="YouTube MM (.xlsx)  — sheet: 'MM'")
    parser.add_argument("matex",      help="Matex file (.xlsx)  — sheet: 'exposures'")
    parser.add_argument("sample",     help="Sample file (.xlsx) — sheet: 'programs'")
    parser.add_argument("--output",   default="QC_Output.xlsx",
                        help="Output file path (default: QC_Output.xlsx)")
    args = parser.parse_args()

    print()
    print("╔══════════════════════════════════════════════════════╗")
    print("║          QC CHECK AUTOMATION PIPELINE               ║")
    print("╚══════════════════════════════════════════════════════╝")
    print()
    print(f"  Adapt      : {args.adapt}")
    print(f"  BSR        : {args.bsr}")
    print(f"  YouTube MM : {args.youtube_mm}")
    print(f"  Matex      : {args.matex}")
    print(f"  Sample     : {args.sample}")
    print(f"  Output     : {args.output}")
    print()

    # ─────────────────────────────────────────────────────────────────────────
    # LOAD PHASE
    # All 5 files are attempted independently.
    # Every failure is reported before the pipeline aborts,
    # so you see all problems in one run.
    # ─────────────────────────────────────────────────────────────────────────
    print("── LOAD ─────────────────────────────────────────────────")
    load_ok  = True
    adapt_df = bsr_df = yt_df = matex_df = sample_df = None

    for label, path, sheet, hdr, target in [
        ("Adapt",      args.adapt,      SHEETS["adapt"],  0, "adapt_df"),
        ("BSR",        args.bsr,        SHEETS["bsr"],    5, "bsr_df"),
        ("YouTube MM", args.youtube_mm, SHEETS["yt"],     0, "yt_df"),
        ("Matex",      args.matex,      SHEETS["matex"],  0, "matex_df"),
        ("Sample",     args.sample,     SHEETS["sample"], 0, "sample_df"),
    ]:
        try:
            df = load_sheet(path, sheet, header_row=hdr)
            if   target == "adapt_df":  adapt_df  = df
            elif target == "bsr_df":    bsr_df    = df
            elif target == "yt_df":     yt_df     = df
            elif target == "matex_df":  matex_df  = df
            elif target == "sample_df": sample_df = df
            log(f"Load {label}", "OK", f"{len(df):,} rows  |  {len(df.columns)} cols")
        except FileNotFoundError as e:
            log(f"Load {label}", "FAILED", f"File not found → {e}")
            load_ok = False
        except ValueError as e:
            log(f"Load {label}", "FAILED", str(e))
            load_ok = False
        except Exception as e:
            log(f"Load {label}", "FAILED", f"Unexpected → {e}")
            traceback.print_exc()
            load_ok = False

    if not load_ok:
        print()
        print("  ✗  PIPELINE ABORTED — fix the load errors above and re-run.")
        print()
        sys.exit(1)

    # ─────────────────────────────────────────────────────────────────────────
    # STAGE PHASE
    # Each stage is isolated — a failure is logged, remaining stages still run.
    # sample_summary is the only cross-stage dependency (Stage 3 → Stage 4).
    # ─────────────────────────────────────────────────────────────────────────
    print()
    print("── STAGES ───────────────────────────────────────────────")

    pipeline_ok    = True
    output_sheets  = {}
    sample_summary = None

    # Stage 1 — RateCheck
    try:
        rate_df = build_rate_check(adapt_df, bsr_df, yt_df)
        output_sheets["RateCheck"] = rate_df
        log("Stage 1 : RateCheck",      "OK", f"{len(rate_df):,} rows")
    except ValueError as e:
        log("Stage 1 : RateCheck",      "FAILED", f"Column error → {e}")
        pipeline_ok = False
    except Exception as e:
        log("Stage 1 : RateCheck",      "FAILED", str(e))
        traceback.print_exc()
        pipeline_ok = False

    # Stage 2 — ExposureCheck  (Sample LEFT vs Matex RIGHT)
    try:
        exp_df = build_exposure_check(sample_df, matex_df)
        output_sheets["ExposureCheck"] = exp_df
        log("Stage 2 : ExposureCheck",  "OK", f"{len(exp_df):,} rows")
    except ValueError as e:
        log("Stage 2 : ExposureCheck",  "FAILED", f"Column error → {e}")
        pipeline_ok = False
    except Exception as e:
        log("Stage 2 : ExposureCheck",  "FAILED", str(e))
        traceback.print_exc()
        pipeline_ok = False

    # Stage 3 — Sample + Summary  (from Sample/programs)
    try:
        sample_sheet, sample_summary = build_sample(sample_df)
        output_sheets["Sample"] = sample_sheet
        log("Stage 3 : Sample",         "OK",
            f"{len(sample_sheet):,} rows  |  {len(sample_summary):,} unique matchdays")
    except ValueError as e:
        log("Stage 3 : Sample",         "FAILED", f"Column error → {e}")
        pipeline_ok = False
    except Exception as e:
        log("Stage 3 : Sample",         "FAILED", str(e))
        traceback.print_exc()
        pipeline_ok = False

    # Stage 4 — ExtrapCheck  (Adapt pivot + XLOOKUP from sample_summary)
    try:
        if sample_summary is None:
            raise RuntimeError(
                "Sample summary unavailable — Stage 3 must succeed before Stage 4 can run."
            )
        extrap_df = build_extrap_check(adapt_df, sample_summary)
        output_sheets["ExtrapCheck"] = extrap_df
        log("Stage 4 : ExtrapCheck",    "OK", f"{len(extrap_df):,} rows")
    except RuntimeError as e:
        log("Stage 4 : ExtrapCheck",    "SKIPPED", str(e))
        pipeline_ok = False
    except ValueError as e:
        log("Stage 4 : ExtrapCheck",    "FAILED", f"Column error → {e}")
        pipeline_ok = False
    except Exception as e:
        log("Stage 4 : ExtrapCheck",    "FAILED", str(e))
        traceback.print_exc()
        pipeline_ok = False

    # Write output
    print()
    print("── OUTPUT ───────────────────────────────────────────────")
    try:
        write_output(args.output, output_sheets)
        log("Write output", "OK",
            f"{len(output_sheets)}/4 sheets written → {args.output}")
    except Exception as e:
        log("Write output", "FAILED", str(e))
        traceback.print_exc()
        pipeline_ok = False

    # Final summary
    print()
    print("╔══════════════════════════════════════════════════════╗")
    if pipeline_ok and len(output_sheets) == 4:
        print(f"║  ✓  COMPLETE — all 4 sheets written                 ║")
        print(f"║     {args.output:<51}║")
    elif output_sheets:
        sheets_done = ", ".join(output_sheets.keys())
        print(f"║  ⚠  PARTIAL — {len(output_sheets)}/4 sheets written               ║")
        print(f"║     {args.output:<51}║")
        print(f"║     Sheets OK : {sheets_done:<36}║")
        print(f"║     Review FAILED stages above                      ║")
    else:
        print(f"║  ✗  FAILED — no output written                      ║")
        print(f"║     Review all FAILED stages above                  ║")
    print("╚══════════════════════════════════════════════════════╝")
    print()

    sys.exit(0 if (pipeline_ok and len(output_sheets) == 4) else 1)


if __name__ == "__main__":
    main()
