"""Data Loader and Processor for KPI Benchmark and Trial Analysis.

Supports Excel (.xlsx, .xls) and CSV (.csv) files, auto-detects date columns,
extracts numeric KPIs/counters, and computes Pre/Post trial statistics.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


@dataclass
class KPIStatistics:
    """Benchmark statistics comparing Pre and Post trial periods."""
    kpi_name: str
    pre_count: int
    pre_mean: float
    pre_median: float
    pre_std: float
    pre_min: float
    pre_max: float
    post_count: int
    post_mean: float
    post_median: float
    post_std: float
    post_min: float
    post_max: float
    abs_delta: float
    pct_delta: float
    has_valid_pre: bool
    has_valid_post: bool


class DataLoader:
    """Loads, validates, and slices KPI data from CSV and Excel files."""

    DATE_CANDIDATES = [
        "recorddate",
        "date",
        "record_date",
        "date_id",
        "dateid",
        "timestamp",
        "time",
        "period_start_time",
        "period_start",
        "start_time",
        "starttime",
        "day",
        "datetime",
        "period",
        "report_date",
        "event_date",
    ]

    EXCLUDE_COL_PATTERNS = [
        r"^objectid$",
        r"^cell$",
        r"^cell_name$",
        r"^cellname$",
        r"^site$",
        r"^sitename$",
        r"^enodeb.*$",
        r"^node.*$",
        r"^id$",
    ]

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.df_raw: Optional[pd.DataFrame] = None
        self.df: Optional[pd.DataFrame] = None
        self.date_col: Optional[str] = None
        self.available_dates: List[str] = []
        self.numeric_kpis: List[str] = []
        self.load_data()

    def load_data(self) -> None:
        """Reads file, standardizes dates, and discovers numeric KPI columns."""
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")

        ext = os.path.splitext(self.file_path)[1].lower()
        if ext in [".xlsx", ".xls"]:
            self.df_raw = pd.read_excel(self.file_path)
        elif ext == ".csv":
            # Try utf-8 first, then fallback to latin1 or cp1252
            try:
                self.df_raw = pd.read_csv(self.file_path, encoding="utf-8")
            except UnicodeDecodeError:
                self.df_raw = pd.read_csv(self.file_path, encoding="latin1")
        else:
            raise ValueError(f"Unsupported file format '{ext}'. Expected .xlsx, .xls, or .csv")

        self.df = self.df_raw.copy()
        self._detect_and_standardize_date()
        self._discover_numeric_kpis()

    def _parse_date_series(self, series: pd.Series) -> pd.Series:
        """Robustly parses date series, supporting standard dates, timestamps, and YYYYMMDD integers."""
        # 1. Direct datetime parse
        parsed = pd.to_datetime(series, errors="coerce")
        # 2. If mostly NaN or parsed into 1970 (integer epoch issue), try parsing as string
        valid_cnt = parsed.notna().sum()
        year_too_early = (valid_cnt > 0 and parsed.dropna().dt.year.min() < 1995)
        if valid_cnt < len(series) * 0.5 or year_too_early:
            str_series = series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            parsed_str = pd.to_datetime(str_series, errors="coerce")
            if parsed_str.notna().sum() > valid_cnt:
                parsed = parsed_str
        return parsed

    def _detect_and_standardize_date(self) -> None:
        """Detects date column, standardizes format to YYYY-MM-DD, and sorts chronologically."""
        cols = list(self.df.columns)
        detected_col = None

        # 1. Check exact or lower match against known candidates
        for cand in self.DATE_CANDIDATES:
            for col in cols:
                clean_name = re.sub(r"[\s_-]+", "", col.strip().lower())
                clean_cand = re.sub(r"[\s_-]+", "", cand.lower())
                if clean_name == clean_cand:
                    detected_col = col
                    break
            if detected_col:
                break

        # 2. Check substring match ("date" or "time")
        if not detected_col:
            for col in cols:
                clean_col = col.strip().lower()
                if "date" in clean_col or "time" in clean_col:
                    try:
                        parsed = self._parse_date_series(self.df[col])
                        if parsed.notna().sum() > len(self.df) * 0.4:
                            detected_col = col
                            break
                    except Exception:
                        continue

        # 3. Fallback: inspect any column for valid dates
        if not detected_col:
            for col in cols:
                try:
                    parsed = self._parse_date_series(self.df[col])
                    if parsed.notna().sum() > len(self.df) * 0.7:
                        detected_col = col
                        break
                except Exception:
                    continue

        if not detected_col:
            raise ValueError(
                "Could not detect a valid Date column in the dataset. "
                "Ensure your file contains a column like 'RecordDate', 'Date', or 'Date_ID'."
            )

        self.date_col = detected_col
        
        # Parse dates
        self.df["_parsed_datetime"] = self._parse_date_series(self.df[self.date_col])
        # Drop rows with invalid dates
        self.df = self.df.dropna(subset=["_parsed_datetime"]).copy()
        # Sort chronologically
        self.df = self.df.sort_values(by="_parsed_datetime").reset_index(drop=True)
        # Create standardized string representation
        self.df["_date_str"] = self.df["_parsed_datetime"].dt.strftime("%Y-%m-%d")
        self.available_dates = sorted(self.df["_date_str"].unique().tolist())

    def _discover_numeric_kpis(self) -> None:
        """Finds all numeric and parsable KPI/counter columns."""
        kpis = []
        for col in self.df.columns:
            if col in ["_parsed_datetime", "_date_str", self.date_col]:
                continue

            # Check if excluded by name pattern
            is_excluded = False
            for pat in self.EXCLUDE_COL_PATTERNS:
                if re.match(pat, col.strip(), re.IGNORECASE):
                    is_excluded = True
                    break
            if is_excluded:
                continue

            # Clean potential commas from numbers like "1,234.56"
            series = self.df[col]
            if series.dtype == object:
                # If values contain semicolons (e.g. vector/distribution counters like '0;0'), skip raw vector
                sample_str = str(series.dropna().iloc[0]) if series.dropna().shape[0] > 0 else ""
                if ";" in sample_str:
                    continue
                series = series.astype(str).str.replace(",", "", regex=False).str.strip()

            numeric_series = pd.to_numeric(series, errors="coerce")
            valid_count = numeric_series.notna().sum()
            if valid_count > 0 and (valid_count >= len(self.df) * 0.3):
                self.df[col] = numeric_series
                kpis.append(col)

        self.numeric_kpis = kpis

    def get_suggested_periods(self, impl_date: Optional[str] = None) -> Dict[str, str]:
        """Auto-suggests Pre and Post periods based on available dates and implementation date."""
        if not self.available_dates:
            return {
                "pre_start": "",
                "pre_end": "",
                "impl_date": "",
                "post_start": "",
                "post_end": "",
            }

        dates = self.available_dates
        if not impl_date or impl_date not in dates:
            # Default implementation date to mid-point
            mid_idx = max(0, len(dates) // 2)
            impl_date = dates[mid_idx]

        impl_idx = dates.index(impl_date)

        # Pre-period: start to day before implementation
        pre_start = dates[0]
        if impl_idx > 0:
            pre_end = dates[impl_idx - 1]
        else:
            pre_end = dates[0]

        # Post-period: day of implementation or day after to end
        if impl_idx < len(dates) - 1:
            post_start = dates[impl_idx + 1]
        else:
            post_start = dates[impl_idx]
        post_end = dates[-1]

        return {
            "pre_start": pre_start,
            "pre_end": pre_end,
            "impl_date": impl_date,
            "post_start": post_start,
            "post_end": post_end,
        }

    def compute_kpi_statistics(
        self,
        kpi_name: str,
        pre_start: str,
        pre_end: str,
        post_start: str,
        post_end: str,
    ) -> KPIStatistics:
        """Computes statistical metrics for pre and post periods for a KPI."""
        if kpi_name not in self.df.columns:
            raise KeyError(f"KPI '{kpi_name}' not found in dataset.")

        pre_mask = (self.df["_date_str"] >= pre_start) & (self.df["_date_str"] <= pre_end)
        post_mask = (self.df["_date_str"] >= post_start) & (self.df["_date_str"] <= post_end)

        # If multiple rows per date exist (e.g. multi-cell/multi-node data), aggregate by date
        if self.df["_date_str"].duplicated().any():
            pre_series = self.df.loc[pre_mask].groupby("_date_str")[kpi_name].mean().dropna().values
            post_series = self.df.loc[post_mask].groupby("_date_str")[kpi_name].mean().dropna().values
        else:
            pre_series = self.df.loc[pre_mask, kpi_name].dropna().values
            post_series = self.df.loc[post_mask, kpi_name].dropna().values

        has_valid_pre = len(pre_series) > 0
        has_valid_post = len(post_series) > 0

        pre_mean = float(np.mean(pre_series)) if has_valid_pre else 0.0
        pre_median = float(np.median(pre_series)) if has_valid_pre else 0.0
        pre_std = float(np.std(pre_series)) if has_valid_pre and len(pre_series) > 1 else 0.0
        pre_min = float(np.min(pre_series)) if has_valid_pre else 0.0
        pre_max = float(np.max(pre_series)) if has_valid_pre else 0.0

        post_mean = float(np.mean(post_series)) if has_valid_post else 0.0
        post_median = float(np.median(post_series)) if has_valid_post else 0.0
        post_std = float(np.std(post_series)) if has_valid_post and len(post_series) > 1 else 0.0
        post_min = float(np.min(post_series)) if has_valid_post else 0.0
        post_max = float(np.max(post_series)) if has_valid_post else 0.0

        abs_delta = post_mean - pre_mean if (has_valid_pre and has_valid_post) else 0.0
        if has_valid_pre and pre_mean != 0:
            pct_delta = (abs_delta / abs(pre_mean)) * 100.0
        else:
            pct_delta = 0.0

        return KPIStatistics(
            kpi_name=kpi_name,
            pre_count=len(pre_series),
            pre_mean=pre_mean,
            pre_median=pre_median,
            pre_std=pre_std,
            pre_min=pre_min,
            pre_max=pre_max,
            post_count=len(post_series),
            post_mean=post_mean,
            post_median=post_median,
            post_std=post_std,
            post_min=post_min,
            post_max=post_max,
            abs_delta=abs_delta,
            pct_delta=pct_delta,
            has_valid_pre=has_valid_pre,
            has_valid_post=has_valid_post,
        )

    def get_timeline_data(
        self,
        kpi_name: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
    ) -> Tuple[List[str], List[Optional[float]]]:
        """Returns the chronological list of date labels and corresponding KPI values."""
        df_sub = self.df
        if start_date:
            df_sub = df_sub[df_sub["_date_str"] >= start_date]
        if end_date:
            df_sub = df_sub[df_sub["_date_str"] <= end_date]

        # Group by date if multiple rows exist per date (e.g. multi-cell / multi-node datasets)
        if df_sub["_date_str"].duplicated().any():
            daily_agg = df_sub.groupby("_date_str")[kpi_name].mean().reset_index()
            daily_agg = daily_agg.sort_values(by="_date_str")
            dates = daily_agg["_date_str"].tolist()
            vals = [None if pd.isna(v) else float(v) for v in daily_agg[kpi_name].tolist()]
        else:
            dates = df_sub["_date_str"].tolist()
            vals = [None if pd.isna(v) else float(v) for v in df_sub[kpi_name].tolist()]
        return dates, vals
