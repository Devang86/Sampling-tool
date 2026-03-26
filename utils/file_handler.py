"""
KKC Audit Sampling Tool – File Handler & Sample Selection
Handles Excel upload, column mapping, population preparation,
stratification, and the three principal sample selection methods.

Reference: KKC Sampling Guide Section 26, Appendix II
"""

import pandas as pd
import numpy as np
from typing import Optional, List, Tuple, Dict


# ─── FILE READING ─────────────────────────────────────────────────────────────

def read_excel_file(uploaded_file) -> Tuple[Optional[pd.DataFrame], str]:
    """
    Read an uploaded Excel file. Tries all sheets if first fails.
    Returns (DataFrame, status_message).
    """
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        if df.empty:
            return None, "The uploaded file appears to be empty."
        # Drop completely empty rows / columns
        df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
        return df, "ok"
    except Exception as exc:
        return None, f"Could not read file: {exc}"


def get_numeric_columns(df: pd.DataFrame) -> List[str]:
    """Return columns that are numeric or can be coerced to numeric."""
    candidates = []
    for col in df.columns:
        series = pd.to_numeric(df[col], errors="coerce")
        if series.notna().sum() > 0:
            candidates.append(col)
    return candidates


# ─── POPULATION PREPARATION ───────────────────────────────────────────────────

def prepare_population(
    df: pd.DataFrame,
    amount_col: str,
    abs_amounts: bool = True,
    exclude_zero: bool = True,
) -> pd.DataFrame:
    """
    Build the working population DataFrame.
    - Coerces the amount column to numeric
    - Optionally takes absolute values
    - Optionally excludes zero-amount rows
    - Adds internal '_amount' column
    """
    pop = df.copy()
    pop["_amount"] = pd.to_numeric(pop[amount_col], errors="coerce").fillna(0.0)
    if abs_amounts:
        pop["_amount"] = pop["_amount"].abs()
    if exclude_zero:
        pop = pop[pop["_amount"] > 0].reset_index(drop=True)
    else:
        pop = pop.reset_index(drop=True)
    pop["_original_row"] = pop.index + 2  # Excel row reference (header = row 1)
    return pop


def population_stats(pop: pd.DataFrame) -> Dict:
    """Return descriptive statistics for the population."""
    amounts = pop["_amount"]
    return {
        "count": len(pop),
        "total": amounts.sum(),
        "mean": amounts.mean(),
        "median": amounts.median(),
        "max": amounts.max(),
        "min": amounts.min(),
        "std": amounts.std(),
    }


# ─── STRATIFICATION ───────────────────────────────────────────────────────────

def apply_stratification(
    df: pd.DataFrame, thresholds: List[float]
) -> pd.DataFrame:
    """
    Divide the population into strata by ascending amount thresholds.
    Adds a '_stratum' column.
    Items above the highest threshold form the top stratum (individually
    significant items recommended for 100% coverage).
    """
    thresholds = sorted([float(t) for t in thresholds if t > 0])
    if not thresholds:
        return df

    df = df.copy()

    def _assign(amount: float) -> str:
        for i, thresh in enumerate(thresholds):
            if amount <= thresh:
                return f"Stratum {i + 1}  (≤ ₹{thresh:,.0f})"
        return f"Stratum {len(thresholds) + 1}  (> ₹{thresholds[-1]:,.0f})  [Individually Significant]"

    df["_stratum"] = df["_amount"].apply(_assign)
    return df


def stratification_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Return a summary table of strata."""
    summary = (
        df.groupby("_stratum", sort=False)["_amount"]
        .agg(
            Count="count",
            Total_Amount="sum",
            Min_Amount="min",
            Max_Amount="max",
            Avg_Amount="mean",
        )
        .reset_index()
    )
    summary.columns = [
        "Stratum", "Count", "Total Amount (₹)", "Min (₹)", "Max (₹)", "Avg (₹)"
    ]
    pop_total = df["_amount"].sum()
    summary["% of Population"] = (summary["Total Amount (₹)"] / pop_total * 100).round(1)
    return summary


# ─── SAMPLE SELECTION METHODS (Appendix II) ───────────────────────────────────

def select_random_sample(df: pd.DataFrame, n: int, seed: int = 42) -> pd.DataFrame:
    """
    Random selection: each sampling unit has an equal chance of selection.
    Uses a reproducible seed (computerised random number generation per Appendix II).
    """
    n = min(n, len(df))
    if n <= 0:
        return df.iloc[0:0]
    return df.sample(n=n, random_state=seed).sort_index().reset_index(drop=True)


def select_systematic_sample(df: pd.DataFrame, n: int, seed: int = 42) -> pd.DataFrame:
    """
    Systematic (every Nth) selection with a random start.
    Interval = population size ÷ sample size.
    Note: Verify no cyclical pattern exists in population before using.
    """
    n = min(n, len(df))
    if n <= 0:
        return df.iloc[0:0]

    interval = max(1, len(df) // n)
    rng = np.random.default_rng(seed)
    start = int(rng.integers(0, interval))
    indices = list(range(start, len(df), interval))[:n]
    return df.iloc[indices].reset_index(drop=True)


def select_mus_sample(df: pd.DataFrame, n: int, seed: int = 42) -> pd.DataFrame:
    """
    Monetary Unit Sampling (MUS / value-weighted selection).
    Probability of selection is proportional to the monetary amount,
    so larger-value items have a higher chance of selection.
    Recommended for TOD where the risk of overstatement is the primary concern.
    """
    n = min(n, len(df))
    if n <= 0 or df["_amount"].sum() == 0:
        return df.iloc[0:0]

    weights = df["_amount"] / df["_amount"].sum()
    rng = np.random.default_rng(seed)
    selected = rng.choice(len(df), size=n, replace=False, p=weights.values)
    return df.iloc[sorted(selected)].reset_index(drop=True)


def select_sample(
    df: pd.DataFrame,
    n: int,
    method: str,
    seed: int = 42,
) -> pd.DataFrame:
    """Dispatch to the appropriate selection method."""
    method_map = {
        "Random":                     select_random_sample,
        "Systematic (Every Nth)":     select_systematic_sample,
        "Monetary Unit Sampling (MUS)": select_mus_sample,
    }
    fn = method_map.get(method, select_random_sample)
    return fn(df, n, seed)


# ─── SAMPLE SUMMARY ───────────────────────────────────────────────────────────

def sample_summary(sample: pd.DataFrame, population_total: float) -> Dict:
    """Return headline statistics for a selected sample."""
    amounts = sample["_amount"]
    return {
        "n": len(sample),
        "total": amounts.sum(),
        "mean": amounts.mean() if len(amounts) > 0 else 0,
        "max": amounts.max() if len(amounts) > 0 else 0,
        "coverage_pct": amounts.sum() / population_total * 100 if population_total > 0 else 0,
    }
