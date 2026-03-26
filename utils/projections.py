"""
KKC Audit Sampling Tool – Projected Misstatement (Section 26.12 / 26.13)
Implements the tainting method for projecting misstatements found in a
sample to the full population, consistent with SA 530 and the KKC
Sampling Guide.

Tainting % = Misstatement ÷ Book Value of the sampled item
Projected Misstatement = Tainting % × Population Amount

Anomalous misstatements (Section 26.11) are excluded from projection
but are added back at their actual amount for the final conclusion.

Conclusion: Total (projected + anomalous) vs Tolerable Misstatement
"""

from __future__ import annotations
import pandas as pd
from typing import List, Dict


def calculate_projected_misstatement(
    results: List[Dict],
    population_amount: float,
    tolerable_misstatement: float,
) -> Dict:
    """
    Project misstatements from sample to population.

    Parameters
    ----------
    results : list of dicts, one per sampled item:
        {
          'ref'        : str   – reference / voucher number
          'narration'  : str   – description (optional)
          'book_value' : float – audited book amount
          'audit_value': float – corrected/audited amount
          'is_anomaly' : bool  – mark as anomalous if demonstrably not representative
        }
    population_amount     : float – total population (₹)
    tolerable_misstatement: float – performance materiality or specific TM (₹)

    Returns
    -------
    dict with keys:
        detail                : pd.DataFrame – item-level breakdown
        projected_misstatement: float
        anomalous_misstatement: float
        total_misstatement    : float
        tolerable_misstatement: float
        is_acceptable         : bool
        conclusion            : str
        items_with_errors     : int
        sample_size           : int
    """
    rows = []
    projected_total = 0.0
    anomalous_total = 0.0

    for item in results:
        book_val  = float(item.get("book_value",  0) or 0)
        audit_val = float(item.get("audit_value", book_val) or book_val)
        misstatement = round(book_val - audit_val, 2)

        if misstatement == 0:
            rows.append({
                "Ref / Voucher": item.get("ref", ""),
                "Narration":     item.get("narration", ""),
                "Book Value (₹)": book_val,
                "Audit Value (₹)": audit_val,
                "Misstatement (₹)": 0.0,
                "Tainting %":     "0.00%",
                "Projected Misstatement (₹)": 0.0,
                "Anomaly": "",
                "Remark": "No misstatement",
            })
            continue

        is_anomaly = bool(item.get("is_anomaly", False))

        if book_val != 0:
            tainting = misstatement / book_val
        else:
            tainting = 0.0

        if is_anomaly:
            projected = 0.0          # Not projected to population
            anomalous_total += misstatement
            remark = "Anomalous – not projected (actual amount included separately)"
        else:
            projected = tainting * population_amount
            projected_total += projected
            remark = "Projected to population"

        rows.append({
            "Ref / Voucher":              item.get("ref", ""),
            "Narration":                  item.get("narration", ""),
            "Book Value (₹)":             round(book_val, 2),
            "Audit Value (₹)":            round(audit_val, 2),
            "Misstatement (₹)":           round(misstatement, 2),
            "Tainting %":                 f"{tainting:.2%}",
            "Projected Misstatement (₹)": round(projected, 2),
            "Anomaly":                    "Yes" if is_anomaly else "No",
            "Remark":                     remark,
        })

    detail_df = pd.DataFrame(rows) if rows else pd.DataFrame()
    items_with_errors = sum(
        1 for r in rows if r.get("Misstatement (₹)", 0) != 0
    )

    total = round(projected_total + anomalous_total, 2)
    is_acceptable = total <= tolerable_misstatement

    if is_acceptable:
        conclusion = (
            "ACCEPTABLE – Projected misstatement (₹{:,.2f}) does not exceed "
            "Tolerable Misstatement (₹{:,.2f}). "
            "Sufficient appropriate audit evidence obtained for this population.".format(
                total, tolerable_misstatement
            )
        )
    else:
        conclusion = (
            "EXCEEDS TOLERABLE MISSTATEMENT – Total misstatement (₹{:,.2f}) "
            "exceeds Tolerable Misstatement (₹{:,.2f}). "
            "Further action required per Section 26.13 of KKC Sampling Guide.".format(
                total, tolerable_misstatement
            )
        )

    return {
        "detail":                  detail_df,
        "projected_misstatement":  round(projected_total, 2),
        "anomalous_misstatement":  round(anomalous_total, 2),
        "total_misstatement":      total,
        "tolerable_misstatement":  tolerable_misstatement,
        "is_acceptable":           is_acceptable,
        "conclusion":              conclusion,
        "items_with_errors":       items_with_errors,
        "sample_size":             len(results),
    }
