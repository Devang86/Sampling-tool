"""
KKC Audit Sampling Tool – Sampling Logic
Implements TOC minimum sample sizes (Appendix V) and TOD Poisson MUS
sample size calculation calibrated to the Inflo Sampling Methodology.

Reference: KKC Sampling Guide Section 26, Appendices III–VI; SA 530 – Audit Sampling
"""

import math
from typing import Tuple


# ─── TOC MINIMUM SAMPLE SIZE TABLE (Appendix V) ───────────────────────────────
# Quarterly note: 1+1 means 1 regular + 1 mandatory Q4 sample = 2 total
TOC_MIN_SAMPLE = {
    "Annual":                       {"Normal": 1,  "Elevated/Significant": 1},
    "Quarterly (Q4 mandatory +1)":  {"Normal": 2,  "Elevated/Significant": 2},
    "Monthly":                       {"Normal": 2,  "Elevated/Significant": 3},
    "Weekly":                        {"Normal": 5,  "Elevated/Significant": 8},
    "Daily":                         {"Normal": 15, "Elevated/Significant": 25},
    "Recurring (multiple/day)":      {"Normal": 25, "Elevated/Significant": 40},
}

TOC_NOTES = {
    "Annual":                       "",
    "Quarterly (Q4 mandatory +1)":  "Q4 (period-end) sample is mandatory as the +1.",
    "Monthly":                       "",
    "Weekly":                        "",
    "Daily":                         "",
    "Recurring (multiple/day)":      "",
}


# ─── TOD POISSON MUS PARAMETERS ───────────────────────────────────────────────
# Base confidence by (Assertion Risk, Entity Risk)
# Calibrated so that:
#   Significant / Normal baseline = 90% (matches Inflo Scenario 1)
BASE_CONFIDENCE: dict = {
    ("Significant", "High"):   0.95,
    ("Significant", "Normal"): 0.90,
    ("Significant", "Low"):    0.85,
    ("Elevated",    "High"):   0.80,
    ("Elevated",    "Normal"): 0.75,
    ("Elevated",    "Low"):    0.70,
    ("Normal",      "High"):   0.60,
    ("Normal",      "Normal"): 0.50,
    ("Normal",      "Low"):    0.40,
}

# Reduction multipliers – calibrated to all four Inflo scenarios exactly:
#   Scenario 1: Significant/Normal, No controls, No SADA, No SAP  → 90%, n=177
#   Scenario 2: Significant/Normal, High controls, No SADA, No SAP → 30%, n=26
#   Scenario 3: Significant/Normal, No controls, No SADA, High SAP → 55%, n=59
#   Scenario 4: Significant/Normal, No controls, High SADA, No SAP →  8%, n=6
CONTROLS_MULT: dict = {"None": 1.000, "Moderate": 0.700, "High": 0.333}
SADA_MULT:     dict = {"None": 1.000, "Low": 0.800, "Moderate": 0.500, "High": 0.089}
SAP_MULT:      dict = {"None": 1.000, "Low": 0.850, "Moderate": 0.700, "High": 0.611}

# Expected error rate uplifts applied to base sample size
EXPECTED_ERROR_UPLIFT: dict = {"None": 0.00, "Low": 0.10, "Medium": 0.25, "High": 0.50}


# ─── PUBLIC FUNCTIONS ─────────────────────────────────────────────────────────

def calculate_confidence_level(
    assertion_risk: str,
    entity_risk: str,
    controls_response: str,
    sada_response: str,
    sap_response: str,
) -> float:
    """
    Derive the required confidence level from risk parameters.
    Returns a value between 0.01 and 0.99.
    """
    base = BASE_CONFIDENCE.get((assertion_risk, entity_risk), 0.75)
    confidence = (
        base
        * CONTROLS_MULT.get(controls_response, 1.0)
        * SADA_MULT.get(sada_response, 1.0)
        * SAP_MULT.get(sap_response, 1.0)
    )
    return round(max(min(confidence, 0.99), 0.01), 4)


def calculate_tod_sample_size(
    population: float,
    performance_materiality: float,
    expected_error_rate: str,
    assertion_risk: str,
    entity_risk: str,
    controls_response: str,
    sada_response: str,
    sap_response: str,
) -> Tuple[int, float, float, float]:
    """
    Calculate TOD sample size using the Poisson MUS method.

    Formula: n = ceil( [-ln(1 - C) / TM_rate] × (1 + uplift) )
    where C = derived confidence level, TM_rate = PM / Population

    Returns:
        (sample_size, confidence_level, tm_rate, reliability_factor)
    """
    if population <= 0:
        raise ValueError("Population amount must be greater than zero.")
    if performance_materiality <= 0:
        raise ValueError("Performance Materiality must be greater than zero.")
    if performance_materiality >= population:
        raise ValueError("Performance Materiality cannot equal or exceed the Population amount.")

    confidence = calculate_confidence_level(
        assertion_risk, entity_risk, controls_response, sada_response, sap_response
    )

    tm_rate = performance_materiality / population
    reliability_factor = -math.log(1 - confidence)   # Poisson R factor
    base_n = reliability_factor / tm_rate

    uplift = EXPECTED_ERROR_UPLIFT.get(expected_error_rate, 0.0)
    # Round to 10 decimal places before ceiling to avoid floating-point
    # artefacts where a mathematically-exact integer (e.g. 177.000000000002)
    # is spuriously rounded up to the next integer.
    final_n = math.ceil(round(base_n * (1 + uplift), 10))

    return final_n, confidence, tm_rate, reliability_factor


def get_toc_min_sample(frequency: str, risk_level: str) -> int:
    """Return minimum TOC sample size per Appendix V."""
    return TOC_MIN_SAMPLE.get(frequency, {}).get(risk_level, 1)


def get_toc_note(frequency: str) -> str:
    """Return any special note for the selected frequency."""
    return TOC_NOTES.get(frequency, "")


def get_toc_table() -> dict:
    """Return the full TOC minimum sample size table."""
    return TOC_MIN_SAMPLE


def describe_confidence_drivers(
    assertion_risk: str,
    entity_risk: str,
    controls_response: str,
    sada_response: str,
    sap_response: str,
) -> list:
    """
    Return a list of human-readable strings explaining how each parameter
    affects the final confidence level. Used in the rationale display.
    """
    base = BASE_CONFIDENCE.get((assertion_risk, entity_risk), 0.75)
    drivers = [
        f"Base confidence ({assertion_risk} assertion risk × {entity_risk} entity risk): {base:.0%}",
    ]
    if controls_response != "None":
        mult = CONTROLS_MULT[controls_response]
        drivers.append(f"Controls Response ({controls_response}): ×{mult:.3f} reduction")
    if sada_response != "None":
        mult = SADA_MULT[sada_response]
        drivers.append(f"SADA Response ({sada_response}): ×{mult:.3f} reduction")
    if sap_response != "None":
        mult = SAP_MULT[sap_response]
        drivers.append(f"SAP Response ({sap_response}): ×{mult:.3f} reduction")
    return drivers
