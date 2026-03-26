# KKC Audit Sampling Tool

**KKC & Associates LLP — Internal Audit Tool**
Statutory Audit | Based on KKC Sampling Guide Section 26 & SA 530 (ICAI)

---

## What This Tool Does

This Streamlit application automates audit sampling for Articles and Managers conducting statutory audits. It implements:

- **Test of Details (TOD)** — Poisson MUS sample size calculation, calibrated to the Inflo Sampling Methodology (Appendix VI)
- **Test of Controls (TOC)** — Minimum sample sizes per Appendix V of the KKC Sampling Guide
- **Projected Misstatement** — Tainting method per Section 26.12–26.13
- **PDF Report** — KKC-branded work paper documentation
- **Excel Export** — Sample list with result-entry columns pre-built for Articles

---

## Setup Instructions

### Prerequisites
- Python 3.10 or later
- pip (Python package manager)

### Step 1 — Create a virtual environment (recommended)

```bash
python -m venv venv
```

Activate it:
- **Windows:**  `venv\Scripts\activate`
- **Mac / Linux:**  `source venv/bin/activate`

### Step 2 — Install dependencies

```bash
pip install -r requirements.txt
```

### Step 3 — Run the tool

```bash
streamlit run app.py
```

The tool will open automatically in your browser at `http://localhost:8501`.

---

## How to Use — Tab by Tab

### Tab 1: Upload & Map
1. Export the relevant ledger or transaction listing from the client's system (Tally / SAP / any ERP) as an Excel file.
2. Upload the file. Any format is accepted.
3. Map the columns: **Amount** (mandatory), Date, Voucher / Ref No, Narration (optional but recommended).
4. Enable Stratification if the population has a wide value range (recommended for large NBFC / bank portfolios).
5. Review the Population Summary statistics. These auto-populate into the Sampling Parameters tab.

### Tab 2: Sampling Parameters
**For TOD:**
- Enter the Performance Materiality set by the Engagement Partner for this area.
- Select all seven risk parameters. Hover over each field for detailed guidance on what each level means.
- The tool calculates the required confidence level and minimum sample size using the Poisson formula.
- A step-by-step calculation walkthrough is displayed below the results.

**For TOC:**
- Select the control frequency and assertion risk level.
- The minimum sample size from Appendix V is displayed immediately.

### Tab 3: Sample Selection
1. Review the required sample size floor (cannot be reduced without Partner approval).
2. Select the selection method — **MUS is recommended for TOD** on monetary populations.
3. If stratification was enabled, allocate samples across strata (proportional or manual).
4. Click **Generate Sample** to produce the sample list.
5. Download the Excel file — it includes pre-built result-entry columns for the Article to complete in the field.

### Tab 4: Results & Projection
1. For each sampled item, enter the audited/corrected amount.
2. If no error, leave the audit value equal to the book value.
3. Check **Anomaly?** only if the misstatement is demonstrably unique and not representative of the population (Section 26.11 requires a high degree of certainty — use this sparingly and document the basis).
4. Click **Calculate Projected Misstatement**.
5. Review the conclusion — acceptable or exceeds Tolerable Misstatement.

### Tab 5: Report & Export
1. Generate the **PDF Work Paper Report** — includes engagement details, all parameters, sample size rationale, and projected misstatement conclusion.
2. Optionally enable the **AI-generated narrative** (requires Anthropic API key entered in the sidebar). This produces a four-paragraph SA 530-style work paper explanation.
3. Generate the **Excel Export** — three sheets: Sampling Parameters, Selected Samples (with result-entry columns), Misstatement Projection.

---

## AI Narrative Feature

The tool can call the Anthropic Claude API to generate a professional work paper narrative explaining:
1. Why this sampling approach was appropriate for the engagement
2. How the sample size was derived (Poisson parameters, confidence level)
3. How sample items were selected and any stratification applied
4. Conclusion on audit sufficiency based on projection results

**To use this feature:**
- Enter your Anthropic API Key in the sidebar (starts with `sk-ant-...`)
- The key is never stored — it is used only for the current session
- The narrative is included in the PDF if the checkbox is selected

---

## Key Reference Standards

| Standard | Application |
|---|---|
| SA 530 – Audit Sampling (ICAI) | Overall audit sampling framework |
| KKC Sampling Guide Section 26 | Firm methodology, parameters, conclusions |
| KKC Sampling Guide Appendix II | Sample selection methods |
| KKC Sampling Guide Appendix V | TOC minimum sample sizes |
| KKC Sampling Guide Appendix VI | Inflo Poisson MUS methodology |
| SA 315 | Significant risk identification (for Assertion Risk parameter) |

---

## File Structure

```
kkc_sampling_tool/
│
├── app.py                      ← Main Streamlit application (5-tab UI)
├── requirements.txt            ← Python dependencies
├── README.md                   ← This file
│
└── utils/
    ├── __init__.py
    ├── sampling.py             ← TOC Appendix V table + TOD Poisson MUS calculator
    ├── file_handler.py         ← Excel upload, column mapping, stratification, sample selection
    ├── projections.py          ← Projected misstatement (tainting method)
    └── report_generator.py     ← PDF (ReportLab, KKC format) + AI narrative
```

---

## Inflo Scenario Calibration Verification

The TOD Poisson calculator is calibrated to all four Inflo reference scenarios:

| Scenario | Parameters | Expected Confidence | Expected n |
|---|---|---|---|
| 1 | Significant/Normal, No controls, No SADA, No SAP | 90% | 177 |
| 2 | Significant/Normal, High controls, No SADA, No SAP | ~30% | 26 |
| 3 | Significant/Normal, No controls, No SADA, High SAP | ~55% | 59 |
| 4 | Significant/Normal, No controls, High SADA, No SAP | ~8% | 6 |

---

## Notes for Engagement Team

- **Sample size floor is hard-enforced** — a warning is displayed if an Article attempts to select fewer items than the computed minimum. Any increase above the minimum requires Engagement Partner approval and documentation.
- **Anomaly classification (Section 26.11)** — requires a high degree of certainty that the misstatement is not representative. Document the basis clearly. Do not use anomaly as a workaround to avoid projection.
- **Seed number** — always document the random seed used for selection so the sample can be reproduced if required during review or quality inspection.
- **Population amount** — confirm with the client/Tally that the ledger balance agrees to the trial balance before finalising the population. Document any reconciling items excluded.

---

*KKC & Associates LLP | Chartered Accountants | Mumbai · Pune · Bengaluru · Ahmedabad*
*Internal use only — not for distribution*
