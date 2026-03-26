"""
KKC Audit Sampling Tool
-----------------------
A Streamlit application for KKC & Associates LLP implementing the
firm's sampling methodology (Section 26 of KKC Sampling Guide)
aligned with SA 530 – Audit Sampling (ICAI).

Tabs:
  1. Upload & Map       – Excel upload, column mapping, population stats
  2. Sampling Parameters – TOC / TOD / Both, risk parameters, sample size
  3. Sample Selection    – Method selection, sample generation, Excel export
  4. Results & Projection – Enter audit results, projected misstatement
  5. Report & Export     – PDF (KKC format) + Excel + AI narrative
"""

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import math
import io
from datetime import date, datetime

import numpy as np
import pandas as pd
import streamlit as st

from utils.sampling import (
    TOC_MIN_SAMPLE,
    calculate_tod_sample_size,
    describe_confidence_drivers,
    get_toc_min_sample,
    get_toc_note,
)
from utils.file_handler import (
    apply_stratification,
    get_numeric_columns,
    population_stats,
    prepare_population,
    read_excel_file,
    sample_summary,
    select_sample,
    stratification_summary,
)
from utils.projections import calculate_projected_misstatement
from utils.report_generator import generate_ai_narrative, generate_pdf_report


# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG  ← must be the very first Streamlit call
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="KKC Audit Sampling Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ─────────────────────────────────────────────────────────────────────────────
# KKC CUSTOM CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

    *, p, div, span, label, h1, h2, h3, h4, button {
        font-family: 'Source Sans Pro', sans-serif !important;
    }

    /* ── App header ── */
    .kkc-hdr   { background:#7CB542; color:#fff; padding:16px 22px 12px;
                  border-radius:6px 6px 0 0; margin-bottom:0; }
    .kkc-subhdr{ background:#808285; color:#fff; padding:7px 22px;
                  border-radius:0 0 6px 6px; margin-bottom:20px; font-size:13px; }
    .kkc-hdr h1{ margin:0; font-size:22px; font-weight:700; }
    .kkc-hdr p { margin:3px 0 0; font-size:12px; opacity:.9; }

    /* ── Metric cards ── */
    .mc-green { background:#EEF7E3; border-left:5px solid #7CB542;
                padding:12px 16px; border-radius:4px; margin:6px 0; }
    .mc-grey  { background:#F5F5F5; border-left:5px solid #808285;
                padding:12px 16px; border-radius:4px; margin:6px 0; }
    .mc-blue  { background:#E8F4F8; border-left:5px solid #17a2b8;
                padding:12px 16px; border-radius:4px; margin:6px 0; }
    .mc-red   { background:#FEE8EA; border-left:5px solid #dc3545;
                padding:12px 16px; border-radius:4px; margin:6px 0; }
    .mc-green h3,.mc-grey h3,.mc-blue h3,.mc-red h3
                { margin:0 0 3px; font-size:11px; color:#555; font-weight:600;
                  text-transform:uppercase; letter-spacing:.5px; }
    .mc-green h2,.mc-grey h2,.mc-blue h2,.mc-red h2
                { margin:0; font-size:20px; font-weight:700; color:#222; }

    /* ── Section titles ── */
    .sec-title { color:#7CB542; font-size:15px; font-weight:700;
                 border-bottom:2px solid #7CB542; padding-bottom:4px;
                 margin:18px 0 10px; }

    /* ── Info / alert boxes ── */
    .box-info  { background:#FFFBE6; border-left:4px solid #FFC107;
                 padding:10px 14px; border-radius:3px; font-size:13px; margin:8px 0; }
    .box-ok    { background:#D4EDDA; border-left:4px solid #28a745;
                 padding:10px 14px; border-radius:3px; font-size:13px; margin:8px 0; }
    .box-err   { background:#F8D7DA; border-left:4px solid #dc3545;
                 padding:10px 14px; border-radius:3px; font-size:13px; margin:8px 0; }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"]
                { gap:3px; border-bottom:2px solid #7CB542; }
    .stTabs [data-baseweb="tab"]
                { background:#EFEFEF; border-radius:5px 5px 0 0;
                  padding:7px 16px; color:#808285; font-weight:600; font-size:13px; }
    .stTabs [aria-selected="true"]
                { background:#7CB542 !important; color:#fff !important; }

    /* ── Primary button ── */
    div.stButton > button[kind="primary"]
                { background:#7CB542 !important; border:none !important;
                  color:#fff !important; font-weight:600; }
    div.stButton > button[kind="primary"]:hover
                { background:#6aa335 !important; }

    /* ── Sidebar ── */
    div[data-testid="stSidebar"] { background:#FAFAFA; }
    div[data-testid="stSidebar"] .stMarkdown h3
                { color:#7CB542; font-size:13px; margin-bottom:4px; }

    /* ── Footer ── */
    .kkc-footer { background:#808285; color:#fff; text-align:center;
                  padding:10px; font-size:11px; border-radius:4px; margin-top:30px; }

    /* ── Status checklist ── */
    .chk-done { color:#7CB542; font-weight:700; }
    .chk-todo { color:#808285; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────────────────────────────────────
# APP HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <div class="kkc-hdr">
      <h1>📊 KKC Audit Sampling Tool</h1>
      <p>Statutory Audit &nbsp;|&nbsp; Based on KKC Sampling Guide Section 26 &amp; SA 530 – Audit Sampling (ICAI)</p>
    </div>
    <div class="kkc-subhdr">
      Supports: Test of Controls (TOC – Appendix V) &nbsp;|&nbsp;
      Test of Details (TOD – Poisson MUS, calibrated to Inflo Methodology) &nbsp;|&nbsp;
      Projected Misstatement (Tainting Method)
    </div>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INITIALISATION
# ─────────────────────────────────────────────────────────────────────────────
_DEFAULTS = {
    "df_raw":           None,
    "df_population":    None,
    "col_map":          {},
    "pop_amount":       0.0,
    "strata_enabled":   False,
    "strata_thresholds":[],
    "sampling_type":    "Test of Details (TOD)",
    "tod_params":       {},
    "toc_params":       {},
    "tod_n":            0,
    "toc_n":            0,
    "sel_method":       "Random",
    "sel_seed":         42,
    "sample_df":        None,
    "results_list":     [],
    "projection":       None,
    "engagement":       {},
}

for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR – ENGAGEMENT DETAILS & API KEY
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Engagement Details")

    _client  = st.text_input("Client Name",      placeholder="e.g. Anand Rathi Wealth Ltd")
    _area    = st.text_input("Audit Area",        placeholder="e.g. Revenue – Trail Commission")
    _fy      = st.text_input("Financial Year",    value="FY 2024-25")
    _prep    = st.text_input("Prepared By",       placeholder="Article / CA Name")
    _rev     = st.text_input("Reviewed By",       placeholder="Partner / Manager")
    _dt      = st.date_input("Date",              value=date.today())

    st.session_state["engagement"] = {
        "client_name":    _client,
        "audit_area":     _area,
        "financial_year": _fy,
        "prepared_by":    _prep,
        "reviewed_by":    _rev,
        "date":           _dt.strftime("%d %B %Y"),
    }

    st.markdown("---")
    st.markdown("### 🔑 AI Narrative (Optional)")
    _api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        placeholder="sk-ant-...",
        help="Required only for the AI-generated work paper narrative in the PDF report. Not stored.",
    )

    st.markdown("---")
    st.markdown(
        """
        <div style="font-size:11px;color:#808285;">
        <b>Standards Reference</b><br>
        SA 530 – Audit Sampling (ICAI)<br>
        KKC Sampling Guide – Section 26<br>
        Appendices II, III, IV, V, VI<br><br>
        <b>Methodology</b><br>
        TOD: Poisson MUS (Inflo-calibrated)<br>
        TOC: Minimum table (Appendix V)
        </div>
        """,
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _metric(cls, label, value):
    st.markdown(
        f'<div class="{cls}"><h3>{label}</h3><h2>{value}</h2></div>',
        unsafe_allow_html=True,
    )


def _inr(x):
    return f"₹ {x:,.2f}"


# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📁  Upload & Map",
    "⚙️  Sampling Parameters",
    "🎯  Sample Selection",
    "📊  Results & Projection",
    "📄  Report & Export",
])


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 – UPLOAD & MAP
# ═════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="sec-title">Step 1 — Upload Ledger / Population Data</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload the client's Excel file (.xlsx / .xls)",
        type=["xlsx", "xls"],
        help="Any format accepted — Tally, SAP, ERP, or custom export. You will map the columns below.",
    )

    if uploaded:
        df_raw, msg = read_excel_file(uploaded)

        if df_raw is None:
            st.error(f"❌ {msg}")
        else:
            st.session_state["df_raw"] = df_raw

            st.markdown('<div class="sec-title">Preview — First 10 Rows</div>', unsafe_allow_html=True)
            st.dataframe(df_raw.head(10), use_container_width=True)

            c1, c2, c3 = st.columns(3)
            with c1: _metric("mc-green", "Rows loaded", f"{len(df_raw):,}")
            with c2: _metric("mc-grey",  "Columns",     f"{len(df_raw.columns)}")
            with c3: _metric("mc-grey",  "Sheet",       "Sheet 1")

            # ── Column mapping ────────────────────────────────────────────────
            st.markdown('<div class="sec-title">Step 2 — Map Columns</div>', unsafe_allow_html=True)
            st.markdown(
                '<div class="box-info">Map columns from your file to the fields below. '
                'Only <b>Amount</b> is mandatory. Articles should select the correct column '
                'for each field. Column names are taken as-is from the uploaded file.</div>',
                unsafe_allow_html=True,
            )

            all_opts  = ["— Not available —"] + list(df_raw.columns)
            num_opts  = ["— Not available —"] + get_numeric_columns(df_raw)

            mc1, mc2, mc3, mc4 = st.columns(4)
            with mc1:
                amt_col = st.selectbox("💰 Amount Column *", num_opts,
                                       help="Transaction / balance amounts (mandatory)")
            with mc2:
                dt_col  = st.selectbox("📅 Date Column",     all_opts)
            with mc3:
                ref_col = st.selectbox("🔢 Ref / Voucher No", all_opts)
            with mc4:
                nar_col = st.selectbox("📝 Narration",        all_opts)

            abs_flag = st.checkbox(
                "Take absolute values of amounts",
                value=True,
                help="Recommended when ledger has negative credits. Converts all amounts to positive.",
            )

            if amt_col != "— Not available —":
                _cm = {
                    "amount":   amt_col,
                    "date":     None if dt_col  == "— Not available —" else dt_col,
                    "ref":      None if ref_col == "— Not available —" else ref_col,
                    "narration":None if nar_col == "— Not available —" else nar_col,
                }
                st.session_state["col_map"] = _cm

                pop_df = prepare_population(df_raw, amt_col, abs_flag)
                st.session_state["df_population"] = pop_df
                st.session_state["pop_amount"]    = float(pop_df["_amount"].sum())

                # ── Population stats ──────────────────────────────────────────
                st.markdown('<div class="sec-title">Step 3 — Population Summary</div>', unsafe_allow_html=True)
                pstat = population_stats(pop_df)
                sc1, sc2, sc3, sc4 = st.columns(4)
                with sc1: _metric("mc-green", "Total Population",  _inr(pstat["total"]))
                with sc2: _metric("mc-grey",  "Transactions",      f"{pstat['count']:,}")
                with sc3: _metric("mc-grey",  "Avg Amount",        _inr(pstat["mean"]))
                with sc4: _metric("mc-grey",  "Largest Item",      _inr(pstat["max"]))

                # ── Stratification ────────────────────────────────────────────
                st.markdown('<div class="sec-title">Step 4 — Stratification (Optional)</div>', unsafe_allow_html=True)
                st.markdown(
                    '<div class="box-info">Stratification divides the population into sub-populations '
                    'to improve audit efficiency and allow greater focus on high-value items. '
                    'The top stratum (above the highest threshold) is typically subject to '
                    '100% examination. Reference: Appendix I of KKC Sampling Guide.</div>',
                    unsafe_allow_html=True,
                )

                enable_strat = st.checkbox("Enable Stratification")
                st.session_state["strata_enabled"] = enable_strat

                if enable_strat:
                    n_strata = st.slider("Number of strata", 2, 5, 2)
                    thresh_cols = st.columns(n_strata - 1)
                    thresholds  = []
                    pop_total   = pstat["total"]

                    for i, col in enumerate(thresh_cols):
                        default_t = pop_total * (i + 1) / n_strata
                        with col:
                            t = st.number_input(
                                f"Stratum {i+1} upper limit (₹)",
                                min_value=0.0,
                                max_value=float(pop_total),
                                value=round(default_t, -3),
                                format="%.2f",
                                key=f"thresh_{i}",
                            )
                            thresholds.append(t)

                    thresholds = sorted(thresholds)
                    st.session_state["strata_thresholds"] = thresholds
                    strat_df = apply_stratification(pop_df, thresholds)
                    st.session_state["df_population"] = strat_df

                    st.markdown("**Stratification Summary**")
                    st.dataframe(
                        stratification_summary(strat_df).style.format({
                            "Total Amount (₹)": "{:,.2f}",
                            "Min (₹)": "{:,.2f}",
                            "Max (₹)": "{:,.2f}",
                            "Avg (₹)": "{:,.2f}",
                            "% of Population": "{:.1f}%",
                        }),
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.session_state["strata_thresholds"] = []

                st.markdown(
                    '<div class="box-ok">✅ Population loaded. Proceed to <b>Sampling Parameters</b>.</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.warning("⚠️ Please select the Amount column to continue.")

    else:
        # Show example format
        st.markdown(
            '<div class="box-info">&#x1F446; Upload the client&#39;s Excel ledger or transaction listing. '
            'Any format is accepted &#8212; you will map the relevant columns in the next step.</div>',
            unsafe_allow_html=True,
        )
        sample_df = pd.DataFrame({
            "Date":       ["01-Apr-24", "05-Apr-24", "10-Apr-24"],
            "Voucher No": ["JV-001",    "JV-002",    "JV-003"],
            "Narration":  ["Trail commission income", "Advisory fees", "Trail fee receipt"],
            "Amount":     [150000, 75000, 230000],
        })
        st.caption("Example format (your file can have different column names):")
        st.dataframe(sample_df, hide_index=True, use_container_width=True)


# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 – SAMPLING PARAMETERS
# ═════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="sec-title">Sampling Type</div>', unsafe_allow_html=True)

    stype = st.radio(
        "Select what to calculate",
        ["Test of Details (TOD)", "Test of Controls (TOC)", "Both TOD and TOC"],
        horizontal=True,
    )
    st.session_state["sampling_type"] = stype

    # ── TOD ───────────────────────────────────────────────────────────────────
    if "TOD" in stype or "Both" in stype:
        st.markdown("---")
        st.markdown('<div class="sec-title">Test of Details – Risk Parameters</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="box-info">Sample size is calculated using the '
            '<b>Poisson Probability Distribution (MUS)</b> method, consistent with the '
            'Inflo Sampling Methodology (Appendix VI). Calibrated to all four Inflo '
            'reference scenarios.</div>',
            unsafe_allow_html=True,
        )

        ta, tb = st.columns(2)

        with ta:
            tod_pop = st.number_input(
                "Population Amount (₹) *",
                min_value=0.0,
                value=max(st.session_state["pop_amount"], 0.0),
                format="%.2f",
                key="tod_pop_input",
                help="Auto-populated from uploaded file. Override if testing a sub-population.",
            )
            # Keep session state in sync when manually overridden
            if tod_pop != st.session_state["pop_amount"] and st.session_state["df_raw"] is None:
                st.session_state["pop_amount"] = tod_pop

            tod_pm = st.number_input(
                "Performance Materiality (₹) *",
                min_value=0.0,
                value=0.0,
                format="%.2f",
                key="tod_pm_input",
                help="Enter the performance materiality set for this audit area.",
            )

            tod_ee = st.selectbox(
                "Expected Error Rate",
                ["None", "Low", "Medium", "High"],
                index=1,
                help=(
                    "None: No errors in prior years, automated controls\n"
                    "Low: Minor errors, controls operating effectively\n"
                    "Medium: Errors identified, partial control reliance\n"
                    "High: Significant/frequent errors, weak controls"
                ),
            )

        with tb:
            tod_er = st.selectbox(
                "Entity Level Risk",
                ["Low", "Normal", "High"],
                index=1,
                help=(
                    "Low: Strong governance, robust controls, no history of misstatements\n"
                    "Normal: Generally effective controls, stable operations\n"
                    "High: Weak governance, deficiencies, history of misstatements/fraud"
                ),
            )

            tod_ar = st.selectbox(
                "Assertion Risk",
                ["Normal", "Elevated", "Significant"],
                index=2,
                help=(
                    "Normal: Controls effective, no prior misstatements\n"
                    "Elevated: Some deficiencies, moderate judgment\n"
                    "Significant: High fraud risk, complex estimates, SA 315 significant risk"
                ),
            )

            tod_cr = st.selectbox(
                "Controls Response (TOC performed?)",
                ["None", "Moderate", "High"],
                index=0,
                help=(
                    "None: No TOC performed / controls ineffective\n"
                    "Moderate: Controls tested, limited samples, minor deficiencies\n"
                    "High: Controls well-designed, operating effectively throughout period"
                ),
            )

        tc, td = st.columns(2)
        with tc:
            tod_sada = st.selectbox(
                "SADA Response",
                ["None", "Low", "Moderate", "High"],
                index=0,
                help=(
                    "None: No SADA performed\n"
                    "Low: Partial coverage, trend only\n"
                    "Moderate: Substantial coverage, exceptions followed up\n"
                    "High: 100% population, assertion-focused, directly addresses risk\n"
                    "(SADA has the most powerful sample reduction effect)"
                ),
            )
        with td:
            tod_sap = st.selectbox(
                "SAP Response",
                ["None", "Low", "Moderate", "High"],
                index=0,
                help=(
                    "None: No SAP performed\n"
                    "Low: High-level, prior-year based\n"
                    "Moderate: Independent expectation developed, variances resolved\n"
                    "High: Highly precise, all variances investigated, clear conclusion"
                ),
            )

        # ── Compute ──────────────────────────────────────────────────────────
        valid_tod = (tod_pop > 0) and (tod_pm > 0) and (tod_pm < tod_pop)

        if valid_tod:
            try:
                n, conf, tm_rate, rf = calculate_tod_sample_size(
                    tod_pop, tod_pm, tod_ee, tod_ar, tod_er, tod_cr, tod_sada, tod_sap
                )
                st.session_state["tod_n"] = n
                st.session_state["tod_params"] = {
                    "population": tod_pop, "pm": tod_pm,
                    "expected_error": tod_ee, "entity_risk": tod_er,
                    "assertion_risk": tod_ar, "controls_response": tod_cr,
                    "sada_response": tod_sada, "sap_response": tod_sap,
                    "sample_size": n, "confidence": conf,
                    "tm_rate": tm_rate, "reliability_factor": rf,
                }

                st.markdown("---")
                st.markdown("**📐 TOD Sample Size Result**")
                rm1, rm2, rm3, rm4 = st.columns(4)
                with rm1: _metric("mc-green", "Required Sample Size", str(n))
                with rm2: _metric("mc-grey",  "Confidence Level",     f"{conf:.0%}")
                with rm3: _metric("mc-grey",  "TM Rate",              f"{tm_rate:.3%}")
                with rm4: _metric("mc-grey",  "Reliability Factor (R)", f"{rf:.4f}")

                # Calculation walkthrough
                base_n = rf / tm_rate
                from utils.sampling import EXPECTED_ERROR_UPLIFT
                uplift = EXPECTED_ERROR_UPLIFT.get(tod_ee, 0)
                st.markdown(
                    f'<div class="box-ok">'
                    f'<b>Calculation walkthrough</b> (Poisson MUS, Inflo-calibrated):<br>'
                    f'R = –ln(1 – {conf:.0%}) = {rf:.4f} &nbsp;|&nbsp; '
                    f'TM Rate = {_inr(tod_pm)} ÷ {_inr(tod_pop)} = {tm_rate:.4%}<br>'
                    f'Base n = R ÷ TM Rate = {base_n:.1f} &nbsp;|&nbsp; '
                    f'Expected error uplift ({tod_ee}) = +{uplift:.0%}<br>'
                    f'<b>Final n = ceil({base_n:.1f} × {1+uplift:.2f}) = {n}</b>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

                # Confidence drivers
                with st.expander("📋 Confidence Level Drivers"):
                    for d in describe_confidence_drivers(tod_ar, tod_er, tod_cr, tod_sada, tod_sap):
                        st.write(f"• {d}")

            except ValueError as e:
                st.error(f"❌ {e}")

        elif tod_pm >= tod_pop and tod_pm > 0:
            st.error("❌ Performance Materiality cannot equal or exceed the Population amount.")
        else:
            st.markdown(
                '<div class="box-info">ℹ️ Enter Population Amount and Performance Materiality to compute sample size.</div>',
                unsafe_allow_html=True,
            )

    # ── TOC ───────────────────────────────────────────────────────────────────
    if "TOC" in stype or "Both" in stype:
        st.markdown("---")
        st.markdown('<div class="sec-title">Test of Controls – Minimum Sample Size (Appendix V)</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="box-info">Minimum sample sizes are mandatory per Appendix V of the KKC Sampling Guide. '
            'The engagement team cannot select fewer samples than these minimums.</div>',
            unsafe_allow_html=True,
        )

        tc1, tc2 = st.columns(2)
        with tc1:
            toc_freq = st.selectbox(
                "Control Frequency",
                list(TOC_MIN_SAMPLE.keys()),
                help="How frequently does this control operate during the audit period?",
            )
        with tc2:
            toc_risk = st.selectbox(
                "Assertion Risk Level",
                ["Normal", "Elevated/Significant"],
                help=(
                    "Normal: Controls designed and operating, no significant prior issues\n"
                    "Elevated/Significant: Some deficiencies, or control area is a significant risk"
                ),
            )

        toc_min = get_toc_min_sample(toc_freq, toc_risk)
        toc_note = get_toc_note(toc_freq)

        st.session_state["toc_n"] = toc_min
        st.session_state["toc_params"] = {
            "frequency": toc_freq,
            "risk_level": toc_risk,
            "min_sample": toc_min,
        }

        _metric("mc-green", "Minimum TOC Sample Size (Appendix V)", str(toc_min))
        if toc_note:
            st.markdown(f'<div class="box-info">📌 {toc_note}</div>', unsafe_allow_html=True)

        with st.expander("📋 Full Minimum Sample Size Table – Appendix V"):
            toc_rows = [
                {
                    "Control Frequency":        freq,
                    "Normal Risk":              v["Normal"],
                    "Elevated / Significant Risk": v["Elevated/Significant"],
                }
                for freq, v in TOC_MIN_SAMPLE.items()
            ]
            st.dataframe(pd.DataFrame(toc_rows), use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════════════════════════════════════
# TAB 3 – SAMPLE SELECTION
# ═════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="sec-title">Sample Selection</div>', unsafe_allow_html=True)

    pop_df = st.session_state.get("df_population")

    if pop_df is None or len(pop_df) == 0:
        st.warning("⚠️ Please complete the Upload & Map step first.")
        st.stop()

    stype   = st.session_state.get("sampling_type", "Test of Details (TOD)")
    tod_n   = st.session_state.get("tod_n", 0)
    toc_n   = st.session_state.get("toc_n", 0)
    min_req = max(tod_n, toc_n)

    # ── Required sizes ────────────────────────────────────────────────────────
    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        if tod_n > 0:
            _metric("mc-green", "TOD Required n", str(tod_n))
    with sc2:
        if toc_n > 0:
            _metric("mc-grey", "TOC Minimum n", str(toc_n))
    with sc3:
        if min_req > 0:
            _metric("mc-blue", "Floor (max of above)", str(min_req))

    # ── Final sample size ─────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">Final Sample Size</div>', unsafe_allow_html=True)

    default_n = max(min_req, 1) if min_req > 0 else 25
    final_n   = st.number_input(
        "Sample size to select",
        min_value=1,
        max_value=len(pop_df),
        value=min(default_n, len(pop_df)),
        help=(
            "Cannot fall below the computed minimum. "
            "Engagement Partner approval required for any increase above the minimum."
        ),
    )

    if min_req > 0 and final_n < min_req:
        st.markdown(
            f'<div class="box-err">❌ <b>Sample size below minimum.</b> '
            f'You entered {final_n} but the minimum is {min_req}. '
            f'As per KKC Sampling Guide, sample size cannot be reduced below the computed minimum.</div>',
            unsafe_allow_html=True,
        )
    elif min_req > 0 and final_n > min_req:
        st.markdown(
            f'<div class="box-info">ℹ️ Sample size increased to {final_n} '
            f'(above floor of {min_req}). Ensure Engagement Partner approval is documented.</div>',
            unsafe_allow_html=True,
        )

    # ── Selection method ──────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">Selection Method (Appendix II)</div>', unsafe_allow_html=True)

    method_opts = {
        "Random":                       "Each item has an equal, known probability of selection. "
                                        "Uses computerised random number generation. Most common for statistical sampling.",
        "Systematic (Every Nth)":       "Select every Nth item after a random start. "
                                        "Efficient for large ordered populations. "
                                        "Verify no cyclical patterns exist before using.",
        "Monetary Unit Sampling (MUS)": "Probability proportional to monetary value. "
                                        "Larger transactions have a higher chance of selection. "
                                        "Recommended for TOD – existence/valuation of large balances.",
    }

    sel_method = st.radio(
        "Principal Method",
        list(method_opts.keys()),
        horizontal=True,
    )
    st.markdown(f'<div class="box-info">ℹ️ {method_opts[sel_method]}</div>', unsafe_allow_html=True)
    st.session_state["sel_method"] = sel_method

    sel_seed = st.number_input(
        "Random Seed",
        min_value=1, max_value=99999, value=42,
        help="Fixes randomisation for reproducibility. Document in work papers.",
    )
    st.session_state["sel_seed"] = int(sel_seed)

    # ── Stratified allocation ─────────────────────────────────────────────────
    strat_alloc = {}
    if st.session_state.get("strata_enabled") and "_stratum" in pop_df.columns:
        st.markdown('<div class="sec-title">Sample Allocation Across Strata</div>', unsafe_allow_html=True)

        alloc_method = st.radio(
            "Allocation method",
            ["Proportional (auto)", "Manual per stratum"],
            horizontal=True,
        )

        strata_list  = sorted(pop_df["_stratum"].unique())
        strata_cnt   = pop_df["_stratum"].value_counts()

        if alloc_method == "Proportional (auto)":
            remaining = final_n
            for i, s in enumerate(strata_list):
                cnt = strata_cnt.get(s, 0)
                prop = cnt / len(pop_df)
                alloc = max(1, round(final_n * prop)) if cnt > 0 else 0
                alloc = min(alloc, cnt)
                strat_alloc[s] = alloc
        else:
            cols_s = st.columns(min(len(strata_list), 3))
            for i, s in enumerate(strata_list):
                cnt = strata_cnt.get(s, 0)
                with cols_s[i % 3]:
                    alloc = st.number_input(
                        f"n for {s[:35]}",
                        min_value=0, max_value=int(cnt),
                        value=min(max(1, final_n // len(strata_list)), cnt),
                        key=f"strat_n_{i}",
                    )
                    strat_alloc[s] = int(alloc)

        st.markdown(
            f'**Total samples across strata: {sum(strat_alloc.values())} '
            f'(target: {final_n})**'
        )

    # ── Generate ──────────────────────────────────────────────────────────────
    st.markdown("---")

    if st.button("🎯 Generate Sample", type="primary"):

        try:
            _col_map = st.session_state.get("col_map", {})

            if st.session_state.get("strata_enabled") and "_stratum" in pop_df.columns and strat_alloc:
                # Stratified sampling
                parts = []
                for s, n_s in strat_alloc.items():
                    stratum_pop = pop_df[pop_df["_stratum"] == s].copy()
                    if n_s > 0 and len(stratum_pop) > 0:
                        parts.append(select_sample(stratum_pop, n_s, sel_method, int(sel_seed)))
                sample_df = (
                    pd.concat(parts).reset_index(drop=True) if parts else pop_df.iloc[:0]
                )
            else:
                sample_df = select_sample(pop_df, final_n, sel_method, int(sel_seed))

            sample_df.index = range(1, len(sample_df) + 1)
            st.session_state["sample_df"] = sample_df

            # ── Build display columns ─────────────────────────────────────────
            disp_cols, col_rename = ["_amount"], {"_amount": "Amount (₹)"}
            for fld, lbl in [("date", "Date"), ("ref", "Ref / Voucher No"), ("narration", "Narration")]:
                c = _col_map.get(fld)
                if c and c in sample_df.columns:
                    if fld == "narration":
                        disp_cols.append(c)
                    else:
                        disp_cols.insert(0, c)
                    col_rename[c] = lbl
            if "_stratum" in sample_df.columns:
                disp_cols.append("_stratum")
                col_rename["_stratum"] = "Stratum"

            disp_df = (
                sample_df[[c for c in disp_cols if c in sample_df.columns]]
                .copy()
                .rename(columns=col_rename)
            )

            # ── Stats ─────────────────────────────────────────────────────────
            stats = sample_summary(sample_df, st.session_state["pop_amount"])
            ss1, ss2, ss3, ss4 = st.columns(4)
            with ss1: _metric("mc-green", "Samples selected",   str(stats["n"]))
            with ss2: _metric("mc-grey",  "Sample amount",      _inr(stats["total"]))
            with ss3: _metric("mc-grey",  "Avg item size",      _inr(stats["mean"]))
            with ss4: _metric("mc-grey",  "Population coverage",f"{stats['coverage_pct']:.1f}%")

            st.markdown("**Selected Sample Items**")
            disp_df_show = disp_df.copy()
            disp_df_show["Amount (₹)"] = disp_df_show["Amount (₹)"].apply(lambda x: f"{x:,.2f}")
            st.dataframe(disp_df_show, use_container_width=True)

            st.markdown(
                '<div class="box-ok">✅ Sample generated. Proceed to <b>Results & Projection</b> to enter audit findings.</div>',
                unsafe_allow_html=True,
            )

            # ── Excel download ─────────────────────────────────────────────────
            xls_buf = io.BytesIO()
            with pd.ExcelWriter(xls_buf, engine="xlsxwriter") as writer:
                wb  = writer.book
                hdr = wb.add_format({"bold": True, "bg_color": "#7CB542", "font_color": "white",
                                      "border": 1, "font_name": "Arial", "font_size": 10})
                dat = wb.add_format({"border": 1, "font_name": "Arial", "font_size": 9})
                amt = wb.add_format({"border": 1, "font_name": "Arial", "font_size": 9,
                                      "num_format": "#,##0.00"})

                disp_df.to_excel(writer, sheet_name="Selected Samples", startrow=1, index=True)
                ws = writer.sheets["Selected Samples"]
                ws.write(0, 0,
                         f"KKC Audit Sampling – {st.session_state['engagement'].get('client_name','')} "
                         f"| {st.session_state['engagement'].get('audit_area','')} | "
                         f"{st.session_state['engagement'].get('financial_year','')}",
                         wb.add_format({"bold": True, "font_color": "#7CB542",
                                         "font_name": "Arial", "font_size": 11}))

                # Result-entry columns for article use
                ncols = len(disp_df.columns) + 1
                for ci, lbl in enumerate(["Audited Value (₹)", "Misstatement (₹)",
                                          "Tainting %", "Anomaly? (Y/N)", "Remarks"]):
                    ws.write(1, ncols + ci, lbl, hdr)
                    ws.set_column(ncols + ci, ncols + ci, 18)

                for ci in range(len(disp_df.columns) + 1):
                    ws.set_column(ci, ci, 20)

            client_slug = st.session_state["engagement"].get("client_name", "Client").replace(" ", "_")
            st.download_button(
                label="⬇️ Download Sample List (Excel)",
                data=xls_buf.getvalue(),
                file_name=f"KKC_Sample_{client_slug}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as exc:
            st.error(f"❌ Error generating sample: {exc}")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 4 – RESULTS & PROJECTION
# ═════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="sec-title">Enter Audit Results per Sampled Item</div>', unsafe_allow_html=True)

    sample_df = st.session_state.get("sample_df")

    if sample_df is None:
        st.warning("⚠️ Please generate a sample in the Sample Selection tab first.")
        st.stop()

    st.markdown(
        '<div class="box-info">For each sampled item, enter the <b>Audited / Corrected Value</b>. '
        'If the book value is correct, leave the audit value unchanged (= book value). '
        'Check <b>Anomaly?</b> only if the misstatement is demonstrably not representative '
        'of the population (Section 26.11 of KKC Sampling Guide — requires high degree of certainty).'
        '</div>',
        unsafe_allow_html=True,
    )

    _col_map = st.session_state.get("col_map", {})
    tod_p    = st.session_state.get("tod_params", {})
    pop_amt  = st.session_state.get("pop_amount", 0.0)
    pm_val   = tod_p.get("pm", 0.0)

    if pm_val == 0:
        pm_val = st.number_input(
            "Performance Materiality (₹) — required for projection",
            min_value=0.0, format="%.2f", key="pm_proj_override",
        )

    st.markdown("---")
    # Table header
    hcols = st.columns([0.5, 2.5, 1.5, 1.5, 1])
    hcols[0].markdown("**Sr.**")
    hcols[1].markdown("**Reference / Narration**")
    hcols[2].markdown("**Book Value (₹)**")
    hcols[3].markdown("**Audit Value (₹)**")
    hcols[4].markdown("**Anomaly?**")

    results_list = []

    for idx in sample_df.index:
        row      = sample_df.loc[idx]
        book_val = float(row["_amount"])
        ref_val  = str(row[_col_map["ref"]]) if _col_map.get("ref") and _col_map["ref"] in row else f"Item {idx}"
        nar_val  = str(row[_col_map["narration"]])[:45] if _col_map.get("narration") and _col_map["narration"] in row else ""
        label    = f"{ref_val}  {('— ' + nar_val) if nar_val else ''}"

        rc = st.columns([0.5, 2.5, 1.5, 1.5, 1])
        rc[0].markdown(f"**{idx}**")
        rc[1].markdown(f"<small>{label}</small>", unsafe_allow_html=True)
        rc[2].markdown(f"₹ {book_val:,.2f}")

        audit_val = rc[3].number_input(
            f"Audit value {idx}",
            min_value=0.0,
            value=book_val,
            format="%.2f",
            label_visibility="collapsed",
            key=f"av_{idx}",
        )
        anomaly = rc[4].checkbox("", key=f"an_{idx}",
                                  help="Section 26.11: Anomaly — demonstrably not representative")

        results_list.append({
            "ref":        ref_val,
            "narration":  nar_val,
            "book_value": book_val,
            "audit_value": audit_val,
            "is_anomaly": anomaly,
        })

    st.markdown("---")

    if st.button("📊 Calculate Projected Misstatement", type="primary"):
        if pm_val <= 0:
            st.error("❌ Please enter a valid Performance Materiality before projecting.")
        else:
            proj = calculate_projected_misstatement(results_list, pop_amt, pm_val)
            st.session_state["projection"]  = proj
            st.session_state["results_list"] = results_list

            # Summary metrics
            is_ok = proj["is_acceptable"]
            pm1, pm2, pm3, pm4 = st.columns(4)
            with pm1: _metric("mc-grey",                     "Projected Misstatement", _inr(proj["projected_misstatement"]))
            with pm2: _metric("mc-grey",                     "Anomalous Misstatement",  _inr(proj["anomalous_misstatement"]))
            with pm3: _metric("mc-green" if is_ok else "mc-red",
                               "Total vs TM",
                               f'{_inr(proj["total_misstatement"])} vs {_inr(proj["tolerable_misstatement"])}')
            with pm4: _metric("mc-green" if is_ok else "mc-red",
                               "Conclusion",
                               "✅ ACCEPTABLE" if is_ok else "❌ EXCEEDS TM")

            if not proj["detail"].empty:
                st.markdown("**Item-Level Misstatement Detail**")
                st.dataframe(
                    proj["detail"].style.format({
                        "Book Value (₹)":              "{:,.2f}",
                        "Audit Value (₹)":             "{:,.2f}",
                        "Misstatement (₹)":            "{:,.2f}",
                        "Projected Misstatement (₹)":  "{:,.2f}",
                    }),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.markdown(
                    '<div class="box-ok">✅ No misstatements identified in the selected sample. '
                    'Projected misstatement = Nil.</div>',
                    unsafe_allow_html=True,
                )

            # Conclusion box
            conc_cls = "box-ok" if is_ok else "box-err"
            st.markdown(
                f'<div class="{conc_cls}">{proj["conclusion"]}</div>',
                unsafe_allow_html=True,
            )

            if not is_ok:
                st.markdown(
                    '<div class="box-info"><b>Next steps (Section 26.13 of KKC Sampling Guide):</b><br>'
                    '(i) Request management to investigate the misstatements identified and the '
                    'potential for further misstatements, and to make any necessary adjustments; or<br>'
                    '(ii) Extend sample size / test an alternative control / modify related substantive '
                    'procedures to obtain additional audit evidence.</div>',
                    unsafe_allow_html=True,
                )


# ═════════════════════════════════════════════════════════════════════════════
# TAB 5 – REPORT & EXPORT
# ═════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="sec-title">Generate Report & Export</div>', unsafe_allow_html=True)

    col_pdf, col_xls = st.columns(2)

    # ── PDF ───────────────────────────────────────────────────────────────────
    with col_pdf:
        st.markdown("**📄 PDF Work Paper Report (KKC Format)**")
        st.caption("Includes: engagement details, sample size rationale, parameters table, projected misstatement, and optional AI narrative.")

        incl_ai = st.checkbox(
            "Include AI-generated work paper narrative",
            value=bool(_api_key),
            help="Requires Anthropic API Key (entered in sidebar). Generates a SA 530-style narrative for documentation.",
        )

        if st.button("Generate PDF Report", type="primary", key="pdf_btn"):
            # Compile report_data dict
            rdata = {
                "engagement": st.session_state.get("engagement", {}),
                "sampling": {
                    "sampling_type":    st.session_state.get("sampling_type", "—"),
                    "population":       st.session_state.get("pop_amount", 0),
                    "selection_method": st.session_state.get("sel_method", "—"),
                },
            }
            if st.session_state.get("tod_params"):
                rdata["tod"] = st.session_state["tod_params"]
            if st.session_state.get("toc_params"):
                rdata["toc"] = st.session_state["toc_params"]

            proj = st.session_state.get("projection")
            if proj:
                rdata["projection"] = {
                    "projected": proj["projected_misstatement"],
                    "anomalous": proj["anomalous_misstatement"],
                    "total":     proj["total_misstatement"],
                    "tolerable": proj["tolerable_misstatement"],
                    "conclusion": proj["conclusion"],
                    "is_acceptable": proj["is_acceptable"],
                }

            if incl_ai and _api_key:
                with st.spinner("Generating AI narrative via Claude API…"):
                    narrative = generate_ai_narrative(rdata, _api_key)
                rdata["ai_narrative"] = narrative
            elif incl_ai and not _api_key:
                st.warning("⚠️ Enter your Anthropic API Key in the sidebar to generate the AI narrative.")

            try:
                pdf_bytes = generate_pdf_report(rdata)
                slug = st.session_state["engagement"].get("client_name", "Client").replace(" ", "_")
                st.download_button(
                    label="⬇️ Download PDF Report",
                    data=pdf_bytes,
                    file_name=f"KKC_Sampling_Report_{slug}_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    key="pdf_dl",
                )
                st.markdown(
                    '<div class="box-ok">✅ PDF generated successfully.</div>',
                    unsafe_allow_html=True,
                )
            except Exception as exc:
                st.error(f"❌ PDF generation failed: {exc}")

    # ── Excel ─────────────────────────────────────────────────────────────────
    with col_xls:
        st.markdown("**📊 Excel Export**")
        st.caption("Includes: Sampling Parameters sheet, Selected Samples sheet with result-entry columns, Projection sheet.")

        if st.button("Generate Excel Export", key="xls_btn"):
            xls_out = io.BytesIO()
            sample_df_x = st.session_state.get("sample_df")

            with pd.ExcelWriter(xls_out, engine="xlsxwriter") as writer:
                wb = writer.book
                grn_hdr = wb.add_format({"bold": True, "bg_color": "#7CB542", "font_color": "white",
                                          "border": 1, "font_name": "Arial", "font_size": 10})
                gry_hdr = wb.add_format({"bold": True, "bg_color": "#808285", "font_color": "white",
                                          "border": 1, "font_name": "Arial", "font_size": 10})
                dat_fmt = wb.add_format({"border": 1, "font_name": "Arial", "font_size": 9})
                amt_fmt = wb.add_format({"border": 1, "font_name": "Arial", "font_size": 9,
                                          "num_format": "\"₹\"#,##0.00"})
                title_fmt = wb.add_format({"bold": True, "font_size": 13,
                                            "font_color": "#7CB542", "font_name": "Arial"})

                # ── Sheet 1: Parameters ───────────────────────────────────────
                ps = wb.add_worksheet("Sampling Parameters")
                ps.write(0, 0, "KKC & Associates LLP – Audit Sampling Documentation", title_fmt)
                ps.set_column(0, 0, 35)
                ps.set_column(1, 1, 30)

                row = 2
                eng = st.session_state.get("engagement", {})
                params = [
                    ("Client Name",       eng.get("client_name", "—")),
                    ("Audit Area",        eng.get("audit_area", "—")),
                    ("Financial Year",    eng.get("financial_year", "—")),
                    ("Prepared By",       eng.get("prepared_by", "—")),
                    ("Reviewed By",       eng.get("reviewed_by", "—")),
                    ("Date",              eng.get("date", "—")),
                    ("", ""),
                    ("Sampling Type",     st.session_state.get("sampling_type", "—")),
                    ("Population Amount", st.session_state.get("pop_amount", 0)),
                    ("Selection Method",  st.session_state.get("sel_method", "—")),
                ]

                tod_p = st.session_state.get("tod_params", {})
                if tod_p:
                    params += [
                        ("", ""),
                        ("── TOD Parameters ──", ""),
                        ("Performance Materiality", tod_p.get("pm", 0)),
                        ("Entity Risk",             tod_p.get("entity_risk", "—")),
                        ("Assertion Risk",          tod_p.get("assertion_risk", "—")),
                        ("Expected Error Rate",     tod_p.get("expected_error", "—")),
                        ("Controls Response",       tod_p.get("controls_response", "—")),
                        ("SADA Response",           tod_p.get("sada_response", "—")),
                        ("SAP Response",            tod_p.get("sap_response", "—")),
                        ("Confidence Level",        f"{tod_p.get('confidence', 0):.0%}"),
                        ("Reliability Factor (R)",  f"{tod_p.get('reliability_factor', 0):.4f}"),
                        ("Calculated Sample Size",  tod_p.get("sample_size", "—")),
                    ]

                toc_p = st.session_state.get("toc_params", {})
                if toc_p:
                    params += [
                        ("", ""),
                        ("── TOC Parameters ──", ""),
                        ("Control Frequency",   toc_p.get("frequency", "—")),
                        ("Assertion Risk Level",toc_p.get("risk_level", "—")),
                        ("Minimum Sample Size", toc_p.get("min_sample", "—")),
                    ]

                for lbl, val in params:
                    if lbl:
                        ps.write(row, 0, lbl, gry_hdr)
                        ps.write(row, 1, val, dat_fmt)
                    row += 1

                # ── Sheet 2: Selected Samples ─────────────────────────────────
                if sample_df_x is not None:
                    _cm = st.session_state.get("col_map", {})
                    exp_cols, rnm = ["_amount"], {"_amount": "Amount (₹)"}
                    for fld, lbl in [("date","Date"), ("ref","Ref/Voucher"), ("narration","Narration")]:
                        c = _cm.get(fld)
                        if c and c in sample_df_x.columns:
                            exp_cols = ([c] + exp_cols) if fld != "narration" else (exp_cols + [c])
                            rnm[c] = lbl
                    if "_stratum" in sample_df_x.columns:
                        exp_cols.append("_stratum")
                        rnm["_stratum"] = "Stratum"

                    samp_export = (
                        sample_df_x[[c for c in exp_cols if c in sample_df_x.columns]]
                        .copy().rename(columns=rnm)
                    )
                    samp_export.index.name = "Sr. No."
                    samp_export.to_excel(writer, sheet_name="Selected Samples", startrow=1)

                    sw = writer.sheets["Selected Samples"]
                    sw.write(0, 0,
                             f"Selected Samples — {eng.get('client_name','')} | "
                             f"{eng.get('audit_area','')} | {eng.get('financial_year','')}",
                             title_fmt)

                    # Add result-entry columns
                    nc = len(samp_export.columns) + 1
                    for ci, hd in enumerate(["Audited Value (₹)", "Misstatement (₹)",
                                              "Tainting %", "Anomaly? (Y/N)", "Remarks"]):
                        sw.write(1, nc + ci, hd, grn_hdr)
                    for ci in range(nc + 5):
                        sw.set_column(ci, ci, 20)

                # ── Sheet 3: Projection ───────────────────────────────────────
                proj = st.session_state.get("projection")
                if proj and not proj["detail"].empty:
                    proj["detail"].to_excel(writer, sheet_name="Misstatement Projection",
                                             index=False, startrow=1)
                    pw = writer.sheets["Misstatement Projection"]
                    pw.write(0, 0, "Projected Misstatement – Tainting Method", title_fmt)

                    lr = len(proj["detail"]) + 3
                    summ = [
                        ("Projected Misstatement (₹)",             proj["projected_misstatement"]),
                        ("Anomalous Misstatement (₹)",             proj["anomalous_misstatement"]),
                        ("Total Misstatement (₹)",                 proj["total_misstatement"]),
                        ("Tolerable Misstatement (₹)",             proj["tolerable_misstatement"]),
                        ("Conclusion",                             proj["conclusion"]),
                    ]
                    for si, (l, v) in enumerate(summ):
                        pw.write(lr + si, 0, l, gry_hdr)
                        pw.write(lr + si, 1, v, dat_fmt)
                        pw.set_column(0, 0, 40)
                        pw.set_column(1, 1, 50)

            slug = st.session_state["engagement"].get("client_name", "Client").replace(" ", "_")
            st.download_button(
                label="⬇️ Download Excel Export",
                data=xls_out.getvalue(),
                file_name=f"KKC_Sampling_{slug}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="xls_dl",
            )
            st.markdown(
                '<div class="box-ok">✅ Excel export ready.</div>',
                unsafe_allow_html=True,
            )

    # ── Completion checklist ──────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="sec-title">Completion Status</div>', unsafe_allow_html=True)

    checks = [
        ("Engagement Details entered",             bool(st.session_state["engagement"].get("client_name"))),
        ("Population loaded from Excel",            st.session_state["pop_amount"] > 0),
        ("Sampling Parameters computed",            bool(st.session_state.get("tod_params") or st.session_state.get("toc_params"))),
        ("Sample generated",                       st.session_state.get("sample_df") is not None),
        ("Audit results entered & projection done", st.session_state.get("projection") is not None),
    ]

    for label, done in checks:
        icon  = "✅" if done else "⬜"
        klass = "chk-done" if done else "chk-todo"
        st.markdown(
            f'<span class="{klass}">{icon}&nbsp; {label}</span>',
            unsafe_allow_html=True,
        )


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <div class="kkc-footer">
    KKC &amp; Associates LLP &nbsp;|&nbsp; Chartered Accountants &nbsp;|&nbsp;
    Mumbai · Pune · Bengaluru · Ahmedabad<br>
    KKC Audit Sampling Tool — Internal Use Only &nbsp;|&nbsp;
    KKC Sampling Guide Section 26 &amp; SA 530 (ICAI)
    </div>
    """,
    unsafe_allow_html=True,
)
