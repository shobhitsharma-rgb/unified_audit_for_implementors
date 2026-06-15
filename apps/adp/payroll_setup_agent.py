import streamlit as st
import pandas as pd
import numpy as np
import os
import itertools

def render_ui():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
        html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
        .main { background-color: #0f1117; }
        .block-container { padding: 2rem 3rem; }
        h1, h2, h3 { font-family: 'IBM Plex Mono', monospace; }
        .title-block {
            background: linear-gradient(135deg, #1a1f2e, #0f1117);
            border-left: 4px solid #00d4aa;
            padding: 1.5rem 2rem; margin-bottom: 2rem; border-radius: 0 8px 8px 0;
        }
        .title-block h1 { color: #00d4aa; font-size: 1.8rem; margin: 0; }
        .title-block p  { color: #8892a4; margin: 0.4rem 0 0; font-size: 0.9rem; }
        .card {
            background: #1a1f2e; border: 1px solid #2a3044;
            border-radius: 10px; padding: 1.2rem 1.5rem; margin-bottom: 1rem;
        }
        .card-title {
            font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem;
            letter-spacing: 0.1em; color: #8892a4; text-transform: uppercase; margin-bottom: 0.5rem;
        }
        .card-value { font-size: 2rem; font-weight: 600; color: #e8ecf4; }
        .tag {
            display: inline-block; padding: 0.25rem 0.7rem; border-radius: 20px;
            font-size: 0.75rem; font-weight: 600; font-family: 'IBM Plex Mono', monospace; margin: 0.2rem;
        }
        .tag-hourly   { background: #0d3d2e; color: #00d4aa; border: 1px solid #00d4aa44; }
        .tag-flat     { background: #2d1f0e; color: #f59e0b; border: 1px solid #f59e0b44; }
        .section-header {
            font-family: 'IBM Plex Mono', monospace; font-size: 1rem; color: #00d4aa;
            border-bottom: 1px solid #2a3044; padding-bottom: 0.5rem; margin: 1.5rem 0 1rem;
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="title-block">
        <h1>📊 ADP Payroll Analyzer</h1>
        <p>Earnings classification · Tax mapping · Uzio-ready output</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ───────────────────────────────────────────────────────────────────

    # ── Load static Uzio state tax reference (bundled alongside the script) ──────
    SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
    STATE_TAX_PATH = os.path.join(SCRIPT_DIR, "state_tax_code.csv")  # committed to repo

    @st.cache_data
    def load_state_tax():
        if not os.path.exists(STATE_TAX_PATH):
            return None
        return pd.read_csv(STATE_TAX_PATH)

    df_state_global = load_state_tax()

    with st.container():
        st.markdown("### 📂 File Uploads")
        col1, col2 = st.columns([2, 1])
        with col1:
            uploaded_payroll = st.file_uploader("ADP Prior Payroll (.xlsx)", type=["xlsx"])
        with col2:
            if df_state_global is not None:
                st.success(f"✅ State Tax Reference loaded  \n`state_tax_code.csv`  \n{len(df_state_global):,} rows")
            else:
                st.error("❌ `state_tax_code.csv` not found.  \nPlace it in the same folder as this script.")
            st.caption("Tax Mapping tab requires the payroll file + state_tax_code.csv in the same folder.")
        st.markdown("---")

    tab1, tab2, tab3 = st.tabs(["📋 Earnings Classifier", "🏛️ Tax Mapping", "💸 Deduction Classifier"])

    # ══════════════════════════════════════════════════════════════════
    # HELPERS
    # ══════════════════════════════════════════════════════════════════
    def extract_code(col):
        return col.split(':')[1].strip().split('-')[0].strip() if ':' in col else col.strip()

    def extract_desc(col):
        return col.split(':')[1].strip() if ':' in col else col.strip()

    # ══════════════════════════════════════════════════════════════════
    # TAB 1 – EARNINGS CLASSIFIER
    # ══════════════════════════════════════════════════════════════════
    with tab1:
        if not uploaded_payroll:
            st.markdown('<div class="card" style="text-align:center;padding:3rem;"><div style="font-size:3rem">📂</div><div style="color:#8892a4;font-family:IBM Plex Mono,monospace;">Upload an ADP Prior Payroll .xlsx file to get started</div></div>', unsafe_allow_html=True)
        else:
            df = pd.read_excel(uploaded_payroll)
            all_cols     = list(df.columns)
            hours_cols   = [c for c in all_cols if 'ADDITIONAL HOURS'    in c.upper()]
            earning_cols = [c for c in all_cols if 'ADDITIONAL EARNINGS' in c.upper()]
            reg_earn_col = next((c for c in all_cols if c.strip().upper() == 'REGULAR EARNINGS'),  None)
            ot_earn_col  = next((c for c in all_cols if c.strip().upper() == 'OVERTIME EARNINGS'), None)
            reg_hrs_col  = next((c for c in all_cols if c.strip().upper() == 'REGULAR HOURS'),     None)
            ot_hrs_col   = next((c for c in all_cols if c.strip().upper() == 'OVERTIME HOURS'),    None)

            hours_codes     = {extract_code(c): c for c in hours_cols}
            hourly_earnings, flat_earnings = [], []
            for ecol in earning_cols:
                code, desc = extract_code(ecol), extract_desc(ecol)
                (hourly_earnings if code in hours_codes else flat_earnings).append(
                    {'code': code, 'description': desc, 'earn_col': ecol,
                     **({'hrs_col': hours_codes[code]} if code in hours_codes else {})}
                )

            def analyze_discretionary(items, df):
                results = []
                if not all([reg_earn_col, ot_earn_col, reg_hrs_col, ot_hrs_col]):
                    return results
                for item in items:
                    ecol = item['earn_col']
                    mask = (df[ot_earn_col].notna() & (df[ot_earn_col] > 0) &
                            df[reg_hrs_col].notna() & (df[ot_hrs_col]  > 0) &
                            df[ecol].notna()         & (df[ecol]        > 0))
                    sub = df[mask].copy()
                    if len(sub) < 2:
                        results.append({**item, 'verdict': 'Insufficient Data', 'avg_diff': None, 'n_rows': len(sub), 'sample': sub})
                        continue
                    sub['base_rate']   = sub[reg_earn_col] / sub[reg_hrs_col]
                    sub['actual_ot']   = sub[ot_earn_col]  / sub[ot_hrs_col]
                    sub['expected_ot'] = sub['base_rate'] * 1.5
                    sub['diff']        = sub['actual_ot'] - sub['expected_ot']
                    avg_diff, med_diff = sub['diff'].mean(), sub['diff'].median()
                    verdict = 'Non-Discretionary' if (avg_diff > 0.15 and med_diff > 0.05) else 'Discretionary'
                    results.append({**item, 'verdict': verdict, 'avg_diff': avg_diff, 'n_rows': len(sub), 'sample': sub})
                return results

            discr_results = analyze_discretionary(hourly_earnings + flat_earnings, df)
            non_discr = sum(1 for r in discr_results if r['verdict'] == 'Non-Discretionary')
            discr_cnt = sum(1 for r in discr_results if r['verdict'] == 'Discretionary')

            c1,c2,c3,c4 = st.columns(4)
            with c1: st.markdown(f'<div class="card"><div class="card-title">Total Earnings</div><div class="card-value">{2+len(earning_cols)}</div></div>', unsafe_allow_html=True)
            with c2: st.markdown(f'<div class="card"><div class="card-title">Hourly Earnings</div><div class="card-value" style="color:#00d4aa">{len(hourly_earnings)+2}</div></div>', unsafe_allow_html=True)
            with c3: st.markdown(f'<div class="card"><div class="card-title">Non-Discretionary</div><div class="card-value" style="color:#a78bfa">{non_discr}</div></div>', unsafe_allow_html=True)
            with c4: st.markdown(f'<div class="card"><div class="card-title">Discretionary</div><div class="card-value" style="color:#60a5fa">{discr_cnt}</div></div>', unsafe_allow_html=True)

            st.markdown('<div class="section-header">① HOURLY vs FLAT EARNINGS</div>', unsafe_allow_html=True)
            lc, rc = st.columns(2)
            with lc:
                st.markdown("**✅ Hourly** *(have matching Hours column)*")
                st.markdown('<span class="tag tag-hourly">REG — Regular Earnings</span>', unsafe_allow_html=True)
                st.markdown('<span class="tag tag-hourly">OT — Overtime Earnings</span>',  unsafe_allow_html=True)
                for i in hourly_earnings:
                    st.markdown(f'<span class="tag tag-hourly">{i["description"]}</span>', unsafe_allow_html=True)
            with rc:
                st.markdown("**💲 Flat / Non-Hourly** *(no hours column)*")
                for i in flat_earnings:
                    st.markdown(f'<span class="tag tag-flat">{i["description"]}</span>', unsafe_allow_html=True)
                if not flat_earnings:
                    st.info("No flat earnings found.")

            st.markdown('<div class="section-header">② DISCRETIONARY vs NON-DISCRETIONARY</div>', unsafe_allow_html=True)
            st.caption("If Actual OT Rate consistently > 1.5× Base Rate when bonus present → Non-Discretionary")

            def sample_table(r):
                sub  = r['sample']
                ecol = r['earn_col']
                aid  = next((c for c in sub.columns if 'ASSOCIATE' in c.upper()), sub.columns[0])
                s = sub[[aid, reg_hrs_col, ot_hrs_col, reg_earn_col, ot_earn_col, ecol,
                          'base_rate','actual_ot','expected_ot','diff']].head(5).copy()
                s.columns = ['Associate ID','Reg Hrs','OT Hrs','Reg Earnings','OT Earnings',
                             'Bonus Amt','Base Rate','Actual OT Rate','Expected OT (1.5x)','Diff']
                for col in ['Reg Earnings','OT Earnings','Bonus Amt','Base Rate','Actual OT Rate','Expected OT (1.5x)','Diff']:
                    s[col] = s[col].apply(lambda x: f"${x:,.4f}" if pd.notna(x) else '')
                return s

            non_d = [r for r in discr_results if r['verdict'] == 'Non-Discretionary']
            d     = [r for r in discr_results if r['verdict'] == 'Discretionary']
            insuf = [r for r in discr_results if r['verdict'] == 'Insufficient Data']

            if non_d:
                st.markdown("#### 🟣 Non-Discretionary")
                for r in non_d:
                    with st.expander(f"**{r['description']}** — avg OT diff: +${r['avg_diff']:.4f} | n={r['n_rows']} rows"):
                        st.dataframe(sample_table(r), use_container_width=True, hide_index=True)
            if d:
                st.markdown("#### 🔵 Discretionary")
                for r in d:
                    with st.expander(f"**{r['description']}** — avg OT diff: ${r['avg_diff']:.4f} | n={r['n_rows']} rows"):
                        st.dataframe(sample_table(r), use_container_width=True, hide_index=True)
            if insuf:
                st.markdown("#### ⚪ Insufficient OT Data")
                for r in insuf:
                    st.markdown(f"- **{r['description']}** ({r['n_rows']} OT rows with this bonus)")

            st.markdown('<div class="section-header">③ FULL SUMMARY TABLE</div>', unsafe_allow_html=True)
            rows = [
                {'Code':'REG','Description':'Regular Earnings', 'Type':'Hourly','Classification':'Non-Discretionary','Avg OT Diff':'—'},
                {'Code':'OT', 'Description':'Overtime Earnings','Type':'Hourly','Classification':'Non-Discretionary','Avg OT Diff':'—'},
            ]
            for r in discr_results:
                is_hourly = any(r['code'] == i['code'] for i in hourly_earnings)
                rows.append({'Code':r['code'],'Description':r['description'],
                             'Type':'Hourly' if is_hourly else 'Flat',
                             'Classification':r['verdict'],
                             'Avg OT Diff': f"${r['avg_diff']:.4f}" if r['avg_diff'] is not None else '—'})
            summary_df = pd.DataFrame(rows)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Download Summary CSV", summary_df.to_csv(index=False).encode(), "earnings_summary.csv","text/csv")


    # ══════════════════════════════════════════════════════════════════
    # TAB 2 – TAX MAPPING
    # ══════════════════════════════════════════════════════════════════
    with tab2:
        if not uploaded_payroll:
            st.markdown('<div class="card" style="text-align:center;padding:3rem;"><div style="font-size:3rem">🏛️</div><div style="color:#8892a4;font-family:IBM Plex Mono,monospace;">Upload an ADP Prior Payroll .xlsx file in the sidebar</div></div>', unsafe_allow_html=True)
        elif df_state_global is None:
            st.error("state_tax_code.csv not found. Place it in the same folder as adp_earnings_analyzer.py and restart the app.")
        else:
            df_pay   = pd.read_excel(uploaded_payroll)
            df_state = df_state_global

            all_cols = list(df_pay.columns)

            # ── ADP tax columns (exclude TAXABLE and TOTAL) ──────────────────────
            def is_actual_tax(col):
                c = col.upper()
                return ('TAX' in c and 'TAXABLE' not in c and 'TOTAL' not in c and col != 'TAX ID')

            raw_tax_cols = [c for c in all_cols if is_actual_tax(c)]

            # ── Auto-detect states from payroll ─────────────────────────────────
            state_col       = next((c for c in all_cols if c.strip().upper() == 'WORKED IN STATE'), None)
            detected_states = sorted(df_pay[state_col].dropna().unique().tolist()) if state_col else []

            # ── Lookup helpers ───────────────────────────────────────────────────
            fed_df = df_state[df_state['state_abbreviation'] == 'FED']

            def lookup_fed(type_code):
                m = fed_df[fed_df['unique_tax_id'].str.contains(f'-{type_code}-', na=False)]
                if not m.empty:
                    r = m.iloc[0]
                    return r['tax_code'], r['unique_tax_id'], r['tax_name']
                return None, None, None

            def lookup_state(state_abbr, type_code):
                st_df = df_state[df_state['state_abbreviation'] == state_abbr]
                m = st_df[st_df['unique_tax_id'].str.contains(f'-{type_code}-', na=False)]
                if not m.empty:
                    r   = m.iloc[0]
                    sub = r.get('sub_tax_desc', None)
                    return r['tax_code'], r['unique_tax_id'], r['tax_name'], (sub if pd.notna(sub) else '')
                return None, None, None, ''

            # ── Keyword → Uzio type_code maps ───────────────────────────────────
            FEDERAL_MAP = {
                'FEDERAL INCOME - EMPLOYEE TAX':        'FIT',
                'MEDICARE - EMPLOYEE TAX':               'MEDI',
                'SOCIAL SECURITY - EMPLOYEE TAX':        'FICA',
                'MEDICARE - EMPLOYER TAX':               'ER_MEDI',
                'SOCIAL SECURITY - EMPLOYER TAX':        'ER_FICA',
                'FUTA - EMPLOYER TAX':                   'ER_FUTA',
            }

            STATE_MAP = {
                'WORKED IN STATE - EMPLOYEE TAX':        'SIT',
                'SUI/SDI - EMPLOYER TAX':                'ER_SUTA',
                'SUI/SDI - EMPLOYEE TAX':                'SDI',
                'WORKED IN LOCAL - EMPLOYEE TAX':        'CITY',
                'LIVED-IN LOCAL - EMPLOYEE TAX':         'CITY',
                'FAMILY LEAVE INSURANCE - EMPLOYEE TAX': 'FLI',
            }

            # ── State selector ───────────────────────────────────────────────────
            st.markdown('<div class="section-header">① SELECT STATES WORKED IN</div>', unsafe_allow_html=True)
            available_states = sorted(df_state[df_state['state_abbreviation'] != 'FED']['state_abbreviation'].unique())
            if detected_states:
                st.info(f"Auto-detected from payroll data: **{', '.join(detected_states)}**")
            selected_states = st.multiselect(
                "Select all states employees worked in (add more if needed):",
                options=available_states,
                default=detected_states
            )

            # ── Build mapping rows ───────────────────────────────────────────────
            mapping_rows = []

            for adp_col in raw_tax_cols:
                col_upper = adp_col.upper()

                # Federal?
                fed_key = next((kw for kw in FEDERAL_MAP if kw in col_upper), None)
                if fed_key:
                    tc = FEDERAL_MAP[fed_key]
                    uzio_code, uid, uname = lookup_fed(tc)
                    mapping_rows.append({
                        'Source Tax Code':            '',
                        'Source Tax Code Name':        adp_col,
                        'Source Tax Code Description': '',
                        'Uzio Tax Code':               uzio_code or '— NOT FOUND —',
                        'Unique Tax ID':               uid       or '—',
                        'Uzio Tax Code Description':   uname     or '—',
                        'Uzio Sub-Tax Description':    '',
                        '_scope': 'Federal', '_state': 'FED', '_mapped': uzio_code is not None,
                    })
                    continue

                # State?
                st_key = next((kw for kw in STATE_MAP if kw in col_upper), None)
                if st_key:
                    tc = STATE_MAP[st_key]
                    if selected_states:
                        for state in selected_states:
                            uzio_code, uid, uname, sub = lookup_state(state, tc)
                            mapping_rows.append({
                                'Source Tax Code':            '',
                                'Source Tax Code Name':        adp_col,
                                'Source Tax Code Description': f'State: {state}',
                                'Uzio Tax Code':               uzio_code or '— NOT FOUND —',
                                'Unique Tax ID':               uid       or '—',
                                'Uzio Tax Code Description':   uname     or '—',
                                'Uzio Sub-Tax Description':    sub,
                                '_scope': 'State', '_state': state, '_mapped': uzio_code is not None,
                            })
                    else:
                        mapping_rows.append({
                            'Source Tax Code':            '',
                            'Source Tax Code Name':        adp_col,
                            'Source Tax Code Description': '⚠️ Select state above',
                            'Uzio Tax Code':               '—',
                            'Unique Tax ID':               '—',
                            'Uzio Tax Code Description':   '—',
                            'Uzio Sub-Tax Description':    '',
                            '_scope': 'State', '_state': '??', '_mapped': False,
                        })
                    continue

                # Unrecognized
                mapping_rows.append({
                    'Source Tax Code':            '',
                    'Source Tax Code Name':        adp_col,
                    'Source Tax Code Description': '',
                    'Uzio Tax Code':               '— MANUAL REVIEW —',
                    'Unique Tax ID':               '—',
                    'Uzio Tax Code Description':   '—',
                    'Uzio Sub-Tax Description':    '',
                    '_scope': 'Unknown', '_state': '—', '_mapped': False,
                })

            # ── Metrics ──────────────────────────────────────────────────────────
            st.markdown('<div class="section-header">② MAPPING RESULTS</div>', unsafe_allow_html=True)
            total    = len(mapping_rows)
            mapped   = sum(1 for r in mapping_rows if r['_mapped'])
            unmapped = total - mapped
            fed_ct   = sum(1 for r in mapping_rows if r['_scope'] == 'Federal')
            st_ct    = sum(1 for r in mapping_rows if r['_scope'] == 'State')

            m1,m2,m3,m4 = st.columns(4)
            with m1: st.markdown(f'<div class="card"><div class="card-title">Total Tax Lines</div><div class="card-value">{total}</div></div>', unsafe_allow_html=True)
            with m2: st.markdown(f'<div class="card"><div class="card-title">Mapped ✓</div><div class="card-value" style="color:#00d4aa">{mapped}</div></div>', unsafe_allow_html=True)
            with m3: st.markdown(f'<div class="card"><div class="card-title">Federal</div><div class="card-value" style="color:#60a5fa">{fed_ct}</div></div>', unsafe_allow_html=True)
            with m4: st.markdown(f'<div class="card"><div class="card-title">State</div><div class="card-value" style="color:#f472b6">{st_ct}</div></div>', unsafe_allow_html=True)

            OUTPUT_COLS = ['Source Tax Code','Source Tax Code Name','Source Tax Code Description',
                           'Uzio Tax Code','Unique Tax ID','Uzio Tax Code Description','Uzio Sub-Tax Description']

            # Federal section
            fed_rows = [r for r in mapping_rows if r['_scope'] == 'Federal']
            if fed_rows:
                st.markdown("#### 🔵 Federal Taxes")
                st.dataframe(pd.DataFrame([{k:r[k] for k in OUTPUT_COLS} for r in fed_rows]),
                             use_container_width=True, hide_index=True)

            # State section — grouped by state
            st_rows = [r for r in mapping_rows if r['_scope'] == 'State']
            if st_rows:
                st.markdown("#### 🩷 State Taxes")
                for state in sorted(set(r['_state'] for r in st_rows)):
                    subset = [r for r in st_rows if r['_state'] == state]
                    st.markdown(f"**{state}**")
                    st.dataframe(pd.DataFrame([{k:r[k] for k in OUTPUT_COLS} for r in subset]),
                                 use_container_width=True, hide_index=True)

            # Unknown / unrecognized
            unk_rows = [r for r in mapping_rows if r['_scope'] == 'Unknown']
            if unk_rows:
                with st.expander(f"⚠️ {len(unk_rows)} unrecognized tax column(s) — manual review needed"):
                    for r in unk_rows:
                        st.markdown(f"- `{r['Source Tax Code Name']}`")

            # Unmapped state taxes
            unm = [r for r in mapping_rows if not r['_mapped'] and r['_scope'] == 'State']
            if unm:
                with st.expander(f"❌ {len(unm)} state tax(es) not found in uploaded CSV"):
                    for r in unm:
                        st.markdown(f"- `{r['Source Tax Code Name']}` — state: **{r['_state']}**")
                    st.caption("The state may not have this tax type, or the CSV may be missing that entry.")

            # ── Full output + download ────────────────────────────────────────────
            st.markdown('<div class="section-header">③ FULL OUTPUT — HANSEN FORMAT</div>', unsafe_allow_html=True)
            out_df = pd.DataFrame([{k:r[k] for k in OUTPUT_COLS} for r in mapping_rows])
            st.dataframe(out_df, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Download Tax Mapping CSV (Hansen Format)",
                out_df.to_csv(index=False).encode(),
                "tax_mapping_output.csv", "text/csv"
            )

    # ══════════════════════════════════════════════════════════════════
    # TAB 3 – DEDUCTION CLASSIFIER (Pre-Tax vs Post-Tax)
    # ══════════════════════════════════════════════════════════════════
    with tab3:
        if not uploaded_payroll:
            st.markdown('<div class="card" style="text-align:center;padding:3rem;"><div style="font-size:3rem">💸</div><div style="color:#8892a4;font-family:IBM Plex Mono,monospace;">Upload an ADP Prior Payroll .xlsx file in the sidebar</div></div>', unsafe_allow_html=True)
        else:
            df_ded = pd.read_excel(uploaded_payroll)
            all_cols_ded = list(df_ded.columns)

            # ── Identify columns ─────────────────────────────────────────────────
            total_earn_col  = next((c for c in all_cols_ded if c.strip().upper() == 'TOTAL EARNINGS'),  None)
            fed_taxable_col = next((c for c in all_cols_ded if 'FEDERAL INCOME - EMPLOYEE TAXABLE' in c.upper()), None)

            ded_cols_raw = [c for c in all_cols_ded
                            if 'VOLUNTARY DEDUCTION' in c.upper()
                            and 'TOTAL' not in c.upper()
                            and 'REV'   not in c.upper()]

            if not total_earn_col or not fed_taxable_col:
                st.error("Could not find TOTAL EARNINGS or FEDERAL INCOME - EMPLOYEE TAXABLE columns in this file.")
            elif not ded_cols_raw:
                st.info("No voluntary deduction columns found in this file.")
            else:
                def get_code(col):
                    return col.split(':')[1].strip().split('-')[0].strip() if ':' in col else col.strip()
                def get_desc(col):
                    return col.split(':')[1].strip() if ':' in col else col.strip()

                # ── Clean and Convert to Numeric ─────────────────────────────────
                df_clean = df_ded.copy()
                
                def clean_currency(series):
                    return pd.to_numeric(series.astype(str).replace(r'[\$,]', '', regex=True), errors='coerce')

                df_clean[total_earn_col]  = clean_currency(df_clean[total_earn_col])
                df_clean[fed_taxable_col] = clean_currency(df_clean[fed_taxable_col])
                
                for col in ded_cols_raw:
                    df_clean[col] = clean_currency(df_clean[col]).fillna(0)

                # ── Filter valid individual rows ─────────────────────────────────
                df_valid = df_clean[
                    df_clean[total_earn_col].notna() &
                    df_clean[fed_taxable_col].notna() &
                    (df_clean[total_earn_col] < 100_000)   # exclude aggregate/summary rows
                ].copy()

                df_valid['_GAP'] = (df_valid[total_earn_col] - df_valid[fed_taxable_col]).round(2)
                df_valid = df_valid[df_valid['_GAP'] >= 0]  # ignore negative gaps

                # ── Core logic ───────────────────────────────────────────────────
                # For each row: find combination of deductions that best explains the GAP
                # Deductions IN the best combo → Pre-Tax
                # Deductions NOT in the best combo → Post-Tax
                TOLERANCE = 5.00   # $5 tolerance for rounding/floating point

                tally = {col: {'pretax': 0, 'posttax': 0, 'total': 0} for col in ded_cols_raw}

                progress = st.progress(0, text="Analysing deductions...")
                total_rows = len(df_valid)

                for i, (_, row) in enumerate(df_valid.iterrows()):
                    gap = row['_GAP']
                    if gap <= 0:
                        continue

                    active = {col: row[col] for col in ded_cols_raw if row[col] > 0}
                    if not active:
                        continue

                    # Find combo of active deductions closest to GAP
                    best_err   = float('inf')
                    best_combo = set()

                    active_cols = list(active.keys())
                    # Sort active cols descending so greedy picks biggest first if we have to truncate
                    active_cols.sort(key=lambda x: active[x], reverse=True)
                    
                    # Safeguard: if >12 deductions on a single row, cap at 12 to avoid 2^n explosion
                    if len(active_cols) > 12:
                        active_cols = active_cols[:12]

                    for r in range(1, len(active_cols) + 1):
                        for combo in itertools.combinations(active_cols, r):
                            s   = sum(active[c] for c in combo)
                            err = abs(s - gap)
                            if err < best_err:
                                best_err   = err
                                best_combo = set(combo)

                    for col in active:
                        tally[col]['total'] += 1
                        if col in best_combo and best_err <= TOLERANCE:
                            tally[col]['pretax']  += 1
                        else:
                            tally[col]['posttax'] += 1

                    if i % 20 == 0:
                        progress.progress(min(i / total_rows, 1.0), text=f"Analysing row {i}/{total_rows}...")

                progress.empty()

                # ── Build results ────────────────────────────────────────────────
                results = []
                for col in ded_cols_raw:
                    t = tally[col]
                    if t['total'] == 0:
                        continue
                    pre_pct  = t['pretax']  / t['total'] * 100
                    post_pct = t['posttax'] / t['total'] * 100
                    if pre_pct >= 60:
                        verdict = 'Pre-Tax'
                    elif post_pct >= 60:
                        verdict = 'Post-Tax'
                    else:
                        verdict = 'Mixed / Unclear'
                    results.append({
                        'Code':         get_code(col),
                        'Description':  get_desc(col),
                        'Total Rows':   t['total'],
                        'Pre-Tax Rows': f"{t['pretax']} ({pre_pct:.0f}%)",
                        'Post-Tax Rows':f"{t['posttax']} ({post_pct:.0f}%)",
                        'Verdict':      verdict,
                        '_pretax_pct':  pre_pct,
                    })

                if not results:
                    st.info("Not enough data rows to classify deductions.")
                else:
                    res_df = pd.DataFrame(results)

                    # ── Metrics ──────────────────────────────────────────────────
                    pre_count  = sum(1 for r in results if r['Verdict'] == 'Pre-Tax')
                    post_count = sum(1 for r in results if r['Verdict'] == 'Post-Tax')
                    mix_count  = sum(1 for r in results if r['Verdict'] == 'Mixed / Unclear')

                    st.markdown('<div class="section-header">① ANALYSIS METHOD</div>', unsafe_allow_html=True)
                    st.markdown("""
                    <div class="card">
                        <div style="color:#c9d1e0; font-size:0.9rem; line-height:1.8;">
                            <b style="color:#00d4aa;">Logic:</b> For each employee row —<br>
                            &nbsp;&nbsp;① Compute <b>GAP = Total Earnings − Federal Income Taxable</b><br>
                            &nbsp;&nbsp;② Find which combination of deductions best explains the GAP<br>
                            &nbsp;&nbsp;③ Deductions <b>inside</b> the best combo → <b style="color:#34d399;">Pre-Tax</b> (they reduce taxable wages)<br>
                            &nbsp;&nbsp;④ Deductions <b>outside</b> the best combo → <b style="color:#f472b6;">Post-Tax</b> (taken after taxes)<br>
                            &nbsp;&nbsp;⑤ Verdict decided by majority across all rows (≥60% threshold, $5 tolerance)
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    m1, m2, m3, m4 = st.columns(4)
                    with m1: st.markdown(f'<div class="card"><div class="card-title">Total Deductions</div><div class="card-value">{len(results)}</div></div>', unsafe_allow_html=True)
                    with m2: st.markdown(f'<div class="card"><div class="card-title">Pre-Tax</div><div class="card-value" style="color:#34d399">{pre_count}</div></div>', unsafe_allow_html=True)
                    with m3: st.markdown(f'<div class="card"><div class="card-title">Post-Tax</div><div class="card-value" style="color:#f472b6">{post_count}</div></div>', unsafe_allow_html=True)
                    with m4: st.markdown(f'<div class="card"><div class="card-title">Unclear</div><div class="card-value" style="color:#f59e0b">{mix_count}</div></div>', unsafe_allow_html=True)

                    # ── Results tables ────────────────────────────────────────────
                    st.markdown('<div class="section-header">② RESULTS</div>', unsafe_allow_html=True)

                    display_cols = ['Code','Description','Total Rows','Pre-Tax Rows','Post-Tax Rows','Verdict']

                    pre_res  = [r for r in results if r['Verdict'] == 'Pre-Tax']
                    post_res = [r for r in results if r['Verdict'] == 'Post-Tax']
                    mix_res  = [r for r in results if r['Verdict'] == 'Mixed / Unclear']

                    if pre_res:
                        st.markdown("#### 🟢 Pre-Tax Deductions")
                        st.dataframe(pd.DataFrame(pre_res)[display_cols],
                                     use_container_width=True, hide_index=True)

                    if post_res:
                        st.markdown("#### 🩷 Post-Tax Deductions")
                        st.dataframe(pd.DataFrame(post_res)[display_cols],
                                     use_container_width=True, hide_index=True)

                    if mix_res:
                        st.markdown("#### 🟡 Mixed / Unclear")
                        st.dataframe(pd.DataFrame(mix_res)[display_cols],
                                     use_container_width=True, hide_index=True)
                        st.caption("These deductions appear in both pre-tax and post-tax positions across different rows. Manual review recommended.")

                    # ── Download ──────────────────────────────────────────────────
                    st.markdown('<div class="section-header">③ FULL SUMMARY</div>', unsafe_allow_html=True)
                    full_df = pd.DataFrame(results)[display_cols]
                    st.dataframe(full_df, use_container_width=True, hide_index=True)
                    st.download_button(
                        "⬇️ Download Deduction Classification CSV",
                        full_df.to_csv(index=False).encode(),
                        "deduction_classification.csv", "text/csv"
                    )

if __name__ == "__main__":
    st.set_page_config(page_title="ADP Payroll Analyzer", page_icon="📊", layout="wide")
    render_ui()
