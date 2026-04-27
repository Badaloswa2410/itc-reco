"""
ITC Reco by Badal — Streamlit Web Application
============================================
Run with:  streamlit run app.py
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import time
import recon_engine as engine

# =============================================================================
# PAGE CONFIG
# =============================================================================

st.set_page_config(
    page_title="ITC Reco by Badal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# PASSWORD / LOGIN
# =============================================================================

# Passwords are stored securely in Streamlit Secrets (not in this code).
# On Streamlit Cloud: go to App Settings → Secrets and add:
#
#   [credentials]
#   badal = "itcreco@2025"
#   youruser = "yourpassword"
#
# For local testing: create .streamlit/secrets.toml with the same content.
# If secrets are not configured yet, falls back to a safe error message.

def get_credentials():
    try:
        return dict(st.secrets["credentials"])
    except Exception:
        return {}

VALID_CREDENTIALS = get_credentials()

def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.markdown("""
        <div style='text-align:center; padding: 60px 0 20px 0;'>
            <h1 style='color:#1a5276; font-size:2.5rem;'>📊 ITC Reco by Badal</h1>
            <p style='color:#555; font-size:1.1rem;'>GSTR 2B vs Purchase Register Reconciliation</p>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3 = st.columns([1, 1.2, 1])
        with col2:
            with st.form("login_form"):
                st.subheader("🔐 Login")
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                submitted = st.form_submit_button("Login", use_container_width=True)
                if submitted:
                    if not VALID_CREDENTIALS:
                        st.error("⚠️ Credentials not configured. Please add secrets in Streamlit Cloud settings.")
                    elif username in VALID_CREDENTIALS and VALID_CREDENTIALS[username] == password:
                        st.session_state.logged_in = True
                        st.session_state.username  = username
                        st.rerun()
                    else:
                        st.error("Invalid username or password.")
        st.stop()

check_login()

# =============================================================================
# SIDEBAR — Settings & Logout
# =============================================================================

with st.sidebar:
    st.markdown(f"### 👤 Logged in as: `{st.session_state.get('username','')}`")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

    st.markdown("---")
    st.markdown("## ⚙️ Settings")
    st.caption("Adjust matching thresholds. Defaults are recommended.")

    fuzzy_inv = st.slider(
        "Invoice Fuzzy Match Threshold",
        min_value=60, max_value=100, value=85, step=1,
        help="How similar two invoice numbers must be to be considered a fuzzy match. Higher = stricter."
    )
    vendor_name = st.slider(
        "Vendor Name Similarity Threshold",
        min_value=50, max_value=100, value=80, step=1,
        help="Similarity required between vendor names to consider GSTIN-mismatch matches. Higher = stricter."
    )
    cross_tol = st.number_input(
        "Cross-Booking Tolerance (₹)",
        min_value=0.0, max_value=10.0, value=1.0, step=0.5,
        help="Maximum ₹ difference allowed when checking IGST vs CGST+SGST cross-booking."
    )
    otm_tol = st.number_input(
        "One-to-Many Combined Tax Tolerance (₹)",
        min_value=0.0, max_value=10.0, value=1.0, step=0.5,
        help="Maximum ₹ difference allowed for combined tax in one-to-many matching."
    )

    st.markdown("---")
    st.markdown("### 📋 Input File Format")
    st.markdown("""
    Upload **Raw Data.xlsx** with two sheets:
    - Sheet **`2B`** — GSTR 2B data
    - Sheet **`PR`** — Purchase Register

    **2B sheet columns:**
    `Name of the Party, GSTIN, Invoice number, Invoice Date, Taxable value, IGST, CGST, SGST, 2B Month`

    **PR sheet columns:**
    `Booking Month, Vendor name, GSTIN, Invoice Number, Invoice Date, Taxable Value, IGST, CGST, SGST`
    """)

# =============================================================================
# MAIN PAGE
# =============================================================================

st.markdown("""
<div style='padding: 10px 0 5px 0;'>
    <h1 style='color:#1a5276; margin-bottom:0;'>📊 ITC Reco by Badal</h1>
    <p style='color:#666; font-size:1rem; margin-top:4px;'>
        GSTR 2B vs Purchase Register — Intelligent Reconciliation
    </p>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# File uploader
uploaded_file = st.file_uploader(
    "📂 Upload your **Raw Data.xlsx** file (must contain sheets: '2B' and 'PR')",
    type=["xlsx"],
    help="File must have exactly two sheets named '2B' and 'PR' with the required columns."
)

if uploaded_file is not None:
    st.success(f"✅ File uploaded: **{uploaded_file.name}**")

    if st.button("▶️ Run Reconciliation", type="primary", use_container_width=True):

        cfg = {
            "fuzzy_inv_threshold"   : fuzzy_inv,
            "vendor_name_threshold" : vendor_name,
            "cross_book_tolerance"  : cross_tol,
            "otm_tolerance"         : otm_tol,
            "potm_max_candidates"   : 15,
        }

        logs     = []
        log_area = st.empty()

        def log_fn(msg):
            logs.append(msg)
            log_area.code("\n".join(logs), language="")

        try:
            with st.spinner("Processing..."):

                # Load data
                log_fn("Loading data...")
                file_bytes = BytesIO(uploaded_file.read())

                gstr2b_df = pd.read_excel(file_bytes, sheet_name="2B")
                file_bytes.seek(0)
                pr_df     = pd.read_excel(file_bytes, sheet_name="PR")

                gstr2b = engine.load_and_prepare(gstr2b_df, None, engine.GSTR2B_COLUMN_MAP, "GSTR 2B",   log_fn=log_fn)
                pr     = engine.load_and_prepare(pr_df,     None, engine.PR_COLUMN_MAP,     "Purchase Register", log_fn=log_fn)

                log_fn(f"\nRunning reconciliation on {len(gstr2b)} 2B records and {len(pr)} PR records...")
                t0 = time.time()

                gstr2b, pr = engine.run_reconciliation(gstr2b, pr, cfg, log_fn=log_fn)

                elapsed = time.time() - t0
                log_fn(f"\n✅ Done in {elapsed:.1f} seconds.")

                # Build output
                output_bytes = engine.build_output(gstr2b, pr, output_path=None, log_fn=log_fn)
                summary_rows = engine.get_summary_dict(gstr2b, pr)

            # Clear log area after done
            log_area.empty()

            st.success(f"✅ Reconciliation complete in **{elapsed:.1f} seconds**!")
            st.markdown("---")

            # ── Live Summary ────────────────────────────────────────────────
            st.markdown("## 📊 Reconciliation Summary")

            total_2b = len(gstr2b)
            total_pr = len(pr)
            um_2b    = int((gstr2b["Match Status"] == engine.UNMATCHED).sum())
            um_pr    = int((pr["Match Status"]     == engine.UNMATCHED).sum())
            matched_2b = total_2b - um_2b
            matched_pr = total_pr - um_pr

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("GSTR 2B Records",    total_2b)
            c2.metric("PR Records",         total_pr)
            c3.metric("Matched in 2B",      matched_2b, delta=f"{matched_2b/total_2b*100:.1f}%")
            c4.metric("Unmatched in 2B",    um_2b,      delta=f"-{um_2b/total_2b*100:.1f}%", delta_color="inverse")

            st.markdown("### Match Status Breakdown")

            # Separate confirmed matches from flags
            matched_rows = [r for r in summary_rows if r["Flag"] == "✅"]
            flag_rows    = [r for r in summary_rows if r["Flag"] == "⚠"]

            if matched_rows:
                st.markdown("**✅ Confirmed Matches**")
                df_matched = pd.DataFrame(matched_rows)[
                    ["Match Status", "Count in GSTR 2B", "Count in PR", "Explanation"]
                ]
                st.dataframe(df_matched, use_container_width=True, hide_index=True)

            if flag_rows:
                st.markdown("**⚠️ Flagged for Manual Review**")
                df_flags = pd.DataFrame(flag_rows)[
                    ["Match Status", "Count in GSTR 2B", "Count in PR", "Explanation"]
                ]
                st.dataframe(df_flags, use_container_width=True, hide_index=True)

            unmatched_2b = next((r for r in summary_rows if r["Match Status"] == engine.UNMATCHED), None)
            if unmatched_2b:
                st.markdown("**❌ Unmatched**")
                st.dataframe(
                    pd.DataFrame([unmatched_2b])[["Match Status","Count in GSTR 2B","Count in PR","Explanation"]],
                    use_container_width=True, hide_index=True
                )

            # ── Download ────────────────────────────────────────────────────
            st.markdown("---")
            st.download_button(
                label="⬇️ Download Reconciliation Output (Excel)",
                data=output_bytes,
                file_name="reconciliation_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

            st.caption("Output contains 3 sheets: **Summary** (first), **2B Recon**, **PR Recon**")

        except ValueError as e:
            st.error(f"❌ Column Error: {e}")
        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.exception(e)

else:
    st.info("👆 Please upload your **Raw Data.xlsx** file to get started.")
    st.markdown("""
    ### How it works
    1. **Upload** your Raw Data.xlsx file with `2B` and `PR` sheets
    2. **Adjust** matching settings in the sidebar if needed
    3. **Click** Run Reconciliation
    4. **Review** the live summary on screen
    5. **Download** the output Excel file

    ### Matching Levels
    The tool applies intelligent matching across multiple levels — exact match,
    fuzzy invoice, variance tolerance, cross-booking, one-to-many, suffix/prefix,
    vendor name fallback, and PAN-based matching — in a strict priority order
    to ensure the highest quality reconciliation.
    """)

# Footer
st.markdown("---")
st.caption("ITC Reco by Badal · Powered by Python & Streamlit")
