"""
=============================================================================
GSTR 2B vs Purchase Register - Reconciliation Engine v5.0
=============================================================================
Dual-use:
  (a) Standalone script  →  python recon_engine.py
  (b) Imported by Streamlit web app  →  import recon_engine

Changes in v5:
  - Input file: Sample.xlsx → Raw Data.xlsx
  - Summary sheet is now the FIRST sheet in output
  - Summary has plain-English 'Explanation' column
  - Engine accepts DataFrames directly (for web app)
  - All settings passed as config dict (adjustable from web app)
=============================================================================
"""

import pandas as pd
import re
import os
from io import BytesIO
from itertools import combinations
from rapidfuzz import fuzz

# =============================================================================
# STANDALONE CONFIG
# =============================================================================

FOLDER      = r"C:\Users\badal\Desktop\Office\Personal Research\2B Reconciliation\Type 2"
INPUT_FILE  = os.path.join(FOLDER, "Raw Data.xlsx")
SHEET_2B    = "2B"
SHEET_PR    = "PR"
OUTPUT_FILE = os.path.join(FOLDER, "reconciliation_output.xlsx")

GSTR2B_COLUMN_MAP = {
    "supplier_name" : "Supplier Name",
    "gstin"         : "Supplier GSTN",
    "invoice_no"    : "Invoice Number",
    "invoice_date"  : "Invoice Date",
    "taxable_value" : "Taxable Value",
    "igst"          : "IGST",
    "cgst"          : "CGST",
    "sgst"          : "SGST",
    "month"         : "2B Month",
}

PR_COLUMN_MAP = {
    "month"         : "Booking Month",
    "supplier_name" : "Supplier Name",
    "gstin"         : "Supplier GSTN",
    "invoice_no"    : "Invoice Number",
    "invoice_date"  : "Invoice Date",
    "taxable_value" : "Taxable Value",
    "igst"          : "IGST",
    "cgst"          : "CGST",
    "sgst"          : "SGST",
}

DEFAULT_CFG = {
    "fuzzy_inv_threshold"   : 85,
    "vendor_name_threshold" : 80,
    "cross_book_tolerance"  : 1.0,
    "otm_tolerance"         : 1.0,
    "potm_max_candidates"   : 15,
}

# =============================================================================
# MATCH LABELS
# =============================================================================

M         = "Matched"
M_FI      = "Matched - Fuzzy Invoice"
M_1       = "Matched - 1% Variance"
M_5       = "Matched - 5% Variance"
M_FI_1    = "Matched - Fuzzy Invoice & 1% Variance"
M_FI_5    = "Matched - Fuzzy Invoice & 5% Variance"
CB        = "Cross Booked"
CB_FI     = "Cross Booked - Fuzzy Invoice"
DFY       = "Possible Match - Different Financial Year"
PAN       = "Matched - PAN Based"
PAN_1     = "Matched - PAN Based - 1% Variance"
PAN_5     = "Matched - PAN Based - 5% Variance"
M_SP      = "Matched - Invoice Suffix/Prefix"
M_SP_1    = "Matched - Invoice Suffix/Prefix - 1% Variance"
M_SP_5    = "Matched - Invoice Suffix/Prefix - 5% Variance"
VN        = "Matched via Vendor Name"
VN_FI     = "Matched via Vendor Name - Fuzzy Invoice"
VN_1      = "Matched via Vendor Name - 1% Variance"
VN_5      = "Matched via Vendor Name - 5% Variance"
CB_GS     = "Cross Booked - GSTIN Mismatch"
OTM       = "One to Many Match"
OTM_1     = "One to Many - 1% Variance"
OTM_5     = "One to Many - 5% Variance"
OTM_FI    = "One to Many - Fuzzy Invoice"
OTM_CB    = "One to Many - Cross Booked"
OTM_SP    = "One to Many - Invoice Suffix/Prefix"
OTM_SP_1  = "One to Many - Invoice Suffix/Prefix - 1% Variance"
OTM_SP_5  = "One to Many - Invoice Suffix/Prefix - 5% Variance"
OTM_GS    = "One to Many - GSTIN Mismatch"
OTM_GS_1  = "One to Many - GSTIN Mismatch & 1% Variance"
OTM_GS_5  = "One to Many - GSTIN Mismatch & 5% Variance"
POTM_M    = "Partial One to Many - Matched"
POTM_M_GS = "Partial One to Many - GSTIN Mismatch - Matched"
POTM_U    = "Partial One to Many - Unmatched Item"
F_GI      = "Possible Match - GSTIN & Invoice Uncertain"
F_VN      = "Possible Match - Vendor Unconfirmed, Same Invoice"
F_CM      = "Possible Match - Component Mismatch"
F_LV      = "Possible Match - Large Variance"
F_OTM_GI  = "One to Many - GSTIN & Invoice Uncertain"
F_OTM_VN  = "One to Many - Vendor Unconfirmed"
UNMATCHED = "Unmatched"

FLAG_LABELS = {F_GI, F_VN, F_OTM_GI, F_OTM_VN, F_CM, F_LV}

ALL_STATUSES = [
    M, M_FI, M_1, M_5, M_FI_1, M_FI_5, CB, CB_FI, DFY,
    PAN, PAN_1, PAN_5, M_SP, M_SP_1, M_SP_5,
    VN, VN_FI, VN_1, VN_5, CB_GS,
    OTM, OTM_1, OTM_5, OTM_FI, OTM_CB,
    OTM_SP, OTM_SP_1, OTM_SP_5, OTM_GS, OTM_GS_1, OTM_GS_5,
    POTM_M, POTM_M_GS, POTM_U,
    F_GI, F_VN, F_OTM_GI, F_OTM_VN, F_CM, F_LV,
    UNMATCHED,
]

STATUS_REMARKS = {
    M         : "Invoice No, GSTIN and all tax components match exactly.",
    M_FI      : "GSTIN and tax match exactly. Invoice No has minor variation (e.g. special characters, spacing). Verify invoice number.",
    M_1       : "Invoice No and GSTIN match exactly. Tax differs by up to 1%. Likely a rounding difference.",
    M_5       : "Invoice No and GSTIN match exactly. Tax differs between 1% and 5%. Check for rate or rounding difference.",
    M_FI_1    : "Invoice No has minor variation AND tax differs by up to 1%. Verify both.",
    M_FI_5    : "Invoice No has minor variation AND tax differs between 1% and 5%. Verify both.",
    CB        : "GSTIN and Invoice match. Supplier filed as IGST but buyer booked CGST+SGST (or vice versa). Verify place of supply.",
    CB_FI     : "Cross-booked (IGST vs CGST+SGST) with minor invoice variation. Verify invoice number and place of supply.",
    DFY       : "GSTIN, Invoice and Tax match exactly but fall in different Financial Years. No ITC risk if intentional; verify booking year.",
    PAN       : "GSTIN differs but PAN matches — same entity, different state registration. Invoice and tax match exactly.",
    PAN_1     : "PAN matches (different state GSTIN). Invoice matches. Tax within 1% variance.",
    PAN_5     : "PAN matches (different state GSTIN). Invoice matches. Tax within 5% variance. Verify tax amount.",
    M_SP      : "GSTIN matches. One invoice contains the other as a prefix/suffix (e.g. SLRO2526-2520004 vs 2520004). Tax matches exactly.",
    M_SP_1    : "Invoice suffix/prefix match with GSTIN match. Tax within 1% variance.",
    M_SP_5    : "Invoice suffix/prefix match with GSTIN match. Tax within 5% variance. Verify tax amount.",
    VN        : "GSTIN differs but Vendor Name is similar — likely wrong GSTIN in PR. Invoice and tax match exactly. Correct the GSTIN in PR.",
    VN_FI     : "Vendor Name similar, GSTIN differs, Invoice has minor variation, Tax exact. Verify GSTIN and invoice.",
    VN_1      : "Vendor Name similar, GSTIN differs, Invoice exact, Tax within 1% variance. Correct GSTIN in PR.",
    VN_5      : "Vendor Name similar, GSTIN differs, Invoice exact, Tax within 5% variance. Correct GSTIN and verify tax.",
    CB_GS     : "Cross-booked (IGST vs CGST+SGST) with GSTIN mismatch. Verify place of supply and correct GSTIN in PR.",
    OTM       : "One invoice in 2B matches multiple PR entries (or vice versa). GSTIN and Invoice exact. Combined tax matches exactly.",
    OTM_1     : "One-to-Many. GSTIN and Invoice exact. Combined tax within 1% variance.",
    OTM_5     : "One-to-Many. GSTIN and Invoice exact. Combined tax within 5% variance. Verify amounts.",
    OTM_FI    : "One-to-Many with minor invoice variation. GSTIN matches. Combined tax exact.",
    OTM_CB    : "One-to-Many with cross-booking (IGST vs CGST+SGST). Verify place of supply.",
    OTM_SP    : "One-to-Many where invoice numbers differ by prefix/suffix (e.g. HGJW2526-2551405 vs 2551405). GSTIN matches. Combined tax exact.",
    OTM_SP_1  : "One-to-Many suffix/prefix. GSTIN matches. Combined tax within 1% variance.",
    OTM_SP_5  : "One-to-Many suffix/prefix. GSTIN matches. Combined tax within 5% variance. Verify amounts.",
    OTM_GS    : "One-to-Many. GSTIN differs but Vendor Name similar. Invoice and combined tax match. Correct GSTIN in PR.",
    OTM_GS_1  : "One-to-Many. GSTIN mismatch, Vendor Name similar. Combined tax within 1% variance.",
    OTM_GS_5  : "One-to-Many. GSTIN mismatch, Vendor Name similar. Combined tax within 5% variance.",
    POTM_M    : "Subset of PR entries matched to one 2B invoice. These rows are part of the matched group.",
    POTM_M_GS : "Subset match with GSTIN mismatch (Vendor Name confirmed). These rows are part of the matched group.",
    POTM_U    : "This entry belongs to a partial one-to-many group but did NOT fit the matched subset. Treat as Unmatched.",
    F_GI      : "REVIEW REQUIRED: GSTIN differs, Vendor Name similar, Invoice fuzzy, Tax not exact. Too risky to auto-match — verify manually.",
    F_VN      : "REVIEW REQUIRED: GSTIN and Vendor Name both differ, Invoice matches, Tax within variance. Risk of wrong match — verify manually.",
    F_CM      : "REVIEW REQUIRED: GSTIN and Invoice match exactly, Total tax matches, but IGST/CGST/SGST breakdown differs. Verify tax components.",
    F_LV      : "REVIEW REQUIRED: GSTIN and Invoice match exactly but Tax differs by more than 5%. Investigate the large difference.",
    F_OTM_GI  : "REVIEW REQUIRED (One-to-Many): GSTIN mismatch, Invoice fuzzy. Could not auto-match safely — verify manually.",
    F_OTM_VN  : "REVIEW REQUIRED (One-to-Many): Vendor not confirmed. Could not auto-match safely — verify manually.",
    UNMATCHED : "No matching invoice found. May be timing difference, missing entry, or data error.",
}

# =============================================================================
# HELPERS
# =============================================================================

def get_financial_year(dt):
    if pd.isnull(dt):
        return "Unknown"
    try:
        m, y = dt.month, dt.year
        return f"{y}-{str(y+1)[2:]}" if m >= 4 else f"{y-1}-{str(y)[2:]}"
    except Exception:
        return "Unknown"

def normalize_invoice(inv_no, aggressive=False):
    if pd.isnull(inv_no): return ""
    inv = re.sub(r"[\s/\-_\.]", "", str(inv_no).strip().upper())
    return inv.replace("O", "0") if aggressive else inv

def within_variance(a, b, pct):
    if a == 0 and b == 0: return True
    base = max(abs(a), abs(b))
    return (abs(a - b) / base) <= pct if base > 0 else True

NAME_EXPANSIONS = [
    (r'\bpvt\.?\b','private'),(r'\bltd\.?\b','limited'),(r'\bco\.?\b','company'),
    (r'\bcorp\.?\b','corporation'),(r'\bllp\.?\b','llp'),(r'\bllc\.?\b','llc'),
    (r'\bent\.?\b','enterprises'),(r'\bbros\.?\b','brothers'),(r'\bmfg\.?\b','manufacturing'),
    (r'\bind\.?\b','industries'),(r'\bintl\.?\b','international'),(r'\bassoc\.?\b','associates'),
]

def normalize_vendor_name(name):
    if pd.isnull(name): return ""
    name = re.sub(r'[^\w\s]', ' ', str(name).lower().strip())
    for pat, rep in NAME_EXPANSIONS:
        name = re.sub(pat, rep, name, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', name).strip()

def extract_pan(gstin):
    g = str(gstin).strip().upper()
    return g[2:12] if len(g) >= 12 else ""

def is_suffix_prefix_match(inv_a, inv_b, min_len=3):
    if not inv_a or not inv_b: return False
    shorter, longer = (inv_a, inv_b) if len(inv_a) <= len(inv_b) else (inv_b, inv_a)
    return len(shorter) >= min_len and shorter in longer

def total_tax_exact(r2b, rpr, tol=0.02):
    t2 = r2b["igst"] + r2b["cgst"] + r2b["sgst"]
    tp = rpr["igst"] + rpr["cgst"] + rpr["sgst"]
    if abs(t2 - tp) > tol: return False
    return not (abs(r2b["igst"]-rpr["igst"]) < 0.01 and abs(r2b["cgst"]-rpr["cgst"]) < 0.01 and abs(r2b["sgst"]-rpr["sgst"]) < 0.01)

def is_cross_booked_pair(r2b, rpr, tol):
    ig2,cg2,sg2 = float(r2b.get("igst",0) or 0), float(r2b.get("cgst",0) or 0), float(r2b.get("sgst",0) or 0)
    igp,cgp,sgp = float(rpr.get("igst",0) or 0), float(rpr.get("cgst",0) or 0), float(rpr.get("sgst",0) or 0)
    return (ig2 > 0 and abs(ig2-(cgp+sgp)) <= tol) or (igp > 0 and abs((cg2+sg2)-igp) <= tol)

def is_cross_booked_group(one_row, many_df, tol):
    ig1,cg1,sg1 = float(one_row.get("igst",0) or 0), float(one_row.get("cgst",0) or 0), float(one_row.get("sgst",0) or 0)
    ig_m,cg_m,sg_m = many_df["igst"].sum(), many_df["cgst"].sum(), many_df["sgst"].sum()
    return (ig1 > 0 and abs(ig1-(cg_m+sg_m)) <= tol) or (ig_m > 0 and abs((cg1+sg1)-ig_m) <= tol)

# =============================================================================
# DATA LOADING
# =============================================================================

def load_and_prepare(source, sheet_name, column_map, source_label, log_fn=print):
    log_fn(f"  Loading [{source_label}]...")
    df = source.copy() if isinstance(source, pd.DataFrame) else pd.read_excel(source, sheet_name=sheet_name)
    df = df.rename(columns={v: k for k, v in column_map.items()})
    missing = [column_map[k] for k in column_map if k not in df.columns]
    if missing:
        raise ValueError(f"[{source_label}] Columns not found: {missing}")
    df["invoice_date"] = pd.to_datetime(df["invoice_date"], dayfirst=True, errors="coerce")
    for col in ["taxable_value","igst","cgst","sgst"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).round(2)
    df["gstin"]       = df["gstin"].astype(str).str.strip().str.upper()
    df["_inv_norm"]   = df["invoice_no"].apply(normalize_invoice)
    df["_inv_aggr"]   = df["invoice_no"].apply(lambda x: normalize_invoice(x, True))
    df["_fy"]         = df["invoice_date"].apply(get_financial_year)
    df["_name_norm"]  = df["supplier_name"].apply(normalize_vendor_name)
    df["_pan"]        = df["gstin"].apply(extract_pan)
    df["_total_tax"]  = (df["igst"] + df["cgst"] + df["sgst"]).round(2)
    df["Match Status"]      = UNMATCHED
    df["_matched_idx"]      = None
    df["_matched_idx_list"] = None
    log_fn(f"    -> {len(df)} records loaded.")
    return df

# =============================================================================
# CLASSIFIERS
# =============================================================================

def classify_pair(r2b, rpr, cfg):
    FI, VN_T, CT = cfg["fuzzy_inv_threshold"], cfg["vendor_name_threshold"], cfg["cross_book_tolerance"]
    gstin_ok = (r2b["gstin"] == rpr["gstin"])
    pan_ok   = (not gstin_ok and r2b["_pan"] != "" and r2b["_pan"] == rpr["_pan"])
    fy_ok    = (r2b["_fy"] == rpr["_fy"])
    inv_exact = (r2b["_inv_norm"] != "" and r2b["_inv_norm"] == rpr["_inv_norm"])
    inv_fuzzy = (not inv_exact and r2b["_inv_norm"] != "" and rpr["_inv_norm"] != ""
                 and fuzz.ratio(r2b["_inv_aggr"], rpr["_inv_aggr"]) >= FI)
    inv_sp    = (not inv_exact and not inv_fuzzy
                 and is_suffix_prefix_match(r2b["_inv_norm"], rpr["_inv_norm"]))
    name_ok   = fuzz.token_sort_ratio(r2b["_name_norm"], rpr["_name_norm"]) >= VN_T
    tax_exact = (abs(r2b["igst"]-rpr["igst"]) < 0.01 and abs(r2b["cgst"]-rpr["cgst"]) < 0.01
                 and abs(r2b["sgst"]-rpr["sgst"]) < 0.01)
    cross    = is_cross_booked_pair(r2b, rpr, CT)
    tax_1pct = (not tax_exact and not cross
                and within_variance(r2b["igst"],rpr["igst"],0.01)
                and within_variance(r2b["cgst"],rpr["cgst"],0.01)
                and within_variance(r2b["sgst"],rpr["sgst"],0.01))
    tax_5pct = (not tax_exact and not cross and not tax_1pct
                and within_variance(r2b["igst"],rpr["igst"],0.05)
                and within_variance(r2b["cgst"],rpr["cgst"],0.05)
                and within_variance(r2b["sgst"],rpr["sgst"],0.05))

    if gstin_ok and fy_ok:
        if inv_exact:
            if cross:     return (1, CB,    True)
            if tax_exact: return (2, M,     True)
            if tax_1pct:  return (3, M_1,   True)
            if tax_5pct:  return (4, M_5,   True)
            if total_tax_exact(r2b, rpr): return (52, F_CM, False)
        if inv_fuzzy:
            if cross:     return (5, CB_FI, True)
            if tax_exact: return (6, M_FI,  True)
            if tax_1pct:  return (7, M_FI_1,True)
            if tax_5pct:  return (8, M_FI_5,True)
        if inv_sp:
            if tax_exact: return (9, M_SP,   True)
            if tax_1pct:  return (10,M_SP_1, True)
            if tax_5pct:  return (11,M_SP_5, True)
    if gstin_ok and not fy_ok and inv_exact and tax_exact and not cross:
        return (12, DFY, True)
    if pan_ok and fy_ok and inv_exact:
        if tax_exact: return (13, PAN,  True)
        if tax_1pct:  return (14, PAN_1,True)
        if tax_5pct:  return (15, PAN_5,True)
    if not gstin_ok and name_ok and fy_ok:
        if inv_exact:
            if cross:     return (16, CB_GS,True)
            if tax_exact: return (17, VN,   True)
            if tax_1pct:  return (18, VN_1, True)
            if tax_5pct:  return (19, VN_5, True)
        if inv_fuzzy:
            if tax_exact: return (20, VN_FI,True)
            if tax_1pct or tax_5pct or cross: return (50, F_GI, False)
    if not gstin_ok and not name_ok and fy_ok:
        if (inv_exact or inv_fuzzy) and (tax_1pct or tax_5pct):
            return (51, F_VN, False)
    if gstin_ok and fy_ok and inv_exact:
        t2 = r2b["igst"]+r2b["cgst"]+r2b["sgst"]
        tp = rpr["igst"]+rpr["cgst"]+rpr["sgst"]
        if t2 > 0 and tp > 0 and not tax_exact and not tax_1pct and not tax_5pct and not cross:
            return (53, F_LV, False)
    return None


def classify_otm(one_row, many_df, cfg):
    FI,VN_T,OT,CT = cfg["fuzzy_inv_threshold"],cfg["vendor_name_threshold"],cfg["otm_tolerance"],cfg["cross_book_tolerance"]
    gstin_ok  = all(g == one_row["gstin"] for g in many_df["gstin"])
    name_ok   = all(fuzz.token_sort_ratio(one_row["_name_norm"],n) >= VN_T for n in many_df["_name_norm"])
    inv_exact = all(n == one_row["_inv_norm"] and n != "" for n in many_df["_inv_norm"])
    inv_fuzzy = (not inv_exact and all(fuzz.ratio(one_row["_inv_aggr"],a) >= FI for a in many_df["_inv_aggr"]))
    inv_sp    = (not inv_exact and not inv_fuzzy
                 and all(is_suffix_prefix_match(one_row["_inv_norm"],a) for a in many_df["_inv_norm"]))
    comb      = many_df["_total_tax"].sum().round(2)
    target    = one_row["_total_tax"]
    tax_exact = abs(comb-target) <= OT
    tax_1pct  = not tax_exact and within_variance(comb,target,0.01)
    tax_5pct  = not tax_exact and not tax_1pct and within_variance(comb,target,0.05)
    cross     = is_cross_booked_group(one_row,many_df,CT)

    if gstin_ok:
        if inv_exact:
            if cross:     return (1, OTM_CB, True)
            if tax_exact: return (2, OTM,    True)
            if tax_1pct:  return (3, OTM_1,  True)
            if tax_5pct:  return (4, OTM_5,  True)
        if inv_fuzzy:
            if cross:     return (5, OTM_CB, True)
            if tax_exact: return (6, OTM_FI, True)
        if inv_sp:
            if tax_exact: return (7, OTM_SP,  True)
            if tax_1pct:  return (8, OTM_SP_1,True)
            if tax_5pct:  return (9, OTM_SP_5,True)
    else:
        if name_ok:
            if inv_exact:
                if tax_exact: return (10,OTM_GS,  True)
                if tax_1pct:  return (11,OTM_GS_1,True)
                if tax_5pct:  return (12,OTM_GS_5,True)
            if inv_fuzzy:
                if tax_exact: return (13,OTM_GS,  True)
                return (50, F_OTM_GI, False)
            if inv_sp:
                if tax_exact: return (14,OTM_SP,  True)
                if tax_1pct:  return (15,OTM_SP_1,True)
                if tax_5pct:  return (16,OTM_SP_5,True)
        else:
            if (inv_exact or inv_fuzzy) and (tax_1pct or tax_5pct):
                return (51, F_OTM_VN, False)
    return None

# =============================================================================
# MARKERS
# =============================================================================

def mark_1to1(i,j,lbl,gstr2b,pr,m2b,mpr):
    gstr2b.at[i,"Match Status"]=lbl; gstr2b.at[i,"_matched_idx"]=j
    pr.at[j,"Match Status"]=lbl;     pr.at[j,"_matched_idx"]=i
    m2b.add(i); mpr.add(j)

def mark_2b_to_many_pr(i,js,lbl,gstr2b,pr,m2b,mpr):
    gstr2b.at[i,"Match Status"]=lbl; gstr2b.at[i,"_matched_idx_list"]=list(js); m2b.add(i)
    for j in js: pr.at[j,"Match Status"]=lbl; pr.at[j,"_matched_idx"]=i; mpr.add(j)

def mark_pr_to_many_2b(j,is_,lbl,gstr2b,pr,m2b,mpr):
    pr.at[j,"Match Status"]=lbl; pr.at[j,"_matched_idx_list"]=list(is_); mpr.add(j)
    for i in is_: gstr2b.at[i,"Match Status"]=lbl; gstr2b.at[i,"_matched_idx"]=j; m2b.add(i)

# =============================================================================
# PHASES
# =============================================================================

VARIANCE_LABELS = {M_FI,M_1,M_5,M_FI_1,M_FI_5,CB_FI,VN_FI,VN_1,VN_5,PAN_1,PAN_5,M_SP_1,M_SP_5}


def _phase1_run(gstr2b,pr,m2b,mpr,cfg,exact_only=False,variance_only=False,label="Phase 1",log_fn=print,progress_fn=None):
    """
    One-to-one matching.

    exact_only mode:
      1. Vectorized pandas merge finds all exact M pairs instantly → assigns immediately
      2. Loop runs on the now-smaller unmatched pool for all other labels
         (fuzzy, variance, PAN, vendor-name, suffix/prefix, DFY)
         Skips M (already done). Skips unmatched rows (already assigned).

    variance_only mode:
      Loop only — finds M_1, M_5, M_FI, VN_1, VN_5 etc.
    """
    log_fn(f"\n  [{label}] One-to-One..."
           + (" (exact)" if exact_only else " (variance)" if variance_only else ""))

    TAX_RATIO_LIMIT = 10.0

    # ── FAST PATH: vectorized merge for M (exact_only only) ───────────────────
    if exact_only and not variance_only:
        pr_um = pr[~pr.index.isin(mpr)]
        g2b   = gstr2b[~gstr2b.index.isin(m2b)]
        merge_keys = ["gstin","_inv_norm","_fy","igst","cgst","sgst"]
        try:
            merged = (
                g2b[merge_keys].reset_index()
                .merge(pr_um[merge_keys].reset_index(),
                       on=merge_keys, suffixes=("_2b","_pr"))
            )
            seen_2b, seen_pr = set(), set()
            for _, row in merged.iterrows():
                i, j = int(row["index_2b"]), int(row["index_pr"])
                if i in m2b or j in mpr: continue
                if i in seen_2b or j in seen_pr: continue
                mark_1to1(i, j, M, gstr2b, pr, m2b, mpr)
                seen_2b.add(i); seen_pr.add(j)
            log_fn(f"    -> {len(seen_2b)} exact matches (fast path).")
            if progress_fn:
                progress_fn(label, len(gstr2b), len(gstr2b))
        except Exception:
            pass  # loop below catches M as fallback

    # ── LOOP PATH: bucket-based, runs on remaining unmatched pool ─────────────
    pr_um = pr[~pr.index.isin(mpr)]   # refreshed after fast path assignments
    bg, bp, bi = {}, {}, {}
    for j, rpr in pr_um.iterrows():
        bg.setdefault(rpr["gstin"], []).append(j)
        if rpr["_pan"]:      bp.setdefault(rpr["_pan"], []).append(j)
        if rpr["_inv_norm"]: bi.setdefault(rpr["_inv_norm"], []).append(j)

    def cands(r2b):
        seen, res = set(), []
        for j in bg.get(r2b["gstin"], []):
            if j not in seen: seen.add(j); res.append(j)
        if r2b["_pan"]:
            for j in bp.get(r2b["_pan"], []):
                if pr_um.loc[j,"gstin"] != r2b["gstin"] and j not in seen:
                    seen.add(j); res.append(j)
        if r2b["_inv_norm"]:
            for j in bi.get(r2b["_inv_norm"], []):
                if pr_um.loc[j,"gstin"] != r2b["gstin"] and j not in seen:
                    seen.add(j); res.append(j)
        return res

    pairs      = []
    total_rows = len(gstr2b)
    processed  = 0

    for i, r2b in gstr2b.iterrows():
        processed += 1
        if progress_fn and processed % 50 == 0:
            progress_fn(label, processed, total_rows)
        if i in m2b:
            continue
        r2b_tax = r2b["_total_tax"]
        r2b_fy  = r2b["_fy"]

        for j in cands(r2b):
            if j in mpr or j not in pr_um.index:
                continue
            rpr = pr_um.loc[j]
            if rpr["_fy"] != r2b_fy and rpr["_inv_norm"] != r2b["_inv_norm"]:
                continue
            rpr_tax = rpr["_total_tax"]
            if r2b_tax > 0 and rpr_tax > 0:
                if max(r2b_tax, rpr_tax) / min(r2b_tax, rpr_tax) > TAX_RATIO_LIMIT:
                    continue
            res = classify_pair(r2b, rpr, cfg)
            if res is None: continue
            p, lbl, ir = res
            if not ir: continue
            if lbl == M: continue  # already handled by fast path
            if exact_only    and lbl in VARIANCE_LABELS: continue
            if variance_only and lbl not in VARIANCE_LABELS: continue
            pairs.append((p, i, j, lbl))

    if progress_fn:
        progress_fn(label, total_rows, total_rows)

    pairs.sort(key=lambda x: x[0])
    lc = {}
    for p, i, j, lbl in pairs:
        if i in m2b or j in mpr: continue
        mark_1to1(i, j, lbl, gstr2b, pr, m2b, mpr)
        lc[lbl] = lc.get(lbl, 0) + 1

    if lc:
        log_fn(f"    -> {sum(lc.values())} additional matched:")
        for lbl, cnt in sorted(lc.items(), key=lambda x: x[1], reverse=True):
            log_fn(f"       {lbl}: {cnt}")




def _build_otm_groups(df, cfg):
    """
    Pre-computes all candidate OTM groups from a DataFrame.
    Returns a dict: row_index → list of candidate groups (each group is a DataFrame).

    Key insight: OTM only makes sense when 2+ rows share the same
    GSTIN+Invoice (or similar). Pre-grouping using pandas groupby is
    O(n log n) — far faster than calling find_groups per row.

    Groups built:
      G1: same GSTIN + same inv_norm + same FY     (exact OTM — most common)
      G2: same GSTIN + same inv_norm (fuzzy aggr)  (fuzzy invoice OTM)
      G3: same GSTIN + suffix/prefix inv            (suffix/prefix OTM)
      G4: same inv_norm + name match                (GSTIN-mismatch OTM)
    Fuzzy/vendor-name groups (G2-G4) are only built for rows where G1 failed.
    """
    FI   = cfg["fuzzy_inv_threshold"]
    VN_T = cfg["vendor_name_threshold"]

    # G1: Exact GSTIN + inv_norm + FY — vectorized groupby
    g1_groups = {}
    for (gstin, inv, fy), grp in df.groupby(["gstin","_inv_norm","_fy"]):
        if inv == "" or len(grp) < 2:
            continue
        for idx in grp.index:
            g1_groups.setdefault(idx, []).append(grp)

    # For rows without G1 candidates, build G2/G3/G4
    # Only process rows that actually have potential — same GSTIN bucket
    gstin_groups = {}
    for j, row in df.iterrows():
        gstin_groups.setdefault(row["gstin"], []).append(j)

    inv_groups = {}  # inv_norm → list of indices
    for j, row in df.iterrows():
        if row["_inv_norm"]:
            inv_groups.setdefault(row["_inv_norm"], []).append(j)

    extra_groups = {}  # row_index → additional candidate groups

    for i, row in df.iterrows():
        if i in g1_groups:
            continue  # G1 already found — skip expensive fuzzy work

        same_gstin = [j for j in gstin_groups.get(row["gstin"], [])
                      if j != i and df.loc[j, "_fy"] == row["_fy"]]

        if len(same_gstin) >= 2:
            # G2: fuzzy invoice within same GSTIN
            fuzzy_ids = [j for j in same_gstin
                         if fuzz.ratio(row["_inv_aggr"], df.loc[j,"_inv_aggr"]) >= FI
                         and df.loc[j,"_inv_norm"] != row["_inv_norm"]]
            if len(fuzzy_ids) >= 2:
                extra_groups.setdefault(i, []).append(df.loc[fuzzy_ids])

            # G3: suffix/prefix within same GSTIN
            sp_ids = [j for j in same_gstin
                      if is_suffix_prefix_match(row["_inv_norm"], df.loc[j,"_inv_norm"])
                      and df.loc[j,"_inv_norm"] != row["_inv_norm"]]
            if len(sp_ids) >= 2:
                extra_groups.setdefault(i, []).append(df.loc[sp_ids])

        # G4: same invoice norm, different GSTIN, vendor name match
        same_inv = [j for j in inv_groups.get(row["_inv_norm"], [])
                    if j != i
                    and df.loc[j,"gstin"] != row["gstin"]
                    and df.loc[j,"_fy"] == row["_fy"]
                    and fuzz.token_sort_ratio(row["_name_norm"], df.loc[j,"_name_norm"]) >= VN_T]
        if len(same_inv) >= 2:
            extra_groups.setdefault(i, []).append(df.loc[same_inv])

    # Merge g1_groups and extra_groups
    all_groups = {}
    for idx, gs in g1_groups.items():
        all_groups.setdefault(idx, []).extend(gs)
    for idx, gs in extra_groups.items():
        all_groups.setdefault(idx, []).extend(gs)

    return all_groups


def phase2_one_to_many(gstr2b,pr,m2b,mpr,cfg,exact_only=False,variance_only=False,label="Phase 2",log_fn=print):
    EXACT_OTM = {OTM, OTM_CB, OTM_FI, OTM_GS, OTM_SP}
    VAR_OTM   = {OTM_1, OTM_5, OTM_GS_1, OTM_GS_5, OTM_SP_1, OTM_SP_5}
    log_fn(f"\n  [{label}] One-to-Many..."
           + (" (exact)" if exact_only else " (variance)" if variance_only else ""))

    pr_um  = pr[~pr.index.isin(mpr)]
    g2b_um = gstr2b[~gstr2b.index.isin(m2b)]
    count, flagged = 0, []

    # Pre-compute candidate groups for both sides — O(n log n)
    pr_candidate_groups  = _build_otm_groups(pr_um,  cfg)
    b2_candidate_groups  = _build_otm_groups(g2b_um, cfg)

    # Direction A: one 2B row → many PR rows
    for i, r2b in g2b_um.iterrows():
        if i in m2b: continue
        # Look up PR groups where this 2B row's counterparts appear
        # We find PR groups keyed by matching GSTIN+inv from the PR side
        gs = pr_candidate_groups.get(
            # Find PR indices that would group with this 2B row
            None, []
        )
        # Direct approach: find groups in PR that match this 2B row's key
        key_gs  = pr_um[(pr_um["gstin"]     == r2b["gstin"]) &
                        (pr_um["_inv_norm"]  == r2b["_inv_norm"]) &
                        (pr_um["_inv_norm"]  != "") &
                        (pr_um["_fy"]        == r2b["_fy"]) &
                        (~pr_um.index.isin(mpr))]
        groups = []
        if len(key_gs) >= 2:
            groups.append(key_gs)

        # Also check extra groups from pr_candidate_groups that match
        for pr_idx, pr_gs_list in pr_candidate_groups.items():
            if pr_idx in mpr: continue
            pr_row = pr_um.loc[pr_idx] if pr_idx in pr_um.index else None
            if pr_row is None: continue
            if pr_row["gstin"] == r2b["gstin"] and pr_row["_inv_norm"] == r2b["_inv_norm"]:
                continue  # already covered by key_gs above
            for g in pr_gs_list:
                if g.index.isin(mpr).any(): continue
                if classify_otm(r2b, g, cfg) is not None:
                    groups.append(g)
                    break

        br, bf = None, None
        for g in groups:
            g_clean = g[~g.index.isin(mpr)]
            if len(g_clean) < 2: continue
            res = classify_otm(r2b, g_clean, cfg)
            if res is None: continue
            p, lbl, ir = res
            if exact_only    and lbl not in EXACT_OTM: continue
            if variance_only and lbl not in VAR_OTM:   continue
            if ir:
                if br is None or p < br[0]: br = (p, g_clean, lbl)
            else:
                if bf is None or p < bf[0]: bf = (p, g_clean, lbl)

        if br:
            _, g, lbl = br
            mark_2b_to_many_pr(i, list(g.index), lbl, gstr2b, pr, m2b, mpr)
            count += 1
        elif bf:
            _, g, lbl = bf
            flagged.append(("2b_to_pr", i, list(g.index), lbl))

    # Direction B: one PR row → many 2B rows
    for j, rpr in pr_um.iterrows():
        if j in mpr: continue
        key_gs = g2b_um[(g2b_um["gstin"]    == rpr["gstin"]) &
                        (g2b_um["_inv_norm"] == rpr["_inv_norm"]) &
                        (g2b_um["_inv_norm"] != "") &
                        (g2b_um["_fy"]       == rpr["_fy"]) &
                        (~g2b_um.index.isin(m2b))]
        groups = []
        if len(key_gs) >= 2:
            groups.append(key_gs)

        br, bf = None, None
        for g in groups:
            g_clean = g[~g.index.isin(m2b)]
            if len(g_clean) < 2: continue
            res = classify_otm(rpr, g_clean, cfg)
            if res is None: continue
            p, lbl, ir = res
            if exact_only    and lbl not in EXACT_OTM: continue
            if variance_only and lbl not in VAR_OTM:   continue
            if ir:
                if br is None or p < br[0]: br = (p, g_clean, lbl)
            else:
                if bf is None or p < bf[0]: bf = (p, g_clean, lbl)

        if br:
            _, g, lbl = br
            mark_pr_to_many_2b(j, list(g.index), lbl, gstr2b, pr, m2b, mpr)
            count += 1
        elif bf:
            _, g, lbl = bf
            flagged.append(("pr_to_2b", j, list(g.index), lbl))

    log_fn(f"    -> {count} groups matched.")
    return flagged


def phase3_partial_otm(gstr2b,pr,m2b,mpr,cfg,log_fn=print):
    VN_T=cfg["vendor_name_threshold"]
    log_fn(f"\n  [Phase 3] Partial One-to-Many...")
    count=0
    def _subset(target,cands,tol=1.0):
        n=len(cands)
        if n<3 or n>15: return None
        vals=[c[1] for c in cands]; idxs=[c[0] for c in cands]
        for r in range(2,n):
            for combo in combinations(range(n),r):
                if abs(sum(vals[k] for k in combo)-target)<=tol:
                    return [idxs[k] for k in combo]
        return None
    for i,r2b in gstr2b.iterrows():
        if i in m2b: continue
        fy_pr=pr[(~pr.index.isin(mpr))&(pr["_fy"]==r2b["_fy"])]; found=False
        for gm,gf in [(True,fy_pr[(fy_pr["gstin"]==r2b["gstin"])&(fy_pr["_inv_norm"]==r2b["_inv_norm"])&(fy_pr["_inv_norm"]!="")]),
                      (False,fy_pr[(fy_pr["gstin"]!=r2b["gstin"])&(fy_pr["_inv_norm"]==r2b["_inv_norm"])&(fy_pr["_inv_norm"]!="")&
                                   fy_pr["_name_norm"].apply(lambda x:fuzz.token_sort_ratio(r2b["_name_norm"],x)>=VN_T)])]:
            if len(gf)<3: continue
            cands=[(j,row["_total_tax"]) for j,row in gf.iterrows()]
            sub=_subset(r2b["_total_tax"],cands)
            if not sub: continue
            lbl=POTM_M if gm else POTM_M_GS; lo=[j for j,_ in cands if j not in sub]
            gstr2b.at[i,"Match Status"]=lbl; gstr2b.at[i,"_matched_idx_list"]=sub; m2b.add(i)
            for j in sub: pr.at[j,"Match Status"]=lbl; pr.at[j,"_matched_idx"]=i; mpr.add(j)
            for j in lo:  pr.at[j,"Match Status"]=POTM_U; pr.at[j,"_matched_idx"]=i; mpr.add(j)
            count+=1; found=True; break
        if found: continue
    for j,rpr in pr.iterrows():
        if j in mpr: continue
        fy_2b=gstr2b[(~gstr2b.index.isin(m2b))&(gstr2b["_fy"]==rpr["_fy"])]
        for gm,gf in [(True,fy_2b[(fy_2b["gstin"]==rpr["gstin"])&(fy_2b["_inv_norm"]==rpr["_inv_norm"])&(fy_2b["_inv_norm"]!="")]),
                      (False,fy_2b[(fy_2b["gstin"]!=rpr["gstin"])&(fy_2b["_inv_norm"]==rpr["_inv_norm"])&(fy_2b["_inv_norm"]!="")&
                                   fy_2b["_name_norm"].apply(lambda x:fuzz.token_sort_ratio(rpr["_name_norm"],x)>=VN_T)])]:
            if len(gf)<3: continue
            cands=[(i,row["_total_tax"]) for i,row in gf.iterrows()]
            sub=_subset(rpr["_total_tax"],cands)
            if not sub: continue
            lbl=POTM_M if gm else POTM_M_GS; lo=[i for i,_ in cands if i not in sub]
            pr.at[j,"Match Status"]=lbl; pr.at[j,"_matched_idx_list"]=sub; mpr.add(j)
            for i in sub: gstr2b.at[i,"Match Status"]=lbl; gstr2b.at[i,"_matched_idx"]=j; m2b.add(i)
            for i in lo:  gstr2b.at[i,"Match Status"]=POTM_U; gstr2b.at[i,"_matched_idx"]=j; m2b.add(i)
            count+=1; break
    log_fn(f"    -> {count} partial groups found.")


def phase4_flags(gstr2b,pr,m2b,mpr,cfg,flagged_otm,log_fn=print):
    """
    Phase 4 — Flags suspicious near-matches + Bucket 4 edge cases.
    Uses same bucket structure as Phase 1 — no full O(n×m) scan.
    """
    log_fn(f"\n  [Phase 4] Flagging and edge-case scan...")
    f2b, fpr, count = set(), set(), 0
    TAX_RATIO_LIMIT = 10.0

    pr_um = pr[~pr.index.isin(mpr)]

    # Build buckets for remaining unmatched PR rows
    bg, bp, bi = {}, {}, {}
    for j, rpr in pr_um.iterrows():
        bg.setdefault(rpr["gstin"], []).append(j)
        if rpr["_pan"]:      bp.setdefault(rpr["_pan"], []).append(j)
        if rpr["_inv_norm"]: bi.setdefault(rpr["_inv_norm"], []).append(j)

    # Also build a suffix-prefix index: map each PR inv_norm to its index
    # for detecting prefix/suffix with full GSTIN+PAN mismatch (Bucket 4)
    pr_inv_list = [(j, rpr["_inv_norm"], rpr["_total_tax"], rpr["gstin"])
                   for j, rpr in pr_um.iterrows() if rpr["_inv_norm"]]

    def cands_phase4(r2b):
        """
        Buckets 1-3 + Bucket 4 (suffix/prefix only, limited to short list).
        """
        seen, res = set(), []
        # Bucket 1
        for j in bg.get(r2b["gstin"], []):
            if j not in seen: seen.add(j); res.append(j)
        # Bucket 2
        if r2b["_pan"]:
            for j in bp.get(r2b["_pan"], []):
                if pr_um.loc[j,"gstin"] != r2b["gstin"] and j not in seen:
                    seen.add(j); res.append(j)
        # Bucket 3
        if r2b["_inv_norm"]:
            for j in bi.get(r2b["_inv_norm"], []):
                if pr_um.loc[j,"gstin"] != r2b["gstin"] and j not in seen:
                    seen.add(j); res.append(j)
        # Bucket 4: suffix/prefix candidates (GSTIN+PAN both miss)
        # Only check if r2b invoice is long enough to have a prefix
        if len(r2b["_inv_norm"]) >= 3:
            for j, inv, tax, gstin in pr_inv_list:
                if j in seen: continue
                if gstin == r2b["gstin"]: continue  # already in Bucket 1
                # Tax pre-filter
                if r2b["_total_tax"] > 0 and tax > 0:
                    if max(r2b["_total_tax"], tax) / min(r2b["_total_tax"], tax) > TAX_RATIO_LIMIT:
                        continue
                if is_suffix_prefix_match(r2b["_inv_norm"], inv):
                    seen.add(j); res.append(j)
        return res

    for i, r2b in gstr2b.iterrows():
        if i in m2b or i in f2b:
            continue
        r2b_tax = r2b["_total_tax"]
        best_real = None
        best_flag = None

        for j in cands_phase4(r2b):
            if j in mpr or j in fpr or j not in pr_um.index:
                continue
            rpr = pr_um.loc[j]
            # Tax pre-filter
            rpr_tax = rpr["_total_tax"]
            if r2b_tax > 0 and rpr_tax > 0:
                if max(r2b_tax, rpr_tax) / min(r2b_tax, rpr_tax) > TAX_RATIO_LIMIT:
                    continue
            res = classify_pair(r2b, rpr, cfg)
            if res is None: continue
            p, lbl, ir = res
            if ir:
                if best_real is None or p < best_real[0]: best_real = (p, j, lbl)
            else:
                if best_flag is None or p < best_flag[0]: best_flag = (p, j, lbl)

        if best_real:
            _, j, lbl = best_real
            mark_1to1(i, j, lbl, gstr2b, pr, m2b, mpr)
            count += 1
        elif best_flag:
            _, j, lbl = best_flag
            gstr2b.at[i,"Match Status"] = lbl; gstr2b.at[i,"_matched_idx"] = j
            pr.at[j,"Match Status"]     = lbl; pr.at[j,"_matched_idx"]     = i
            f2b.add(i); fpr.add(j); count += 1

    # OTM flags from Phase 2
    for direction, anchor, others, lbl in flagged_otm:
        if direction == "2b_to_pr":
            i  = anchor
            js = [j for j in others if j not in mpr and j not in fpr]
            if i in m2b or i in f2b or not js: continue
            gstr2b.at[i,"Match Status"]      = lbl
            gstr2b.at[i,"_matched_idx_list"] = js
            f2b.add(i)
            for j in js:
                pr.at[j,"Match Status"] = lbl; pr.at[j,"_matched_idx"] = i; fpr.add(j)
            count += 1
        else:
            j   = anchor
            is_ = [i for i in others if i not in m2b and i not in f2b]
            if j in mpr or j in fpr or not is_: continue
            pr.at[j,"Match Status"]      = lbl
            pr.at[j,"_matched_idx_list"] = is_
            fpr.add(j)
            for i in is_:
                gstr2b.at[i,"Match Status"] = lbl; gstr2b.at[i,"_matched_idx"] = j; f2b.add(i)
            count += 1

    log_fn(f"    -> {count} matched/flagged.")

# =============================================================================
# ORCHESTRATOR
# =============================================================================

def run_reconciliation(gstr2b, pr, cfg=None, log_fn=print, progress_fn=None):
    if cfg is None: cfg = DEFAULT_CFG.copy()
    m2b, mpr = set(), set()
    _phase1_run(gstr2b,pr,m2b,mpr,cfg,exact_only=True,    label="Phase 1A",log_fn=log_fn,progress_fn=progress_fn)
    flagged=phase2_one_to_many(gstr2b,pr,m2b,mpr,cfg,exact_only=True,    label="Phase 2A",log_fn=log_fn)
    _phase1_run(gstr2b,pr,m2b,mpr,cfg,variance_only=True, label="Phase 1B",log_fn=log_fn,progress_fn=progress_fn)
    flagged+=phase2_one_to_many(gstr2b,pr,m2b,mpr,cfg,variance_only=True, label="Phase 2B",log_fn=log_fn)
    phase3_partial_otm(gstr2b,pr,m2b,mpr,cfg,log_fn=log_fn)
    phase4_flags(gstr2b,pr,m2b,mpr,cfg,flagged,log_fn=log_fn)
    return gstr2b, pr

# =============================================================================
# OUTPUT
# =============================================================================

def build_output(gstr2b, pr, output_path=None, log_fn=print):
    """
    output_path=None  → returns BytesIO (web app)
    output_path=str   → writes to file (standalone)
    Sheet order: Summary FIRST, then 2B Recon, then PR Recon.
    """
    def get_val(row, col, src):
        il = row["_matched_idx_list"]
        if il is not None:
            if col in ("igst","cgst","sgst"): return round(sum(float(src.at[k,col] or 0) for k in il),2)
            if col in ("invoice_no","month"):  return " | ".join(dict.fromkeys(str(src.at[k,col]) for k in il))
            return src.at[il[0],col]
        idx=row["_matched_idx"]
        return src.at[idx,col] if idx is not None else ""

    gstr2b["PR Booking Month"]  = gstr2b.apply(lambda r: get_val(r,"month",     pr),    axis=1)
    gstr2b["PR Invoice Number"] = gstr2b.apply(lambda r: get_val(r,"invoice_no",pr),    axis=1)
    gstr2b["PR IGST"]           = gstr2b.apply(lambda r: get_val(r,"igst",      pr),    axis=1)
    gstr2b["PR CGST"]           = gstr2b.apply(lambda r: get_val(r,"cgst",      pr),    axis=1)
    gstr2b["PR SGST"]           = gstr2b.apply(lambda r: get_val(r,"sgst",      pr),    axis=1)
    pr["2B Month"]          = pr.apply(lambda r: get_val(r,"month",     gstr2b),axis=1)
    pr["2B Invoice Number"] = pr.apply(lambda r: get_val(r,"invoice_no",gstr2b),axis=1)
    pr["2B IGST"]           = pr.apply(lambda r: get_val(r,"igst",      gstr2b),axis=1)
    pr["2B CGST"]           = pr.apply(lambda r: get_val(r,"cgst",      gstr2b),axis=1)
    pr["2B SGST"]           = pr.apply(lambda r: get_val(r,"sgst",      gstr2b),axis=1)

    drop=["_inv_norm","_inv_aggr","_fy","_name_norm","_pan","_total_tax","_matched_idx","_matched_idx_list"]
    go=gstr2b.drop(columns=drop,errors="ignore").rename(columns={k:v for k,v in GSTR2B_COLUMN_MAP.items()})
    po=pr.drop(columns=drop,errors="ignore").rename(columns={k:v for k,v in PR_COLUMN_MAP.items()})

    def pr_(df,cols):
        ex=[c for c in cols if c in df.columns]
        return df[[c for c in df.columns if c not in ex]+ex]
    go=pr_(go,["Match Status","PR Booking Month","PR Invoice Number","PR IGST","PR CGST","PR SGST"])
    po=pr_(po,["Match Status","2B Month","2B Invoice Number","2B IGST","2B CGST","2B SGST"])

    srows=[]
    for s in ALL_STATUSES:
        c2=int((gstr2b["Match Status"]==s).sum()); cp=int((pr["Match Status"]==s).sum())
        if c2>0 or cp>0:
            srows.append({"Match Status":s,"Count in GSTR 2B":c2,"Count in PR":cp,
                          "Action Required":"⚠ Review Required" if s in FLAG_LABELS else "✅ Matched",
                          "Explanation":STATUS_REMARKS.get(s,"")})
    if not srows:
        srows.append({"Match Status":UNMATCHED,"Count in GSTR 2B":0,"Count in PR":0,
                      "Action Required":"","Explanation":STATUS_REMARKS[UNMATCHED]})
    sdf=pd.DataFrame(srows)

    target = BytesIO() if output_path is None else output_path
    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        sdf.to_excel(writer, sheet_name="Summary",  index=False)   # FIRST
        go.to_excel( writer, sheet_name="2B Recon", index=False)
        po.to_excel( writer, sheet_name="PR Recon", index=False)
        for sheet in writer.sheets.values():
            sheet.freeze_panes="A2"
            for cc in sheet.columns:
                ml=max((len(str(c.value or "")) for c in cc),default=10)
                sheet.column_dimensions[cc[0].column_letter].width=min(ml+4,60)

    log_fn("  Output ready.")
    if output_path is None:
        target.seek(0); return target


def create_template():
    """
    Creates a downloadable Excel template with correct column headers.
    Returns BytesIO object.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    b2_cols = [
        "Sr. No", "Supplier GSTN", "Supplier Name", "Invoice Number",
        "Invoice Date", "Taxable Value", "IGST", "CGST", "SGST", "2B Month"
    ]
    pr_cols = [
        "Sr. No", "Supplier GSTN", "Supplier Name", "Invoice Number",
        "Invoice Date", "Taxable Value", "IGST", "CGST", "SGST", "Booking Month"
    ]

    wb = openpyxl.Workbook()

    for sheet_name, cols in [("2B", b2_cols), ("PR", pr_cols)]:
        ws = wb.create_sheet(title=sheet_name)

        # Write headers
        for col_idx, col_name in enumerate(cols, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font      = Font(bold=True, color="FFFFFF")
            cell.fill      = PatternFill("solid", fgColor="1A5276")
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[get_column_letter(col_idx)].width = 20

        # Write one sample row
        sample = [
            1, "27AABCU9603R1ZX", "ABC Traders", "INV/2025/001",
            "01-04-2025", 100000, 18000, 0, 0,
            "Apr-2025" if sheet_name == "2B" else "Apr-2025"
        ]
        for col_idx, val in enumerate(sample, start=1):
            ws.cell(row=2, column=col_idx, value=val)

    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def get_summary_dict(gstr2b, pr):
    rows=[]
    for s in ALL_STATUSES:
        c2=int((gstr2b["Match Status"]==s).sum()); cp=int((pr["Match Status"]==s).sum())
        if c2>0 or cp>0:
            rows.append({"Match Status":s,"Count in GSTR 2B":c2,"Count in PR":cp,
                         "Flag":"⚠" if s in FLAG_LABELS else "✅",
                         "Explanation":STATUS_REMARKS.get(s,"")})
    return rows

# =============================================================================
# STANDALONE
# =============================================================================

def main():
    print("="*68)
    print("  GSTR 2B vs Purchase Register — Reconciliation Tool v5.0")
    print("="*68)
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"\nFile not found: '{INPUT_FILE}'\nEnsure 'Raw Data.xlsx' is in:\n{FOLDER}")
    cfg = DEFAULT_CFG.copy()
    print("\n[1] Loading Data...")
    gstr2b = load_and_prepare(INPUT_FILE, SHEET_2B, GSTR2B_COLUMN_MAP, "GSTR 2B")
    pr     = load_and_prepare(INPUT_FILE, SHEET_PR,  PR_COLUMN_MAP,     "Purchase Register")
    print("\n[2] Running Reconciliation...")
    gstr2b, pr = run_reconciliation(gstr2b, pr, cfg)
    print("\n[3] Generating Output...")
    build_output(gstr2b, pr, output_path=OUTPUT_FILE)
    print("\n"+"="*68+"  SUMMARY\n"+"="*68)
    for row in get_summary_dict(gstr2b, pr):
        print(f"  {row['Match Status']:<54}  2B:{row['Count in GSTR 2B']:>3}  PR:{row['Count in PR']:>3}  {row['Flag']}")
    print("="*68+"\nDone.\n")

if __name__ == "__main__":
    main()
