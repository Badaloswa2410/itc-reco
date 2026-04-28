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
    log_fn(f"\n  [{label}] One-to-One..."
           + (" (exact)" if exact_only else " (variance)" if variance_only else ""))
    pr_um = pr[~pr.index.isin(mpr)]

    # ── Optimisation 2: Pre-group PR rows by Financial Year ───────────────────
    # We only ever compare rows within the same FY (DFY is the only cross-FY
    # case and is handled separately). Pre-grouping avoids FY checks inside
    # the hot loop.
    pr_by_fy = {}
    for j, rpr in pr_um.iterrows():
        pr_by_fy.setdefault(rpr["_fy"], []).append(j)

    # ── Buckets 1-3 (Bucket 4 removed — moved to Phase 4) ────────────────────
    # Bucket 1: GSTIN          → all GSTIN-match cases
    # Bucket 2: PAN            → PAN-match, GSTIN differs
    # Bucket 3: Invoice norm   → vendor-name match (invoice must align)
    # Bucket 4 removal: suffix/prefix with full GSTIN+PAN mismatch is now
    # flagged in Phase 4 only — safer (uncertain match) and much faster.
    bg, bp, bi = {}, {}, {}
    for j, rpr in pr_um.iterrows():
        bg.setdefault(rpr["gstin"], []).append(j)
        if rpr["_pan"]:
            bp.setdefault(rpr["_pan"], []).append(j)
        if rpr["_inv_norm"]:
            bi.setdefault(rpr["_inv_norm"], []).append(j)

    # ── Optimisation 3: Tax pre-filter threshold ──────────────────────────────
    # If total tax of the two rows differs by more than 10x, they can never
    # match under any of our rules (max variance is 5%). Skip classify_pair
    # entirely. This avoids expensive fuzzy comparisons on clearly wrong pairs.
    TAX_RATIO_LIMIT = 10.0

    def tax_plausible(tax_a, tax_b):
        """Returns True if taxes are close enough to bother comparing."""
        if tax_a == 0 and tax_b == 0:
            return True
        if tax_a == 0 or tax_b == 0:
            return False
        ratio = max(tax_a, tax_b) / min(tax_a, tax_b)
        return ratio <= TAX_RATIO_LIMIT

    def cands(r2b):
        """
        Returns candidate PR indices for a given 2B row.
        Uses buckets 1-3 only. No full O(n×m) scan.
        """
        seen, res = set(), []
        # Bucket 1: same GSTIN
        for j in bg.get(r2b["gstin"], []):
            if j not in seen:
                seen.add(j); res.append(j)
        # Bucket 2: same PAN, GSTIN differs
        if r2b["_pan"]:
            for j in bp.get(r2b["_pan"], []):
                if pr_um.loc[j, "gstin"] != r2b["gstin"] and j not in seen:
                    seen.add(j); res.append(j)
        # Bucket 3: same invoice norm, GSTIN differs (vendor name cases)
        if r2b["_inv_norm"]:
            for j in bi.get(r2b["_inv_norm"], []):
                if pr_um.loc[j, "gstin"] != r2b["gstin"] and j not in seen:
                    seen.add(j); res.append(j)
        return res

    pairs = []
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

            # ── Optimisation 2: FY check using pre-group ──────────────────
            # Allow same FY or DFY (cross FY exact match) — skip everything else
            rpr_fy = rpr["_fy"]
            if rpr_fy != r2b_fy:
                # Only worth checking for DFY: needs exact invoice + exact tax
                if rpr["_inv_norm"] != r2b["_inv_norm"]:
                    continue

            # ── Optimisation 3: Tax pre-filter ────────────────────────────
            if not tax_plausible(r2b_tax, rpr["_total_tax"]):
                continue

            res = classify_pair(r2b, rpr, cfg)
            if res is None:
                continue
            p, lbl, ir = res
            if not ir:
                continue
            if exact_only    and lbl in VARIANCE_LABELS: continue
            if variance_only and lbl not in VARIANCE_LABELS: continue
            pairs.append((p, i, j, lbl))

    if progress_fn:
        progress_fn(label, total_rows, total_rows)

    pairs.sort(key=lambda x: x[0])
    lc = {}
    for p, i, j, lbl in pairs:
        if i in m2b or j in mpr:
            continue
        mark_1to1(i, j, lbl, gstr2b, pr, m2b, mpr)
        lc[lbl] = lc.get(lbl, 0) + 1
    log_fn(f"    -> {sum(lc.values())} matched.")


def _find_groups(one_row,other_df,cfg):
    FI,VN_T=cfg["fuzzy_inv_threshold"],cfg["vendor_name_threshold"]
    pool=other_df[other_df["_fy"]==one_row["_fy"]]; gs=[]
    g=pool[(pool["gstin"]==one_row["gstin"])&(pool["_inv_norm"]==one_row["_inv_norm"])&(pool["_inv_norm"]!="")]
    if len(g)>=2: gs.append(g)
    g=pool[(pool["gstin"]==one_row["gstin"])&(pool["_inv_norm"]!=one_row["_inv_norm"])&
           pool["_inv_aggr"].apply(lambda x:fuzz.ratio(one_row["_inv_aggr"],x)>=FI)]
    if len(g)>=2: gs.append(g)
    g=pool[(pool["gstin"]==one_row["gstin"])&(pool["_inv_norm"]!=one_row["_inv_norm"])&
           pool["_inv_norm"].apply(lambda x:is_suffix_prefix_match(one_row["_inv_norm"],x))]
    if len(g)>=2: gs.append(g)
    g=pool[(pool["gstin"]!=one_row["gstin"])&(pool["_inv_norm"]==one_row["_inv_norm"])&(pool["_inv_norm"]!="")&
           pool["_name_norm"].apply(lambda x:fuzz.token_sort_ratio(one_row["_name_norm"],x)>=VN_T)]
    if len(g)>=2: gs.append(g)
    g=pool[(pool["gstin"]!=one_row["gstin"])&
           pool["_name_norm"].apply(lambda x:fuzz.token_sort_ratio(one_row["_name_norm"],x)>=VN_T)&
           pool["_inv_aggr"].apply(lambda x:fuzz.ratio(one_row["_inv_aggr"],x)>=FI)]
    if len(g)>=2: gs.append(g)
    g=pool[(pool["gstin"]!=one_row["gstin"])&
           pool["_name_norm"].apply(lambda x:fuzz.token_sort_ratio(one_row["_name_norm"],x)>=VN_T)&
           pool["_inv_norm"].apply(lambda x:is_suffix_prefix_match(one_row["_inv_norm"],x))]
    if len(g)>=2: gs.append(g)
    return gs


def phase2_one_to_many(gstr2b,pr,m2b,mpr,cfg,exact_only=False,variance_only=False,label="Phase 2",log_fn=print):
    EXACT_OTM={OTM,OTM_CB,OTM_FI,OTM_GS,OTM_SP}
    VAR_OTM  ={OTM_1,OTM_5,OTM_GS_1,OTM_GS_5,OTM_SP_1,OTM_SP_5}
    log_fn(f"\n  [{label}] One-to-Many..."
           + (" (exact)" if exact_only else " (variance)" if variance_only else ""))
    count,flagged=0,[]
    for i,r2b in gstr2b.iterrows():
        if i in m2b: continue
        gs=_find_groups(r2b,pr[~pr.index.isin(mpr)],cfg)
        br,bf=None,None
        for g in gs:
            res=classify_otm(r2b,g,cfg)
            if res is None: continue
            p,lbl,ir=res
            if exact_only    and lbl not in EXACT_OTM: continue
            if variance_only and lbl not in VAR_OTM:   continue
            if ir:
                if br is None or p<br[0]: br=(p,g,lbl)
            else:
                if bf is None or p<bf[0]: bf=(p,g,lbl)
        if br: _,g,lbl=br; mark_2b_to_many_pr(i,list(g.index),lbl,gstr2b,pr,m2b,mpr); count+=1
        elif bf: _,g,lbl=bf; flagged.append(("2b_to_pr",i,list(g.index),lbl))
    for j,rpr in pr.iterrows():
        if j in mpr: continue
        gs=_find_groups(rpr,gstr2b[~gstr2b.index.isin(m2b)],cfg)
        br,bf=None,None
        for g in gs:
            res=classify_otm(rpr,g,cfg)
            if res is None: continue
            p,lbl,ir=res
            if exact_only    and lbl not in EXACT_OTM: continue
            if variance_only and lbl not in VAR_OTM:   continue
            if ir:
                if br is None or p<br[0]: br=(p,g,lbl)
            else:
                if bf is None or p<bf[0]: bf=(p,g,lbl)
        if br: _,g,lbl=br; mark_pr_to_many_2b(j,list(g.index),lbl,gstr2b,pr,m2b,mpr); count+=1
        elif bf: _,g,lbl=bf; flagged.append(("pr_to_2b",j,list(g.index),lbl))
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
    Phase 4 — Flags suspicious near-matches for manual review.
    Also handles Bucket 4 (full O(n×m) scan on remaining rows) which was
    removed from Phase 1 for performance. By Phase 4 the unmatched pool is
    small so the full scan is fast.
    """
    log_fn(f"\n  [Phase 4] Flagging and edge-case scan...")
    f2b, fpr, count = set(), set(), 0

    # ── One-to-one flags + Bucket 4 edge cases ───────────────────────────────
    # We scan all remaining unmatched pairs. classify_pair returns:
    #   is_real=True  → confirmed match missed by buckets 1-3 (e.g. suffix/prefix
    #                   with full GSTIN+PAN+name mismatch) → mark as confirmed
    #   is_real=False → flag for manual review
    TAX_RATIO_LIMIT = 10.0

    for i, r2b in gstr2b.iterrows():
        if i in m2b or i in f2b:
            continue
        r2b_tax = r2b["_total_tax"]
        best_real = None   # (priority, j, label) — confirmed late match
        best_flag = None   # (priority, j, label) — flag

        for j, rpr in pr[~pr.index.isin(mpr | fpr)].iterrows():
            # Tax pre-filter
            if r2b_tax > 0 and rpr["_total_tax"] > 0:
                ratio = max(r2b_tax, rpr["_total_tax"]) / min(r2b_tax, rpr["_total_tax"])
                if ratio > TAX_RATIO_LIMIT:
                    continue
            res = classify_pair(r2b, rpr, cfg)
            if res is None:
                continue
            p, lbl, ir = res
            if ir:
                if best_real is None or p < best_real[0]:
                    best_real = (p, j, lbl)
            else:
                if best_flag is None or p < best_flag[0]:
                    best_flag = (p, j, lbl)

        if best_real:
            # Confirmed match found in Phase 4 (Bucket 4 edge case)
            _, j, lbl = best_real
            mark_1to1(i, j, lbl, gstr2b, pr, m2b, mpr)
            count += 1
        elif best_flag:
            _, j, lbl = best_flag
            gstr2b.at[i, "Match Status"] = lbl
            gstr2b.at[i, "_matched_idx"] = j
            pr.at[j,     "Match Status"] = lbl
            pr.at[j,     "_matched_idx"] = i
            f2b.add(i); fpr.add(j); count += 1

    # ── One-to-many flags from Phase 2 ───────────────────────────────────────
    for direction, anchor, others, lbl in flagged_otm:
        if direction == "2b_to_pr":
            i  = anchor
            js = [j for j in others if j not in mpr and j not in fpr]
            if i in m2b or i in f2b or not js: continue
            gstr2b.at[i, "Match Status"]      = lbl
            gstr2b.at[i, "_matched_idx_list"] = js
            f2b.add(i)
            for j in js:
                pr.at[j, "Match Status"] = lbl
                pr.at[j, "_matched_idx"] = i
                fpr.add(j)
            count += 1
        else:
            j   = anchor
            is_ = [i for i in others if i not in m2b and i not in f2b]
            if j in mpr or j in fpr or not is_: continue
            pr.at[j,     "Match Status"]      = lbl
            pr.at[j,     "_matched_idx_list"] = is_
            fpr.add(j)
            for i in is_:
                gstr2b.at[i, "Match Status"] = lbl
                gstr2b.at[i, "_matched_idx"] = j
                f2b.add(i)
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
