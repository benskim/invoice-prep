import streamlit as st
import pandas as pd
import re
import unicodedata
from collections import defaultdict

st.title("🔍 PO vs Invoice Matching (Excel Only)")

# -----------------------------
# Normalize
# -----------------------------
def normalize(s):
    if pd.isna(s) or not str(s).strip():
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.upper()
    s = " ".join(s.split())
    s = re.sub(r'[-_/:\s]+', '|', s)
    s = re.sub(r'[^A-Z0-9|]', '', s)
    return re.sub(r'\|+', '|', s).strip('|')

def tokenize(s):
    return s.split("|") if s else []

# -----------------------------
# Column Detection
# -----------------------------
def detect_columns(df):
    part_k = ["PART", "P/N", "MODEL", "ITEM"]
    qty_k = ["QTY", "QUANTITY", "EA", "PCS", "수량"]
    po_k = ["PO", "ORDER"]
    vendor_k = ["VENDOR", "SUPPLIER"]

    cols = {"part": None, "qty": None, "po": None, "vendor": None}

    for col in df.columns:
        c = col.upper()
        if any(k in c for k in part_k):
            cols["part"] = col
        elif any(k in c for k in qty_k):
            cols["qty"] = col
        elif any(k in c for k in po_k):
            cols["po"] = col
        elif any(k in c for k in vendor_k):
            cols["vendor"] = col

    return cols

# -----------------------------
# Similarity
# -----------------------------
def prefix_sim(a, b):
    score = 0
    total = 0
    max_len = max(len(a), len(b))
    for i in range(min(len(a), len(b))):
        w = max_len - i
        total += w
        if a[i] == b[i]:
            score += w
        else:
            break
    return score / total if total else 0

def token_sim(t1, t2):
    match = 0
    for i in range(min(len(t1), len(t2))):
        if t1[i] == t2[i]:
            match += 1
        else:
            break
    return match / max(len(t1), len(t2)) if t1 and t2 else 0

def sim_score(a, b):
    t1, t2 = tokenize(a), tokenize(b)
    return 0.7 * prefix_sim(a, b) + 0.3 * token_sim(t1, t2)

# -----------------------------
# Upload
# -----------------------------
po_file = st.file_uploader("Upload PO Excel", type=["xlsx"])
inv_file = st.file_uploader("Upload Invoice Excel", type=["xlsx"])

if po_file and inv_file:

    po_df = pd.read_excel(po_file)
    inv_df = pd.read_excel(inv_file)

    po_map = detect_columns(po_df)
    inv_map = detect_columns(inv_df)

    po_df["Part Number"] = po_df[po_map["part"]]
    po_df["Qty"] = po_df[po_map["qty"]]
    po_df["Vendor"] = po_df[po_map["vendor"]] if po_map["vendor"] else "UNKNOWN"
    po_df["PO"] = po_df[po_map["po"]] if po_map["po"] else "UNKNOWN"

    inv_df["Part Number"] = inv_df[inv_map["part"]]
    inv_df["Qty"] = inv_df[inv_map["qty"]]
    inv_df["Vendor"] = inv_df[inv_map["vendor"]] if inv_map["vendor"] else "UNKNOWN"
    inv_df["PO"] = inv_df[inv_map["po"]] if inv_map["po"] else "UNKNOWN"

    po_df["norm"] = po_df["Part Number"].apply(normalize)
    inv_df["norm"] = inv_df["Part Number"].apply(normalize)

    used = set()
    results = []

    for _, po in po_df.iterrows():

        candidates = []

        for i, inv in inv_df.iterrows():

            if po["PO"] != inv["PO"]:
                continue
            if po["Vendor"] != inv["Vendor"]:
                continue

            s = sim_score(po["norm"], inv["norm"])
            if s > 0.75:
                candidates.append((i, inv, s))

        candidates.sort(key=lambda x: x[2], reverse=True)

        matched = False

        # 1:1
        for i, inv, s in candidates:
            if i in used:
                continue
            if po["Qty"] == inv["Qty"] and s > 0.9:
                used.add(i)
                results.append([po["Part Number"], "AUTO", s])
                matched = True
                break

        if matched:
            continue

        # 1:N
        total = 0
        selected = []
        for i, inv, s in candidates:
            if i in used:
                continue
            total += inv["Qty"]
            selected.append(i)
            if total == po["Qty"]:
                for x in selected:
                    used.add(x)
                results.append([po["Part Number"], "AUTO(1:N)", s])
                matched = True
                break

        if not matched:
            results.append([po["Part Number"], "REVIEW", 0])

    result_df = pd.DataFrame(results, columns=["Part", "Status", "Score"])
    st.dataframe(result_df)
