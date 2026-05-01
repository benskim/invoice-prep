import streamlit as st
import pandas as pd
import re
import unicodedata
from collections import defaultdict
import pytesseract
import pdf2image

st.set_page_config(layout="wide")
st.title("🔍 PO vs Invoice Matching (Real PoC)")

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
# Column Auto Detection
# -----------------------------
def detect_columns(df):
    part_keywords = ["PART", "P/N", "PN", "MODEL", "ITEM"]
    qty_keywords = ["QTY", "QUANTITY", "EA", "PCS", "수량"]

    part_col, qty_col = None, None

    for col in df.columns:
        c = col.upper()
        if any(k in c for k in part_keywords):
            part_col = col
        if any(k in c for k in qty_keywords):
            qty_col = col

    return part_col, qty_col

# -----------------------------
# OCR PDF
# -----------------------------
def pdf_to_df(pdf_file):
    images = pdf2image.convert_from_bytes(pdf_file.read())
    rows = []

    for img in images:
        text = pytesseract.image_to_string(img)
        for line in text.split("\n"):
            parts = re.findall(r'[A-Z0-9\-/]+|\d+', line.upper())
            if len(parts) >= 2 and parts[-1].isdigit():
                rows.append((parts[0], int(parts[-1])))

    return pd.DataFrame(rows, columns=["Part Number", "Qty"])

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

def score(a, b):
    t1, t2 = tokenize(a), tokenize(b)
    ps = prefix_sim(a, b)
    ts = token_sim(t1, t2)
    return 0.7 * ps + 0.3 * ts

# -----------------------------
# UI Upload
# -----------------------------
po_file = st.file_uploader("Upload PO (Excel)", type=["xlsx"])
inv_excel = st.file_uploader("Upload Invoice (Excel)", type=["xlsx"])
inv_pdf = st.file_uploader("Upload Invoice (PDF)", type=["pdf"])

if po_file:
    po_df = pd.read_excel(po_file)
    p_col, q_col = detect_columns(po_df)

    po_df["Part Number"] = po_df[p_col]
    po_df["Qty"] = po_df[q_col]

    if inv_excel:
        inv_df = pd.read_excel(inv_excel)
        p_col, q_col = detect_columns(inv_df)

        inv_df["Part Number"] = inv_df[p_col]
        inv_df["Qty"] = inv_df[q_col]

    elif inv_pdf:
        inv_df = pdf_to_df(inv_pdf)
    else:
        st.stop()

    # normalize
    po_df["norm"] = po_df["Part Number"].apply(normalize)
    inv_df["norm"] = inv_df["Part Number"].apply(normalize)

    # index
    def key(x):
        return x[:6]

    inv_index = defaultdict(list)
    for i, r in inv_df.iterrows():
        inv_index[key(r["norm"])].append((i, r))

    used = set()
    results = []

    for _, po in po_df.iterrows():
        candidates = []

        for i, inv in inv_index[key(po["norm"])]:
            if i in used:
                continue
            s = score(po["norm"], inv["norm"])
            if s > 0.75:
                candidates.append((i, inv, s))

        candidates.sort(key=lambda x: x[2], reverse=True)

        matched = False

        # 1:1
        for i, inv, s in candidates:
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

    df = pd.DataFrame(results, columns=["Part", "Status", "Score"])
    st.dataframe(df)
