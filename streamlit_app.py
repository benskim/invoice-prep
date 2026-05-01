import streamlit as st
import pandas as pd
import re
import unicodedata
from collections import defaultdict

# PDF optional
try:
    import pytesseract
    import pdf2image
    PDF_ENABLED = True
except:
    PDF_ENABLED = False

st.title("🔍 PO vs Invoice Matching (PDF + Excel)")

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
# Column Detection (Excel)
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
# PDF → DataFrame
# -----------------------------
def pdf_to_df(pdf_file):
    images = pdf2image.convert_from_bytes(pdf_file.read())
    rows = []
    vendor = None
    po_number = None

    for img in images:
        text = pytesseract.image_to_string(img)

        for line in text.split("\n"):
            line = line.strip()

            # Vendor 추출
            if "VENDOR" in line.upper():
                vendor = line.split(":")[-1].strip()

            # PO 번호 추출
            if "PO" in line.upper():
                po_number = line.split(":")[-1].strip()

            # Part + Qty 추출
            match = re.match(r'(.+?)\s+(\d+)$', line.upper())
            if match:
                part = match.group(1).strip()
                qty = int(match.group(2))
                rows.append((part, qty, vendor, po_number))

    df = pd.DataFrame(rows, columns=["Part Number", "Qty", "Vendor", "PO"])
    return df

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
# File Upload
# -----------------------------
st.subheader("Upload PO")
po_excel = st.file_uploader("PO Excel", type=["xlsx"], key="po_excel")
po_pdf = st.file_uploader("PO PDF", type=["pdf"], key="po_pdf")

st.subheader("Upload Invoice")
inv_excel = st.file_uploader("Invoice Excel", type=["xlsx"], key="inv_excel")
inv_pdf = st.file_uploader("Invoice PDF", type=["pdf"], key="inv_pdf")

# -----------------------------
# Load PO
# -----------------------------
po_df = None

if po_excel:
    po_df = pd.read_excel(po_excel)
    m = detect_columns(po_df)
    po_df["Part Number"] = po_df[m["part"]]
    po_df["Qty"] = po_df[m["qty"]]
    po_df["Vendor"] = po_df[m["vendor"]] if m["vendor"] else "UNKNOWN"
    po_df["PO"] = po_df[m["po"]] if m["po"] else "UNKNOWN"

elif po_pdf and PDF_ENABLED:
    po_df = pdf_to_df(po_pdf)

# -----------------------------
# Load Invoice
# -----------------------------
inv_df = None

if inv_excel:
    inv_df = pd.read_excel(inv_excel)
    m = detect_columns(inv_df)
    inv_df["Part Number"] = inv_df[m["part"]]
    inv_df["Qty"] = inv_df[m["qty"]]
    inv_df["Vendor"] = inv_df[m["vendor"]] if m["vendor"] else "UNKNOWN"
    inv_df["PO"] = inv_df[m["po"]] if m["po"] else "UNKNOWN"

elif inv_pdf and PDF_ENABLED:
    inv_df = pdf_to_df(inv_pdf)

# -----------------------------
# Matching
# -----------------------------
if po_df is not None and inv_df is not None:

    po_df["norm"] = po_df["Part Number"].apply(normalize)
    inv_df["norm"] = inv_df["Part Number"].apply(normalize)

    used = set()
    results = []

    for _, po in po_df.iterrows():

        candidates = []

        for i, inv in inv_df.iterrows():

            # PO / Vendor 필터
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

    df = pd.DataFrame(results, columns=["Part", "Status", "Score"])
    st.dataframe(df)

elif (po_pdf or inv_pdf) and not PDF_ENABLED:
    st.error("PDF 기능을 사용하려면 pytesseract + pdf2image 설치 필요")
