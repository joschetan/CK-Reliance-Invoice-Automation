import streamlit as st
import pandas as pd
import re
import io
import pypdf
from collections import defaultdict

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="CK Reliance Invoice Automation",
    page_icon="📊",
    layout="centered"
)

st.markdown("""
    <style>
    .main-title {
        color: #1e3a8a;
        text-align: center;
        font-family: 'Segoe UI', Arial, sans-serif;
        font-weight: 700;
        margin-bottom: 2px;
    }
    .sub-title {
        color: #64748b;
        text-align: center;
        font-size: 14px;
        margin-bottom: 30px;
    }
    .stButton>button {
        background-color: #10b981 !important;
        color: white !important;
        font-weight: bold !important;
        width: 100%;
        height: 50px;
        border-radius: 8px;
        border: none;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-title">CK Reliance Invoice Automation</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Reliance Proforma & Plant Certificate Data Extractor Tool</p>', unsafe_allow_html=True)

# =========================================================
# SESSION STATE INIT
# =========================================================
if "compiled_excel" not in st.session_state:
    st.session_state.compiled_excel = None
if "compiled_df" not in st.session_state:
    st.session_state.compiled_df = None
if "last_file_signature" not in st.session_state:
    st.session_state.last_file_signature = None
if "debug_logs" not in st.session_state:
    st.session_state.debug_logs = []

# =========================================================
# CONFIG / MASTER DATA
# =========================================================
PREFIX_MAPPING = {
    "A2": "HZ",
    "A8": "SIL",
    "A4": "BARODA",
    "OO": "RIL SAILY NEW OO",
    "24": "HZ",
    "QX": "ALOK QX",
    "XX": "HZ",
    "QT": "ALOK QT",
    "A6": "DAHEJ",
    "BG": "SARIGAM BG",
    "QB": "RIL DEPOT NAVSARI",
    "NA": "BARODA VCD",
    "BN": "BN"
}

PKG_TYPES_LIST = [
    "PALLETS",
    "BALES",
    "CARTON",
    "BAGS",
    "DRUMS",
    "ISO TANK",
    "PALLETISED BULK UNIT"
]

# =========================================================
# HELPERS
# =========================================================
def log_debug(msg: str):
    st.session_state.debug_logs.append(msg)

def normalize_space(text: str) -> str:
    return " ".join(text.split()) if text else ""

def safe_str(x):
    return "" if x is None else str(x).strip()

def make_file_signature(uploaded_files):
    """
    Unique signature so that same uploaded files are not reprocessed on download click rerun.
    """
    sig_parts = []
    for f in uploaded_files:
        # name + size = enough for most cases
        try:
            content = f.getvalue()
            sig_parts.append(f"{f.name}|{len(content)}")
        except Exception:
            sig_parts.append(f"{f.name}|0")
    return "||".join(sorted(sig_parts))

def extract_pdf_text(file_bytes: bytes) -> str:
    """
    Extract full text from PDF safely.
    """
    pdf_text = ""
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            pdf_text += txt + "\n"
    return pdf_text

def extract_invoice_no(file_name: str) -> str:
    """
    Invoice number extraction.
    Current fallback logic:
    - Try exact 10-char alphanumeric token from filename
    - Ignore words like PROFORMA / CERTIFICATE if possible

    IMPORTANT:
    If you know exact invoice format, replace this with exact regex.
    """
    base = file_name.upper()

    # Remove extension for safer parsing
    base = re.sub(r"\.PDF$", "", base, flags=re.IGNORECASE)

    # Split on common separators and check tokens first
    tokens = re.split(r"[_\-\s\.]+", base)
    for tok in tokens:
        tok = tok.strip()
        if re.fullmatch(r"[A-Z0-9]{10}", tok):
            return tok

    # fallback: any 10-char alnum chunk
    m = re.search(r"\b([A-Z0-9]{10})\b", base)
    return m.group(1) if m else ""

def extract_first_match(pattern, text, flags=0, group=1, default=""):
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else default

def parse_float_indian(num_str: str):
    """
    Converts strings like:
    12,345.678
    12.345.678
    12345.678
    into float safely where possible.
    """
    if num_str is None:
        return None

    s = str(num_str).strip()
    if not s:
        return None

    # Remove spaces
    s = s.replace(" ", "")

    # Case: multiple dots like 12.345.678 => make last dot decimal
    if s.count(".") > 1:
        parts = s.split(".")
        s = "".join(parts[:-1]) + "." + parts[-1]

    # Remove commas
    s = s.replace(",", "")

    try:
        return float(s)
    except Exception:
        return None

def merge_non_empty(old_val, new_val):
    """
    Keep old if already populated, else new.
    """
    if safe_str(old_val):
        return old_val
    return new_val

def dedupe_containers(container_list):
    """
    Remove duplicate container entries by container_no.
    Keeps first non-empty data, merges missing fields if later row has them.
    """
    temp = {}
    for c in container_list:
        c_no = safe_str(c.get("container_no"))
        if not c_no:
            continue

        if c_no not in temp:
            temp[c_no] = {
                "container_no": c_no,
                "ot_seal": safe_str(c.get("ot_seal")),
                "line_seal": safe_str(c.get("line_seal")),
                "gst_inv": safe_str(c.get("gst_inv")),
            }
        else:
            # merge blanks
            temp[c_no]["ot_seal"] = merge_non_empty(temp[c_no]["ot_seal"], safe_str(c.get("ot_seal")))
            temp[c_no]["line_seal"] = merge_non_empty(temp[c_no]["line_seal"], safe_str(c.get("line_seal")))
            temp[c_no]["gst_inv"] = merge_non_empty(temp[c_no]["gst_inv"], safe_str(c.get("gst_inv")))

    return list(temp.values())

def parse_proforma(file_name: str, pdf_text: str):
    """
    Parse proforma PDF into invoice-level details.
    """
    pdf_text_clean = normalize_space(pdf_text)
    inv_no = extract_invoice_no(file_name)
    if not inv_no:
        raise ValueError(f"Invoice number not found in filename: {file_name}")

    file_details = {
        "inv_no": inv_no,
        "containers": [],
        "single_c_fallback": None,
        "prefix_code": "",
        "division": "",
        "sto": "",
        "date": "",
        "hsn": "",
        "consignee": "",
        "bank": "",
        "port": "",
        "other_ref": "",
        "fcl_20": "",
        "fcl_40": ""
    }

    # ---------------- Prefix ----------------
    prefix = inv_no[:2].upper()
    if prefix == "QB" and "NAVSARI QB" in pdf_text_clean.upper():
        file_details["prefix_code"] = "RIL NAVSARI QB"
    else:
        file_details["prefix_code"] = PREFIX_MAPPING.get(prefix, "")

    # ---------------- Division ----------------
    # Example: Division HAZIRA / Division Hazira
    div_m = re.search(r'Division\s+([A-Za-z\s]+)', pdf_text_clean, flags=re.IGNORECASE)
    if div_m:
        file_details["division"] = div_m.group(1).strip()[:3].upper()

    # ---------------- STO ----------------
    file_details["sto"] = extract_first_match(
        r'Stock\s+Transfer\s+Order\s+(\d{9})',
        pdf_text_clean,
        flags=re.IGNORECASE
    )

    # ---------------- Date ----------------
    # Example: Invoice No. & Date. XXXXX 12.05.2025
    date_val = extract_first_match(
        r'Invoice\s*No\.?\s*&\s*Date\.?\s+[A-Z0-9\/\-]+\s+(\d{2}[./-]\d{2}[./-]\d{4})',
        pdf_text_clean,
        flags=re.IGNORECASE
    )
    if date_val:
        file_details["date"] = date_val.replace(".", "/").replace("-", "/")

    # ---------------- HSN ----------------
    hsn = extract_first_match(
        r'HSN\.?\s*(\d{4}\s*\d{2}\s*\d{2})',
        pdf_text_clean,
        flags=re.IGNORECASE
    )
    file_details["hsn"] = hsn

    # ---------------- Consignee ----------------
    lines = pdf_text.splitlines()
    consignee_name = ""
    for idx, line in enumerate(lines):
        if "consignee" in line.lower():
            test_line = re.sub(r'(?i)consignee', '', line).strip()
            if test_line:
                consignee_name = test_line
                break
            else:
                for next_line in lines[idx+1:]:
                    n = next_line.strip()
                    if n and "reliance" not in n.lower():
                        consignee_name = n
                        break
                break
    file_details["consignee"] = consignee_name

    # ---------------- Bank ----------------
    bank_m = re.search(
        r'Negotiating\s+Bank\s*:?\s*([A-Za-z\s\d\.,&\-\/]+?)(?=\s*Port|\s*AD|\s*Division|$)',
        pdf_text_clean,
        flags=re.IGNORECASE
    )
    if bank_m:
        b_name = bank_m.group(1).strip()
        if "THE HONGKONG AND SHANGHAI BANKING" in b_name.upper():
            file_details["bank"] = "THE HONGKONG AND SHANGHAI BANKING"
        else:
            file_details["bank"] = b_name

    # ---------------- Port ----------------
    port_m = re.search(
        r'Port\s+of\s+Discharge\s+([A-Za-z\s\-,\/]+?)(?=\s*Final|\s*AD|\s*Division|$)',
        pdf_text_clean,
        flags=re.IGNORECASE
    )
    if port_m:
        raw_port = port_m.group(1).strip()
        raw_port = re.split(r'(?i)Place of Receipt|Port of Loading|Country of Origin', raw_port)[0].strip()
        file_details["port"] = raw_port

    # ---------------- Ref No ----------------
    ref_m = re.search(r'REF\s*NO\.?\s*(PC/\d{4})', pdf_text_clean, flags=re.IGNORECASE)
    if ref_m:
        file_details["other_ref"] = ref_m.group(1).strip()
    else:
        ref_alt = re.search(r'REF\s*NO\.?\s*([A-Z]{2}/\d+)', pdf_text_clean, flags=re.IGNORECASE)
        if ref_alt:
            file_details["other_ref"] = ref_alt.group(1).strip()

    # ---------------- FCL ----------------
    upper_text = pdf_text_clean.upper()
    file_details["fcl_20"] = 1 if ("20'FCL" in upper_text or "2 * 20'" in upper_text or "1*20'FCL" in upper_text) else ""
    file_details["fcl_40"] = 1 if ("40'FCL" in upper_text or "1*40'FCL" in upper_text or "2 * 40'" in upper_text) else ""

    # =====================================================
    # CONTAINER EXTRACTION
    # =====================================================
    # Approach:
    # 1) Try line-wise extraction first
    # 2) fallback token-based extraction
    # 3) dedupe container numbers
    # =====================================================

    containers = []

    # -------- Pass 1: line-wise --------
    for line in lines:
        line_clean = normalize_space(line)
        c_matches = re.findall(r'\b([A-Z]{4}\d{7})\b', line_clean)
        if not c_matches:
            continue

        # Try extracting all useful tokens from same line
        tokens = line_clean.split()

        for c_no in c_matches:
            try:
                c_idx = tokens.index(c_no)
            except ValueError:
                c_idx = -1

            ot_seal = ""
            line_seal = ""
            gst_inv = ""

            if c_idx != -1:
                # Nearby tokens heuristic
                nearby = tokens[c_idx+1:c_idx+8]

                # Seal / invoice candidates
                # first alnum tokens after container can be seals/invoice
                alnums = [x for x in nearby if re.fullmatch(r'[A-Z0-9\-\/]+', x, flags=re.IGNORECASE)]

                if len(alnums) >= 1:
                    ot_seal = alnums[0]
                if len(alnums) >= 2:
                    line_seal = alnums[1]
                if len(alnums) >= 3:
                    gst_inv = alnums[2]

            containers.append({
                "container_no": c_no,
                "ot_seal": ot_seal,
                "line_seal": line_seal,
                "gst_inv": gst_inv
            })

    # -------- Pass 2: token fallback if nothing found --------
    if not containers:
        tokens = pdf_text_clean.split()
        for idx, t in enumerate(tokens):
            if re.fullmatch(r'[A-Z]{4}\d{7}', t):
                ot_seal = tokens[idx+1] if idx+1 < len(tokens) else ""
                line_seal = tokens[idx+2] if idx+2 < len(tokens) else ""
                gst_inv = tokens[idx+4] if idx+4 < len(tokens) else ""

                containers.append({
                    "container_no": t,
                    "ot_seal": ot_seal,
                    "line_seal": line_seal,
                    "gst_inv": gst_inv
                })

    containers = dedupe_containers(containers)
    file_details["containers"] = containers

    if file_details["containers"]:
        file_details["single_c_fallback"] = file_details["containers"][0]["container_no"]

    return inv_no, file_details

def parse_plant_certificate(file_name: str, pdf_text: str):
    """
    Parse plant certificate and return container-level packing / weight data.
    """
    pdf_text_clean = normalize_space(pdf_text)
    lines_cert = pdf_text.splitlines()

    # unit detect
    is_kg_unit = "WT. (KG)" in pdf_text_clean.upper() or "WT.(KG)" in pdf_text_clean.upper()

    # package type detect
    matched_pkg = "BAGS"
    for p_type in PKG_TYPES_LIST:
        if p_type in pdf_text_clean.upper():
            matched_pkg = p_type
            if p_type == "PALLETISED BULK UNIT":
                matched_pkg = "Palletised bulk unit"
            break

    # container-wise temp aggregation
    temp_agg = defaultdict(lambda: {
        "bags": 0,
        "gross_wt": 0.0,
        "net_wt": 0.0,
        "pkg_type": matched_pkg
    })

    for line in lines_cert:
        line_clean = normalize_space(line)
        if not line_clean:
            continue

        c_match = re.search(r'\b([A-Z]{4}\d{7})\b', line_clean)
        if not c_match:
            continue

        c_no = c_match.group(1)

        # Container ke baad wale numeric segments nikaalo
        tail = line_clean.split(c_no, 1)[1].strip() if c_no in line_clean else ""
        segs = tail.split()

        # Keep numeric-looking tokens only
        metrics = [s for s in segs if re.fullmatch(r'[\d,\.]+', s)]

        # We expect last 3 metrics = bags, gross, net (as per your original logic)
        if len(metrics) >= 3:
            bags_raw = metrics[-3]
            gross_raw = metrics[-2]
            net_raw = metrics[-1]

            try:
                bags_val = int(float(bags_raw.replace(",", "")))
            except Exception:
                bags_val = 0

            gross_val = parse_float_indian(gross_raw)
            net_val = parse_float_indian(net_raw)

            if gross_val is None:
                gross_val = 0.0
            if net_val is None:
                net_val = 0.0

            # If values are in MT and not KG, convert to KG
            if not is_kg_unit:
                gross_val *= 1000
                net_val *= 1000

            temp_agg[c_no]["bags"] += bags_val
            temp_agg[c_no]["gross_wt"] += gross_val
            temp_agg[c_no]["net_wt"] += net_val
            temp_agg[c_no]["pkg_type"] = matched_pkg

    return temp_agg

def merge_proforma(existing: dict, new_data: dict):
    """
    Merge multiple proforma records for same invoice safely.
    """
    # merge scalar fields only if existing blank
    scalar_fields = [
        "prefix_code", "division", "sto", "date", "hsn",
        "consignee", "bank", "port", "other_ref",
        "fcl_20", "fcl_40"
    ]

    for fld in scalar_fields:
        existing[fld] = merge_non_empty(existing.get(fld, ""), new_data.get(fld, ""))

    # merge containers
    existing_containers = existing.get("containers", [])
    new_containers = new_data.get("containers", [])
    merged = existing_containers + new_containers
    existing["containers"] = dedupe_containers(merged)

    if existing["containers"] and not existing.get("single_c_fallback"):
        existing["single_c_fallback"] = existing["containers"][0]["container_no"]

    return existing

def build_output_dataframe(proforma_data: dict, cert_data: dict) -> pd.DataFrame:
    """
    Build final output rows.
    Rule:
    - 1 row per container if containers exist
    - if no container found, 1 row with blank container
    - duplicates removed by invoice + container
    """
    final_rows = []

    for inv_no, p_info in proforma_data.items():
        containers_to_process = p_info.get("containers", [])

        if not containers_to_process:
            containers_to_process = [{
                "container_no": "",
                "ot_seal": "",
                "line_seal": "",
                "gst_inv": ""
            }]

        for c_info in containers_to_process:
            c_no = safe_str(c_info.get("container_no"))

            # Try exact container cert match
            c_cert = cert_data.get(c_no)

            # fallback to first container cert if exact not found
            if not c_cert and p_info.get("single_c_fallback"):
                c_cert = cert_data.get(p_info["single_c_fallback"])

            if not c_cert:
                c_cert = {"bags": "", "pkg_type": "", "gross_wt": "", "net_wt": ""}

            row_dict = {
                "J": safe_str(p_info.get("division")),
                "K": safe_str(p_info.get("sto")),
                "L": safe_str(p_info.get("prefix_code")),
                "O": p_info.get("fcl_20", ""),
                "P": p_info.get("fcl_40", ""),
                "U": c_no,
                "Z": safe_str(c_info.get("line_seal")),
                "AA": safe_str(c_info.get("ot_seal")),
                "AB": safe_str(c_info.get("gst_inv")),
                "AC": safe_str(p_info.get("other_ref")),
                "AD": safe_str(p_info.get("date")),
                "AE": inv_no,
                "AF": safe_str(p_info.get("date")),
                "AG": c_cert["bags"] if c_cert["bags"] != "" else "",
                "AH": safe_str(c_cert.get("pkg_type")),
                "AI": f"{c_cert['gross_wt']:.3f}" if c_cert.get("gross_wt", "") != "" else "",
                "AJ": f"{c_cert['net_wt']:.3f}" if c_cert.get("net_wt", "") != "" else "",
                "AK": safe_str(p_info.get("port")),
                "AN": safe_str(p_info.get("consignee")),
                "AX": safe_str(p_info.get("bank")),
                "AZ": safe_str(p_info.get("hsn"))
            }
            final_rows.append(row_dict)

    df = pd.DataFrame(final_rows)

    if df.empty:
        return df

    # Remove exact duplicate rows first
    df = df.drop_duplicates()

    # Remove duplicate invoice + container rows
    # if container blank, still keep first invoice row only
    df = df.drop_duplicates(subset=["AE", "U"], keep="first")

    # Optional sort
    df = df.sort_values(by=["AE", "U"], na_position="last").reset_index(drop=True)

    return df

def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    buffer.seek(0)
    return buffer.getvalue()

def process_files(uploaded_files):
    """
    Main processor
    """
    proforma_data = {}
    cert_data = defaultdict(lambda: {
        "bags": 0,
        "gross_wt": 0.0,
        "net_wt": 0.0,
        "pkg_type": ""
    })

    total_files = len(uploaded_files)
    processed_count = 0

    for file in uploaded_files:
        file_name = file.name
        try:
            file_content = file.getvalue()
            if not file_content:
                log_debug(f"Skipped empty file: {file_name}")
                continue

            pdf_text = extract_pdf_text(file_content)
            if not pdf_text.strip():
                log_debug(f"No extractable text found in: {file_name}")
                continue

            lower_name = file_name.lower()

            # ---------------- PROFORMA ----------------
            if "_proforma" in lower_name:
                inv_no, details = parse_proforma(file_name, pdf_text)

                if inv_no in proforma_data:
                    proforma_data[inv_no] = merge_proforma(proforma_data[inv_no], details)
                else:
                    proforma_data[inv_no] = details

                processed_count += 1

            # ---------------- PLANT CERTIFICATE ----------------
            elif "_plant_certificate" in lower_name:
                cert_rows = parse_plant_certificate(file_name, pdf_text)

                for c_no, vals in cert_rows.items():
                    cert_data[c_no]["bags"] += vals["bags"]
                    cert_data[c_no]["gross_wt"] += vals["gross_wt"]
                    cert_data[c_no]["net_wt"] += vals["net_wt"]

                    # pkg type preserve if blank
                    if not cert_data[c_no]["pkg_type"]:
                        cert_data[c_no]["pkg_type"] = vals["pkg_type"]

                processed_count += 1

            else:
                log_debug(f"Skipped (name pattern not matched): {file_name}")

        except Exception as e:
            log_debug(f"Error in {file_name}: {str(e)}")

    df = build_output_dataframe(proforma_data, cert_data)
    return df, processed_count, total_files

# =========================================================
# UI
# =========================================================
uploaded_files = st.file_uploader(
    "📂 Saari PDF Files Ek Sath Select Ya Drop Karein (Ctrl+A)",
    type="pdf",
    accept_multiple_files=True
)

col1, col2 = st.columns(2)
with col1:
    process_btn = st.button("⚙️ Process Files")
with col2:
    clear_btn = st.button("🗑️ Clear Session")

if clear_btn:
    st.session_state.compiled_excel = None
    st.session_state.compiled_df = None
    st.session_state.last_file_signature = None
    st.session_state.debug_logs = []
    st.success("Session cleared.")
    st.stop()

# =========================================================
# PROCESSING
# =========================================================
if uploaded_files:
    current_signature = make_file_signature(uploaded_files)

    # Process only when button clicked OR files changed and no output exists yet
    if process_btn:
        st.session_state.debug_logs = []
        with st.spinner("PDF files process ho rahi hain..."):
            try:
                df, processed_count, total_files = process_files(uploaded_files)

                if df.empty:
                    st.warning("Koi valid output row generate nahi hui. PDFs ka format ya filename pattern check karo.")
                    st.session_state.compiled_excel = None
                    st.session_state.compiled_df = None
                    st.session_state.last_file_signature = current_signature
                else:
                    excel_bytes = dataframe_to_excel_bytes(df)
                    st.session_state.compiled_excel = excel_bytes
                    st.session_state.compiled_df = df
                    st.session_state.last_file_signature = current_signature

                    st.success(
                        f"🎉 Processing complete! "
                        f"{processed_count}/{total_files} files successfully parsed. "
                        f"Output rows: {len(df)}"
                    )

            except Exception as e:
                st.error(f"Processing failed: {e}")

    # If same files already processed, show cached result without reprocessing
    if (
        st.session_state.compiled_excel is not None
        and st.session_state.compiled_df is not None
        and st.session_state.last_file_signature == current_signature
    ):
        st.success(f"Ready! Output rows: {len(st.session_state.compiled_df)}")

        st.download_button(
            label="📥 Download Compiled Excel File",
            data=st.session_state.compiled_excel,
            file_name="Reliance_Invoice_Data_Compiled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        with st.expander("📋 Preview Output Data"):
            st.dataframe(st.session_state.compiled_df, use_container_width=True)

        if st.session_state.debug_logs:
            with st.expander("🛠 Debug Log"):
                for item in st.session_state.debug_logs:
                    st.write(item)

    else:
        if not process_btn:
            st.info("Files upload ho gayi hain. Ab **Process Files** button dabao.")

else:
    st.info("PDF files upload karo, phir Process Files dabao.")
