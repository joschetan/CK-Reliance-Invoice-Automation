import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="CK Reliance Invoice Automation", page_icon="📊", layout="centered")

st.markdown("""
    <style>
    .main-title { color: #1e3a8a; text-align: center; font-family: 'Segoe UI', Arial, sans-serif; font-weight: 700; margin-bottom: 2px; }
    .sub-title { color: #64748b; text-align: center; font-size: 14px; margin-bottom: 30px; }
    .stButton>button { background-color: #10b981 !important; color: white !important; font-weight: bold !important; width: 100%; height: 50px; border-radius: 8px; border: none; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-title">CK Reliance Invoice Automation</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Reliance Proforma & Plant Certificate Data Extractor Tool</p>', unsafe_allow_html=True)

import pypdf

uploaded_files = st.file_uploader("📂 Saari PDF Files Ek Sath Select Ya Drop Karein (Ctrl+A)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.info(f"⚡ Total {len(uploaded_files)} files uploaded. Processing...")
    
    proforma_data = {}
    cert_data = {}
    
    prefix_mapping = {
        "A2": "HZ", "A8": "SIL", "A4": "BARODA", "OO": "RIL SAILY NEW OO",
        "24": "HZ", "QX": "ALOK QX", "XX": "HZ", "QT": "ALOK QT",
        "A6": "DAHEJ", "BG": "SARIGAM BG", "QB": "RIL DEPOT NAVSARI",
        "NA": "BARODA VCD", "BN": "BN"
    }
    pkg_types_list = ["PALLETS", "BALES", "CARTON", "BAGS", "DRUMS", "ISO TANK", "PALLETISED BULK UNIT"]

    for file in uploaded_files:
        file_name = file.name
        file_content = file.read()
        
        inv_match = re.search(r'([A-Z0-9]{10})', file_name)
        if not inv_match:
            continue
        inv_no = inv_match.group(1)
        
        pdf_text = ""
        try:
            reader = pypdf.PdfReader(io.BytesIO(file_content))
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    pdf_text += text + "\n"
        except:
            continue
            
        pdf_text_clean = " ".join(pdf_text.split())

        # --- SECTION 1: PROFORMA PROCESSING ---
        if "_proforma" in file_name.lower():
            file_details = {"inv_no": inv_no, "containers": [], "single_c_fallback": None}
            prefix = inv_no[:2].upper()
            if prefix == "QB" and "NAVSARI QB" in pdf_text_clean.upper():
                file_details["prefix_code"] = "RIL NAVSARI QB"
            else:
                file_details["prefix_code"] = prefix_mapping.get(prefix, "")
                
            div_m = re.search(r'Division\s+([A-Za-z\s]+)', pdf_text_clean)
            file_details["division"] = div_m.group(1).strip()[:3].upper() if div_m else ""
            
            sto_m = re.search(r'Stock Transfer Order\s+(\d{9})', pdf_text_clean)
            file_details["sto"] = sto_m.group(1) if sto_m else ""
            
            date_m = re.search(r'Invoice No\. & Date\.\s+[A-Z0-9]+\s+(\d{2}\.\d{2}\.\d{4})', pdf_text_clean)
            file_details["date"] = date_m.group(1).replace('.', '/') if date_m else ""
            
            hsn_m = re.search(r'HSN\.\s*(\d{4}\s*\d{2}\s*\d{2})', pdf_text_clean)
            file_details["hsn"] = hsn_m.group(1).strip() if hsn_m else ""
            
            lines = pdf_text.split('\n')
            consignee_name = ""
            for idx, line in enumerate(lines):
                if "Consignee" in line:
                    test_line = line.replace("Consignee", "").strip()
                    if test_line:
                        consignee_name = test_line
                    else:
                        for next_line in lines[idx+1:]:
                            if next_line.strip() and "Reliance" not in next_line:
                                consignee_name = next_line.strip()
                                break
                        break
            file_details["consignee"] = consignee_name
                
            bank_m = re.search(r'Negotiating Bank\s*:\s*([A-Za-z\s\d\.,]+?)(?=\s*Port|\s*AD|$)', pdf_text_clean)
            if bank_m:
                b_name = bank_m.group(1).strip()
                file_details["bank"] = "THE HONGKONG AND SHANGHAI BANKING" if "THE HONGKONG AND SHANGHAI BANKING" in b_name.upper() else b_name
            else:
                file_details["bank"] = ""
            
            port_m = re.search(r'Port of Discharge\s+([A-Za-z\s\-,\/]+?)(?=\s*Final|\s*AD|\s*Division|$)', pdf_text_clean)
            if port_m:
                raw_port = port_m.group(1).strip()
                raw_port = re.split(r'(?i)Place of Receipt|Port of Loading|Country of Origin', raw_port)[0].strip()
                file_details["port"] = raw_port
            else:
                file_details["port"] = ""
            
            ref_m = re.search(r'REF NO\.(PC/\d{4})', pdf_text_clean)
            if ref_m:
                file_details["other_ref"] = ref_m.group(1).strip()
            else:
                ref_alt = re.search(r'REF NO\.([A-Z]{2}/\d+)', pdf_text_clean)
                file_details["other_ref"] = ref_alt.group(1).strip() if ref_alt else ""
            
            file_details["fcl_20"] = 1 if "20'FCL" in pdf_text_clean or "2 * 20'" in pdf_text_clean else ""
            file_details["fcl_40"] = 1 if "40'FCL" in pdf_text_clean or "1*40'FCL" in pdf_text_clean or "2 * 40'" in pdf_text_clean else ""
            
            tokens = pdf_text_clean.split()
            seen_containers = set()
            for idx, t in enumerate(tokens):
                # ISO Standard container rules applied
                if re.match(r'^[A-Z]{3}U\d{7}$', t):
                    try:
                        if t not in seen_containers:
                            seen_containers.add(t)
                            file_details["containers"].append({
                                "container_no": t,
                                "ot_seal": tokens[idx+1],
                                "line_seal": tokens[idx+2],
                                "gst_inv": tokens[idx+4]
                            })
                    except:
                        pass
            
            if file_details["containers"]:
                file_details["single_c_fallback"] = file_details["containers"][0]["container_no"]
            proforma_data[inv_no] = file_details

        # --- SECTION 2: PLANT CERTIFICATE PROCESSING ---
        elif "_plant_certificate" in file_name.lower():
            is_kg_unit = "WT. (KG)" in pdf_text_clean.upper() or "WT.(KG)" in pdf_text_clean.upper()
            matched_pkg = "BAGS"
            for p_type in pkg_types_list:
                if p_type in pdf_text_clean.upper():
                    matched_pkg = p_type
                    if p_type == "PALLETISED BULK UNIT":
                        matched_pkg = "Palletised bulk unit"
                    break
            
            lines_cert = pdf_text.split('\n')
            temp_rows = []
            
            for line in lines_cert:
                line_clean = " ".join(line.split())
                c_match = re.search(r'\b([A-Z]{3}U\d{7})\b', line_clean)
                if c_match:
                    c_no = c_match.group(1)
                    num_segments = line_clean.split(c_no)[1].strip().split()
                    metrics = [s for s in num_segments if re.match(r'^[\d,\.]+$', s)]
                    
                    if len(metrics) >= 3:
                        try:
                            net_raw = metrics[-1]
                            gross_raw = metrics[-2]
                            bags_raw = metrics[-3]
                            
                            bags_val = int(bags_raw)
                            
                            if gross_raw.count('.') > 1:
                                gross_final = float("".join(gross_raw.split('.')[:-1]) + "." + gross_raw.split('.')[-1])
                            else:
                                gross_final = float(gross_raw.replace(',', ''))
                                
                            if net_raw.count('.') > 1:
                                net_final = float("".join(net_raw.split('.')[:-1]) + "." + net_raw.split('.')[-1])
                            else:
                                net_final = float(net_raw.replace(',', ''))
                            
                            if not is_kg_unit:
                                gross_final *= 1000
                                net_final *= 1000
                                
                            temp_rows.append({"c_no": c_no, "bags": bags_val, "gross": gross_final, "net": net_final})
                        except:
                            pass

            for r in temp_rows:
                key_c = r["c_no"]
                if key_c not in cert_data:
                    cert_data[key_c] = {"bags": 0, "gross_wt": 0.0, "net_wt": 0.0, "pkg_type": matched_pkg}
                cert_data[key_c]["bags"] += r["bags"]
                cert_data[key_c]["gross_wt"] += r["gross"]
                cert_data[key_c]["net_wt"] += r["net"]

    # --- SECTION 3: EXACT ALIGNED EXCEL DATA GRID GENERATION ---
    # Creating strict mapping for standard master alphabet rows A to AZ
    columns_list = [chr(65 + i) for i in range(26)] + ["A" + chr(65 + i) for i in range(26)]
    final_rows = []
    used_columns = set()
    strictly_processed_keys = set()

    for inv_no, p_info in proforma_data.items():
        if len(p_info["containers"]) >= 1:
            containers_to_process = p_info["containers"]
        else:
            containers_to_process = [{"container_no": "UNKNOWN", "ot_seal": "", "line_seal": "", "gst_inv": ""}]
        
        for idx, c_info in enumerate(containers_to_process):
            c_no = c_info["container_no"]
            
            master_row_key = f"{inv_no}_{c_no}"
            if master_row_key in strictly_processed_keys:
                continue
            strictly_processed_keys.add(master_row_key)
            
            c_cert = cert_data.get(c_no, cert_data.get(p_info["single_c_fallback"], {"bags": "", "pkg_type": "", "gross_wt": "", "net_wt": ""}))
            
            # Map values strictly into their designated master layout grid keys
            row_dict = {col: "" for col in columns_list}
            row_dict["J"] = p_info["division"]
            row_dict["K"] = p_info["sto"]
            row_dict["L"] = p_info["prefix_code"]
            row_dict["O"] = p_info["fcl_20"]
            row_dict["P"] = p_info["fcl_40"]
            row_dict["U"] = c_no if c_no != "UNKNOWN" else ""
            row_dict["Z"] = c_info["line_seal"]
            row_dict["AA"] = c_info["ot_seal"]
            row_dict["AB"] = c_info["gst_inv"]
            row_dict["AC"] = p_info["other_ref"]
            row_dict["AD"] = p_info["date"]
            row_dict["AE"] = inv_no
            row_dict["AF"] = p_info["date"]
            row_dict["AG"] = c_cert["bags"]
            row_dict["AH"] = c_cert["pkg_type"]
            row_dict["AI"] = f"{c_cert['gross_wt']:.3f}" if c_cert['gross_wt'] != "" else ""
            row_dict["AJ"] = f"{c_cert['net_wt']:.3f}" if c_cert['net_wt'] != "" else ""
            row_dict["AK"] = p_info["port"]
            row_dict["AN"] = p_info["consignee"]
            row_dict["AX"] = p_info["bank"]
            row_dict["AZ"] = p_info["hsn"]
            
            for k, v in row_dict.items():
                if v != "":
                    used_columns.add(k)
            final_rows.append(row_dict)

    if final_rows:
        df = pd.DataFrame(final_rows, columns=columns_list)
        excel_buffer = io.BytesIO()
        
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            
            # FIXED LOGIC: Loops through structure and safely hides empty spaces without overloading memory
            for idx, col_name in enumerate(columns_list, start=1):
                if col_name not in used_columns:
                    col_letter = worksheet.cell(row=1, column=idx).column_letter
                    worksheet.column_dimensions[col_letter].hidden = True
            
        st.success("🎉 All columns re-aligned perfectly to their original positions!")
        st.download_button(
            label="📥 Download Compiled Excel File",
            data=excel_buffer.getvalue(),
            file_name="Reliance_Invoice_Data_Compiled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
