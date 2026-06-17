import streamlit as st
import io
import re
import requests
import logging
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

logging.getLogger("pypdf").setLevel(logging.CRITICAL)

st.set_page_config(page_title="Component Engine", page_icon="⚡", layout="wide")

CLIENT_ID = '295EGpEwJEuPCaTsslUztQdBUXQOCLvGztU2UlEkqGfcIyur'
CLIENT_SECRET = 'X80tbzNKh50mx6IieAOoJcWl57jhE3dmiycn2jYj74XTVZisrUGyJHKumFiB1wDr'

SPEC_ALIASES = {
    "power": ["power dissipation", "power consumption", "p_diss", "max. power", "power"],
    "voltage": ["supply voltage", "voltage - supply", "operating voltage", "rated voltage", "voltage"],
    "current": ["continuous drain current", "drain current", "id", "supply current", "operating current", "current"],
    "temperature": ["operating temperature", "storage temperature", "temp range", "case temperature"],
    "min temperature": ["min temperature", "minimum temperature", "min temp"], 
    "max temperature": ["max temperature", "maximum temperature", "max temp"], 
    "wavelength": ["wavelength", "wave length", "optical wavelength"],
    "data rate": ["data rate", "speed"],
    "vds": ["drain to source voltage", "drain-source voltage", "vds", "vdss"],
    "vgs": ["gate to source voltage", "gate-source voltage", "vgs", "vgss"],
    "rds on": ["rds(on)", "drain-source on-state resistance", "on-resistance", "rds on (max)"],
    "gate charge": ["gate charge", "qg"],
    "forward voltage": ["voltage - forward", "forward voltage", "vf"],
    "luminous intensity": ["millicandela rating", "luminous intensity", "brightness", "mcd"],
    "viewing angle": ["viewing angle", "beam angle"],
    "color": ["emitted color", "color"],
    "sound level": ["sound pressure level", "spl", "db"]
}

@st.cache_data(ttl=3600)
def get_digikey_token():
    url = "https://api.digikey.com/v1/oauth2/token"
    payload = {'grant_type': 'client_credentials', 'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET}
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = requests.post(url, data=payload, headers=headers)
    if response.status_code != 200: 
        st.error("Authentication failed! Check credentials.")
        return None
    return response.json()['access_token']

def extract_price(prod_dict):
    try:
        variations = prod_dict.get('ProductVariations', [])
        if variations and variations[0].get('StandardPricing'):
            return variations[0]['StandardPricing'][0].get('UnitPrice', 0.0)
        pricing = prod_dict.get('StandardPricing', [])
        if pricing: return pricing[0].get('UnitPrice', 0.0)
    except Exception: pass
    return 0.0

def split_temperature_ranges(part_dict):
    for k, v in list(part_dict.items()):
        if "temperature" in k and "~" in str(v):
            match = re.search(r"([-+]?\d+(?:\.\d+)?)[^0-9\-+]*~[^0-9\-+]*([-+]?\d+(?:\.\d+)?)", str(v))
            if match:
                part_dict['min temperature'] = match.group(1) + "°C"
                part_dict['max temperature'] = match.group(2) + "°C"
            break

def auto_extract_specs_from_pdf(pdf_url):
    if not pdf_url: return {}
    if pdf_url.startswith("//"): pdf_url = "https:" + pdf_url
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(pdf_url, headers=headers, timeout=5)
        if response.status_code != 200: return {}
        reader = PdfReader(io.BytesIO(response.content))
        max_pages = min(3, len(reader.pages))
        full_text = "".join(reader.pages[i].extract_text() or "" for i in range(max_pages))
        extracted_params = {}
        for category, aliases in SPEC_ALIASES.items():
            pattern = re.compile(rf"({'|'.join(sorted(aliases, key=len, reverse=True))})[^\d]*([0-9\.\-\~]+\s*[a-zA-Z]*)", re.IGNORECASE)
            match = pattern.search(full_text)
            if match: extracted_params[category] = match.group(2).split('\n')[0].strip()
        split_temperature_ranges(extracted_params)
        return extracted_params
    except Exception: return {}

def fetch_part_data(token, part_number):
    url = "https://api.digikey.com/products/v4/search/keyword"
    headers = {"Authorization": f"Bearer {token}", "X-DIGIKEY-Client-Id": CLIENT_ID, "Content-Type": "application/json"}
    response = requests.post(url, json={"Keywords": part_number, "Limit": 1}, headers=headers)
    if response.status_code != 200 or not response.json().get('Products'): return None 
    
    prod = response.json()['Products'][0]
    datasheet_url = prod.get('DatasheetUrl')
    part_profile = {
        "Part Number": prod.get('ManufacturerProductNumber') or prod.get('ProductCode') or part_number,
        "Manufacturer": prod.get('Manufacturer', {}).get('Name'),
        "Description": prod.get('ProductDescription') or "",
        "Price": extract_price(prod),
        "DatasheetUrl": datasheet_url
    }
    for param in prod.get('Parameters', []):
        name = param.get('ParameterText', '').strip().lower()
        if name: part_profile[name] = param.get('ValueText', '').strip()
    split_temperature_ranges(part_profile)
    if datasheet_url:
        pdf_specs = auto_extract_specs_from_pdf(datasheet_url)
        for k, v in pdf_specs.items():
            if k not in part_profile: part_profile[k] = v
    return part_profile

def map_parameter_name(user_term, available_keys):
    user_term = user_term.lower().strip()
    if user_term in available_keys: return user_term
    for alias in SPEC_ALIASES.get(user_term, [user_term]):
        for key in available_keys:
            if alias in key: return key
    for key in available_keys:
        if user_term in key: return key
    return None

def find_similar_parts(token, search_keyword, constants_dict, variables_list):
    url = "https://api.digikey.com/products/v4/search/keyword"
    headers = {"Authorization": f"Bearer {token}", "X-DIGIKEY-Client-Id": CLIENT_ID, "Content-Type": "application/json"}
    response = requests.post(url, json={"Keywords": search_keyword, "Limit": 50}, headers=headers)
    
    if response.status_code != 200: return None
    products = response.json().get('Products', [])
    if not products: return []

    matching_candidates = []
    for prod in products:
        cand_params = {p.get('ParameterText', '').strip().lower(): p.get('ValueText', '').strip() for p in prod.get('Parameters', []) if p.get('ParameterText')}
        split_temperature_ranges(cand_params)
        datasheet_url = prod.get('DatasheetUrl')
        pdf_scanned, match_failed = False, False

        keys_to_check = list(constants_dict.keys()) + variables_list
        if datasheet_url and any(k not in cand_params for k in keys_to_check):
            pdf_data = auto_extract_specs_from_pdf(datasheet_url)
            for k, v in pdf_data.items():
                if k not in cand_params: cand_params[k] = v

        for c_param, required_val in constants_dict.items():
            cand_val = cand_params.get(c_param)
            if not cand_val: match_failed = True; break
            
            req_str, cand_str = str(required_val).lower(), str(cand_val).lower()
            if req_str in cand_str or cand_str in req_str: continue
            req_nums = "".join(c for c in req_str if c.isdigit() or c == '.')
            cand_nums = "".join(c for c in cand_str if c.isdigit() or c == '.')
            if req_nums and cand_nums and (req_nums in cand_nums or cand_nums in req_nums): continue
            match_failed = True; break

        if match_failed: continue

        matching_candidates.append({
            "Part Number": prod.get('ManufacturerProductNumber') or prod.get('ProductCode'),
            "Manufacturer": prod.get('Manufacturer', {}).get('Name'),
            "Price": extract_price(prod),
            "DatasheetUrl": datasheet_url,
            "all_params": cand_params  
        })
    return matching_candidates

def generate_advanced_excel_buffer(ref_part, candidates, constants, variables):
    wb = Workbook()
    ws = wb.active
    center_align = Alignment(horizontal='center', vertical='center')
    price_col = 1 + len(constants) + len(variables) + 1 

    ws.cell(row=1, column=1, value="Name / Part Number").alignment = center_align
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    if constants:
        ws.cell(row=1, column=2, value="Constant Aspects").alignment = center_align
        ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=1+len(constants))
    if variables:
        ws.cell(row=1, column=2+len(constants), value="Varying Aspects").alignment = center_align
        ws.merge_cells(start_row=1, start_column=2+len(constants), end_row=1, end_column=1+len(constants)+len(variables))
    ws.cell(row=1, column=price_col, value="Price (USD)").alignment = center_align
    ws.merge_cells(start_row=1, start_column=price_col, end_row=2, end_column=price_col)

    col = 2
    for c in constants + variables:
        ws.cell(row=2, column=col, value=str(c))
        col += 1

    ws.cell(row=3, column=1, value=f"ORIGINAL: {ref_part['Part Number']}")
    col = 2
    for c in constants + variables:
        ws.cell(row=3, column=col, value=str(ref_part.get(c, 'N/A')))
        col += 1
    ws.cell(row=3, column=price_col, value=round(float(ref_part.get('Price', 0.0)), 2))
    ws.cell(row=3, column=price_col).number_format = '0.00'

    current_row = 4
    for cand in candidates:
        ws.cell(row=current_row, column=1, value=str(cand['Part Number']))
        col = 2
        for c in constants + variables:
            val = cand['all_params'].get(c, 'N/A')
            ws.cell(row=current_row, column=col, value=str(val))
            if str(val) != str(ref_part.get(c, 'N/A')):
                ws.cell(row=current_row, column=col).font = Font(underline="single")
            col += 1
        ws.cell(row=current_row, column=price_col, value=round(float(cand.get('Price', 0.0)), 2))
        ws.cell(row=current_row, column=price_col).number_format = '0.00'
        current_row += 1

    ws.column_dimensions['A'].width = 40
    for col_idx in range(2, price_col + 1):
        ws.column_dimensions[ws.cell(row=3, column=col_idx).column_letter].width = 20
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generate_batch_extract_excel_buffer(part_profiles, all_params):
    wb = Workbook()
    ws = wb.active
    ws.title = "Batch Extraction"
    bold_font = Font(bold=True)

    ws.cell(row=1, column=1, value="Part Number").font = bold_font
    ws.cell(row=1, column=2, value="Manufacturer").font = bold_font
    ws.cell(row=1, column=3, value="Price (USD)").font = bold_font
    
    col = 4
    for param in all_params:
        ws.cell(row=1, column=col, value=str(param)).font = bold_font
        col += 1

    current_row = 2
    for profile in part_profiles:
        ws.cell(row=current_row, column=1, value=str(profile.get('Part Number', 'N/A')))
        ws.cell(row=current_row, column=2, value=str(profile.get('Manufacturer', 'N/A')))
        ws.cell(row=current_row, column=3, value=round(float(profile.get('Price', 0.0)), 2))
        ws.cell(row=current_row, column=3).number_format = '0.00'

        col = 4
        for param in all_params:
            ws.cell(row=current_row, column=col, value=str(profile.get(param, 'N/A')))
            col += 1
        current_row += 1

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    for col_idx in range(4, 4 + len(all_params)):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 20

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ==========================================
# STREAMLIT UI
# ==========================================
st.title("⚡ Component Optimization & Extraction Engine")
token = get_digikey_token()

if not token:
    st.stop()

tab1, tab2 = st.tabs(["Find Alternatives", "Batch Extract Characteristics"])

# --- TAB 1: FIND ALTERNATIVES ---
with tab1:
    st.markdown("Find alternative components based on strict rules.")
    ref_part_num = st.text_input("Enter Reference Part Number:", key="ref_input")
    
    if st.button("Lookup Part", type="primary"):
        if ref_part_num:
            with st.spinner("Fetching data..."):
                profile = fetch_part_data(token, ref_part_num)
                if profile:
                    st.session_state['ref_profile'] = profile
                else:
                    st.error("Part not found on DigiKey.")
                    
    if 'ref_profile' in st.session_state:
        ref = st.session_state['ref_profile']
        st.success(f"**Loaded:** {ref['Part Number']} ({ref['Manufacturer']}) - ${round(float(ref['Price']), 2)}")
        
        keys = [k for k in ref.keys() if k not in ["Part Number", "Manufacturer", "Description", "Price", "DatasheetUrl", "Category"]]
        with st.expander("View Available Parameters"):
            st.write(", ".join(keys))
            
        desc = str(ref.get("Description", ""))
        smart_sugg = " ".join(desc.split()[:2]) if desc else ""
        
        with st.form("compare_form"):
            search_cat = st.text_input("Search Keyword (DigiKey shorthand):", value=smart_sugg)
            constants_input = st.text_input("MUST remain the same (comma-separated):")
            variants_input = st.text_input("Can vary (comma-separated):")
            
            if st.form_submit_button("Generate Comparison"):
                constants = [map_parameter_name(t, keys) for t in constants_input.split(",") if t.strip() and map_parameter_name(t, keys)]
                variants = [map_parameter_name(t, keys) for t in variants_input.split(",") if t.strip() and map_parameter_name(t, keys)]
                
                with st.spinner("Scanning database & PDFs..."):
                    target_consts = {c: ref[c] for c in constants}
                    cands = find_similar_parts(token, search_cat, target_consts, variants)
                    
                    if not cands:
                        st.error("No matches found. Filters are too strict!")
                    else:
                        st.success(f"Found {len(cands)} matches!")
                        buffer = generate_advanced_excel_buffer(ref, cands, constants, variants)
                        st.download_button("📥 Download Excel", data=buffer, file_name=f"Comparison_{ref['Part Number']}.xlsx")

# --- TAB 2: BATCH EXTRACT ---
with tab2:
    st.markdown("Paste a list of part numbers (or PDF filenames) to extract all their characteristics into one master spreadsheet.")
    batch_input = st.text_area("Enter part numbers (comma-separated or one per line):", height=150)
    
    if st.button("Extract Data", type="primary"):
        if batch_input:
            # Handle both commas and new lines
            parts = [p.strip() for p in batch_input.replace('\n', ',').split(',') if p.strip()]
            
            with st.spinner("Vacuuming data from DigiKey and PDFs..."):
                extracted = []
                all_params = set()
                
                for part in parts:
                    clean = part.replace('.pdf', '').replace('.PDF', '')
                    clean = re.sub(r'(?i)(^DS_|^infineon-|-datasheet-en|_en)', '', clean)
                    
                    prof = fetch_part_data(token, clean)
                    if prof:
                        extracted.append(prof)
                        for k in prof.keys():
                            if k not in ["Part Number", "Manufacturer", "Description", "Price", "DatasheetUrl", "Category"]:
                                all_params.add(k)
                
                if extracted:
                    st.success(f"Successfully processed {len(extracted)} parts!")
                    buffer = generate_batch_extract_excel_buffer(extracted, sorted(list(all_params)))
                    st.download_button("📥 Download Master Spreadsheet", data=buffer, file_name="Batch_Extraction.xlsx")
                else:
                    st.error("Failed to fetch any parts.")