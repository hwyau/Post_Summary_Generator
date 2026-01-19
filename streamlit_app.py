import streamlit as st
import pandas as pd
import re
import json
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import io

st.set_page_config(page_title="HKPF Posting Summary Generator", layout="wide")

st.title("📊 HKPF Posting Summary Generator")
st.write("Upload an Excel file to generate a professional posting summary Word document")

# ========= VOCAB + CONFIGURATION =========
STARTER_ROLE_EXPANSIONS = {
    "CSP": "Chief Superintendent of Police",
    "SSP": "Senior Superintendent of Police",
    "SP": "Superintendent of Police",
    "CIP": "Chief Inspector of Police",
    "SIP": "Senior Inspector of Police",
    "IP": "Inspector of Police",
    "PI": "Probationary Inspector of Police",
    "SSGT": "Station Sergeant",
    "SGT": "Sergeant",
    "SPC": "Senior Police Constable",
    "PC": "Police Constable",
    "DC": "District Commander",
    "DDC": "Deputy District Commander",
    "ADC": "Assistant District Commander",
    "ADM": "Administration",
    "A&S": "Administration and Support",
    "CRM": "Crime",
    "ES": "Efficiency Studies",
    "RI": "Research and Inspections",
    "CTRL": "Command and Control (Control Room)",
    "GEN": "General",
    "FLD": "Field",
    "PSU 1": "Patrol Sub-unit 1",
    "PSU 2": "Patrol Sub-unit 2",
    "PSU 3": "Patrol Sub-unit 3",
    "PSU 4": "Patrol Sub-unit 4",
    "TFSU": "Task Force Sub-unit",
    "DVIT 1": "Divisional Investigation Team 1",
    "DVIT 2": "Divisional Investigation Team 2",
    "DVIT 3": "Divisional Investigation Team 3",
    "DVIT 4": "Divisional Investigation Team 4",
    "DVIT 5": "Divisional Investigation Team 5",
    "DVIT 6": "Divisional Investigation Team 6",
    "DVIT 7": "Divisional Investigation Team 7",
    "DVIT 8": "Divisional Investigation Team 8",
    "SDS 1": "Special Duties Squad 1",
    "DSDS 2": "District Special Duties Squad 2",
    "SYMPOSIUM": "Symposium",
    "RPC TRG (INTAKE)": "Recruit Police Constable Training (Intake)",
    "CS&INT": "Counterfeit, Support and Intelligence",
    "INP 2": "Inspection 2",
    "AUX": "Auxiliary",
    "SCIU": "Security Company and Guarding Services Bill-Police Inspection Team",
    "SA": "Security Advisory Section",
    "PCRO": "Police Community Relations Office",
}

STARTER_LOCATION_ALIASES = {
    "CCB": "COMMERCIAL CRIME BUREAU",
    "C DIV CCB": "COMMERCIAL CRIME BUREAU",
    "C DIVISION COMMERCIAL CRIME BUREAU": "COMMERCIAL CRIME BUREAU",
    "CPB": "CRIME PREVENTION BUREAU",
    "PPRB": "POLICE PUBLIC RELATIONS BRANCH",
    "RCCC HKI": "REGIONAL COMMAND AND CONTROL CENTRE HONG KONG ISLAND",
    "HKI": "HONG KONG ISLAND REGIONAL HEADQUARTERS",
    "OPS": "OPERATIONS WING",
    "PTU": "POLICE TACTICAL UNIT",
    "CDIST": "CENTRAL DISTRICT",
    "WDIST": "WESTERN DISTRICT",
    "EDIST": "EASTERN DISTRICT",
    "WCH DIV": "WAN CHAI DIVISION",
    "WCH DIST": "WAN CHAI DISTRICT",
    "CDIV": "CENTRAL DIVISION",
    "WDIV": "WESTERN DIVISION",
    "NPDIV": "NORTH POINT DIVISION",
    "STYSDIV": "STANLEY SUB-DIVISION",
    "ABDDIV": "ABERDEEN DIVISION",
    "PTU A": "POLICE TACTICAL UNIT (A COMPANY)",
    "PTS": "POLICE TRAINING SCHOOL",
    "CAPO": "COMPLAINTS AGAINST POLICE OFFICE",
    "CAPO HKI": "COMPLAINTS AGAINST POLICE OFFICE HONG KONG ISLAND",
    "SQ": "SERVICE QUALITY WING",
    "SUPPORT": "SUPPORT WING",
    "IST": "IN-SERVICE TRAINING",
    "PC TRG": "POLICE CONSTABLE TRAINING DIVISION",
    "TRVE SUP": "TRAINING RESERVE SUPPORT WING",
    "RATU KW": "REGIONAL ANTI TRIAD UNIT KOWLOON WEST",
    "RIU KW": "REGIONAL INTELLIGENCE UNIT KOWLOON WEST",
    "RCCC NTN": "REGIONAL COMMAND AND CONTROL CENTRE NEW TERRITORIES NORTH",
    "EU NTN": "EMERGENCY UNIT NEW TERRITORIES NORTH",
    "KW": "KOWLOON WEST REGIONAL HEADQUARTERS",
    "CRM KW": "CRIME KOWLOON WEST REGIONAL HEADQUARTERS",
    "SMPDIST": "SAU MAU PING DISTRICT",
    "TWDIST": "TSUEN WAN DISTRICT",
    "MKDIST": "MONG KOK DISTRICT",
}

# ========= CORE PROCESSING FUNCTIONS =========
rank_order = ['PC', 'SPC', 'SGT', 'SSGT', 'PI', 'IP', 'SIP', 'CIP', 'SP', 'SSP', 'CSP', 'ACP', 'SACP', 'DCP', 'CP']
rank_index = {r: i for i, r in enumerate(rank_order)}
rank_index['IP/SIP'] = rank_index['IP']

rank_map = {
    'pc': 'PC', 'police constable': 'PC',
    'spc': 'SPC', 'senior police constable': 'SPC',
    'sgt': 'SGT', 'sergeant': 'SGT',
    'ssgt': 'SSGT', 'station sergeant': 'SSGT',
    'pi': 'PI', 'probationary inspector': 'PI', 'probationary inspector of police': 'PI',
    'ip': 'IP', 'insp': 'IP', 'inspector': 'IP', 'inspector of police': 'IP',
    'sip': 'SIP', 'sr insp': 'SIP', 'sen insp': 'SIP', 'senior insp': 'SIP', 'senior inspector': 'SIP', 'senior inspector of police': 'SIP',
    'cip': 'CIP', 'ch insp': 'CIP', 'chief insp': 'CIP', 'chief inspector': 'CIP', 'chief inspector of police': 'CIP',
    'sp': 'SP', 'superintendent': 'SP', 'superintendent of police': 'SP',
    'ssp': 'SSP', 'senior superintendent': 'SSP', 'senior superintendent of police': 'SSP',
    'csp': 'CSP', 'chief superintendent': 'CSP', 'chief superintendent of police': 'CSP',
}

acting_tokens_pattern = re.compile(r'\b(acting|actg|a/|ag\.|temp|temporary|acting up)\b', flags=re.IGNORECASE)

def looks_like_ip_sip(text: str) -> bool:
    s = str(text or "").lower()
    s = acting_tokens_pattern.sub('', s)
    s = s.replace('\\', '/')
    s = re.sub(r'[().,;]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    s2 = s
    s2 = re.sub(r'\bsenior\s+inspector(?:\s+of\s+police)?\b', 'sip', s2)
    s2 = re.sub(r'\bsr\s+insp\b', 'sip', s2)
    s2 = re.sub(r'\bsen\s+insp\b', 'sip', s2)
    s2 = re.sub(r'\binspector(?:\s+of\s+police)?\b', 'ip', s2)
    s2 = re.sub(r'\binsp\b', 'ip', s2)
    tokens = set(re.split(r'[^a-z/]+', s2))
    tokens.discard('')
    return ('ip' in tokens and 'sip' in tokens) or ('ip/sip' in s2)

def map_rank(text: str):
    if looks_like_ip_sip(text):
        return 'IP/SIP'
    s = str(text or "").strip().lower()
    s = acting_tokens_pattern.sub('', s)
    s = s.replace('\\', '/')
    s = re.sub(r'[().,;]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    if not s:
        return None
    if s in rank_map:
        v = rank_map[s]
    else:
        v = None
        for k, val in rank_map.items():
            if re.search(r'\b' + re.escape(k) + r'\b', s):
                v = val
                break
    if v in {'IP', 'SIP'}:
        return 'IP/SIP'
    return v

def is_acting(row) -> bool:
    fields = [
        str(row.get('designation', '') or ''),
        str(row.get('designation_desc', '') or ''),
        str(row.get('post_type', '') or ''),
        str(row.get('post_type_desc', '') or ''),
    ]
    haystack = ' || '.join(fields)
    return bool(acting_tokens_pattern.search(haystack))

def detect_header_row(df_raw: pd.DataFrame) -> int:
    target_headers = {'date start', 'date end', 'post type'}
    header_row = None
    for i in range(min(60, len(df_raw))):
        row_vals = df_raw.iloc[i].astype(str).str.strip().str.lower().tolist()
        if target_headers.issubset(set(row_vals)):
            header_row = i
            break
    if header_row is None:
        raise ValueError("Couldn't find header row with Date Start/Date End/Post Type.")
    return header_row

def snake(s: str) -> str:
    return re.sub(r'\s+', '_', str(s).strip().lower())

def _is_blankish(s: str) -> bool:
    return s is None or (isinstance(s, float) and pd.isna(s)) or str(s).strip().lower() in {"", "nan", "none", "null"}

def clean_role_token(txt: str) -> str:
    if _is_blankish(txt):
        return ""
    s = str(txt).strip()
    s = re.sub(r'\s+', ' ', s).strip()
    if s.lower() in {"nan", "none", "null", "-"}:
        return ""
    return s

def clean_loc_label(txt: str) -> str:
    if _is_blankish(txt):
        return ""
    s = str(txt).strip()
    s = re.sub(r'\s+', ' ', s).strip()
    if s.lower() in {"nan", "none", "null"}:
        return ""
    return s

def normalize_location(label: str, loc_alias) -> str:
    l = clean_loc_label(label)
    if not l:
        return ""
    if l.upper() == "LEAVE RESERVE":
        return ""
    if l in loc_alias:
        return loc_alias[l]
    l_up = l.upper()
    if l_up in loc_alias:
        return loc_alias[l_up]
    return l

def is_rank_text(text: str) -> bool:
    if not text:
        return False
    mapped = map_rank(text)
    return mapped in rank_index

def expand_role(token: str, role_exp) -> str:
    t = clean_role_token(token)
    if not t:
        return ""
    if t in role_exp:
        return role_exp[t]
    t_up = t.upper()
    if t_up in role_exp:
        return role_exp[t_up]
    return t

def process_excel_file(uploaded_file):
    """Process the uploaded Excel file and return enhanced ranges."""
    try:
        # Read Excel file
        xls = pd.ExcelFile(uploaded_file)
        df = pd.read_excel(uploaded_file, sheet_name=xls.sheet_names[0], header=None)
        
        # Detect header
        header_row = detect_header_row(df)
        df = pd.read_excel(uploaded_file, sheet_name=xls.sheet_names[0], header=header_row)
        df = df.loc[:, ~df.columns.astype(str).str.startswith('Unnamed')]
        df = df.dropna(how='all')
        
        # Normalize columns
        aliases = {
            'date_start': 'date_start', 'date_start_(description)': 'date_start_desc',
            'date_end': 'date_end', 'date_end_(description)': 'date_end_desc',
            'post_type': 'post_type', 'post_type_(description)': 'post_type_desc',
            'designation': 'designation', 'designation_(description)': 'designation_desc',
            'location': 'location', 'location_(description)': 'location_desc',
        }
        df = df.rename(columns={c: aliases.get(snake(c), snake(c)) for c in df.columns})
        
        # Parse dates
        for col in ['date_start', 'date_end']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Ensure columns exist
        for col in ['post_type', 'post_type_desc', 'designation', 'designation_desc', 'location', 'location_desc']:
            if col not in df.columns:
                df[col] = ""
        
        # Sort
        if 'date_start' in df.columns and df['date_start'].notna().any():
            df = df.sort_values(by=['date_start', 'date_end'], ascending=[True, True], na_position='last').reset_index(drop=True)
        
        # Map ranks
        df['reported_rank'] = (df['post_type'].astype(str) + ' || ' + df.get('post_type_desc', '').astype(str)).apply(map_rank)
        df['acting_flag'] = df.apply(is_acting, axis=1)
        
        # True rank calculation
        rows = df.to_dict(orient='records')
        
        def future_has_lower_than(rows, from_idx, ref_rank):
            if ref_rank is None:
                return False
            ref_idx = rank_index.get(ref_rank, -1)
            for j in range(from_idx, len(rows)):
                if bool(rows[j].get('acting_flag')):
                    continue
                rj = rows[j].get('reported_rank')
                if rj is None:
                    continue
                if rank_index.get(rj, -1) < ref_idx:
                    return True
            return False
        
        current_substantive_rank = None
        for i, r in enumerate(rows):
            rep = r.get('reported_rank')
            acting = bool(r.get('acting_flag'))
            
            if current_substantive_rank is None:
                if rep is None:
                    r['true_rank'] = None
                else:
                    looks_acting = acting or future_has_lower_than(rows, i + 1, rep)
                    if looks_acting:
                        r['true_rank'] = None
                    else:
                        current_substantive_rank = rep
                        r['true_rank'] = current_substantive_rank
                continue
            
            if rep is None:
                r['true_rank'] = current_substantive_rank
                continue
            
            if acting:
                r['true_rank'] = current_substantive_rank
                continue
            
            cur_idx = rank_index.get(current_substantive_rank, -1)
            rep_idx = rank_index.get(rep, -1)
            
            if rep_idx > cur_idx:
                if future_has_lower_than(rows, i + 1, rep):
                    r['true_rank'] = current_substantive_rank
                else:
                    current_substantive_rank = rep
                    r['true_rank'] = current_substantive_rank
            else:
                r['true_rank'] = current_substantive_rank
        
        first_true = next((rr.get('true_rank') for rr in rows if rr.get('true_rank') is not None), None)
        if first_true is not None:
            for r in rows:
                if r.get('true_rank') is None:
                    r['true_rank'] = first_true
        
        df = pd.DataFrame(rows)
        
        # Year ranges by rank
        segments = []
        current_rank = None
        seg_start = None
        seg_end = None
        
        def max_dt(a, b):
            if pd.isna(a): return b
            if pd.isna(b): return a
            return max(a, b)
        
        def min_dt(a, b):
            if pd.isna(a): return b
            if pd.isna(b): return a
            return min(a, b)
        
        for _, row in df.iterrows():
            tr = row.get('true_rank')
            ds = row.get('date_start')
            de = row.get('date_end')
            if current_rank is None:
                current_rank = tr
                seg_start = ds
                seg_end = de
                continue
            if tr != current_rank:
                segments.append({'true_rank': current_rank, 'start': seg_start, 'end': seg_end})
                current_rank = tr
                seg_start = ds
                seg_end = de
            else:
                seg_start = min_dt(seg_start, ds)
                seg_end = max_dt(seg_end, de)
        
        if current_rank is not None:
            segments.append({'true_rank': current_rank, 'start': seg_start, 'end': seg_end})
        
        def fmt_year_range(start_dt, end_dt):
            if pd.isna(start_dt) and pd.isna(end_dt):
                return "Unknown"
            if pd.isna(start_dt):
                return f"–{end_dt.year}"
            sy = start_dt.year
            if pd.isna(end_dt):
                return f"{sy}–Present"
            ey = end_dt.year
            if ey < sy:
                return f"{sy}"
            return f"{sy}–{ey}"
        
        year_ranges = [
            {'true_rank': seg['true_rank'], 'start': seg.get('start'), 'end': seg.get('end'),
             'year_range': fmt_year_range(seg.get('start'), seg.get('end'))}
            for seg in segments if seg.get('true_rank') is not None
        ]
        
        # Enhanced ranges with locations and roles
        enhanced_ranges = []
        for seg in year_ranges:
            tr = seg['true_rank']
            start_dt = seg['start']
            end_dt = seg['end']
            
            mask = (df['true_rank'] == tr)
            if pd.notna(start_dt):
                mask &= (df['date_start'].isna() | (df['date_start'] >= start_dt))
            if pd.notna(end_dt):
                mask &= (df['date_end'].isna() | (df['date_end'] <= end_dt))
            
            sub = df.loc[mask].copy()
            
            loc_desc = sub['location_desc']
            loc_series = loc_desc.where(
                loc_desc.notna() & (loc_desc.astype(str).str.strip() != ''),
                sub['location']
            ).apply(lambda x: normalize_location(x, STARTER_LOCATION_ALIASES))
            
            def extract_roles_from_row(r):
                roles_out = []
                dd_raw = r.get('designation_desc') or ''
                dd = clean_role_token(expand_role(dd_raw, STARTER_ROLE_EXPANSIONS))
                if dd:
                    roles_out.append(dd)
                
                d_raw = r.get('designation') or ''
                if d_raw and not is_rank_text(d_raw):
                    d = clean_role_token(expand_role(d_raw, STARTER_ROLE_EXPANSIONS))
                    if d and (not dd or d.casefold() != dd.casefold()):
                        roles_out.append(d)
                
                pt = r.get('post_type') or ''
                if pt and not is_rank_text(pt):
                    p = clean_role_token(expand_role(pt, STARTER_ROLE_EXPANSIONS))
                    if p:
                        roles_out.append(p)
                
                seen = set()
                clean_list = []
                for x in roles_out:
                    if not x or x.lower() in {"nan", "none", "null"}:
                        continue
                    key = x.casefold()
                    if key not in seen:
                        seen.add(key)
                        clean_list.append(x)
                return clean_list
            
            sub['role_list'] = sub.apply(extract_roles_from_row, axis=1)
            
            roles_by_loc = {}
            seen_by_loc = {}
            
            for l, roles_here in zip(loc_series, sub['role_list']):
                if not l:
                    continue
                roles_by_loc.setdefault(l, [])
                seen_by_loc.setdefault(l, set())
                for role in roles_here:
                    if not role or role.upper() == "LEAVE RESERVE":
                        continue
                    key = role.casefold()
                    if key not in seen_by_loc[l]:
                        roles_by_loc[l].append(role)
                        seen_by_loc[l].add(key)
            
            unique_locs = []
            seen_locs = set()
            for l in loc_series.tolist():
                if not l:
                    continue
                if l not in seen_locs:
                    seen_locs.add(l)
                    unique_locs.append(l)
            
            enhanced_ranges.append({
                'true_rank': tr,
                'year_range': seg['year_range'],
                'locations': unique_locs,
                'roles_by_location': roles_by_loc
            })
        
        return enhanced_ranges, None
    
    except Exception as e:
        return None, str(e)

def generate_word_document(enhanced_ranges):
    """Generate Word document with a table format."""
    doc = Document()
    
    # Create table with 3 columns (Year Range, Rank, Posting Info)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Light Grid Accent 1'
    
    # Set column widths (narrow for Year Range and Rank, wide for Posting Info)
    from docx.shared import Inches
    table.columns[0].width = Inches(1.0)  # Year Range - narrow
    table.columns[1].width = Inches(1.2)  # Rank - narrow
    table.columns[2].width = Inches(4.3)  # Posting Location & Roles - wide
    
    # Add header row
    header_cells = table.rows[0].cells
    header_cells[0].text = "Year Range"
    header_cells[1].text = "Rank"
    header_cells[2].text = "Posting Location & Roles"
    
    # Format header row - bold and size 13
    for cell in header_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(13)
    
    # Add data rows
    for item in enhanced_ranges:
        year_range = item['year_range']
        rank = item['true_rank']
        locations = item['locations']
        roles_by_location = item['roles_by_location']
        
        # Build posting info text
        posting_info = []
        if locations:
            for loc in locations:
                posting_info.append(f"{loc}")
                roles = roles_by_location.get(loc, [])
                if roles:
                    for role in roles:
                        posting_info.append(f"  • {role}")
        else:
            posting_info.append("(No locations recorded)")
        
        posting_text = "\n".join(posting_info)
        
        # Add row
        row_cells = table.add_row().cells
        row_cells[0].text = year_range
        row_cells[1].text = rank
        row_cells[2].text = posting_text
        
        # Set font size 13 for all cells
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(13)
    
    return doc

# ========= STREAMLIT UI =========
st.markdown("---")

# File uploader
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("📁 Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"✓ File uploaded: {uploaded_file.name}")
    
    # Process button
    if st.button("🔄 Process File", use_container_width=True):
        with st.spinner("Processing your Excel file..."):
            enhanced_ranges, error = process_excel_file(uploaded_file)
            
            if error:
                st.error(f"Error processing file: {error}")
            else:
                st.success(f"✓ Processing complete! Found {len(enhanced_ranges)} rank periods.")
                
                # Generate Word document
                doc = generate_word_document(enhanced_ranges)
                
                # Create downloadable file
                doc_bytes = io.BytesIO()
                doc.save(doc_bytes)
                doc_bytes.seek(0)
                
                # Download button
                st.download_button(
                    label="📥 Download Word Document",
                    data=doc_bytes.getvalue(),
                    file_name="HKPF_Posting_Summary.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                # Show preview
                with st.expander("📋 Preview Summary"):
                    for item in enhanced_ranges:
                        st.write(f"**{item['true_rank']}: {item['year_range']}**")
                        for loc in item['locations']:
                            st.write(f"  • {loc}")
                            roles = item['roles_by_location'].get(loc, [])
                            for role in roles:
                                st.write(f"    - {role}")
else:
    st.info("👆 Please upload an Excel file to get started")
