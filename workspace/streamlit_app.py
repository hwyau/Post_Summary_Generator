import streamlit as st
import pandas as pd
import re
import json
from pathlib import Path

st.set_page_config(page_title="HKPF Posting Summary Generator", layout="wide")
st.title("\U0001F4CA HKPF Posting Summary Generator")

st.markdown("""
**Instructions:** Upload your Excel file. Only the most complete form of each role will be shown for each location. Abbreviations will never appear if a full form exists.
""")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if not uploaded_file:
    st.info("Please upload an Excel file to begin.")
    st.stop()

def clean_role_token(txt):
    if pd.isna(txt) or str(txt).strip().lower() in {"", "nan", "none", "null", "-"}:
        return ""
    s = str(txt).strip()
    s = re.sub(r'\s+', ' ', s).strip()
    return s

CANON_SYNONYMS = {
    r'\\bCMU\\s*REL\\b': 'Community Relations',
    r'^CMU REL$': 'Community Relations',
    r'^C M U REL$': 'Community Relations',
    r'^COMMUNITY RELATIONS$': 'Community Relations',
    r'^COMM REL$': 'Community Relations',
}

def canonicalize_role(name):
    s = clean_role_token(name)
    if not s:
        return ""
    for pattern, repl in CANON_SYNONYMS.items():
        s = re.sub(pattern, repl, s, flags=re.IGNORECASE)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def get_final_designation(row):
    desc = str(row.get('designation_desc', '') or '').strip()
    if desc and desc.lower() not in {"", "nan", "none", "null", "-"}:
        return desc
    desig = str(row.get('designation', '') or '').strip()
    if desig and desig.lower() not in {"", "nan", "none", "null", "-"}:
        return desig
    return ''

def deduplicate_roles(role_list):
    # Canonicalize and keep only the longest unique form for each normalized key
    norm_map = {}
    for r in role_list:
        canon = canonicalize_role(r)
        key = re.sub(r'[^a-z0-9]', '', canon.casefold())
        if not key:
            continue
        if key not in norm_map or len(canon) > len(norm_map[key]):
            norm_map[key] = canon
    # Remove abbreviations if a full form exists
    final_roles = set(norm_map.values())
    to_remove = set()
    for r1 in final_roles:
        for r2 in final_roles:
            if r1 == r2:
                continue
            r1_norm = re.sub(r'\s+', '', r1).lower()
            r2_norm = re.sub(r'\s+', '', r2).lower()
            if r1_norm != r2_norm and r1_norm in r2_norm:
                to_remove.add(r1)
    deduped = [r for r in sorted(final_roles - to_remove)]
    # Final aggressive dedup: remove any role that matches (case-insensitive, ignoring spaces) any other role already in the list
    truly_unique = []
    seen_norms = set()
    for r in deduped:
        norm = re.sub(r'\s+', '', r).lower()
        if norm not in seen_norms:
            truly_unique.append(r)
            seen_norms.add(norm)
    return truly_unique

@st.cache_data(show_spinner=False)
def process_excel(file):
    df = pd.read_excel(file)
    # Ensure columns
    for col in ['designation', 'designation_desc', 'location', 'location_desc']:
        if col not in df.columns:
            df[col] = ''
    df['final_designation'] = df.apply(get_final_designation, axis=1)
    df['designation'] = df['final_designation']
    df['designation_desc'] = df['final_designation']
    # Use location_desc if present, else location
    df['final_location'] = df['location_desc'].where(
        df['location_desc'].astype(str).str.strip().ne(''),
        df['location']
    )
    # Group by location, collect roles
    loc_roles = {}
    for _, row in df.iterrows():
        loc = str(row['final_location']).strip()
        if not loc:
            continue
        role = row['final_designation']
        if not role:
            continue
        loc_roles.setdefault(loc, []).append(role)
    # Deduplicate roles for each location
    for loc in loc_roles:
        loc_roles[loc] = deduplicate_roles(loc_roles[loc])
    return loc_roles

loc_roles = process_excel(uploaded_file)

st.header("Summary by Location")
for loc, roles in loc_roles.items():
    st.markdown(f"**{loc}**")
    for r in roles:
        st.markdown(f"- {r}")
    st.markdown("")
