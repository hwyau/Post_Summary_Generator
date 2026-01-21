import streamlit as st
import pandas as pd
import re
import json
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from datetime import datetime
import io

st.set_page_config(page_title="HKPF Posting Summary Generator", layout="wide")

st.title("📊 HKPF Posting Summary Generator")
st.write("Upload an Excel file to generate a professional posting summary Word document")

# ========= VOCAB + CONFIGURATION =========
STARTER_ROLE_EXPANSIONS = {
	"CSP": "Chief Superintendent",
	"SSP": "Senior Superintendent",
	"SP": "Superintendent",
	"CIP": "Chief Inspector",
	"SIP": "Senior Inspector",
	"IP": "Inspector",
	"PI": "Probationary Inspector",
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
	# ...existing code...
}

# ========= CORE PROCESSING FUNCTIONS =========
# ...existing code...

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
				st.success("✓ Processing complete!")
                
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
import streamlit as st
import pandas as pd
import re
import json
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from datetime import datetime
import io

# ...existing code...
