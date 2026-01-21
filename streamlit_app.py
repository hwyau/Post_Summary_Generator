
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
...existing code from /workspaces/Post_Summary_Generator/workspace/streamlit_app.py (lines 61-1175) pasted here...
...existing code from workspace/streamlit_app.py continues...
