from flask import Flask, render_template, request, redirect, url_for, send_file, flash, after_this_request, session
import json
import os
import re
import zipfile
import time
from docx import Document
from docxcompose.composer import Composer
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
app.secret_key = 'police_gujarat_secure_key'

PREVIEW_HIGHLIGHT_START = "[[[HIGHLIGHT_START]]]"
PREVIEW_HIGHLIGHT_END = "[[[HIGHLIGHT_END]]]"

# Configuration
import tempfile
TEMPLATE_DIR = 'word_templates'
PREVIEW_DIR = 'word_previews'
GENERATED_DIR = tempfile.gettempdir()

# --------------------------------------------------
# DOCUMENT MAPPING LOGIC
# --------------------------------------------------
OFFENCE_MAPPING = {
    "281": [
        "Arrest_memo_281.docx",
        "Bail_Bond_281.docx",
        "Notice_281.docx"
    ],
    "302": [
        "Arrest_Memo_Major.docx",
        "Remand_Application.docx",
        "Panchnama_Scene.docx"
    ],
    "66_1B": [
        "kalam66_b1.docx"
    ],
    "BNS_285": [
        "kalamBNS285.docx"
    ],
    "KALAM_185": [
        "kalam185.docx"
    ],
    "KALAM_65": [
        "kalam_65.docx"
    ],
    "GENERAL": [
        "Standard_Intimation.docx"
    ]
}

def get_template_path(filename, section=None):
    """Get full path to template file, using section if provided"""
    if section and section in OFFENCE_MAPPING:
        return os.path.join(TEMPLATE_DIR, section, filename)
    else:
        for sec in OFFENCE_MAPPING:
            if filename in OFFENCE_MAPPING[sec]:
                return os.path.join(TEMPLATE_DIR, sec, filename)
        return os.path.join(TEMPLATE_DIR, filename)

def get_preview_html_path(filename):
    """Get HTML preview path for a document (if available)."""
    safe_name = os.path.basename(filename)
    base_name = os.path.splitext(safe_name)[0] + '.html'
    return os.path.join(PREVIEW_DIR, base_name)

# --------------------------------------------------
# DATA VARIABLES LIST (For validation and cleaning)
# --------------------------------------------------
REQUIRED_FIELDS = [
    # Accused
    "acc_name", "acc_father", "acc_surname", "acc_alias",
    "acc_gender", "acc_age", "acc_dob",
    "acc_religion", "acc_caste", "acc_subcaste", "acc_nationality", "acc_marital",
    # Occupation
    "occ_type", "occ_place", "occ_income",
    # Perm Address
    "perm_house", "perm_area", "perm_village", "perm_district", "perm_taluka", "perm_state", "perm_pin",
    # Curr Address
    "curr_address", "curr_city", "curr_district", "curr_taluka", "curr_state", "curr_pin",
    # Contact & ID
    "mobile_1", "mobile_2", "id_type", "id_number",
    # Physical
    "phy_height", "phy_build", "phy_complexion", "phy_eyes", "phy_hair", "phy_facial_hair",
    "mark_1", "mark_2", "old_wounds", "other_id_marks",
    # Relative
    "rel_name", "rel_relation", "rel_mobile", "rel_address",
    # Case
    "case_ps", "case_district", "case_taluka",
    "case_banv_date", "case_jann_date", "case_banv_time", "case_jann_time",
    "crime_no", "crime_type", "crime_year",
    "offence_desc", "offence_section",
    "offence_place", "offence_date", "offence_time",
    "vehicle", "vehicle_name", "vehicle_type",
    "investigated_places", "investigated_date", "investigated_time","investigated_end_time"
    "case_diary",
    # Arrest
    "is_arrested", "arrest_date", "arrest_time", "arrest_place", "arrest_entry_no",
    "is_bailed", "bail_authority", "bail_date", "bail_time", "bail_conditions",
    # Status
    "status_proven", "status_chargesheet", "status_case", "status_release_date",
    "intimation_method", "intimation_date", "intimation_time", "intimation_entry_no", "intimation_entry_time",
    # Officer
    "remarks", "io_name", "io_designation", "io_buckle", "io_police_station",
    # Complainant & Bail
    "complain_person", "complain_person_work", "bail_person_name",
    # Panchnama
    "panchanamu_time", "panch1_name", "panch2_name", "pso_name", "pso_police_station", "panch_clerk",
    # Court & Review Officer
    "court_place", "ro_name", "ro_designation", "ro_buckle", "ro_police_station",
    # Authentication
    "auth_date", "auth_place", "auth_print_name",
    # Panchnama Additional
    "panchnumu_date", "panchnamu_time",
    # Complainant Additional
    "comp_per_mobile", "comp_per_address",
    # Panch 1 Additional
    "panch1_surname", "panch1_age", "panch1_work", "panch1_address",
    # Panch 2 Additional
    "panch2_surname", "panch2_age", "panch2_work", "panch2_address"
]

# --------------------------------------------------
# FIELD LABELS MAPPING (For displaying Gujarati labels)
# --------------------------------------------------
FIELD_LABELS = {
    # Accused Details
    "acc_name": "નામ (Name)",
    "acc_father": "પિતાનું નામ (Father's Name)",
    "acc_surname": "અટક (Surname)",
    "acc_alias": "ઉર્ફે / ઉપનામ (Alias)",
    "acc_gender": "લિંગ (Gender)",
    "acc_age": "ઉંમર (Age)",
    "acc_dob": "જન્મ તારીખ (DOB)",
    "acc_religion": "ધર્મ (Religion)",
    "acc_caste": "જાતિ (Caste)",
    "acc_subcaste": "પેટાજાતિ (Subcaste)",
    "acc_nationality": "રાષ્ટ્રીયતા (Nationality)",
    "acc_marital": "વૈવાહિક સ્થિતિ (Marital Status)",
    # Occupation
    "occ_type": "ધંધો (Occupation Type)",
    "occ_place": "ધંધાનું સ્થળ (Place of Work)",
    "occ_income": "આવક (Income)",
    # Permanent Address
    "perm_house": "મકાન નં (House No)",
    "perm_area": "વિસ્તાર (Area)",
    "perm_village": "ગામ (Village)",
    "perm_taluka": "તાલુકો (Taluka)",
    "perm_district": "જીલ્લો (District)",
    "perm_state": "રાજ્ય (State)",
    "perm_pin": "પિનકોડ (Pincode)",
    # Current Address
    "curr_address": "હાલનું સરનામું (Full Current Address Line)",
    "curr_city": "શહેર/ગામ (City/Village)",
    "curr_taluka": "તાલુકો (Taluka)",
    "curr_district": "જીલ્લો (District)",
    "curr_state": "રાજ્ય (State)",
    "curr_pin": "પિનકોડ (Pincode)",
    # Contact & ID
    "mobile_1": "મોબાઈલ ૧ (Primary Mobile)",
    "mobile_2": "મોબાઈલ ૨ (Alt Mobile)",
    "id_type": "ઓળખ પત્ર પ્રકાર (ID Type)",
    "id_number": "ઓળખ પત્ર નંબર (ID Number)",
    # Physical Description
    "phy_height": "ઊંચાઈ (Height)",
    "phy_build": "બાંધો (Build)",
    "phy_complexion": "વર્ણ/રંગ (Complexion)",
    "phy_eyes": "આંખોનો રંગ (Eye Color)",
    "phy_hair": "વાળનો રંગ (Hair Color)",
    "phy_facial_hair": "દાઢી/મૂછ (Beard/Mustache)",
    # Marks
    "mark_1": "નિશાન ૧ (Id Mark 1)",
    "mark_2": "નિશાન ૨ (Id Mark 2)",
    "old_wounds": "જુના ઘા (Old Wounds)",
    "other_id_marks": "અન્ય ઓળખ (Other Marks)",
    # Relative Intimation
    "rel_name": "સગાનું નામ (Relative Name)",
    "rel_relation": "સંબંધ (Relation)",
    "rel_mobile": "મોબાઈલ (Relative Mobile)",
    "rel_address": "સરનામું (Relative Address)",
    # Case Details
    "case_ps": "પોલીસ સ્ટેશન (Police Station)",
    "case_district": "જીલ્લો (District)",
    "case_taluka": "તાલુકો (Taluka)",
    "case_banv_date": "ગુ.બ.તા.સ. તારીખ",
    "case_jann_date": "ગુ.જા.તા.સ. તારીખ",
    "case_banv_time": "ગુ.બ.તા.સ. સમય",
    "case_jann_time": "ગુ.જા.તા.સ. સમય",
    "crime_no": "ગુન્હા રજી. નંબર (Crime No)",
    "crime_type": "ગુન્હા પ્રકાર (Crime Type)",
    "crime_year": "વર્ષ (Year)",
    # Offence
    "offence_desc": "ગુન્હો / તહોમત (Offence Description)",
    "offence_section": "લાગુ કલમ (Applicable Section)",
    "offence_place": "ગુન્હાનું સ્થળ (Place of Offence)",
    "offence_date": "ગુન્હો તારીખ (Date of Offence)",
    "offence_time": "ગુન્હો સમય (Time of Offence)",
    "vehicle": "વાહનનો નંબર (Vehicle No)",
    "vehicle_name": "વાહનનુ નામ (Vehicle Name)",
    "vehicle_type": "વાહન પ્રકાર (Vehicle Type)",
    "investigated_places": "તપાસેલ જગ્યાઓ (Investigated Places)",
    "investigated_date": "તપાસ તારીખ (Investigated Date)",
    "investigated_time": "તપાસ સમય (Investigated Time)",
    "investigated_end_time":"તપાસ પુરી કર્યાનો સમય",
    "case_diary": "કેસ ડાયરી (Case Diary)",
    # Arrest
    "is_arrested": "અટકાયત છે? (Is Arrested?)",
    "arrest_date": "અટકાયત તારીખ (Arrest Date)",
    "arrest_time": "અટકાયત સમય (Arrest Time)",
    "arrest_place": "અટકાયત સ્થળ (Place of Arrest)",
    "arrest_entry_no": "સ્ટેશન ડાયરી એન્ટ્રી (Station Diary Entry No)",
    # Bail
    "is_bailed": "જામીન પર? (Is Bailed?)",
    "bail_authority": "જામીન આપનાર (Bail Authority)",
    "bail_date": "જામીન તારીખ (Bail Date)",
    "bail_time": "જામીન સમય (Bail Time)",
    "bail_conditions": "શરતો (Conditions)",
    # Officer & Remarks
    "remarks": "શેરો / નોંધ (Remarks)",
    "io_name": "અમલદારનું નામ (Investigating Officer Name)",
    "io_designation": "હોદ્દો (Designation)",
    "io_buckle": "બકલ નં (Buckle No)",
    "io_police_station": "પોલીસ સ્ટેશન (IO Police Station)",
    # Complainant & Bail Person
    "complain_person": "ફરીયાદીનુ નામ (Complainant Name)",
    "complain_person_work": "નોકરી / પોલીસ સ્ટેશન (Work / Police Station)",
    "bail_person_name": "જામીનનુ નામ (Bail Person Name)",
    # Panchnama
    "panchanamu_time": "પંચનામા નો સમય (Panchnama Time)",
    "panchnumu_date": "પંચનામા તારીખ (Panchnama Date)",
    "panchnamu_time": "પંચનામા સમય (Panchnama Time)",
    "panch1_name": "પંચ-૧ (Panch 1)",
    "panch1_surname": "પંચ-૧ અટક (Panch 1 Surname)",
    "panch1_age": "પંચ-૧ ઉંમર (Panch 1 Age)",
    "panch1_work": "પંચ-૧ ધંધો (Panch 1 Work)",
    "panch1_address": "પંચ-૧ સરનામું (Panch 1 Address)",
    "panch2_name": "પંચ-૨ (Panch 2)",
    "panch2_surname": "પંચ-૨ અટક (Panch 2 Surname)",
    "panch2_age": "પંચ-૨ ઉંમર (Panch 2 Age)",
    "panch2_work": "પંચ-૨ ધંધો (Panch 2 Work)",
    "panch2_address": "પંચ-૨ સરનામું (Panch 2 Address)",
    "pso_name": "PSO નુ નામ (PSO Name)",
    "pso_police_station": "PSO પોલીસ સ્ટેશન (PSO Police Station)",
    "panch_clerk": "પંચનામુ કલાક (Panchnama Clerk)",
    # Complainant Additional
    "comp_per_mobile": "ફરીયાદી મોબાઈલ (Complainant Mobile)",
    "comp_per_address": "ફરીયાદી સરનામું (Complainant Address)",
    # Court & Review Officer
    "court_place": "કોર્ટ (Court Place)",
    "ro_name": "RO નામ (Review Officer Name)",
    "ro_designation": "હોદ્દો (Designation)",
    "ro_buckle": "બકલ નં (Buckle No)",
    "ro_police_station": "પોલીસ સ્ટેશન (Police Station)",
    # Authentication
    "auth_date": "દસ્તાવેજ તારીખ (Document Date)",
    "auth_place": "સ્થળ (Place)",
    "auth_print_name": "પ્રિન્ટ નામ (Name for Signature)",
    # Status
    "status_proven": "સ્થિતિ (પ્રમાણિત) (Status)",
    "status_chargesheet": "ચાર્જશીટ સ્થિતિ (Chargesheet Status)",
    "status_case": "કેસ સ્થિતિ (Case Status)",
    "status_release_date": "રિલીઝ તારીખ (Release Date)",
    "intimation_method": "સૂચના રીત (Intimation Method)",
    "intimation_date": "સૂચના તારીખ (Intimation Date)",
    "intimation_time": "સૂચના સમય (Intimation Time)",
    "intimation_entry_no": "સૂચના એન્ટ્રી નં (Intimation Entry No)",
    "intimation_entry_time": "સૂચના એન્ટ્રી સમય (Intimation Entry Time)",
}

# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------
def load_data():
    if 'case_data' in session:
        return session['case_data']
    return {}

@app.before_request
def sync_session():
    if 'case_data' not in session:
        session['case_data'] = {}

def save_data(data):
    session['case_data'] = data
    session.modified = True

GUJ_DIGIT_TRANS = str.maketrans("0123456789", "૦૧૨૩૪૫૬૭૮૯")

# Date fields that need DD-MM-YYYY formatting
DATE_FIELDS = [
    'offence_date', 'arrest_date', 'bail_date', 'status_release_date',
    'intimation_date', 'panchnumu_date', 'auth_date', 'acc_dob',
    'case_banv_date', 'case_jann_date', 'investigated_date'
]

def format_date_dd_mm_yyyy(date_str):
    """Convert date from YYYY-MM-DD to DD-MM-YYYY format"""
    if not date_str or not isinstance(date_str, str):
        return date_str
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        parts = date_str.split('-')
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
    return date_str

def to_gujarati_digits(value):
    if value is None:
        return ""
    return str(value).translate(GUJ_DIGIT_TRANS)

def convert_data_to_gujarati(data):
    """Convert data to Gujarati digits, with date formatting to DD-MM-YYYY"""
    result = {}
    for key, value in (data or {}).items():
        if key in DATE_FIELDS:
            value = format_date_dd_mm_yyyy(value)
        result[key] = to_gujarati_digits(value)
    return result

def replace_text_in_paragraph(paragraph, data):
    """Optimized regex-based placeholder replacement for faster processing"""
    if not paragraph.runs:
        return

    full_text = "".join(run.text for run in paragraph.runs)
    modified = False

    for key in REQUIRED_FIELDS:
        placeholder = f"[{key}]"
        value = str(data.get(key, ""))
        if placeholder in full_text:
            full_text = full_text.replace(placeholder, value)
            modified = True

    if modified:
        for run in paragraph.runs:
            run._element.getparent().remove(run._element)
        paragraph.add_run(full_text)

# --------------------------------------------------
# OCCURRENCE-BASED TEXT EDIT HELPERS (NEW)
# --------------------------------------------------

def replace_nth_in_text(text, old, new, n):
    """Replace only the nth occurrence (0-indexed) of `old` in `text`.
    Returns text unchanged if the nth occurrence is not found."""
    count = 0
    start = 0
    while True:
        idx = text.find(old, start)
        if idx == -1:
            return text
        if count == n:
            return text[:idx] + new + text[idx + len(old):]
        count += 1
        start = idx + len(old)

def normalize_text_edits(text_edits):
    """Normalize text edits into a dict mapping original -> replacement (legacy use only)."""
    if isinstance(text_edits, list):
        normalized = {}
        for edit in text_edits:
            if not isinstance(edit, dict):
                continue
            original = edit.get('originalText')
            replacement = edit.get('newText')
            if isinstance(original, str) and isinstance(replacement, str):
                normalized[original] = replacement
        return normalized

    if isinstance(text_edits, dict):
        return {
            key: str(value)
            for key, value in text_edits.items()
            if isinstance(key, str)
        }

    return {}

def get_case_edits_list_for_section(case_edits, section):
    """Return edits as a list of {originalText, newText, occurrenceIndex} for the section.
    Handles both new array format and legacy flat object format."""
    # Try to get section-specific edits first
    if isinstance(case_edits, dict) and section in case_edits:
        raw = case_edits[section]
    else:
        raw = case_edits

    if isinstance(raw, list):
        result = []
        for e in raw:
            if not isinstance(e, dict):
                continue
            original = e.get('originalText')
            replacement = e.get('newText', '')
            if isinstance(original, str):
                result.append({
                    'originalText': original,
                    'newText': str(replacement),
                    'occurrenceIndex': int(e.get('occurrenceIndex', 0))
                })
        return result

    if isinstance(raw, dict):
        # Legacy flat object format — treat every edit as occurrenceIndex 0
        return [
            {'originalText': k, 'newText': str(v), 'occurrenceIndex': 0}
            for k, v in raw.items() if isinstance(k, str)
        ]

    return []

def get_case_edits_for_section(case_edits, section):
    """Legacy: Return edits as a dict. Kept for any old callers."""
    if isinstance(case_edits, dict) and section in case_edits:
        return normalize_text_edits(case_edits.get(section))
    return normalize_text_edits(case_edits)

def collect_all_paragraphs_from_doc(doc):
    """Recursively collect ALL paragraphs in document order,
    including paragraphs inside nested tables at any depth."""
    from docx.text.paragraph import Paragraph as DocxPara
    from docx.table import Table as DocxTable

    result = []

    def recurse(element):
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        if tag == 'p':
            result.append(DocxPara(element, doc))
        elif tag == 'tbl':
            tbl = DocxTable(element, doc)
            for row in tbl.rows:
                for cell in row.cells:
                    for child in cell._element:
                        recurse(child)
        else:
            # Recurse into any other container (e.g. w:sdt content controls)
            for child in element:
                recurse(child)

    for element in doc.element.body:
        recurse(element)

    return result

def apply_occurrence_replacements_to_doc(doc, text_edits_list):
    """Apply occurrence-indexed text edits across the entire document.
    Each edit replaces ONLY the Nth occurrence (0-indexed) of originalText globally,
    scanning all paragraphs in document order including nested tables."""
    if not text_edits_list:
        return

    all_paragraphs = collect_all_paragraphs_from_doc(doc)

    for edit in text_edits_list:
        original = edit.get('originalText', '')
        replacement = edit.get('newText', '')
        target_n = int(edit.get('occurrenceIndex', 0))

        if not original:
            continue

        global_count = 0
        replaced = False

        for para in all_paragraphs:
            if replaced:
                break
            if not para.runs:
                continue

            full_text = "".join(run.text for run in para.runs)
            local_occurrences = full_text.count(original)

            if local_occurrences == 0:
                continue

            if global_count + local_occurrences > target_n:
                # The target occurrence falls inside this paragraph
                local_target = target_n - global_count
                new_full_text = replace_nth_in_text(full_text, original, replacement, local_target)
                if new_full_text != full_text:
                    for run in para.runs:
                        run._element.getparent().remove(run._element)
                    para.add_run(new_full_text)
                replaced = True
            else:
                global_count += local_occurrences

def generate_document(template_name, data, output_path, section=None):
    template_path = get_template_path(template_name, section)
    if not os.path.exists(template_path):
        return False

    doc = Document(template_path)
    data_for_doc = convert_data_to_gujarati(data)

    # Step 1: Replace [placeholder] variables across all paragraphs
    # Uses recursive collector to handle nested tables at any depth
    all_paragraphs = collect_all_paragraphs_from_doc(doc)
    for p in all_paragraphs:
        replace_text_in_paragraph(p, data_for_doc)

    # Step 2: Apply occurrence-specific user edits across the whole document
    case_edits = session.get('case_edits', {})
    text_edits_list = get_case_edits_list_for_section(case_edits, section)

    print(f"=== generate_document: section={section} ===")
    print(f"=== case_edits from session: {case_edits} ===")
    print(f"=== text_edits_list resolved: {text_edits_list} ===")

    if text_edits_list:
        apply_occurrence_replacements_to_doc(doc, text_edits_list)

    doc.save(output_path)
    return True

# --------------------------------------------------
# PREVIEW HELPERS
# --------------------------------------------------

def replace_text_in_paragraph_for_preview(paragraph, data):
    """Replace placeholders and wrap values with highlight markers for preview."""
    if not paragraph.runs:
        return

    full_text = "".join(run.text for run in paragraph.runs)
    modified = False

    for key in REQUIRED_FIELDS:
        placeholder = f"[{key}]"
        if placeholder in full_text:
            value = str(data.get(key, ""))
            full_text = full_text.replace(
                placeholder,
                f"{PREVIEW_HIGHLIGHT_START}{value}{PREVIEW_HIGHLIGHT_END}"
            )
            modified = True

    if modified:
        for run in paragraph.runs:
            run._element.getparent().remove(run._element)
        paragraph.add_run(full_text)

def extract_placeholders_from_docx(template_path):
    if not os.path.exists(template_path):
        return []

    doc = Document(template_path)
    found = []
    seen = set()
    required_set = set(REQUIRED_FIELDS)
    pattern = re.compile(r"\[([A-Za-z0-9_]+)\]")

    def scan_text(text):
        for match in pattern.finditer(text or ""):
            key = match.group(1)
            if key in required_set and key not in seen:
                seen.add(key)
                found.append(key)

    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for element in doc.element.body:
        if isinstance(element, CT_P):
            paragraph = Paragraph(element, doc)
            scan_text(paragraph.text)
        elif isinstance(element, CT_Tbl):
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        scan_text(p.text)

    return found

def extract_preview_from_docx(template_path, data):
    if not os.path.exists(template_path):
        return ""

    doc = Document(template_path)
    html_content = "<style>"
    html_content += """
    .docx-preview-para { margin: 0.5em 0; line-height: 1.4; }
    .docx-preview-table { width: 100%; margin: 0.5em 0; border-collapse: collapse; border: 1px solid #999; }
    .docx-preview-table td { border: 1px solid #999; padding: 8px; vertical-align: top; }
    .docx-preview-table th { border: 1px solid #999; padding: 8px; background-color: #e9ecef; font-weight: bold; }
    .docx-preview-run { display: inline; }
    .docx-preview-highlight { background: #FFFF00; padding: 0 2px; border-radius: 2px; }
    </style>"""

    for p in doc.paragraphs:
        p_copy = Document().add_paragraph()
        for run in p.runs:
            p_copy.add_run(run.text)
        replace_text_in_paragraph_for_preview(p_copy, data)

        p_html = "<p class='docx-preview-para' style='"

        alignment = p.alignment
        if alignment == 1:
            p_html += "text-align: center; "
        elif alignment == 2:
            p_html += "text-align: right; "
        elif alignment == 3:
            p_html += "text-align: justify; "

        if p.paragraph_format.left_indent:
            p_html += f"margin-left: {p.paragraph_format.left_indent.pt}pt; "
        if p.paragraph_format.right_indent:
            p_html += f"margin-right: {p.paragraph_format.right_indent.pt}pt; "
        if p.paragraph_format.space_before:
            p_html += f"margin-top: {p.paragraph_format.space_before.pt}pt; "
        if p.paragraph_format.space_after:
            p_html += f"margin-bottom: {p.paragraph_format.space_after.pt}pt; "

        p_html += "'>"

        for run in p_copy.runs:
            run_html = "<span class='docx-preview-run' style='"
            if run.font.size:
                run_html += f"font-size: {run.font.size.pt}pt; "
            if run.font.bold:
                run_html += "font-weight: bold; "
            if run.font.italic:
                run_html += "font-style: italic; "
            if run.font.underline:
                run_html += "text-decoration: underline; "
            if run.font.color and run.font.color.rgb:
                run_html += f"color: #{str(run.font.color.rgb)}; "

            run_text = run.text or ""
            run_text = run_text.replace(PREVIEW_HIGHLIGHT_START, "<span class='docx-preview-highlight'>")
            run_text = run_text.replace(PREVIEW_HIGHLIGHT_END, "</span>")
            run_html += "'>" + run_text + "</span>"
            p_html += run_html

        p_html += "</p>"

        if p_copy.text.strip():
            html_content += p_html
        else:
            html_content += "<p class='docx-preview-para' style='height: 0.5em;'></p>"

    for table in doc.tables:
        html_content += "<table class='docx-preview-table'>"

        for row_idx, row in enumerate(table.rows):
            html_content += "<tr>"

            for cell in row.cells:
                is_header = row_idx == 0
                tag = "th" if is_header else "td"
                cell_html = f"<{tag}>"

                for p in cell.paragraphs:
                    p_copy = Document().add_paragraph()
                    for run in p.runs:
                        p_copy.add_run(run.text)
                    replace_text_in_paragraph_for_preview(p_copy, data)

                    cell_p_html = "<div style='"
                    if p.alignment == 1:
                        cell_p_html += "text-align: center; "
                    elif p.alignment == 2:
                        cell_p_html += "text-align: right; "
                    cell_p_html += "'>"

                    for run in p_copy.runs:
                        run_text = run.text or ""
                        run_text = run_text.replace(PREVIEW_HIGHLIGHT_START, "<span class='docx-preview-highlight'>")
                        run_text = run_text.replace(PREVIEW_HIGHLIGHT_END, "</span>")
                        if run.font.bold:
                            run_text = f"<b>{run_text}</b>"
                        if run.font.italic:
                            run_text = f"<i>{run_text}</i>"
                        cell_p_html += run_text

                    cell_p_html += "</div>"
                    cell_html += cell_p_html

                cell_html += f"</{tag}>"
                html_content += cell_html

            html_content += "</tr>"

        html_content += "</table>"

    return html_content

def generate_pdf_from_docx(docx_path, pdf_path):
    """Convert DOCX to PDF preserving exact Word formatting"""
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from docx.shared import RGBColor
        import sys

        gujarati_font_name = 'GujaratiFont'
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))

            embedded_fonts = [
                os.path.join(base_dir, 'fonts', 'gujarati', 'NotoSansGujarati-VariableFont_wdth,wght.ttf'),
                os.path.join(base_dir, 'fonts', 'gujarati', 'lohit_gu.ttf'),
            ]

            if sys.platform == 'win32':
                system_fonts = [
                    r'C:\Windows\Fonts\nirmala.ttf',
                    r'C:\Windows\Fonts\shruti.ttf',
                ]
                font_paths = embedded_fonts + system_fonts
            else:
                font_paths = embedded_fonts

            font_registered = False
            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont(gujarati_font_name, font_path))
                    print(f"Gujarati font registered: {os.path.basename(font_path)}")
                    font_registered = True
                    break

            if not font_registered:
                print("Warning: No Gujarati font found, falling back to Helvetica")
                gujarati_font_name = 'Helvetica'
        except Exception as e:
            print(f"Font registration error: {e}")
            gujarati_font_name = 'Helvetica'

        doc = Document(docx_path)

        section = doc.sections[0]
        left_margin = section.left_margin.inches if section.left_margin else 1
        right_margin = section.right_margin.inches if section.right_margin else 1
        top_margin = section.top_margin.inches if section.top_margin else 1
        bottom_margin = section.bottom_margin.inches if section.bottom_margin else 1

        pdf = SimpleDocTemplate(pdf_path, pagesize=A4,
                                rightMargin=right_margin*inch, leftMargin=left_margin*inch,
                                topMargin=top_margin*inch, bottomMargin=bottom_margin*inch)

        story = []

        def get_color(color_obj):
            if color_obj and isinstance(color_obj, RGBColor):
                return colors.Color(color_obj.r/255, color_obj.g/255, color_obj.b/255)
            return colors.black

        def build_formatted_text(para):
            html_parts = []
            for run in para.runs:
                text = run.text.replace('<', '&lt;').replace('>', '&gt;').replace('&', '&amp;')
                font_size = run.font.size.pt if run.font.size else 11
                is_bold = run.bold
                is_italic = run.italic
                is_underline = run.underline
                color = get_color(run.font.color.rgb if run.font.color else None)

                if is_bold:
                    text = f'<b>{text}</b>'
                if is_italic:
                    text = f'<i>{text}</i>'
                if is_underline:
                    text = f'<u>{text}</u>'
                if font_size != 11:
                    text = f'<font size="{int(font_size)}">{text}</font>'
                if color != colors.black:
                    text = f'<font color="#{int(color.red*255):02x}{int(color.green*255):02x}{int(color.blue*255):02x}">{text}</font>'

                html_parts.append(text)

            return ''.join(html_parts)

        for para in doc.paragraphs:
            if not para.text.strip():
                story.append(Spacer(1, 0.05*inch))
                continue

            alignment = TA_LEFT
            if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                alignment = TA_CENTER
            elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                alignment = TA_RIGHT
            elif para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                alignment = TA_JUSTIFY

            space_before = para.paragraph_format.space_before.pt if para.paragraph_format.space_before else 0
            space_after = para.paragraph_format.space_after.pt if para.paragraph_format.space_after else 6

            font_size = 11
            if para.runs and para.runs[0].font.size:
                font_size = para.runs[0].font.size.pt

            line_spacing = para.paragraph_format.line_spacing
            if line_spacing:
                leading = font_size * line_spacing if isinstance(line_spacing, (int, float)) else font_size * 1.2
            else:
                leading = font_size * 1.2

            left_indent = para.paragraph_format.left_indent.inches if para.paragraph_format.left_indent else 0
            first_line_indent = para.paragraph_format.first_line_indent.inches if para.paragraph_format.first_line_indent else 0

            para_style = ParagraphStyle(
                f'Para_{id(para)}',
                fontName=gujarati_font_name,
                fontSize=font_size,
                leading=leading,
                alignment=alignment,
                spaceBefore=space_before,
                spaceAfter=space_after,
                leftIndent=left_indent*inch,
                firstLineIndent=first_line_indent*inch
            )

            formatted_text = build_formatted_text(para)
            p = Paragraph(formatted_text, para_style)
            story.append(p)

        for table in doc.tables:
            table_data = []
            table_styles = [
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), gujarati_font_name),
            ]

            for row_idx, row in enumerate(table.rows):
                row_data = []
                for col_idx, cell in enumerate(row.cells):
                    cell_paragraphs = []
                    for p in cell.paragraphs:
                        if p.text.strip():
                            cell_paragraphs.append(p.text)
                    cell_text = '\n'.join(cell_paragraphs) if cell_paragraphs else ''
                    row_data.append(cell_text)

                    try:
                        cell_color = cell._element.get_or_add_tcPr().shd_val
                        if cell_color:
                            bg_color = colors.HexColor(f'#{cell_color}')
                            table_styles.append(('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), bg_color))
                    except:
                        pass

                table_data.append(row_data)

            if table_data:
                table_styles.extend([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ])

                t = Table(table_data)
                t.setStyle(TableStyle(table_styles))
                story.append(t)
                story.append(Spacer(1, 0.15*inch))

        pdf.build(story)
        return os.path.exists(pdf_path)
    except Exception as e:
        print(f"PDF generation error: {e}")
        import traceback
        traceback.print_exc()
        return False

# --------------------------------------------------
# API ENDPOINTS FOR LOCALSTORAGE SYNC
# --------------------------------------------------

@app.route('/api/save_data', methods=['POST'])
def api_save_data():
    payload = request.get_json() or {}

    if 'case_data' in payload:
        save_data(payload['case_data'])
        if 'case_edits' in payload:
            session['case_edits'] = payload['case_edits']
            session.modified = True
            print(f"=== api_save_data: case_edits saved ===")
            print(session['case_edits'])
    else:
        # Backwards compatibility: payload is direct data
        save_data(payload)

    return {'status': 'success'}, 200

@app.route('/api/load_data', methods=['GET'])
def api_load_data():
    data = load_data()
    return data, 200

# --------------------------------------------------
# ROUTES
# --------------------------------------------------

@app.route('/', methods=['GET', 'POST'])
def home():
    data = load_data()
    if request.method == 'POST':
        for field in REQUIRED_FIELDS:
            data[field] = request.form.get(field, "")
        save_data(data)
        if 'proceed' in request.form:
            return redirect(url_for('documents'))
    return render_template('index.html', data=data, fields=REQUIRED_FIELDS)

@app.route('/documents', methods=['GET', 'POST'])
def documents():
    data = load_data()

    if request.method == 'POST':
        for key in request.form:
            if key in REQUIRED_FIELDS:
                data[key] = request.form[key]
        save_data(data)

    section = data.get('offence_section', 'GENERAL')
    doc_list = OFFENCE_MAPPING.get(section, OFFENCE_MAPPING['GENERAL'])

    required_in_docs = set()
    ordered_required = []
    seen_required = set()
    for filename in doc_list:
        template_path = get_template_path(filename, section)
        placeholders = extract_placeholders_from_docx(template_path)
        required_in_docs.update(placeholders)
        for key in placeholders:
            if key not in seen_required:
                seen_required.add(key)
                ordered_required.append(key)
    missing_fields = [f for f in ordered_required if not data.get(f)]

    required_fields_list = sorted(list(required_in_docs))

    return render_template('documents.html',
                           data=data,
                           docs=doc_list,
                           missing_fields=missing_fields,
                           section=section,
                           field_labels=FIELD_LABELS,
                           required_fields=required_fields_list,
                           required_in_docs=required_in_docs)

@app.route('/download_single/<filename>')
def download_single(filename):
    data = load_data()
    section = data.get('offence_section', 'GENERAL')
    output_path = os.path.join(GENERATED_DIR, f"Filled_{filename}")

    success = generate_document(filename, data, output_path, section)
    if success:
        @after_this_request
        def cleanup(response):
            try:
                os.remove(output_path)
            except:
                pass
            return response
        return send_file(output_path, as_attachment=True)
    return "Template not found", 404

@app.route('/download_single_pdf/<filename>')
def download_single_pdf(filename):
    data = load_data()
    section = data.get('offence_section', 'GENERAL')
    output_docx = os.path.join(GENERATED_DIR, f"Filled_{filename}")
    output_pdf = os.path.join(GENERATED_DIR, f"Filled_{os.path.splitext(filename)[0]}.pdf")

    success = generate_document(filename, data, output_docx, section)
    if not success:
        return "Template not found", 404

    if not generate_pdf_from_docx(output_docx, output_pdf):
        return "PDF generation failed", 500

    @after_this_request
    def cleanup(response):
        try:
            os.remove(output_docx)
            os.remove(output_pdf)
        except:
            pass
        return response
    return send_file(output_pdf, as_attachment=True)

@app.route('/download_all_zip')
def download_all_zip():
    data = load_data()
    section = data.get('offence_section', 'GENERAL')
    doc_list = OFFENCE_MAPPING.get(section, OFFENCE_MAPPING['GENERAL'])

    zip_path = os.path.join(GENERATED_DIR, "All_Documents.zip")
    generated_files = []
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for filename in doc_list:
            output_name = f"Filled_{filename}"
            output_path = os.path.join(GENERATED_DIR, output_name)
            if generate_document(filename, data, output_path, section):
                zipf.write(output_path, arcname=output_name)
                generated_files.append(output_path)

    @after_this_request
    def cleanup(response):
        try:
            os.remove(zip_path)
            for f in generated_files:
                try:
                    os.remove(f)
                except:
                    pass
        except:
            pass
        return response
    return send_file(zip_path, as_attachment=True)

@app.route('/download_all_pdf_zip')
def download_all_pdf_zip():
    data = load_data()
    section = data.get('offence_section', 'GENERAL')
    doc_list = OFFENCE_MAPPING.get(section, OFFENCE_MAPPING['GENERAL'])

    zip_path = os.path.join(GENERATED_DIR, "All_Documents_PDF.zip")
    generated_files = []
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for filename in doc_list:
            output_docx = os.path.join(GENERATED_DIR, f"Filled_{filename}")
            output_pdf = os.path.join(GENERATED_DIR, f"Filled_{os.path.splitext(filename)[0]}.pdf")
            if generate_document(filename, data, output_docx, section):
                if generate_pdf_from_docx(output_docx, output_pdf):
                    zipf.write(output_pdf, arcname=os.path.basename(output_pdf))
                    generated_files.extend([output_docx, output_pdf])

    @after_this_request
    def cleanup(response):
        try:
            os.remove(zip_path)
            for f in generated_files:
                try:
                    os.remove(f)
                except:
                    pass
        except:
            pass
        return response
    return send_file(zip_path, as_attachment=True)

@app.route('/download_merged')
def download_merged():
    data = load_data()
    section = data.get('offence_section', 'GENERAL')
    doc_list = OFFENCE_MAPPING.get(section, OFFENCE_MAPPING['GENERAL'])

    if not doc_list:
        return "No documents to merge", 400

    base_path = os.path.join(GENERATED_DIR, "Merged_Master.docx")
    generate_document(doc_list[0], data, base_path, section)

    master_doc = Document(base_path)
    composer = Composer(master_doc)

    temp_files = []
    for filename in doc_list[1:]:
        temp_path = os.path.join(GENERATED_DIR, f"temp_{filename}")
        generate_document(filename, data, temp_path, section)
        doc_to_append = Document(temp_path)
        master_doc.add_page_break()
        composer.append(doc_to_append)
        temp_files.append(temp_path)

    composer.save(base_path)

    @after_this_request
    def cleanup(response):
        try:
            os.remove(base_path)
            for f in temp_files:
                try:
                    os.remove(f)
                except:
                    pass
        except:
            pass
        return response
    return send_file(base_path, as_attachment=True)

@app.route('/download_merged_pdf')
def download_merged_pdf():
    data = load_data()
    section = data.get('offence_section', 'GENERAL')
    doc_list = OFFENCE_MAPPING.get(section, OFFENCE_MAPPING['GENERAL'])

    if not doc_list:
        return "No documents to merge", 400

    base_path = os.path.join(GENERATED_DIR, "Merged_Master.docx")
    generate_document(doc_list[0], data, base_path, section)

    master_doc = Document(base_path)
    composer = Composer(master_doc)

    temp_files = []
    for filename in doc_list[1:]:
        temp_path = os.path.join(GENERATED_DIR, f"temp_{filename}")
        generate_document(filename, data, temp_path, section)
        doc_to_append = Document(temp_path)
        master_doc.add_page_break()
        composer.append(doc_to_append)
        temp_files.append(temp_path)

    composer.save(base_path)

    merged_pdf = os.path.join(GENERATED_DIR, "Merged_Master.pdf")
    if not generate_pdf_from_docx(base_path, merged_pdf):
        return "PDF generation failed", 500

    @after_this_request
    def cleanup(response):
        try:
            os.remove(base_path)
            os.remove(merged_pdf)
            for f in temp_files:
                try:
                    os.remove(f)
                except:
                    pass
        except:
            pass
        return response
    return send_file(merged_pdf, as_attachment=True)

@app.route('/word_previews/<filename>')
def serve_preview_html(filename):
    """Serve HTML preview files from word_previews directory"""
    preview_path = os.path.join(PREVIEW_DIR, filename)
    if not os.path.exists(preview_path) or not filename.endswith('.html'):
        return "Preview not found", 404
    return send_file(preview_path)

@app.route('/preview/<filename>')
def preview_document(filename):
    data = load_data()
    data_for_preview = convert_data_to_gujarati(data)
    section = data.get('offence_section', 'GENERAL')
    template_path = get_template_path(filename, section)
    if not os.path.exists(template_path):
        return "Template not found", 404
    preview_html = extract_preview_from_docx(template_path, data_for_preview)
    return preview_html

@app.route('/preview_merged')
def preview_merged():
    data = load_data()
    data_for_preview = convert_data_to_gujarati(data)
    section = data.get('offence_section', 'GENERAL')
    doc_list = OFFENCE_MAPPING.get(section, OFFENCE_MAPPING['GENERAL'])

    merged_html = ""
    for filename in doc_list:
        template_path = get_template_path(filename, section)
        if os.path.exists(template_path):
            merged_html += f"<div style='page-break-after: always; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 3px solid #ccc;'>"
            merged_html += f"<h4 style='color: #003366; margin-bottom: 15px;'>📄 {filename}</h4>"
            merged_html += extract_preview_from_docx(template_path, data_for_preview)
            merged_html += "</div>"

    if not merged_html:
        return "<p class='text-muted'>No documents to preview</p>"

    return merged_html

if __name__ == '__main__':
    app.run(debug=True, port=5000)