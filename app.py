import os
import ssl
import urllib3

# Disable SSL verification for corporate proxy/firewall
os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["SSL_CERT_FILE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import io
import json
import streamlit as st
import pandas as pd
from docx import Document
from pptx import Presentation
from pptx.util import Pt
from PIL import Image
from google import genai
from google.genai import types

# Try importing pypdf for PDF reading
try:
    import pypdf
except ImportError:
    pypdf = None

# ---------------------------------------------------------
# Page Configuration & Visual Theme
# ---------------------------------------------------------
st.set_page_config(
    page_title="ReqAssist - BA/PM Requirement Engine",
    page_icon="🚀",
    layout="wide"
)

# Custom CSS: Light theme matching file uploader buttons, uniform height (38px), non-bold text
st.markdown("""
<style>
    /* Safe top padding so content is never hidden under the Streamlit top header */
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 2.5rem !important;
    }
    
    /* Universal Button Styling: Light theme, matches uploader button style & compact height */
    div.stButton > button {
        width: 100% !important;
        height: 38px !important;
        min-height: 38px !important;
        max-height: 38px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        border-radius: 6px !important;
        font-size: 15px !important;
        font-weight: 500 !important;
        padding: 0 12px !important;
        margin: 0 !important;
        line-height: 1 !important;
        white-space: nowrap !important;
        text-align: center !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05) !important;
        transition: all 0.2s ease-in-out !important;
    }

    /* Inner text & paragraphs: strictly non-bold and zero vertical margins */
    div.stButton > button p,
    div.stButton > button div,
    div.stButton > button span {
        font-size: 15px !important;
        font-weight: 500 !important;
        margin: 0 !important;
        padding: 0 !important;
        line-height: 1 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        color: inherit !important;
    }

    /* Hover State for Inactive Buttons */
    div.stButton > button:hover {
        background-color: #f1f5f9 !important;
        border-color: #94a3b8 !important;
        color: #0f172a !important;
    }

    /* Active / Selected Button State: Distinct Royal Blue Highlight */
    div.stButton > button[kind="primary"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: 1.5px solid #1d4ed8 !important;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.35) !important;
    }

    /* Language Buttons with Compact Flag Icons */
    div.st-key-btn_lang_en button::before {
        content: "" !important;
        display: inline-block !important;
        width: 18px !important;
        height: 12px !important;
        margin-right: 6px !important;
        background-image: url('https://flagcdn.com/w40/gb.png') !important;
        background-size: cover !important;
        background-position: center !important;
        border-radius: 2px !important;
    }

    div.st-key-btn_lang_it button::before {
        content: "" !important;
        display: inline-block !important;
        width: 18px !important;
        height: 12px !important;
        margin-right: 6px !important;
        background-image: url('https://flagcdn.com/w40/it.png') !important;
        background-size: cover !important;
        background-position: center !important;
        border-radius: 2px !important;
    }

    /* Generate Action Button Styling */
    div.st-key-main_generate_btn button {
        height: 42px !important;
        min-height: 42px !important;
        font-size: 16px !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# UI Dictionary (English & Italiano)
# ---------------------------------------------------------
T = {
    "English": {
        "sidebar_title": "🚀 ReqAssist Settings",
        "lang_label": "Language / Lingua:",
        "name_label": "Your Name:",
        "role_label": "Select Your Role:",
        "roles": ["BA / PM / PO (Full Authoring)", "Developer / Tester / Viewer"],
        "api_key_label": "Gemini API Key:",
        "api_key_warn": "⚠️ Enter your Gemini API Key in the left sidebar to activate ReqAssist.",
        "welcome": "Let's turn your raw requirements into high-impact deliverables.",
        "step1_title": "📂 STEP 1: Upload Source Documentation & Assets",
        "raw_title": "📄 1. Functional Notes RAW",
        "excel_title": "📊 2. Field Validations",
        "figma_title": "🎨 3. Figma UI Screens (.png, .jpg)",
        "supported_formats_help": "Supported: .docx, .txt, .pdf, .xlsx, .csv, .xls",
        "step2_title": "⚡ STEP 2: Select & Generate Deliverable",
        "generate_btn": "✨ Generate Deliverable",
        "tab_preview": "📄 Preview Output",
        "tab_download": "📥 Download Formatted Files",
        "docx_btn": "📄 Download as Word Document (.docx)",
        "xlsx_btn": "📊 Download Test Cases as Excel (.xlsx)",
        "pptx_btn": "📊 Download PowerPoint Deck (.pptx)",
        "viewer_err": "🚫 Access Restricted: Options 1 to 4 are restricted to BA & PM roles. Please contact your project PM or PO.",
        "popup_title": "⚠️ Missing Required Documents",
        "popup_msg": "ReqAssist strictly requires the following document(s) before generating this artifact. Please upload them in Step 1:",
        "popup_btn": "OK, Got It",
        "options": [
            "1️⃣ 📝 Acceptance Criteria (AC)",
            "2️⃣ 🧪 Test Cases",
            "3️⃣ 📑 Detailed Functional Analysis",
            "4️⃣ 🎥 Demo Video",
            "5️⃣ 📊 Presentation / PPT (5 to 10 slides)",
            "6️⃣ ❓ FAQs",
            "7️⃣ 🧠 Quiz"
        ]
    },
    "Italiano": {
        "sidebar_title": "🚀 Impostazioni ReqAssist",
        "lang_label": "Language / Lingua:",
        "name_label": "Il Tuo Nome:",
        "role_label": "Seleziona il Tuo Ruolo:",
        "roles": ["BA / PM / PO (Autore Completo)", "Sviluppatore / Tester / Viewer"],
        "api_key_label": "Chiave API Gemini:",
        "api_key_warn": "⚠️ Inserisci la tua chiave API Gemini nella barra laterale per attivare ReqAssist.",
        "welcome": "Trasformiamo i tuoi requisiti in deliverable di alto impatto.",
        "step1_title": "📂 PASSAGGIO 1: Carica la Documentazione Sorgente & Asset",
        "raw_title": "📄 1. Analisi Funzionale RAW",
        "excel_title": "📊 2. Validazioni Campi",
        "figma_title": "🎨 3. Schermate UI Figma (.png, .jpg)",
        "supported_formats_help": "Formati supportati: .docx, .txt, .pdf, .xlsx, .csv, .xls",
        "step2_title": "⚡ PASSAGGIO 2: Seleziona & Genera Deliverable",
        "generate_btn": "✨ Genera Deliverable",
        "tab_preview": "📄 Anteprima Output",
        "tab_download": "📥 Scarica File Formattati",
        "docx_btn": "📄 Scarica Documento Word (.docx)",
        "xlsx_btn": "📊 Scarica Test Case in Excel (.xlsx)",
        "pptx_btn": "📊 Scarica Presentazione PowerPoint (.pptx)",
        "viewer_err": "🚫 Accesso Limitato: Le opzioni da 1 a 4 sono riservate a BA e PM. Contatta il PM o PO del progetto.",
        "popup_title": "⚠️ Documenti Obbligatori Mancanti",
        "popup_msg": "ReqAssist richiede obbligatoriamente i seguenti documenti prima di procedere. Caricali nel Passaggio 1:",
        "popup_btn": "OK, Ho Capito",
        "options": [
            "1️⃣ 📝 Acceptance Criteria (AC)",
            "2️⃣ 🧪 Test Cases (Casi di Test)",
            "3️⃣ 📑 Analisi Funzionale Dettagliata",
            "4️⃣ 🎥 Video Demo",
            "5️⃣ 📊 Presentazione / PPT (5-10 slide)",
            "6️⃣ ❓ FAQ (Domande Frequenti)",
            "7️⃣ 🧠 Quiz"
        ]
    }
}

# ---------------------------------------------------------
# Sidebar Setup: Language & Credentials (State Persistence)
# ---------------------------------------------------------
if "selected_lang" not in st.session_state:
    st.session_state.selected_lang = "English"

if "saved_api_key" not in st.session_state:
    st.session_state.saved_api_key = ""

st.sidebar.markdown("**Language / Lingua:**")

col_en, col_it = st.sidebar.columns(2)
with col_en:
    uk_selected = (st.session_state.selected_lang == "English")
    if st.button("English", key="btn_lang_en", type="primary" if uk_selected else "secondary", use_container_width=True):
        st.session_state.selected_lang = "English"
        st.rerun()

with col_it:
    it_selected = (st.session_state.selected_lang == "Italiano")
    if st.button("Italiano", key="btn_lang_it", type="primary" if it_selected else "secondary", use_container_width=True):
        st.session_state.selected_lang = "Italiano"
        st.rerun()

lang_key = st.session_state.selected_lang
ui = T[lang_key]

st.sidebar.title(ui["sidebar_title"])
user_name = st.sidebar.text_input(ui["name_label"], value="Mayank", key="user_name_input")
user_role = st.sidebar.selectbox(ui["role_label"], ui["roles"], key="user_role_select")

# Check Streamlit Secrets first, then session_state, else blank
secret_key = st.secrets.get("GEMINI_API_KEY", "")
default_key = secret_key if secret_key else st.session_state.get("saved_api_key", "")

api_key_input = st.sidebar.text_input(
    ui["api_key_label"], 
    type="password", 
    value=default_key,
    key="gemini_api_key_input",
    help="Pre-configured via Cloud Secrets or enter manually at aistudio.google.com"
)

if api_key_input:
    st.session_state.saved_api_key = api_key_input

api_key = st.session_state.saved_api_key

# ---------------------------------------------------------
# Header & Welcome
# ---------------------------------------------------------
st.title("🚀 ReqAssist")
st.caption(f"**{'Benvenuto' if lang_key == 'Italiano' else 'Welcome'}, {user_name}!** {ui['welcome']}")

# Check API Key after rendering header
if not api_key:
    st.sidebar.warning(ui["api_key_warn"])
    st.info(f"👈 **{ui['api_key_warn']}**")
    st.stop()


# Initialize Gemini Client with SSL verification bypass
client = genai.Client(
    api_key=api_key,
    http_options=types.HttpOptions(
        client_args={"verify": False}
    )
)

# ---------------------------------------------------------
# Modal Alert Popup Component
# ---------------------------------------------------------
@st.dialog(ui["popup_title"])
def show_missing_data_popup(missing_list):
    st.error(ui["popup_msg"])
    for item in missing_list:
        st.markdown(f"❌ **{item}**")
    st.write("")
    if st.button(ui["popup_btn"], type="primary", use_container_width=True):
        st.rerun()

# ---------------------------------------------------------
# Universal Multi-Format File Reader (docx, pdf, xlsx, xls, csv, txt, md)
# ---------------------------------------------------------
def parse_uploaded_file(uploaded_file) -> str:
    """Extracts clean text or tabular representations from docx, pdf, xlsx, xls, csv, txt, md."""
    if not uploaded_file:
        return ""
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".docx"):
            doc = Document(uploaded_file)
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        
        elif name.endswith(".pdf"):
            if pypdf is None:
                return "PDF parser (pypdf) is not installed. Please run: pip install pypdf"
            reader = pypdf.PdfReader(uploaded_file)
            pages_text = []
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    pages_text.append(extracted)
            return "\n".join(pages_text)
        
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            df = pd.read_excel(uploaded_file)
            return df.to_string(index=False)
            
        elif name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            return df.to_string(index=False)
            
        else: # txt, md, etc.
            return uploaded_file.read().decode("utf-8", errors="ignore")
            
    except Exception as e:
        return f"Error reading {uploaded_file.name}: {str(e)}"

# ---------------------------------------------------------
# Helper Functions: Exporters
# ---------------------------------------------------------
def create_docx(title: str, content: str) -> io.BytesIO:
    doc = Document()
    doc.add_heading(title, 0)
    for line in content.split("\n"):
        if line.startswith("## "):
            doc.add_heading(line.replace("## ", ""), level=1)
        elif line.startswith("### "):
            doc.add_heading(line.replace("### ", ""), level=2)
        elif line.startswith("* ") or line.startswith("- "):
            doc.add_paragraph(line[2:], style="List Bullet")
        elif line.strip():
            doc.add_paragraph(line)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def create_pptx(slides_json: list) -> io.BytesIO:
    prs = Presentation()
    for slide_data in slides_json:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_data.get("title", "Slide")
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame
        tf.clear()
        
        for bullet in slide_data.get("bullets", []):
            p = tf.add_paragraph()
            p.text = bullet
            p.font.size = Pt(18)
            
        if "notes" in slide_data and slide_data["notes"]:
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = slide_data["notes"]

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

def parse_markdown_table_to_excel(markdown_text: str) -> io.BytesIO:
    lines = [line.strip() for line in markdown_text.split("\n") if line.strip().startswith("|")]
    if len(lines) >= 3:
        headers = [c.strip() for c in lines[0].split("|")[1:-1]]
        data = []
        for line in lines[2:]:
            row = [c.strip() for c in line.split("|")[1:-1]]
            if len(row) == len(headers):
                data.append(row)
        if data:
            df = pd.DataFrame(data, columns=headers)
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Test Cases")
            bio.seek(0)
            return bio
    return None

# ---------------------------------------------------------
# STEP 1: Upload Source Materials (Accepts docx, txt, xlsx, csv, xls, pdf)
# ---------------------------------------------------------
ALLOWED_EXTENSIONS = ["docx", "txt", "md", "xlsx", "xls", "csv", "pdf"]

with st.container():
    st.markdown(f"#### {ui['step1_title']}")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"**{ui['raw_title']}**")
        raw_file = st.file_uploader(
            "Upload FA RAW", 
            type=ALLOWED_EXTENSIONS, 
            key="raw_file_uploader", 
            label_visibility="collapsed",
            help=ui["supported_formats_help"]
        )
        raw_text = parse_uploaded_file(raw_file) if raw_file else ""
        if raw_file:
            st.caption(f"✅ Loaded: `{raw_file.name}`")

    with col2:
        st.markdown(f"**{ui['excel_title']}**")
        excel_file = st.file_uploader(
            "Upload Validations", 
            type=ALLOWED_EXTENSIONS, 
            key="excel_file_uploader", 
            label_visibility="collapsed",
            help=ui["supported_formats_help"]
        )
        excel_summary = parse_uploaded_file(excel_file) if excel_file else ""
        if excel_file:
            st.caption(f"✅ Loaded: `{excel_file.name}`")

    with col3:
        st.markdown(f"**{ui['figma_title']}**")
        figma_images = st.file_uploader(
            "Upload Screens", 
            type=["png", "jpg", "jpeg"], 
            accept_multiple_files=True, 
            key="figma_uploader", 
            label_visibility="collapsed"
        )
        if figma_images:
            st.caption(f"✅ Loaded `{len(figma_images)}` screen(s).")

st.markdown("---")

# ---------------------------------------------------------
# STEP 2: Deliverable Selection (Compact Light Theme Buttons)
# ---------------------------------------------------------
st.markdown(f"### {ui['step2_title']}")

if "selected_opt_idx" not in st.session_state:
    st.session_state.selected_opt_idx = 0

btn_col1, btn_col2 = st.columns(2)
for idx, opt_name in enumerate(ui["options"]):
    target_col = btn_col1 if idx < 4 else btn_col2
    with target_col:
        is_active = (st.session_state.selected_opt_idx == idx)
        btn_type = "primary" if is_active else "secondary"
        if st.button(opt_name, key=f"btn_opt_{idx}", type=btn_type):
            st.session_state.selected_opt_idx = idx
            st.rerun()

current_idx = st.session_state.selected_opt_idx
current_option = ui["options"][current_idx]

# ---------------------------------------------------------
# Role Restriction Guardrail
# ---------------------------------------------------------
is_viewer = user_role == ui["roles"][1]
if is_viewer and current_idx < 4:
    st.error(ui["viewer_err"])
    st.stop()

# ---------------------------------------------------------
# Updated Document Matrix Checking
# ---------------------------------------------------------
# 0 (AC): Notes RAW + Field Validation
# 1 (Test Cases): Field Validation
# 2 (Detailed FA): Notes RAW + Field Validation + FIGMA
# 3 (Demo Video): Notes RAW + Field Validation + FIGMA
# 4 (Presentation/PPT): Notes RAW + Field Validation + FIGMA
# 5 (FAQs): Notes RAW + Field Validation
# 6 (Quiz): Notes RAW + Field Validation

needs_raw = current_idx in [0, 2, 3, 4, 5, 6]
needs_excel = current_idx in [0, 1, 2, 3, 4, 5, 6]
needs_figma = current_idx in [2, 3, 4]

missing_core_items = []
if needs_raw and not raw_text:
    missing_core_items.append(ui["raw_title"])
if needs_excel and not excel_summary:
    missing_core_items.append(ui["excel_title"])
if needs_figma and not figma_images:
    missing_core_items.append(ui["figma_title"])

# ---------------------------------------------------------
# Generation Trigger & AI Execution
# ---------------------------------------------------------
st.markdown("<br>", unsafe_allow_html=True)
if st.button(f"{ui['generate_btn']} : {current_option}", key="main_generate_btn", type="primary", use_container_width=True):
    # Trigger modal popup if any required document is missing
    if missing_core_items:
        show_missing_data_popup(missing_core_items)
        st.stop()

    with st.spinner(f"ReqAssist is generating ({lang_key})..."):
        
        system_prompt = f"""
[ROLE & PERSONA]
You are "ReqAssist," an energetic Principal Business Analyst and Product Requirements Architect. Address the user as {user_name}.

[CRUCIAL LANGUAGE RULE]
You MUST respond ENTIRELY and STRICTLY in {lang_key}. All headings, titles, descriptions, table columns, test steps, bullet points, speaker notes, and quiz questions MUST be written in {lang_key}.

[ANTI-HALLUCINATION & STRICT GROUNDING RULES]
1. ZERO INVENTED BUSINESS LOGIC: Do not invent rules not found in inputs.
2. STRICT SOURCE GROUNDING: Ground all answers strictly on the uploaded Functional Notes, Figma UI screens, and Validation sheet.
3. EXPLICIT ASSUMPTION FLAGGING: If an assumption is forced, tag it clearly as '[ASSUMPTION: ...]' (or '[ASSUNZIONE: ...]' if Italian).
4. UNKNOWN DATA: If detail is missing, state: "{'⚠️ Questo dettaglio non è specificato nei documenti forniti. Si prega di consultare il PM/PO.' if lang_key == 'Italiano' else '⚠️ This detail is not specified in the provided documents. Please consult your PM/PO.'}"

[OUTPUT TEMPLATES]
- Acceptance Criteria: BDD format (GIVEN/WHEN/THEN or DATO/QUANDO/ALLORA) grouped by User Story.
- Test Cases: Markdown Table with columns: | Test Case ID | Scenario | Pre-conditions | Test Steps | Expected Result | Field Validation Check | Functional Test case or Technical Test case |
- Detailed Functional Analysis: MS Word structure with H2 and H3 headings.
- Demo Video: Markdown Table with columns: | Timestamp | UI Screen / Area | Visual Action | Voiceover Script |
- PPT Slides: 5 to 10 slides maximum (Title, Visual Reference, Bullet Points, Speaker Notes).
- FAQs: Split into Developer Technical FAQs and Stakeholder Business FAQs.
- Quiz: 7 to 12 Multiple-Choice Questions (10 ideal) with complete Answer Key at the end.
"""

        user_content = f"""
DELIVERABLE REQUESTED: {current_option}
LANGUAGE: {lang_key}

SOURCE MATERIALS:
- FUNCTIONAL NOTES RAW:
{raw_text if raw_text else 'None required or provided'}

- FIELD VALIDATIONS:
{excel_summary if excel_summary else 'None provided'}

Generate the complete artifact strictly adhering to the requested templates and grounding rules.
"""

        contents = [user_content]
        if figma_images and needs_figma:
            for img in figma_images:
                contents.append(Image.open(img))

        try:
            response = client.models.generate_content(
                model="gemini-3.7-flash",
                contents=contents,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=0.2
                )
            )
            
            output_text = response.text
            st.success("🎉 Generation Complete!")
            
            tab_prev, tab_down = st.tabs([ui["tab_preview"], ui["tab_download"]])
            
            with tab_prev:
                st.markdown(output_text)
                
            with tab_down:
                docx_bio = create_docx(current_option, output_text)
                st.download_button(
                    label=ui["docx_btn"],
                    data=docx_bio,
                    file_name=f"ReqAssist_{user_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Excel Export (.xlsx)
                if current_idx == 1:
                    excel_bio = parse_markdown_table_to_excel(output_text)
                    if excel_bio:
                        st.download_button(
                            label=ui["xlsx_btn"],
                            data=excel_bio,
                            file_name=f"ReqAssist_TestCases_{user_name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                # PPT Export (.pptx)
                if current_idx == 4:
                    ppt_parser = client.models.generate_content(
                        model="gemini-3.7-flash",
                        contents=[f"Convert this presentation into a JSON array of slide objects with keys 'title', 'bullets' (list of strings), 'notes' (string). Keep text in {lang_key}:\n\n{output_text}"],
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    try:
                        slides_data = json.loads(ppt_parser.text)
                        pptx_bio = create_pptx(slides_data)
                        st.download_button(
                            label=ui["pptx_btn"],
                            data=pptx_bio,
                            file_name=f"ReqAssist_Presentation_{user_name}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    except Exception:
                        st.info("You can copy the slide Markdown directly from the Preview tab.")

        except Exception as e:
            st.error(f"Error: {str(e)}")
