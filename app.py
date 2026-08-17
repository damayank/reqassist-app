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

# Custom CSS: Sidebar nowrap fix, standard buttons, and prominent #10347d CTA button
st.markdown("""
<style>
    /* 1. Safe top & bottom padding */
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 4rem !important;
    }
    
    /* 2. Prevent Sidebar title from wrapping in Italian */
    section[data-testid="stSidebar"] h1 {
        font-size: 1.25rem !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        line-height: 1.3 !important;
        padding-bottom: 0.4rem !important;
    }

    /* 3. Universal Secondary Button Styling: Light theme & standard compact height (38px) */
    div.stButton > button {
        width: 100% !important;
        height: 38px !important;
        min-height: 38px !important;
        max-height: 38px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        border-radius: 6px !important;
        font-size: 14px !important;
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

    /* Inner text for generic buttons */
    div.stButton > button p,
    div.stButton > button div,
    div.stButton > button span {
        font-size: 14px !important;
        font-weight: 500 !important;
        margin: 0 !important;
        padding: 0 !important;
        line-height: 1 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        color: inherit !important;
    }

    /* Hover State for Inactive Selection Buttons */
    div.stButton > button:hover {
        background-color: #f1f5f9 !important;
        border-color: #94a3b8 !important;
        color: #0f172a !important;
    }

    /* Active / Selected Option Buttons */
    div.stButton > button[kind="primary"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: 1.5px solid #1d4ed8 !important;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.30) !important;
    }

    /* Flag icons on language selector buttons */
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

    /* ---------------------------------------------------------
       4. Primary Action Button: #10347d (Dark Blue) -> #f56642 (Hover/Click)
          +15% larger size (56px height, 18px text, 34px padding)
       --------------------------------------------------------- */
    div.st-key-main_generate_btn button {
        background-color: #10347d !important;   /* Target Dark Blue */
        color: #ffffff !important;
        border: 1px solid #10347d !important;
        font-size: 18px !important;            /* +15% font enlargement */
        font-weight: 700 !important;
        height: 56px !important;               /* +15% height enlargement */
        min-height: 56px !important;
        max-height: 56px !important;
        padding: 0 34px !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 16px rgba(16, 52, 125, 0.35) !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }

    div.st-key-main_generate_btn button p,
    div.st-key-main_generate_btn button div,
    div.st-key-main_generate_btn button span {
        color: #ffffff !important;
        font-weight: 700 !important;
        font-size: 18px !important;
    }

    /* Hover & Active / Clicked State */
    div.st-key-main_generate_btn button:hover,
    div.st-key-main_generate_btn button:active,
    div.st-key-main_generate_btn button:focus:active {
        background-color: #f56642 !important;   /* Hover / Click: #f56642 */
        border-color: #f56642 !important;
        color: #ffffff !important;
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(245, 102, 66, 0.45) !important;
    }

    div.st-key-main_generate_btn button:hover p,
    div.st-key-main_generate_btn button:hover div,
    div.st-key-main_generate_btn button:hover span,
    div.st-key-main_generate_btn button:active p,
    div.st-key-main_generate_btn button:active div,
    div.st-key-main_generate_btn button:active span {
        color: #ffffff !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# UI Dictionary (English & Italiano Only - No Number Icons)
# ---------------------------------------------------------
T = {
    "English": {
        "sidebar_title": "🚀 ReqAssist Settings",
        "lang_label": "Language / Lingua:",
        "name_label": "Your Name:",
        "role_label": "Select Your Role:",
        "roles": ["BA / PM / PO (Full Authoring)", "Developer / Tester / Viewer"],
        "api_key_warn": "⚠️ Gemini API Key missing. Please configure GEMINI_API_KEY in Streamlit Cloud Secrets.",
        "welcome": "Let's turn your raw requirements into high-impact deliverables.",
        "step1_title": "📂 STEP 1: Upload Source Documentation & Assets",
        "raw_title": "📄 Functional Notes RAW",
        "excel_title": "📊 Field Validations",
        "figma_title": "🎨 Figma UI Screens (.png, .jpg)",
        "supported_formats_help": "Supported: .docx, .txt, .pdf, .xlsx, .csv, .xls",
        "step2_title": "⚡ STEP 2: Select & Generate Deliverable",
        "generate_btn": "✨ Generate Deliverable",
        "tab_preview": "📄 Preview Output",
        "tab_download": "📥 Download Formatted Files",
        "docx_btn": "📄 Download Word (.docx)",
        "xlsx_btn": "📊 Download Test Cases (.xlsx)",
        "pptx_btn": "📊 Download Presentation (.pptx)",
        "md_btn": "📝 Download Markdown (.md)",
        "viewer_err": "🚫 Access Restricted: These options are restricted to BA & PM roles. Please contact your project PM or PO.",
        "popup_title": "⚠️ Missing Required Documents",
        "popup_msg": "ReqAssist strictly requires the following document(s) before generating this artifact. Please upload them in Step 1:",
        "popup_btn": "OK, Got It",
        "options": [
            "📝 Acceptance Criteria (AC)",
            "🧪 Test Cases",
            "📑 Detailed Functional Analysis",
            "🎥 Demo Video",
            "📊 Presentation / PPT (5 to 10 slides)",
            "❓ FAQs",
            "🧠 Quiz"
        ]
    },
    "Italiano": {
        "sidebar_title": "🚀 Impostazioni ReqAssist",
        "lang_label": "Language / Lingua:",
        "name_label": "Il Tuo Nome:",
        "role_label": "Seleziona il Tuo Ruolo:",
        "roles": ["BA / PM / PO (Autore Completo)", "Sviluppatore / Tester / Viewer"],
        "api_key_warn": "⚠️ Chiave API Gemini mancante. Configura GEMINI_API_KEY nei Secrets di Streamlit Cloud.",
        "welcome": "Trasformiamo i tuoi requisiti in deliverable di alto impatto.",
        "step1_title": "📂 PASSAGGIO 1: Carica la Documentazione Sorgente & Asset",
        "raw_title": "📄 Analisi Funzionale RAW",
        "excel_title": "📊 Validazioni Campi",
        "figma_title": "🎨 Schermate UI Figma (.png, .jpg)",
        "supported_formats_help": "Formati supportati: .docx, .txt, .pdf, .xlsx, .csv, .xls",
        "step2_title": "⚡ PASSAGGIO 2: Seleziona & Genera Deliverable",
        "generate_btn": "✨ Genera Deliverable",
        "tab_preview": "📄 Anteprima Output",
        "tab_download": "📥 Scarica File Formattati",
        "docx_btn": "📄 Scarica Documento Word (.docx)",
        "xlsx_btn": "📊 Scarica Test Case in Excel (.xlsx)",
        "pptx_btn": "📊 Scarica Presentazione (.pptx)",
        "md_btn": "📝 Scarica Markdown (.md)",
        "viewer_err": "🚫 Accesso Limitato: Queste opzioni sono riservate a BA e PM. Contatta il PM o PO del progetto.",
        "popup_title": "⚠️ Documenti Obbligatori Mancanti",
        "popup_msg": "ReqAssist richiede obbligatoriamente i seguenti documenti prima di procedere. Caricali nel Passaggio 1:",
        "popup_btn": "OK, Ho Capito",
        "options": [
            "📝 Acceptance Criteria (AC)",
            "🧪 Test Cases (Casi di Test)",
            "📑 Analisi Funzionale Dettagliata",
            "🎥 Video Demo",
            "📊 Presentazione / PPT (5-10 slide)",
            "❓ FAQ (Domande Frequenti)",
            "🧠 Quiz"
        ]
    }
}

# ---------------------------------------------------------
# Sidebar Setup: Language & Role Selection
# ---------------------------------------------------------
if "selected_lang" not in st.session_state:
    st.session_state.selected_lang = "English"

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

# 🔒 Silent Secret Loading (Backend only)
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))

# ---------------------------------------------------------
# Header & Welcome
# ---------------------------------------------------------
st.title("🚀 ReqAssist")
st.caption(f"**{'Benvenuto' if lang_key == 'Italiano' else 'Welcome'}, {user_name}!** {ui['welcome']}")

# Check API Key validity silently
if not api_key:
    st.error(ui["api_key_warn"])
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
# Helper Functions: Exporters (DOCX, PPTX, XLSX)
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
# STEP 1: Upload Source Materials
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
# STEP 2: Deliverable Selection (Clean Names without numbers)
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
# Document Matrix Checking
# ---------------------------------------------------------
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
# Centered, Prominent #10347d Action Button (Positioned at Bottom)
# ---------------------------------------------------------
# Generous vertical space to position the button near the bottom of the screen
st.markdown("<div style='margin-top: 4.5rem; margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)

col_b1, col_b2, col_b3 = st.columns([1, 2, 1])
with col_b2:
    generate_clicked = st.button(
        f"{ui['generate_btn']} : {current_option}",
        key="main_generate_btn",
        type="primary",
        use_container_width=True
    )

if generate_clicked:
    # Trigger modal popup if required documents are missing
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
                model="gemini-2.5-flash",
                contents=contents,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=0.2
                )
            )
            
            output_text = response.text
            st.session_state["generated_output"] = output_text
            st.session_state["generated_option"] = current_option
            st.session_state["generated_idx"] = current_idx
            st.success("🎉 Generation Complete!")

        except Exception as e:
            st.error(f"Error: {str(e)}")

# ---------------------------------------------------------
# Output Display & Direct Exports (.docx, .pptx, .md, .xlsx)
# ---------------------------------------------------------
if "generated_output" in st.session_state and st.session_state["generated_output"]:
    output_text = st.session_state["generated_output"]
    active_opt_name = st.session_state.get("generated_option", current_option)
    active_opt_idx = st.session_state.get("generated_idx", current_idx)

    st.markdown("---")
    tab_prev, tab_down = st.tabs([ui["tab_preview"], ui["tab_download"]])
    
    with tab_prev:
        st.markdown(output_text)
        
    with tab_down:
        # Determine clean filename
        base_name = active_opt_name.split(" ", 1)[-1].replace(" ", "_").replace("/", "_")
        
        col_d1, col_d2, col_d3 = st.columns(3)
        
        # 1. Word Document (.docx) Export
        with col_d1:
            docx_bio = create_docx(active_opt_name, output_text)
            st.download_button(
                label=ui["docx_btn"],
                data=docx_bio,
                file_name=f"ReqAssist_{base_name}_{user_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        # 2. PowerPoint (.pptx) Export
        with col_d2:
            ppt_parser = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[f"Convert this deliverable into a structured presentation as a JSON array of slide objects with keys 'title', 'bullets' (list of strings), 'notes' (string). Keep text strictly in {lang_key}:\n\n{output_text}"],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            try:
                slides_data = json.loads(ppt_parser.text)
                pptx_bio = create_pptx(slides_data)
                st.download_button(
                    label=ui["pptx_btn"],
                    data=pptx_bio,
                    file_name=f"ReqAssist_{base_name}_{user_name}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            except Exception:
                st.info("Direct Markdown available below.")

        # 3. Markdown (.md) Export
        with col_d3:
            st.download_button(
                label=ui["md_btn"],
                data=output_text.encode("utf-8"),
                file_name=f"ReqAssist_{base_name}_{user_name}.md",
                mime="text/markdown",
                use_container_width=True
            )
            
        # 4. Optional Excel (.xlsx) Export for Test Cases
        if active_opt_idx == 1:
            excel_bio = parse_markdown_table_to_excel(output_text)
            if excel_bio:
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label=ui["xlsx_btn"],
                    data=excel_bio,
                    file_name=f"ReqAssist_TestCases_{user_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
