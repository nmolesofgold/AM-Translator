import streamlit as st
import fitz  # PyMuPDF
from deep_translator import GoogleTranslator
from docx import Document
import openpyxl
import io
import time

# --- Page Configuration ---
st.set_page_config(
    page_title="Universal Layout Translator",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom CSS ---
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: 700; color: #4F8BF9; text-align: center; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1.1rem; color: #666; text-align: center; margin-bottom: 2rem; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; font-weight: 600; }
    .success-box { padding: 1rem; background-color: #d4edda; border-radius: 5px; margin-top: 1rem; }
</style>
""", unsafe_allow_html=True)

# --- Session State Management ---
if 'processed_output_data' not in st.session_state:
    st.session_state['processed_output_data'] = None
    st.session_state['processed_output_name'] = ""
    st.session_state['processed_output_mime'] = ""
if 'text_output_state' not in st.session_state:
    st.session_state['text_output_state'] = ""

# --- Header ---
st.markdown('<div class="main-header">🌐 Smart Multi-Format Translator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Translate PDFs (Layout Preserved), Word Docs, Excel Sheets, and Text.</div>', unsafe_allow_html=True)

# --- Language Mapping ---
LANGUAGES = {
    "Auto Detect": "auto", "German": "de", "English": "en", "Spanish": "es",
    "French": "fr", "Italian": "it", "Portuguese": "pt", "Chinese (Simplified)": "zh-CN",
    "Japanese": "ja", "Korean": "ko", "Russian": "ru", "Arabic": "ar", "Hindi": "hi",
    "Dutch": "nl", "Polish": "pl", "Turkish": "tr", "Swedish": "sv"
}

# --- Sidebar ---
with st.sidebar:
    st.header("⚙️ Configuration")
    
    source_language = st.selectbox("Source Language", options=list(LANGUAGES.keys()), index=0)
    target_language = st.selectbox("Target Language", options=list(LANGUAGES.keys()), index=2) # Default English
    
    src_code = LANGUAGES[source_language]
    tgt_code = LANGUAGES[target_language]
    
    st.divider()
    st.info("💡 **PDFs:** Layout is preserved by redrawing text blocks.\n\n**Word/Excel:** Content is translated while keeping styles intact.")

# --- Main Logic ---
mode = st.radio("Select Input Mode", ["📁 Upload Files (.pdf, .docx, .xlsx)", "📝 Plain Text"], horizontal=True)

col1, col2 = st.columns([1, 1])

# Initialize Translator once per run
translator = GoogleTranslator(source=src_code, target=tgt_code)

if mode == "📁 Upload Files (.pdf, .docx, .xlsx)":
    with col1:
        st.subheader("Upload Document")
        uploaded_file = st.file_uploader("Drop PDF, Word, or Excel file", type=["pdf", "docx", "xlsx"], label_visibility="collapsed")

        if uploaded_file:
            file_ext = uploaded_file.name.split(".")[-1].lower()
            st.success(f"Loaded: **{uploaded_file.name}**")

            if st.button("🚀 Process & Translate", type="primary"):
                try:
                    st.session_state['processed_output_data'] = None
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    # --- PDF PROCESSING ---
                    if file_ext == "pdf":
                        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                        total_pages = len(doc)
                        
                        for page_num in range(total_pages):
                            status_text.text(f"Processing PDF page {page_num + 1}/{total_pages}...")
                            page = doc[page_num]
                            
                            # Enhanced extraction flags
                            blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)["blocks"]

                            for block in blocks:
                                if "lines" not in block: continue
                                
                                # Extract text
                                block_text = " ".join(
                                    span["text"].strip() 
                                    for line in block["lines"] 
                                    for span in line["spans"] 
                                    if span["text"].strip()
                                )
                                if not block_text: continue

                                bbox = fitz.Rect(block["bbox"])

                                try:
                                    translated_text = translator.translate(block_text)
                                    time.sleep(0.1) # Rate limiting

                                    # Dynamic Font Sizing
                                    font_size = 10
                                    if len(translated_text) > len(block_text) * 1.3:
                                        font_size = 8
                                    elif len(translated_text) > len(block_text) * 1.5:
                                        font_size = 7

                                    # Redraw
                                    page.draw_rect(bbox, color=(1, 1, 1), fill=(1, 1, 1))
                                    page.insert_textbox(bbox, translated_text, fontsize=font_size, color=(0, 0, 0))
                                except Exception:
                                    continue # Skip failed blocks
                            
                            progress_bar.progress((page_num + 1) / total_pages)

                        output_buffer = io.BytesIO()
                        doc.save(output_buffer)
                        doc.close()
                        st.session_state['processed_output_data'] = output_buffer.getvalue()
                        st.session_state['processed_output_mime'] = "application/pdf"

                    # --- WORD PROCESSING ---
                    elif file_ext == "docx":
                        status_text.text("Processing Word Document...")
                        doc = Document(io.BytesIO(uploaded_file.read()))
