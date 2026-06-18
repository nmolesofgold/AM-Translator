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
                        
                        for para in doc.paragraphs:
                            if para.text.strip():
                                try:
                                    para.text = translator.translate(para.text)
                                    time.sleep(0.05)
                                except: pass
                                
                        for table in doc.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    if cell.text.strip():
                                        try:
                                            cell.text = translator.translate(cell.text)
                                            time.sleep(0.05)
                                        except: pass

                        output_buffer = io.BytesIO()
                        doc.save(output_buffer)
                        st.session_state['processed_output_data'] = output_buffer.getvalue()
                        st.session_state['processed_output_mime'] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                    # --- EXCEL PROCESSING ---
                    elif file_ext == "xlsx":
                        status_text.text("Processing Excel Spreadsheet...")
                        wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.read()))
                        total_sheets = len(wb.worksheets)
                        
                        for idx, sheet in enumerate(wb.worksheets):
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(cell.value, str) and cell.value.strip():
                                        try:
                                            cell.value = translator.translate(cell.value)
                                            time.sleep(0.05)
                                        except: pass
                            progress_bar.progress((idx + 1) / total_sheets)

                        output_buffer = io.BytesIO()
                        wb.save(output_buffer)
                        st.session_state['processed_output_data'] = output_buffer.getvalue()
                        st.session_state['processed_output_mime'] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                    status_text.empty()
                    progress_bar.empty()
                    
                    st.session_state['processed_output_name'] = f"translated_{uploaded_file.name}"
                    st.balloons()
                    st.success("✅ Processing Complete!")

                except Exception as e:
                    st.error(f"❌ Critical Error: {str(e)}")
                    st.info("💡 Tip: Ensure the file is not password protected.")
        else:
            st.info("👆 Upload a file to begin")
            
    with col2:
        st.subheader("Download Result")
        if st.session_state['processed_output_data']:
            st.markdown('<div class="success-box">🎉 File ready for download!</div>', unsafe_allow_html=True)
            st.download_button(
                label="📥 Download Translated File",
                data=st.session_state['processed_output_data'],
                file_name=st.session_state['processed_output_name'],
                mime=st.session_state['processed_output_mime'],
                use_container_width=True
            )
        else:
            st.info("Translated file will appear here.")

elif mode == "📝 Plain Text":
    with col1:
        st.subheader("Enter Text")
        input_text = st.text_area("Paste text here", height=250, placeholder="Type or paste content...", label_visibility="collapsed")

        if st.button("🔄 Translate Text", type="primary"):
            if not input_text.strip():
                st.warning("Please enter text first.")
            else:
                try:
                    with st.spinner("Translating..."):
                        st.session_state['text_output_state'] = translator.translate(input_text)
                    st.success("Done!")
                except Exception as e:
                    st.error(f"Error: {str(e)}")

    with col2:
        st.subheader("Result")
        if st.session_state['text_output_state']:
            st.text_area("Output", value=st.session_state['text_output_state'], height=250, label_visibility="collapsed")
            st.download_button(
                label="📥 Download .txt",
                data=st.session_state['text_output_state'],
                file_name="translated_text.txt",
                mime="text/plain",
                use_container_width=True
            )
        else:
            st.info("Translation appears here.")

# Footer
st.markdown("---")
st.markdown("<div style='text-align: center; color: #888; font-size: 0.9rem;'>Powered by Streamlit, PyMuPDF, python-docx & openpyxl</div>", unsafe_allow_html=True)
