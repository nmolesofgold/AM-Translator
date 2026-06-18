import streamlit as st
import fitz  # PyMuPDF
from deep_translator import GoogleTranslator
from docx import Document
import openpyxl
import io
import time
import re

# --- Page Configuration ---
st.set_page_config(
    page_title="PDF & Text Translator",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- Session State Management ---
if 'processed_output_data' not in st.session_state:
    st.session_state['processed_output_data'] = None
    st.session_state['processed_output_name'] = ""
if 'text_input' not in st.session_state:
    st.session_state['text_input'] = ""
if 'translated_text_result' not in st.session_state:
    st.session_state['translated_text_result'] = ""
if 'debug_log' not in st.session_state:
    st.session_state['debug_log'] = []

# --- Header ---
st.title("🌐 Universal Document Translator")

# --- Language Mapping ---
LANGUAGES = {
    "Auto Detect": "auto", "German": "de", "English": "en", "Spanish": "es",
    "French": "fr", "Italian": "it", "Portuguese": "pt", "Chinese (Simplified)": "zh-CN",
    "Japanese": "ja", "Korean": "ko", "Russian": "ru", "Arabic": "ar", "Hindi": "hi",
    "Dutch": "nl", "Polish": "pl", "Turkish": "tr", "Swedish": "sv"
}

# --- Sidebar ---
with st.sidebar:
    st.header("Settings")
    
    source_language = st.selectbox("Source Language", options=list(LANGUAGES.keys()), index=0)
    target_language = st.selectbox("Target Language", options=list(LANGUAGES.keys()), index=2)
    
    src_code = LANGUAGES[source_language]
    tgt_code = LANGUAGES[target_language]
    
    st.divider()
    debug_mode = st.checkbox("Enable Error Logging", value=False)

# --- Helper Functions ---

def sanitize_text(text):
    if not text:
        return text
    replacements = {
        '\u00A0': ' ', '\u200B': '', '\uFEFF': '', '\u200C': '', '\u200D': '',
    }
    sanitized = text
    for old, new in replacements.items():
        sanitized = sanitized.replace(old, new)
    return sanitized

def split_into_sentences(text):
    sentences = re.split(r'(?<=[.!?])\s+|\n', text)
    return [s.strip() for s in sentences if s.strip()]

def translate_robustly(text, translator, block_id=""):
    if not text or not text.strip():
        return text, None
    
    clean_text = sanitize_text(text)
    if not clean_text.strip():
        return "[INFO: Text contained only control characters]", None
    
    if len(clean_text) > 4500:
        sentences = split_into_sentences(clean_text)
        chunks = []
        current_chunk = ""
        translated_chunks = []
        
        for sent in sentences:
            if len(current_chunk) + len(sent) < 4000:
                current_chunk += sent + " "
            else:
                chunks.append(current_chunk)
                current_chunk = sent + " "
        if current_chunk:
            chunks.append(current_chunk)
            
        for chunk in chunks:
            try:
                translated_chunks.append(translator.translate(chunk))
                time.sleep(0.15)
            except Exception:
                translated_chunks.append(f"[ERROR: Chunk too complex]")
        return " ".join(translated_chunks), None

    try:
        result = translator.translate(clean_text)
        return result, None
    except Exception:
        pass
    
    sentences = split_into_sentences(text)
    translated_parts = []
    errors_found = []
    
    for i, sentence in enumerate(sentences):
        clean_sentence = sanitize_text(sentence)
        if not clean_sentence.strip():
            continue
        try:
            translated = translator.translate(clean_sentence)
            translated_parts.append(translated)
            time.sleep(0.15)
        except Exception as se:
            err_str = f"[ERROR: Sentence {i+1} failed]"
            translated_parts.append(err_str)
            errors_found.append(f"Sentence '{sentence[:30]}...': {str(se)[:50]}")
            
    final_text = " ".join(translated_parts)
    if errors_found and len(errors_found) > len(sentences) * 0.5:
        final_text = f"[WARNING: Multiple translation errors] {final_text}"
        
    return final_text, ", ".join(errors_found) if errors_found else None

def instant_translate():
    input_text = st.session_state['text_input']
    if not input_text.strip():
        st.session_state['translated_text_result'] = ""
        return
    
    try:
        translator = GoogleTranslator(source=src_code, target=tgt_code)
        result, _ = translate_robustly(input_text, translator)
        st.session_state['translated_text_result'] = result
    except Exception as e:
        st.session_state['translated_text_result'] = f"Error: {str(e)}"

# --- Main Logic ---
mode = st.radio("Select Mode", ["Upload Files", "Plain Text"], horizontal=True)

if mode == "Upload Files":
    st.subheader("Upload PDF, Word, or Excel File")
    uploaded_file = st.file_uploader("Choose a file", type=["pdf", "docx", "xlsx"], label_visibility="collapsed")

    if uploaded_file:
        st.success(f"Loaded: **{uploaded_file.name}**")
        
        if st.button("Process & Translate", type="primary"):
            try:
                st.session_state['debug_log'] = []
                st.session_state['processed_output_data'] = None
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                original_data = uploaded_file.read()
                file_ext = uploaded_file.name.split(".")[-1].lower()
                
                if file_ext != "pdf":
                    st.error("Only PDF layout translation is supported in this version. Please use Plain Text mode for other formats or convert to PDF first.")
                    st.stop()

                doc = fitz.open(stream=original_data, filetype="pdf")
                total_pages = len(doc)
                translator = GoogleTranslator(source=src_code, target=tgt_code)
                
                for page_num in range(total_pages):
                    status_text.text(f"Processing page {page_num + 1}/{total_pages}...")
                    page = doc[page_num]
                    
                    blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT | fitz.TEXT_ACCURATE_BBOXES)["blocks"]
                    processed_count = 0

                    for idx, block in enumerate(blocks):
                        if "lines" not in block: continue
                        
                        block_text = " ".join(
                            span["text"].strip() 
                            for line in block["lines"] 
                            for span in line["spans"] 
                            if span["text"].strip()
                        )
                        if not block_text: continue

                        bbox = fitz.Rect(block["bbox"])
                        translated_text, error = translate_robustly(block_text, translator, block_id=f"P{page_num}-B{idx}")
                        
                        if error and debug_mode:
                            st.session_state['debug_log'].append(f"Page {page_num+1}, Block {idx}: {error}")

                        # Erase and Redraw with Auto-Fit
                        page.draw_rect(bbox, color=(1, 1, 1), fill=(1, 1, 1))
                        font_size = 11.0
                        while font_size > 4.0:
                            rc = page.insert_textbox(bbox, translated_text, fontsize=font_size, color=(0,0,0), render_mode=3)
                            if rc >= 0:
                                break
                            font_size -= 0.5
                        page.insert_textbox(bbox, translated_text, fontsize=font_size, color=(0, 0, 0))
                        
                        processed_count += 1
                        time.sleep(0.02)
                    
                    progress_bar.progress((page_num + 1) / total_pages)

                output_buffer = io.BytesIO()
                doc.save(output_buffer)
                doc.close()
                
                st.session_state['processed_output_data'] = output_buffer.getvalue()
                st.session_state['processed_output_name'] = f"translated_{uploaded_file.name}"
                
                st.balloons()
                st.success("Translation Complete!")
                
                if debug_mode and st.session_state['debug_log']:
                    with st.expander("View Error Log"):
                        for log in st.session_state['debug_log']:
                            st.code(log)

            except Exception as e:
                st.error(f"Error: {str(e)}")
    
    # Result Section (Only appears after processing)
    if st.session_state['processed_output_data']:
        st.divider()
        st.subheader("Download Result")
        st.download_button(
            label="📥 Download Translated PDF",
            data=st.session_state['processed_output_data'],
            file_name=st.session_state['processed_output_name'],
            mime="application/pdf",
            use_container_width=True
        )

elif mode == "Plain Text":
    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("Input")
        st.text_area(
            "Paste text here", 
            value=st.session_state['text_input'], 
            height=400, 
            placeholder="Type or paste text to translate instantly...", 
            label_visibility="collapsed",
            key="text_input",
            on_change=instant_translate
        )

    with col2:
        st.subheader("Translation")
        if st.session_state['translated_text_result']:
            st.text_area("Output", value=st.session_state['translated_text_result'], height=400, label_visibility="collapsed")
            st.download_button(
                "Download .txt", 
                data=st.session_state['translated_text_result'], 
                file_name="translated.txt", 
                mime="text/plain", 
                use_container_width=True
            )
        else:
            st.info("Translation will appear here automatically.")
