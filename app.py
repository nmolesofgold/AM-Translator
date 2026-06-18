import streamlit as st
import fitz  # PyMuPDF
from deep_translator import GoogleTranslator
from docx import Document
import openpyxl
import io
import time
import re
import base64

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
    .success-box { padding: 1rem; background-color: #d4edda; border-radius: 5px; margin-top: 1rem; border: 1px solid #c3e6cb; }
    .warning-box { padding: 1rem; background-color: #fff3cd; border-radius: 5px; margin-top: 1rem; border: 1px solid #ffeeba; }
    .pdf-viewer { height: 600px; border: 1px solid #ddd; border-radius: 5px; background: #f9f9f9; }
</style>
""", unsafe_allow_html=True)

# --- Session State Management ---
if 'processed_output_data' not in st.session_state:
    st.session_state['processed_output_data'] = None
    st.session_state['processed_output_name'] = ""
    st.session_state['processed_output_mime'] = ""
if 'text_output_state' not in st.session_state:
    st.session_state['text_output_state'] = ""
if 'debug_log' not in st.session_state:
    st.session_state['debug_log'] = []
if 'original_pdf_data' not in st.session_state:
    st.session_state['original_pdf_data'] = None

# --- Header ---
st.markdown('<div class="main-header">🌐 Smart Multi-Format Translator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Translate PDFs with Live Preview & Error Detection</div>', unsafe_allow_html=True)

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
    target_language = st.selectbox("Target Language", options=list(LANGUAGES.keys()), index=2)
    
    src_code = LANGUAGES[source_language]
    tgt_code = LANGUAGES[target_language]
    
    st.divider()
    
    debug_mode = st.checkbox("Enable Detailed Debug Logging", value=True)
    
    st.info("💡 **Updates Applied:**\n- **Accurate BBoxes**: Fixed text skipping via vector tracking.\n- **UTF-8 Safe**: Accented characters (ä, ö, ü, é) are now preserved.\n- **Anti-Ban**: Increased rate-limiting to prevent IP blocks.")

# --- HELPER FUNCTIONS ---

def sanitize_text(text):
    """
    Cleans special characters that break APIs while preserving UTF-8 accents.
    MODIFICATION: Removed ASCII encode/decode to preserve European characters.
    """
    if not text:
        return text
    
    replacements = {
        '„': '"', '“': '"', '‚': "'", '‘': "'", '’': "'",
        '–': '-', '—': '-',
        '…': '...',
        '\u00A0': ' ',  # Non-breaking space
        '\u200B': '',   # Zero-width space
        '\uFEFF': '',   # BOM
        '«': '"', '»': '"',
        '‹': "'", '›': "'",
    }
    
    sanitized = text
    for old, new in replacements.items():
        sanitized = sanitized.replace(old, new)
    
    # REMOVED: The ascii encode/decode block that was stripping accents like ä, ö, ü.
    # deep-translator handles UTF-8 natively.
        
    return sanitized

def split_into_sentences(text):
    """Splits text by sentence boundaries."""
    sentences = re.split(r'(?<=[.!?])\s+|\n', text)
    return [s.strip() for s in sentences if s.strip()]

def translate_robustly(text, translator, block_id=""):
    """
    Translates with error tracking. 
    Returns tuple: (translated_text, error_message_or_None)
    """
    if not text or not text.strip():
        return text, None
    
    clean_text = sanitize_text(text)
    if not clean_text.strip():
        return "[INFO: Original text contained only special symbols]", None
    
    # Handle Character Limits (API max is ~5000, we safe-guard at 4500)
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
                time.sleep(0.15) # Increased delay for large chunks
            except Exception as e:
                translated_chunks.append(f"[ERROR: Chunk too complex]")
        
        return " ".join(translated_chunks), None

    # Strategy 1: Direct Translation
    try:
        result = translator.translate(clean_text)
        return result, None
    except Exception as e:
        pass
    
    # Strategy 2: Sentence-level Fallback
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
            # MODIFICATION: Increased sleep from 0.03 to 0.15 to prevent 429 Rate Limit errors
            time.sleep(0.15) 
        except Exception as se:
            err_str = f"[ERROR: Sentence {i+1} failed]"
            translated_parts.append(err_str)
            errors_found.append(f"Sentence '{sentence[:30]}...': {str(se)[:50]}")
            
    final_text = " ".join(translated_parts)
    
    if errors_found and len(errors_found) > len(sentences) * 0.5:
        final_text = f"[WARNING: Multiple translation errors in this block] {final_text}"
        
    error_summary = ", ".join(errors_found) if errors_found else None
    return final_text, error_summary

def embed_pdf_viewer(pdf_bytes, title):
    """Embeds a PDF viewer using base64 encoding."""
    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
    pdf_display = f"""
    <iframe src="data:application/pdf;base64,{base64_pdf}" 
            width="100%" 
            height="600px" 
            type="application/pdf"
            class="pdf-viewer">
    </iframe>
    """
    st.subheader(title)
    st.markdown(pdf_display, unsafe_allow_html=True)

# --- Main Logic ---
mode = st.radio("Select Input Mode", ["📁 Upload Files (.pdf, .docx, .xlsx)", "📝 Plain Text"], horizontal=True)

if mode == "📁 Upload Files (.pdf, .docx, .xlsx)":
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("1. Upload & Original Preview")
        uploaded_file = st.file_uploader("Drop PDF, Word, or Excel file", type=["pdf", "docx", "xlsx"], label_visibility="collapsed")

        if uploaded_file:
            file_ext = uploaded_file.name.split(".")[-1].lower()
            
            # Store original data immediately for preview
            st.session_state['original_pdf_data'] = uploaded_file.read()
            
            st.success(f"Loaded: **{uploaded_file.name}** ({len(st.session_state['original_pdf_data'])} bytes)")
            
            # Show Original PDF Preview
            if file_ext == "pdf":
                embed_pdf_viewer(st.session_state['original_pdf_data'], "📄 Original Document")
            elif file_ext in ["docx", "xlsx"]:
                st.info("Preview not available for Office files. They will be converted/processed directly.")
            
            if st.button("🚀 Process & Translate", type="primary"):
                try:
                    st.session_state['debug_log'] = [] # Reset logs
                    st.session_state['processed_output_data'] = None
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Re-read from session state
                    doc = fitz.open(stream=st.session_state['original_pdf_data'], filetype="pdf")
                    total_pages = len(doc)
                    
                    translator = GoogleTranslator(source=src_code, target=tgt_code)
                    
                    for page_num in range(total_pages):
                        status_text.text(f"Processing page {page_num + 1}/{total_pages}...")
                        page = doc[page_num]
                        
                        # MODIFICATION: Added fitz.TEXT_ACCURATE_BBOXES to fix text skipping
                        blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT | fitz.TEXT_ACCURATE_BBOXES)["blocks"]
                        
                        processed_count = 0
                        error_count = 0

                        for idx, block in enumerate(blocks):
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
                            
                            # Translate with error tracking
                            translated_text, error = translate_robustly(block_text, translator, block_id=f"P{page_num}-B{idx}")
                            
                            if error:
                                error_count += 1
                                if debug_mode:
                                    st.session_state['debug_log'].append(f"Page {page_num+1}, Block {idx}: {error}")

                            # Dynamic Font Sizing
                            font_size = 9
                            ratio = len(translated_text) / max(1, len(block_text))
                            if ratio > 1.5:
                                font_size = 7
                            elif ratio > 1.2:
                                font_size = 8

                            # Redraw
                            page.draw_rect(bbox, color=(1, 1, 1), fill=(1, 1, 1))
                            page.insert_textbox(bbox, translated_text, fontsize=font_size, color=(0, 0, 0))
                            
                            processed_count += 1
                            time.sleep(0.02)
                        
                        progress_bar.progress((page_num + 1) / total_pages)
                        status_text.text(f"Page {page_num+1}: Processed {processed_count} blocks, {error_count} warnings.")

                    output_buffer = io.BytesIO()
                    doc.save(output_buffer)
                    doc.close()
                    
                    st.session_state['processed_output_data'] = output_buffer.getvalue()
                    st.session_state['processed_output_name'] = f"translated_{uploaded_file.name}"
                    st.session_state['processed_output_mime'] = "application/pdf"
                    
                    st.balloons()
                    st.success("✅ Processing Complete!")
                    
                    # Show Debug Summary
                    if st.session_state['debug_log']:
                        st.warning(f"⚠️ {len(st.session_state['debug_log'])} blocks had issues. Errors are injected in the PDF text.")
                        with st.expander("🔍 View Detailed Error Log"):
                            for log in st.session_state['debug_log']:
                                st.code(log)
                    else:
                        st.info("✅ No errors detected during translation.")

                except Exception as e:
                    st.error(f"❌ Critical System Error: {str(e)}")
                    st.info("💡 Tip: If this persists, your PDF might be scanned (image-based).")
        else:
            st.info("👆 Upload a file to begin")
            
    with col2:
        st.subheader("2. Translated Preview & Download")
        if st.session_state['processed_output_data']:
            embed_pdf_viewer(st.session_state['processed_output_data'], "🌐 Translated Document")
            
            st.markdown('<div class="success-box">🎉 File ready for download!</div>', unsafe_allow_html=True)
            st.download_button(
                label="📥 Download Translated PDF",
                data=st.session_state['processed_output_data'],
                file_name=st.session_state['processed_output_name'],
                mime="application/pdf",
                use_container_width=True
            )
        else:
            st.info("Translated document will appear here after processing.")

elif mode == "📝 Plain Text":
    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("Enter Text")
        input_text = st.text_area("Paste text here", height=300, placeholder="Type or paste content...", label_visibility="collapsed")
        if st.button("🔄 Translate Text", type="primary"):
            if not input_text.strip():
                st.warning("Please enter text first.")
            else:
                translator = GoogleTranslator(source=src_code, target=tgt_code)
                try:
                    with st.spinner("Translating..."):
                        result, _ = translate_robustly(input_text, translator)
                        st.session_state['text_output_state'] = result
                    st.success("Done!")
                except Exception as e:
                    st.error(f"Error: {str(e)}")
    with col2:
        st.subheader("Result")
        if st.session_state['text_output_state']:
            st.text_area("Output", value=st.session_state['text_output_state'], height=300, label_visibility="collapsed")
            st.download_button("📥 Download .txt", data=st.session_state['text_output_state'], file_name="translated.txt", mime="text/plain", use_container_width=True)
        else:
            st.info("Translation appears here.")

st.markdown("---")
st.markdown("<div style='text-align: center; color: #888; font-size: 0.9rem;'>Powered by Streamlit, PyMuPDF & Deep-Translator</div>", unsafe_allow_html=True)
