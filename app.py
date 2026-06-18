import streamlit as st
import fitz  # PyMuPDF
from deep_translator import GoogleTranslator
import io
import time
import re

# --- Page Configuration ---
st.set_page_config(page_title="PDF Translator", page_icon="🌐", layout="wide")

# --- Session State ---
if 'processed_output_data' not in st.session_state:
    st.session_state['processed_output_data'] = None
    st.session_state['processed_output_name'] = ""
if 'text_input' not in st.session_state:
    st.session_state['text_input'] = ""
if 'translated_text_result' not in st.session_state:
    st.session_state['translated_text_result'] = ""

# --- Header ---
st.title("🌐 Professional PDF Translator")

# --- Sidebar ---
with st.sidebar:
    st.header("Settings")
    LANGUAGES = {
        "Auto Detect": "auto", "German": "de", "English": "en", "Spanish": "es",
        "French": "fr", "Italian": "it", "Portuguese": "pt", "Chinese (Simplified)": "zh-CN",
        "Japanese": "ja", "Korean": "ko", "Russian": "ru", "Arabic": "ar", "Hindi": "hi"
    }
    src_code = st.selectbox("Source", options=list(LANGUAGES.values()), index=0)
    tgt_code = st.selectbox("Target", options=list(LANGUAGES.values()), index=2)
    
    # Map values back to keys if needed, but API uses codes directly
    # Note: selectbox returns the value if we pass values, but let's be explicit
    # Actually, simpler to just use the codes directly in the dict values above.
    
    debug_mode = st.checkbox("Show Debug Logs", value=False)

# --- Helper Functions ---

def sanitize_text(text):
    if not text: return text
    # Only remove invisible control chars, keep visible punctuation like quotes/dashes
    replacements = {'\u00A0': ' ', '\u200B': '', '\uFEFF': '', '\u200C': '', '\u200D': ''}
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text

def translate_robustly(text, translator):
    if not text or not text.strip(): return text, None
    clean_text = sanitize_text(text)
    if len(clean_text) > 4500:
        # Simple chunking for very large blocks
        try:
            return translator.translate(clean_text[:4500]), None
        except: return "[Error: Text too long]", None
    
    try:
        return translator.translate(clean_text), None
    except Exception as e:
        # Fallback: split by sentence
        sentences = re.split(r'(?<=[.!?])\s+', clean_text)
        result = []
        for s in sentences:
            if s.strip():
                try:
                    result.append(translator.translate(s))
                    time.sleep(0.1)
                except:
                    result.append(s) # Keep original if fails
        return " ".join(result), "Partial Failure"

def instant_translate():
    txt = st.session_state['text_input']
    if not txt.strip():
        st.session_state['translated_text_result'] = ""
        return
    try:
        translator = GoogleTranslator(source=src_code, target=tgt_code)
        res, _ = translate_robustly(txt, translator)
        st.session_state['translated_text_result'] = res
    except Exception as e:
        st.session_state['translated_text_result'] = f"Error: {e}"

# --- Main Logic ---
mode = st.radio("Mode", ["Upload PDF", "Plain Text"], horizontal=True)

if mode == "Upload PDF":
    uploaded_file = st.file_uploader("Choose PDF", type="pdf")
    
    if uploaded_file:
        if st.button("Translate PDF", type="primary"):
            try:
                pdf_bytes = uploaded_file.read()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                translator = GoogleTranslator(source=src_code, target=tgt_code)
                
                progress_bar = st.progress(0)
                
                for i, page in enumerate(doc):
                    # Get blocks with accurate boxes
                    blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT | fitz.TEXT_ACCURATE_BBOXES)["blocks"]
                    
                    for b_idx, block in enumerate(blocks):
                        if "lines" not in block: continue
                        
                        # --- FIX 1: Better Text Extraction (Preserve Lists) ---
                        block_text_parts = []
                        for line in block["lines"]:
                            line_text = ""
                            for span in line["spans"]:
                                line_text += span["text"]
                            if line_text.strip():
                                block_text_parts.append(line_text.strip())
                        
                        block_text = "\n".join(block_text_parts) # Preserve internal newlines
                        
                        if not block_text.strip(): continue
                        
                        bbox = fitz.Rect(block["bbox"])
                        
                        # Translate
                        trans_text, err = translate_robustly(block_text, translator)
                        if err and debug_mode:
                            st.warning(f"Page {i+1}: {err}")
                        
                        # --- FIX 2: Double-Cover Erase to prevent Ghosting ---
                        # Expand the box slightly by 2 points in all directions to ensure full coverage
                        erase_bbox = bbox + (-2, -2, 2, 2) 
                        page.draw_rect(erase_bbox, color=(1, 1, 1), fill=(1, 1, 1))
                        
                        # --- FIX 3: Smart Font Sizing & Insertion ---
                        font_size = 10.0
                        # Check fitment
                        while font_size > 5.0:
                            # render_mode=3 checks fit without drawing
                            rc = page.insert_textbox(bbox, trans_text, fontsize=font_size, color=(0,0,0), render_mode=3)
                            if rc >= 0: # Fits!
                                break
                            font_size -= 0.5
                        
                        # Final Draw
                        page.insert_textbox(bbox, trans_text, fontsize=font_size, color=(0, 0, 0))
                    
                    progress_bar.progress((i + 1) / len(doc))
                
                # Save
                out_buf = io.BytesIO()
                doc.save(out_buf)
                doc.close()
                
                st.session_state['processed_output_data'] = out_buf.getvalue()
                st.session_state['processed_output_name'] = f"translated_{uploaded_file.name}"
                st.balloons()
                st.success("Done!")
                
            except Exception as e:
                st.error(f"Critical Error: {e}")

    if st.session_state['processed_output_data']:
        st.download_button(
            label="📥 Download Translated PDF",
            data=st.session_state['processed_output_data'],
            file_name=st.session_state['processed_output_name'],
            mime="application/pdf",
            use_container_width=True
        )

elif mode == "Plain Text":
    c1, c2 = st.columns(2)
    with c1:
        st.text_area("Input", key="text_input", height=400, on_change=instant_translate, placeholder="Type here...")
    with c2:
        if st.session_state['translated_text_result']:
            st.text_area("Output", value=st.session_state['translated_text_result'], height=400, label_visibility="collapsed")
            st.download_button("Download .txt", st.session_state['translated_text_result'], "translated.txt")
        else:
            st.info("Translation appears here...")
