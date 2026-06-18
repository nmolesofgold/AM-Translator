import streamlit as st
import fitz  # PyMuPDF
from deep_translator import GoogleTranslator
import io
import time

# --- Page Configuration & Custom CSS ---
st.set_page_config(
    page_title="Smart PDF & Text Translator",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #4F8BF9;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- Initialize Global Session States ---
if 'translated_pdf_data' not in st.session_state:
    st.session_state['translated_pdf_data'] = None
    st.session_state['translated_pdf_name'] = ""
if 'translated_result' not in st.session_state:
    st.session_state['translated_result'] = ""

# --- Header ---
st.markdown('<div class="main-header">🌐 Smart Translator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Translate PDFs (with layout preservation) or Plain Text instantly.</div>', unsafe_allow_html=True)

# --- Sidebar Settings ---
with st.sidebar:
    st.header("⚙️ Settings")

    LANGUAGES = {
        "English": "en", "Spanish": "es", "French": "fr", "German": "de",
        "Italian": "it", "Portuguese": "pt", "Chinese (Simplified)": "zh-CN",
        "Japanese": "ja", "Korean": "ko", "Russian": "ru", "Arabic": "ar",
        "Hindi": "hi", "Dutch": "nl", "Polish": "pl", "Turkish": "tr",
        "Vietnamese": "vi", "Thai": "th", "Indonesian": "id", "Swedish": "sv",
        "Norwegian": "no", "Danish": "da", "Finnish": "fi", "Greek": "el",
        "Czech": "cs", "Romanian": "ro", "Hungarian": "hu", "Ukrainian": "uk"
    }

    target_language = st.selectbox(
        "Target Language",
        options=list(LANGUAGES.keys()),
        index=1  # Default Spanish
    )
    target_lang_code = LANGUAGES[target_language]

    st.divider()
    st.info("💡 **Tip:** For PDFs, this tool preserves layout by redrawing text blocks. For best results, use text-based PDFs (not scanned images).")

# --- Main Content Area ---
mode = st.radio("Choose Input Mode", ["📄 PDF File", "📝 Plain Text"], horizontal=True)

col1, col2 = st.columns([1, 1])

if mode == "📄 PDF File":
    with col1:
        st.subheader("Upload PDF")
        uploaded_file = st.file_uploader("Drop your PDF here", type=["pdf"], label_visibility="collapsed")

        if uploaded_file:
            st.success(f"Loaded: **{uploaded_file.name}** ({round(uploaded_file.size/1024, 2)} KB)")

            if st.button("🚀 Translate PDF", type="primary"):
                try:
                    pdf_bytes = uploaded_file.read()
                    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    translator = GoogleTranslator(source='auto', target=target_lang_code)

                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    total_pages = len(doc)

                    for page_num in range(total_pages):
                        status_text.text(f"Processing page {page_num + 1}/{total_pages}...")
                        page = doc[page_num]

                        try:
                            blocks = page.get_text("dict")["blocks"]
                        except KeyError:
                            continue

                        for block in blocks:
                            if "lines" not in block: continue

                            block_text = "".join(span["text"] for line in block["lines"] for span in line["spans"])
                            if not block_text.strip(): continue

                            bbox = fitz.Rect(block["bbox"])

                            try:
                                translated_text = translator.translate(block_text)
                                time.sleep(0.15)  # Anti-throttling micro-delay

                                # Dynamic Font Sizing
                                orig_len = max(1, len(block_text))
                                trans_len = max(1, len(translated_text))
                                font_size = 10

                                if trans_len > orig_len * 1.5:
                                    font_size = max(6, font_size * (orig_len / trans_len))
                                elif trans_len > orig_len * 1.2:
                                    font_size = max(7, font_size * 0.9)

                                # Redraw Layering
                                page.draw_rect(bbox, color=(1, 1, 1), fill=(1, 1, 1))
                                page.insert_textbox(bbox, translated_text, fontsize=font_size, color=(0, 0, 0))

                            except Exception:
                                continue

                        progress_bar.progress((page_num + 1) / total_pages)

                    status_text.empty()
                    progress_bar.empty()

                    # Store variables inside memory states
                    output_buffer = io.BytesIO()
                    doc.save(output_buffer)
                    doc.close()

                    st.session_state['translated_pdf_data'] = output_buffer.getvalue()
                    st.session_state['translated_pdf_name'] = f"translated_{uploaded_file.name}"
                    st.balloons()

                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
        else:
            # Clear storage states if file is removed
            st.session_state['translated_pdf_data'] = None
            st.info("👆 Upload a PDF to begin")
            
    with col2:
        st.subheader("Action Status")
        if st.session_state['translated_pdf_data'] is not None:
            st.success("🎉 Translation ready for download!")
            st.download_button(
                label="📥 Download Translated PDF",
                data=st.session_state['translated_pdf_data'],
                file_name=st.session_state['translated_pdf_name'],
                mime="application/pdf",
                use_container_width=True
            )
        else:
            st.info("No compiled operations queue found. Upload and process a PDF on the left column to run the file assembler.")

elif mode == "📝 Plain Text":
    with col1:
        st.subheader("Enter Text")
        input_text = st.text_area("Paste your text here", height=300, placeholder="Type or paste text to translate...")

        if st.button("🔄 Translate Text", type="primary"):
            if not input_text.strip():
                st.warning("Please enter some text first.")
            else:
                try:
                    with st.spinner("Translating..."):
                        translator = GoogleTranslator(source='auto', target=target_lang_code)
                        translated_text = translator.translate(input_text)

                    st.session_state['translated_result'] = translated_text
                    st.success("Translation complete!")
                except Exception as e:
                    st.error(f"Translation failed: {str(e)}")

    with col2:
        st.subheader("Result")
        if st.session_state['translated_result']:
            st.text_area("Translated Output", value=st.session_state['translated_result'], height=300, label_visibility="collapsed")

            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    label="📥 Download .txt",
                    data=st.session_state['translated_result'],
                    file_name="translated_text.txt",
                    mime="text/plain",
                    use_container_width=True
                )
            with c2:
                st.code(st.session_state['translated_result'], language=None)
                st.caption("Click the copy icon in the top-right corner to copy the text.")
        else:
            st.info("Translation will appear here.")

# Footer
st.markdown("---")
st.markdown("<div style='text-align: center; color: #888; font-size: 0.9rem;'>Powered by Streamlit, PyMuPDF & Deep-Translator</div>", unsafe_allow_html=True)
