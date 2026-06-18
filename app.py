import streamlit as st
import fitz  # PyMuPDF
from deep_translator import GoogleTranslator
import io
import time

# Page configuration
st.set_page_config(
    page_title="Free PDF Translator",
    page_icon="🌐",
    layout="centered"
)

st.title("🌐 Free PDF Translator")
st.markdown("""
Upload a PDF file and translate it to your desired language. 
This tool preserves the original layout while translating text blocks.
""")

# Language options
LANGUAGES = {
    "English": "en",
    "Spanish": "es",
    "French": "fr",
    "German": "de",
    "Italian": "it",
    "Portuguese": "pt",
    "Chinese (Simplified)": "zh-CN",
    "Japanese": "ja",
    "Korean": "ko",
    "Russian": "ru",
    "Arabic": "ar",
    "Hindi": "hi",
    "Dutch": "nl",
    "Polish": "pl",
    "Turkish": "tr",
    "Vietnamese": "vi",
    "Thai": "th",
    "Indonesian": "id",
    "Swedish": "sv",
    "Norwegian": "no",
    "Danish": "da",
    "Finnish": "fi",
    "Greek": "el",
    "Czech": "cs",
    "Romanian": "ro",
    "Hungarian": "hu",
    "Ukrainian": "uk",
    "Bengali": "bn",
    "Tamil": "ta",
    "Telugu": "te",
    "Marathi": "mr",
    "Gujarati": "gu",
    "Kannada": "kn",
    "Malayalam": "ml",
    "Punjabi": "pa",
    "Urdu": "ur",
    "Persian": "fa",
    "Hebrew": "he",
    "Swahili": "sw",
    "Afrikaans": "af",
    "Zulu": "zu",
    "Xhosa": "xh",
    "Yoruba": "yo",
    "Igbo": "ig",
    "Amharic": "am",
    "Somali": "so",
    "Hausa": "ha",
    "Shona": "sn",
    "Chichewa": "ny",
    "Sesotho": "st",
    "Setswana": "tn",
    "Xitsonga": "ts",
    "Tshivenda": "ve",
    "IsiNdebele": "nr",
    "SiSwati": "ss",
    "IsiXhosa": "xh",
    "IsiZulu": "zu",
    "Sesotho sa Leboa": "nso",
    "Sepedi": "nso",
    "Setswana": "tn",
    "Xitsonga": "ts",
    "Tshivenda": "ve",
    "IsiNdebele": "nr",
    "SiSwati": "ss"
}

# Sidebar for language selection
st.sidebar.header("Translation Settings")
target_language = st.sidebar.selectbox(
    "Select Target Language",
    options=list(LANGUAGES.keys()),
    index=1  # Default to Spanish
)

target_lang_code = LANGUAGES[target_language]

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file is not None:
    st.info(f"📄 Uploaded: {uploaded_file.name}")
    
    if st.button("🚀 Start Translation", type="primary"):
        try:
            # Read PDF into memory
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            
            # Initialize translator
            translator = GoogleTranslator(source='auto', target=target_lang_code)
            
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            total_pages = len(doc)
            
            for page_num in range(total_pages):
                status_text.text(f"Processing page {page_num + 1} of {total_pages}...")
                page = doc[page_num]
                
                # Get text blocks with coordinates
                blocks = page.get_text("dict")["blocks"]
                
                for block in blocks:
                    if "lines" not in block:
                        continue
                    
                    # Extract text from block
                    block_text = ""
                    for line in block["lines"]:
                        for span in line["spans"]:
                            block_text += span["text"]
                    
                    if not block_text.strip():
                        continue
                    
                    # Get bounding box
                    bbox = fitz.Rect(block["bbox"])
                    
                    try:
                        # Translate the text
                        translated_text = translator.translate(block_text)
                        
                        # Calculate font size based on original text length vs translated text length
                        original_length = len(block_text)
                        translated_length = len(translated_text)
                        
                        # Base font size
                        font_size = 10
                        
                        # If translated text is significantly longer, reduce font size
                        if translated_length > original_length * 1.5:
                            font_size = max(6, font_size * (original_length / translated_length))
                        elif translated_length > original_length * 1.2:
                            font_size = max(7, font_size * 0.9)
                        
                        # Draw white rectangle to erase original text
                        page.draw_rect(bbox, color=(1, 1, 1), fill=(1, 1, 1))
                        
                        # Insert translated text
                        page.insert_textbox(
                            bbox,
                            translated_text,
                            fontsize=font_size,
                            color=(0, 0, 0)
                        )
                        
                    except Exception as e:
                        # Skip this block if translation fails
                        st.warning(f"⚠️ Could not translate a text block on page {page_num + 1}: {str(e)}")
                        continue
                
                # Update progress
                progress_bar.progress((page_num + 1) / total_pages)
            
            status_text.text("✅ Translation complete!")
            
            # Save translated PDF to bytes
            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            doc.close()
            output_buffer.seek(0)
            
            # Success message
            st.success("🎉 PDF translation completed successfully!")
            
            # Download button
            st.download_button(
                label="📥 Download Translated PDF",
                data=output_buffer.getvalue(),
                file_name=f"translated_{uploaded_file.name}",
                mime="application/pdf"
            )
            
        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
            st.error("Please try again with a different PDF file.")

else:
    st.info("👆 Please upload a PDF file to begin translation.")

# Footer
st.markdown("---")
st.markdown("""
**Note:** This tool uses free translation APIs. For very large PDFs or rare languages, 
some text blocks may not translate perfectly. The layout preservation works best with 
text-based PDFs (not scanned images).
""")