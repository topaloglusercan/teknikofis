import streamlit as st
import streamlit.components.v1 as components
import os

st.set_page_config(
    page_title="Şantiye Tutanak",
    page_icon="📋",
    layout="wide"
)

# Sayfanın kendi padding'ini kaldır — tam ekran iframe için
st.markdown("""
    <style>
        .block-container { padding: 0 !important; margin: 0 !important; max-width: 100% !important; }
        header { display: none !important; }
        #MainMenu { display: none !important; }
        footer { display: none !important; }
    </style>
""", unsafe_allow_html=True)

# HTML dosyasını oku
HTML_DOSYASI = os.path.join(os.path.dirname(__file__), "santiye_tutanak.html")

if os.path.exists(HTML_DOSYASI):
    with open(HTML_DOSYASI, "r", encoding="utf-8") as f:
        html_icerik = f.read()

    # Streamlit components.html ile göm — scrolling=True ve yeterli yükseklik ver
    components.html(
        html_icerik,
        height=950,
        scrolling=True
    )
else:
    st.error("❌ `santiye_tutanak.html` dosyası bulunamadı.")
    st.info("""
    **Kurulum:**
    1. `santiye_tutanak.html` dosyasını bu `.py` dosyasıyla **aynı klasöre** koyun
    2. Sayfayı yenileyin
    """)
    st.code("""
    pages/
    ├── santiye_tutanak.py       ← bu dosya
    ├── santiye_tutanak.html     ← HTML aracı (aynı klasörde)
    ├── idari_hakedis.py
    └── pursantaj.py
    """)
