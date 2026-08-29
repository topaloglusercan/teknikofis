import streamlit as st

st.set_page_config(page_title="Teknik Ofis Portalı", layout="wide", page_icon="🏗️")

def ana_sayfa():
    st.title("🏗️ Teknik Ofis Portalı")
    st.markdown("""
    ### Hoş Geldiniz!
    Bu portal şantiye ve teknik ofis süreçlerinizi dijitalleştirmek için tasarlanmıştır.
    Sol taraftaki menüyü kullanarak ilgili modüle geçiş yapabilirsiniz.
    """)

giris_sayfasi = st.Page(ana_sayfa, title="Ana Sayfa", icon="🏠", default=True)

# --- ESKİ MODÜLLER ---
eski_moduller = [
    st.Page("pages/1_idari_hakedis.py", title="İdari Hakediş", icon="📄"),
    st.Page("pages/2_pursantaj.py", title="Pursantaj", icon="📊"),
    st.Page("pages/3_santiye_tutanak.py", title="Şantiye Tutanak", icon="📝"),
    st.Page("pages/4_performans_analizi.py", title="Performans Analizi", icon="📈"),
    st.Page("pages/5_fiyat_farki_simulatoru.py", title="Fiyat Farkı Simülatörü", icon="💰"),
    st.Page("pages/6_pursantaj_revize.py", title="Pursantaj Revize", icon="🔄"),
    # İsimdeki tüm harfler İngilizce karakter (i) yapıldı:
    st.Page("pages/7_teklif_karsilastirma.py", title="Teklif Karşılaştırma", icon="⚖️") 
]

# --- YENİ P6 MODÜLLERİ ---
p6_modulleri = [
    st.Page("pages/p6_adam_saat.py", title="P6 Adam-Saat Analizi", icon="👷‍♂️"),
    st.Page("pages/p6_lag_analizi.py", title="P6 Lag Analizi", icon="⏳")
]

# --- SOL MENÜ KATEGORİLERİ ---
sayfalar = {
    "Giriş": [giris_sayfasi],
    "Teknik Ofis Modülleri": eski_moduller,
    "P6 İnceleme": p6_modulleri
}

pg = st.navigation(sayfalar)
st.sidebar.markdown("<h2 style='text-align: center; color: #1f77b4;'>TEKNİK OFİS PORTALI</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")
pg.run()
