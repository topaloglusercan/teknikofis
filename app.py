import streamlit as st
import os

st.set_page_config(page_title="Teknik Ofis Portalı", layout="wide", page_icon="🏗️")

def ana_sayfa():
    st.title("🏗️ Teknik Ofis Portalı")
    st.markdown("""
    ### Hoş Geldiniz!
    Bu portal şantiye ve teknik ofis süreçlerinizi dijitalleştirmek için tasarlanmıştır.
    Sol taraftaki menüyü kullanarak ilgili modüle geçiş yapabilirsiniz.
    """)

giris_sayfasi = st.Page(ana_sayfa, title="Ana Sayfa", icon="🏠", default=True)

# pages klasöründeki dosyaları otomatik buluyoruz (harf hatası riski kalmaz)
pages_dir = "pages"
bulunan_dosyalar = os.listdir(pages_dir) if os.path.exists(pages_dir) else []

eski_moduller = []
p6_modulleri = []

for dosya in sorted(bulunan_dosyalar):
    if dosya.endswith(".py"):
        dosya_yolu = f"pages/{dosya}"
        # Dosya adı temizlenerek menü ismi yapılır (p6_adam_saat -> P6 Adam Saat)
        temiz_isim = dosya.replace(".py", "").replace("_", " ").title()
        
        if "p6" in dosya.lower():
            p6_modulleri.append(st.Page(dosya_yolu, title=temiz_isim, icon="👷‍♂️"))
        else:
            eski_moduller.append(st.Page(dosya_yolu, title=temiz_isim, icon="📄"))

sayfalar = {
    "Giriş": [giris_sayfasi],
    "Teknik Ofis Modülleri": eski_moduller,
    "P6 İnceleme": p6_modulleri
}

pg = st.navigation(sayfalar)
st.sidebar.markdown("<h2 style='text-align: center; color: #1f77b4;'>TEKNİK OFİS PORTALI</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")
pg.run()
