import streamlit as st

st.set_page_config(page_title="Teknik Ofis Portalı", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
        .block-container {
            max-width: 95% !important;
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }
    </style>
""", unsafe_allow_html=True)

def ana_sayfa():
    st.markdown("<h1 style='text-align: center; color: #2c3e50;'>TEKNİK OFİS PORTALI</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #7f8c8d; font-size: 18px;'>Şantiye ve Teknik Ofis Dijital Süreç Yönetimi</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("""
    ### Hoş Geldiniz!
    Bu portal şantiye ve teknik ofis süreçlerinizi dijitalleştirmek için tasarlanmıştır.
    Sol taraftaki menüyü kullanarak ilgili modüle geçiş yapabilirsiniz.
    <br><br>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📁 Teknik Ofis Modülleri")
        st.markdown("""
        * **İdari Hakediş:** Taşeron ve ana firma hakedişleri için sayısal kontrol.
        * **Pursantaj:** Sözleşme bedelinin kalemlere göre dağılımı ve analizi.
        * **Şantiye Tutanak:** Eklenti/kesinti tutanakları hazırlama ve PDF arşivi.
        * **Performans Analizi:** ESA ve EVA yöntemleriyle proje performans ölçümü.
        * **Fiyat Farkı Simülatörü:** Gecikme matrisi ile fiyat farkı simülasyonu.
        * **Pursantaj Revize:** İş artışlarında pursantajların otomatik yeni dağıtımı.
        * **Teklif Karşılaştırma:** TOPSIS algoritmasıyla ideal alt yüklenici seçimi.
        """)

    with c2:
        st.markdown("#### 🔍 P6 İnceleme Modülleri")
        st.markdown("""
        * **P6 Adam-Saat Analizi:** XER veritabanından kaynak dağılımlarını aylara bölen analiz.
        * **P6 S-Eğrisi (İlerleme):** Bütçe/maliyet verileriyle kümülatif S-Eğrisi grafikleri.
        * **P6 Lag Analizi:** Aktivite ilişkilerindeki gizlenmiş Lag (Bekleme) değerlerinin tespiti.
        """)
        
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    st.caption("🔒 *Güvenliğiniz için girdiğiniz veriler sunucuda tutulmaz. İşlemleriniz sadece tarayıcınızın belleğinde gerçekleşir.*")

giris_sayfasi = st.Page(ana_sayfa, title="Ana Sayfa", icon=":material/home:", default=True)

eski_moduller = [
    st.Page("pages/1_idari_hakedis.py", title="İdari Hakediş", icon=":material/receipt_long:"),
    st.Page("pages/2_pursantaj.py", title="Pursantaj", icon=":material/pie_chart:"),
    st.Page("pages/3_santiye_tutanak.py", title="Şantiye Tutanak", icon=":material/edit_document:"),
    st.Page("pages/4_performans_analizi.py", title="Performans Analizi", icon=":material/insights:"),
    st.Page("pages/5_fiyat_farki_simulatoru.py", title="Fiyat Farkı Simülatörü", icon=":material/price_change:"),
    st.Page("pages/6_pursantaj_revize.py", title="Pursantaj Revize", icon=":material/sync_alt:"),
    st.Page("pages/7_teklif_karsılastırma.py", title="Teklif Karşılaştırma", icon=":material/balance:") 
]

p6_modulleri = [
    st.Page("pages/p6_adam_saat.py", title="P6 Adam-Saat Analizi", icon=":material/engineering:"),
    st.Page("pages/p6_s_egrisi.py", title="P6 S-Eğrisi", icon=":material/show_chart:"),
    st.Page("pages/p6_lag_analizi.py", title="P6 Lag Analizi", icon=":material/hourglass_empty:"),
    st.Page("pages/p6_gereksiz_baglar.py", title="P6 Gereksiz Bağ Analizi", icon=":material/link_off:"),
    # YENİ MODÜL BURAYA EKLENDİ:
    st.Page("pages/p6_aktivite_kodu_analizi.py", title="P6 Aktivite Kodu Analizi", icon=":material/dashboard:")
]

sayfalar = {
    "Giriş": [giris_sayfasi],
    "Teknik Ofis Modülleri": eski_moduller,
    "P6 İnceleme": p6_modulleri
}

pg = st.navigation(sayfalar)
st.sidebar.markdown("<h3 style='text-align: center; color: #1f77b4;'>TEKNİK OFİS PORTALI</h3>", unsafe_allow_html=True)
st.sidebar.markdown("---")
pg.run()
