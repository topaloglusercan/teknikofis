import streamlit as st

# 1. TÜM PORTAL İÇİN GENİŞ EKRAN (WIDE) AYARI 
st.set_page_config(page_title="Teknik Ofis Portalı", layout="wide", initial_sidebar_state="expanded")

# --- KARŞILAMA EKRANI VE MODÜL ÖZETLERİ ---
def ana_sayfa():
    st.markdown("<h1 style='text-align: center; color: #2c3e50;'>TEKNİK OFİS PORTALI</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #7f8c8d; font-size: 18px;'>Şantiye ve Teknik Ofis Dijital Süreç Yönetimi</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    c1, c2 = st.columns(2)
    
    with c1:
        st.markdown("### 📁 Teknik Ofis Modülleri")
        st.info("**İdari Hakediş:** Taşeron ve ana firma hakedişleri için sayısal kontrol sistem aracıdır.")
        st.info("**Pursantaj:** Sözleşme bedelinin iş kalemlerine göre dağılımını planlayan, izleyen ve analiz eden bütünleşik bir dijital araçtır.")
        st.info("**Şantiye Tutanak:** Eklenti ve kesinti tutanakları hazırlayıp PDF olarak indirebileceğiniz, filtrelenebilir dijital arşivdir.")
        st.info("**Performans Analizi:** Proje harcamalarını Kazanılmış Takvim (ESA) ve Kazanılmış Değer (EVA) ile kıyaslayan performans ölçüm aracıdır.")
        st.info("**Fiyat Farkı Simülatörü:** Kova sistemi ve gecikme matrisi mantığıyla fiyat farkı senaryolarını anlık simüle edip karşılaştıran araçtır.")
        st.info("**Pursantaj Revize:** Keşif/iş artışı durumunda tüm kalemlerin pursantaj ve parasal ağırlıklarını %100'e kilitlenecek şekilde otomatik paylaştıran motordur.")
        st.info("**Teklif Karşılaştırma:** Fiyat, Kalite ve Finansal Güç kriterlerini TOPSIS algoritmasıyla ağırlıklandırarak en ideal alt yükleniciyi bulan ve raporlayan karar destek asistanıdır.")

    with c2:
        st.markdown("### 🔍 P6 İnceleme Modülleri")
        st.success("**P6 Adam-Saat Analizi:** Primavera P6 (XER) veritabanını tarayarak iş kalemlerinin kaynak dağılımlarını şelale (spread) yöntemiyle aylara bölen ve ekip ihtiyacını görselleştiren analiz aracıdır.")
        st.success("**P6 Lag Analizi:** XER dosyasındaki aktivite ilişkilerini (Network) tarayarak aralara gizlenmiş Lag (Bekleme Süresi) değerlerini anında tespit edip listeleyen kontrol modülüdür.")
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.caption("🔒 *Güvenliğiniz için girdiğiniz veriler sunucuda tutulmaz. İşlemleriniz sadece tarayıcınızın belleğinde gerçekleşir.*")

# Material ikonlar (tek renkli, kurumsal vektörel çizimler) kullanıldı
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
    st.Page("pages/p6_lag_analizi.py", title="P6 Lag Analizi", icon=":material/hourglass_empty:")
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
