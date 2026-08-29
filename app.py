import streamlit as st

st.set_page_config(page_title="Teknik Ofis Portalı", layout="wide", page_icon="🏗️")

# --- ANA SAYFA (Karşılama Ekranı) FONKSİYONU ---
# Senin yazdığın o uzun açıklama metnini burada koruyoruz
def ana_sayfa():
    st.title("🏗️ Teknik Ofis Portalı")
    st.markdown("""
    ### Hoş Geldiniz!
    Bu portal şantiye ve teknik ofis süreçlerinizi dijitalleştirmek için tasarlanmıştır.
        
    **Nasıl Kullanılır?**
    Sol taraftaki menüyü kullanarak modüller arasında geçiş yapabilirsiniz:
    * 📂 **İdari Hakediş:** İdari Hakediş sayısal kontrol sistem aracıdır.
    * 📂 **Pursantaj:** Pursantaj Yönetim Sistemi, inşaat projelerinde sözleşme bedelinin iş kalemlerine göre dağılımını planlayan, izleyen ve analiz eden bütünleşik bir dijital araçtır.
    * 📂 **Şantiye Tutanak:** İnşaat projelerinde eklenti ve kesinti tutanakları hazırlayıp, pdf dosyası indirebileceğiniz, filtre yapabileceğiniz bir dijital araçtır.
    * 📂 **Performans Analizi:** Projelerdeki harcama ve ilerleme durumunu Kazanılmış Takvim (ESA) ve Kazanılmış Değer (EVA) yöntemleriyle kıyaslayan performans ölçüm aracıdır.
    * 📂 **Fiyat Farkı Simülatörü:** Kova sistemi ve gecikme matrisi mantığıyla hesaplanan fiyat farkını, gecikme/hızlanma, imalat hızı, endeks artışı ve alt endeks ağırlıkları gibi değişkenleri slider'larla anlık değiştirerek simüle etmenizi ve farklı senaryoları yan yana karşılaştırmanızı sağlayan bir dijital araçtır.
    * 📂 **Revize Pursantaj Dağıtım:** Revize Pursantaj Dağıtım Motoru, sözleşme bedelinin değişmesi (keşif artışı) veya yeni mukayese işlerinin (iş artışı/eksilişi) sisteme dahil olması durumunda, projedeki tüm iş kalemlerinin pursantaj (yüzde) ve parasal (TL) ağırlıklarını belirli kısıtlar altında ve %100'e kilitlenecek şekilde otomatik olarak yeniden paylaştıran gelişmiş bir teknik ofis aracıdır.
    * 📂 **Dinamik Teklif Kıyaslama ve Karar Destek Motoru:** Teklif alma süreçlerindeki manuel tablo birleştirme ve salt fiyata dayalı yüzeysel karar verme hatalarını ortadan kaldıran; firmaların Fiyat, Kalite, Finansal Güç ve Teslimat Hızı kriterlerini TOPSIS (Çok Kriterli Karar Verme) algoritmasıyla ağırlıklandırarak size en ideal ("Optimum") alt yükleniciyi matematiksel olarak sunan gelişmiş bir teknik ofis asistanıdır. Sistem, kalem bazlı min/max fiyat ısı haritaları çıkartır, "Karma Dağılım" senaryosu ile maksimum tasarruf tutarını hesaplar ve tüm analizleri saniyeler içinde A3 boyutunda, yönetime sunulmaya hazır, dilerseniz isimlerin gizlendiği (anonimleştirilmiş) profesyonel E-Tablo raporlarına dönüştürür.

    *Not: Güvenliğiniz için girdiğiniz veriler sunucuda tutulmaz, yan menüden projelerinizi kendi bilgisayarınıza indirebilir ve yükleyebilirsiniz.*
    """)

# Giriş ekranını navigasyona bir sayfa olarak ekliyoruz
giris_sayfasi = st.Page(ana_sayfa, title="Ana Sayfa", icon="🏠", default=True)

# --- ESKİ MODÜLLER (Senin GitHub'daki dosyaların) ---
eski_moduller = [
    st.Page("pages/1_idari_hakedis.py", title="İdari Hakediş", icon="📄"),
    st.Page("pages/2_pursantaj.py", title="Pursantaj", icon="📊"),
    st.Page("pages/3_santiye_tutanak.py", title="Şantiye Tutanak", icon="📝"),
    st.Page("pages/4_performans_analizi.py", title="Performans Analizi", icon="📈"),
    st.Page("pages/5_fiyat_farki_simulatoru.py", title="Fiyat Farkı Simülatörü", icon="💰"),
    st.Page("pages/6_pursantaj_revize.py", title="Pursantaj Revize", icon="🔄"),
    st.Page("pages/7_teklif_karsilastirma.py", title="Teklif Karşılaştırma", icon="⚖️")
]

# --- YENİ EKLENEN P6 MODÜLLERİ ---
p6_modulleri = [
    st.Page("pages/p6_adam_saat.py", title="P6 Adam-Saat Analizi", icon="👷‍♂️"),
    st.Page("pages/p6_lag_analizi.py", title="P6 Lag (Bekleme) Analizi", icon="⏳")
]

# --- SOL MENÜ KATEGORİLERİ ---
sayfalar = {
    "Giriş": [giris_sayfasi],
    "Teknik Ofis Modülleri": eski_moduller,
    "P6 İnceleme": p6_modulleri
}

pg = st.navigation(sayfalar)

# Sol Menü Üst Başlığı (Firma isminden bağımsız olarak Teknik Ofis Portalı yazıyor)
st.sidebar.markdown("<h2 style='text-align: center; color: #1f77b4;'>TEKNİK OFİS PORTALI</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")

# Sayfayı Çalıştır
pg.run()
