"""
Revize Pursantaj Dağıtım Motoru (Üretim / Prod Sürümü - Esnek Hedefli)
============================================================
- Kullanıcıya kuruş ve yuvarlama farklarını (residual error) 
  istediği kaleme yedirme özgürlüğü eklendi.
- Sistem en büyük bütçeli esnek kalemi otomatik önerir ancak 
  son söz tamamen kullanıcıdadır.
"""

import streamlit as st
import pandas as pd
import json
import io
import re

# ==========================================
# PROD AYARLARI: Sayfa Yapısı
# ==========================================
st.set_page_config(page_title="Revize Pursantaj Dağıtımı", page_icon="⚖️", layout="wide", initial_sidebar_state="expanded")

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# ==========================================
# 0. YARDIMCI FONKSİYONLAR
# ==========================================
def get_template_df():
    return pd.DataFrame([
        {"Kalem": "A3 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 11.01, "Eski TL": 129369962.66, "Durum": "TL Sabit"},
        {"Kalem": "B1 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 10.22, "Eski TL": 120087285.96, "Durum": "TL Sabit"},
        {"Kalem": "D1 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 3.93, "Eski TL": 46178379.04, "Durum": "TL Sabit"},
        {"Kalem": "D2 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 3.93, "Eski TL": 46178379.04, "Durum": "TL Sabit"},
        {"Kalem": "E1 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 4.71, "Eski TL": 55309567.27, "Durum": "TL Sabit"},
        {"Kalem": "E2 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 4.71, "Eski TL": 55309567.27, "Durum": "TL Sabit"},
        {"Kalem": "F1 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 9.43, "Eski TL": 110925588.52, "Durum": "TL Sabit"},
        {"Kalem": "G1 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 9.43, "Eski TL": 110925588.52, "Durum": "TL Sabit"},
        {"Kalem": "G2 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 11.01, "Eski TL": 129369962.66, "Durum": "TL Sabit"},
        {"Kalem": "G3 TİPİ KONUT BLOĞU İNŞAATI", "Eski %": 10.22, "Eski TL": 120087285.96, "Durum": "TL Sabit"},
        {"Kalem": "ADAİÇİ ALTYAPI VE ÇEVRE DÜZENLEME İŞLERİ", "Eski %": 2.49, "Eski TL": 29258056.95, "Durum": "TL Sabit"},
        {"Kalem": "ZEMİN İYİLEŞTİRME VE İKSA İŞLERİ", "Eski %": 3.89, "Eski TL": 45708370.10, "Durum": "TL Sabit"},
        {"Kalem": "01- NOLU MUKAYESE İŞLERİ", "Eski %": 0.22, "Eski TL": 2585049.21, "Durum": "TL Sabit"},
        {"Kalem": "02- NOLU MUKAYESE İŞLERİ", "Eski %": 0.86, "Eski TL": 10105192.36, "Durum": "TL Sabit"},
        {"Kalem": "03- NOLU MUKAYESE İŞLERİ", "Eski %": 3.85, "Eski TL": 45238361.15, "Durum": "TL Sabit"},
        {"Kalem": "04- NOLU MUKAYESE İŞLERİ", "Eski %": 5.16, "Eski TL": 60631154.16, "Durum": "TL Sabit"},
        {"Kalem": "05- NOLU MUKAYESE İŞLERİ", "Eski %": 0.49, "Eski TL": 5757609.60, "Durum": "TL Sabit"},
        {"Kalem": "KAT İRTİFAKININ KURULMASI", "Eski %": 0.43, "Eski TL": 5052596.18, "Durum": "% Sabit"},
        {"Kalem": "GEÇİCİ KABULÜN ONAYLANMASI", "Eski %": 1.70, "Eski TL": 19975380.25, "Durum": "% Sabit"},
        {"Kalem": "İSKAN RAPORLARININ ALINMASI (YAPI KULLANMA İZİN BELGESİ)", "Eski %": 0.43, "Eski TL": 5052596.18, "Durum": "% Sabit"},
        {"Kalem": "KAT MÜLKİYETİNİN KURULMASI", "Eski %": 0.43, "Eski TL": 5052596.18, "Durum": "% Sabit"},
        {"Kalem": "KESİN HESABIN ONAYLANMASI", "Eski %": 0.43, "Eski TL": 5052596.18, "Durum": "% Sabit"},
        {"Kalem": "İŞ SAĞLIĞI VE GÜVENLİĞİ HİZMETİ BEDELİ", "Eski %": 0.17, "Eski TL": 1997538.02, "Durum": "% Sabit"},
        {"Kalem": "KESİN KABULÜN ONAYLANMASI", "Eski %": 0.85, "Eski TL": 9987690.12, "Durum": "% Sabit"}
    ])

def df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Pursantaj')
    return output.getvalue()

def get_next_mukayese_info(df):
    max_num = 0
    last_idx = -1
    
    for idx, row in df.iterrows():
        kalem_adi = str(row['Kalem']).upper()
        match = re.search(r'(\d+)\s*-\s*NOLU MUKAYESE', kalem_adi)
        if match:
            num = int(match.group(1))
            if num > max_num:
                max_num = num
            if idx > last_idx:
                last_idx = idx
                
    next_num = max_num + 1
    next_name = f"{next_num:02d}- NOLU MUKAYESE İŞLERİ"
    
    if last_idx == -1:
        sabit_mask = df['Durum'] == '% Sabit'
        if sabit_mask.any():
            last_idx = df[sabit_mask].index[0] - 1
        else:
            last_idx = len(df) - 1
            
    insert_idx = last_idx + 1
    return next_name, insert_idx

# ==========================================
# 1. VERİ BAŞLATMA VE YÜKLEME (DOSYA KİLİTLİ)
# ==========================================
if "data" not in st.session_state:
    st.session_state.data = get_template_df()

if "editor_key" not in st.session_state:
    st.session_state.editor_key = 0

st.sidebar.header("📂 Dosya İşlemleri")
st.sidebar.info("Proje verilerinizi Google E-Tablo (Excel) veya JSON olarak yükleyin.")

uploaded_file = st.sidebar.file_uploader("Veri Yükle (.xlsx, .json)", type=["xlsx", "json"])
if uploaded_file is not None:
    if "yuklenen_dosya_adi" not in st.session_state or st.session_state.yuklenen_dosya_adi != uploaded_file.name:
        try:
            if uploaded_file.name.endswith('.json'):
                data = json.load(uploaded_file)
                st.session_state.data = pd.DataFrame(data)
            elif uploaded_file.name.endswith('.xlsx'):
                st.session_state.data = pd.read_excel(uploaded_file)
            
            st.session_state.editor_key += 1
            st.session_state.yuklenen_dosya_adi = uploaded_file.name
            st.sidebar.success("Dosya başarıyla yüklendi!")
        except Exception as e:
            st.sidebar.error("Geçersiz dosya formatı.")

st.sidebar.divider()
st.sidebar.download_button(
    label="📄 Boş Şablon İndir (E-Tablo Uyumlu)",
    data=df_to_excel(get_template_df()),
    file_name="pursantaj_sablon.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.title("⚖️ Revize Pursantaj Dağıtımı (Sistem Uyumlu)")
st.caption("Yüzdeler üzerinden tersine çalışan hakediş ödeme mantığı koda entegre edilmiştir.")

# ==========================================
# 2. OTOMATİK FARK BULUCU VE KONTROL PANELİ
# ==========================================
aktif_eski_toplam = round(st.session_state.data["Eski TL"].sum(), 2)

col1, col2, col3 = st.columns(3)

with col1:
    yeni_toplam_tl = st.number_input(
        "Yeni Sözleşme Bedeli (TL)", 
        value=float(aktif_eski_toplam), 
        step=50000.0, 
        format="%.2f"
    )
    st.caption(f"Mevcut Tablo Toplamı: {aktif_eski_toplam:,.2f} TL")

otomatik_fark = round(yeni_toplam_tl - aktif_eski_toplam, 2)
tahmin_isim, eklenecek_index = get_next_mukayese_info(st.session_state.data)

with col2:
    girilen_kalem_adi = st.text_input("Eklenecek Kalemin Adı:", value=tahmin_isim)
    if otomatik_fark > 0:
        st.info(f"Hesaplanan Tutar: **{otomatik_fark:,.2f} TL**")
    else:
        st.warning("Artış yok (Fark: 0 TL)")

with col3:
    st.write("")
    st.write("")
    if st.button("➕ Otomatik Tutarla Listeye Ekle", use_container_width=True):
        if otomatik_fark > 0 and girilen_kalem_adi.strip() != "":
            yeni_satir = pd.DataFrame([{
                "Kalem": girilen_kalem_adi.strip(), 
                "Eski %": 0.00, 
                "Eski TL": otomatik_fark, 
                "Durum": "Serbest" 
            }])
            
            df_kopya = st.session_state.data.copy()
            non_empty_index = min(eklenecek_index, len(df_kopya))
            
            df_ust = df_kopya.iloc[:non_empty_index]
            df_alt = df_kopya.iloc[non_empty_index:]
            
            st.session_state.data = pd.concat([df_ust, yeni_satir, df_alt]).reset_index(drop=True)
            st.session_state.editor_key += 1
            st.rerun()

st.divider()

# ==========================================
# 3. VERİ GİRİŞ TABLOSU VE YEDİRME AYARI
# ==========================================
st.subheader("İş Kalemleri ve Dağıtım Kısıtları")
edited_df = st.data_editor(
    st.session_state.data,
    column_config={
        "Durum": st.column_config.SelectboxColumn("Durum", options=["Serbest", "TL Sabit", "% Sabit"], required=True),
        "Eski TL": st.column_config.NumberColumn(format="%.2f"),
        "Eski %": st.column_config.NumberColumn(format="%.2f")
    },
    use_container_width=True, 
    num_rows="dynamic",
    key=f"editor_{st.session_state.editor_key}"
)
st.session_state.data = edited_df.copy()

# Kullanıcının Kuruş Farkı Kalemini Seçmesi İçin Özel Alan
st.markdown("### ⚙️ Hesaplama Ayarları")
esnek_df = edited_df[edited_df["Durum"] != "% Sabit"]

# Eğer tabloda esnek (TL Sabit veya Serbest) kalem varsa onlardan liste yap, yoksa hepsini al
if not esnek_df.empty:
    en_buyuk_kalem_adi = esnek_df.loc[esnek_df["Eski TL"].astype(float).idxmax(), "Kalem"]
    esnek_kalemler = esnek_df["Kalem"].tolist()
    default_idx = esnek_kalemler.index(en_buyuk_kalem_adi) if en_buyuk_kalem_adi in esnek_kalemler else 0
else:
    esnek_kalemler = edited_df["Kalem"].tolist()
    default_idx = 0

secili_duzeltme_kalemi = st.selectbox(
    "Kuruş ve Yuvarlama Farklarının Yedirileceği Hedef Kalem:",
    options=esnek_kalemler,
    index=default_idx,
    help="Yüzde hesaplamalarından kaynaklanan 0.01'lik kuruş ve sapma farkları bu kaleme eklenir/çıkarılır."
)

st.divider()

# ==========================================
# 4. HESAPLAMA MOTORU (ESNEK KORUMALI)
# ==========================================
if st.button("🚀 Yeni Pursantajı Dağıt ve Ödemeyi Hesapla", type="primary"):
    df = edited_df.copy()
    
    df["Yeni TL"] = 0.00
    df["Yeni %"] = 0.00
    df["İdare Ödeme (TL)"] = 0.00
    
    # ADIM 1: TL'leri netleştir
    kullanilan_tl = 0.00
    for index, row in df.iterrows():
        durum = row["Durum"]
        if durum == "TL Sabit":
            df.at[index, "Yeni TL"] = round(float(row["Eski TL"]), 2)
            kullanilan_tl += df.at[index, "Yeni TL"]
        elif durum == "% Sabit":
            df.at[index, "Yeni TL"] = round((float(row["Eski %"]) / 100.0) * yeni_toplam_tl, 2)
            kullanilan_tl += df.at[index, "Yeni TL"]

    kalan_tl = round(yeni_toplam_tl - kullanilan_tl, 2)
    serbest_mask = df["Durum"] == "Serbest"
    serbest_eski_tl_toplam = df.loc[serbest_mask, "Eski TL"].astype(float).sum()

    if serbest_eski_tl_toplam > 0:
        for index, row in df[serbest_mask].iterrows():
            agirlik = float(row["Eski TL"]) / serbest_eski_tl_toplam
            df.at[index, "Yeni TL"] = round(kalan_tl * agirlik, 2)
    elif kalan_tl != 0:
        st.error("Havuzda dağıtılacak tutar kaldı ancak 'Serbest' kalem bulunamadı!")

    # Hedef Kalemi İndeksini Bul (Kullanıcı Seçimi)
    hedef_idx = df[df["Kalem"] == secili_duzeltme_kalemi].index[0]

    # TL Kuruş Düzeltmesi
    fark_tl = round(yeni_toplam_tl - df["Yeni TL"].sum(), 2)
    if fark_tl != 0:
        df.at[hedef_idx, "Yeni TL"] = round(df.at[hedef_idx, "Yeni TL"] + fark_tl, 2)

    # ADIM 2: Yeni Oranları (%) bul (Erken yuvarlama yasak)
    for index, row in df.iterrows():
        durum = row["Durum"]
        if durum == "% Sabit":
            df.at[index, "Yeni %"] = round(float(row["Eski %"]), 2)
        else:
            df.at[index, "Yeni %"] = round((df.at[index, "Yeni TL"] / yeni_toplam_tl) * 100, 2) if yeni_toplam_tl > 0 else 0.00

    # YÜZDE Kuruş Düzeltmesi (Kullanıcı Seçimi)
    fark_yuzde = round(100.00 - df["Yeni %"].sum(), 2)
    if fark_yuzde != 0:
        df.at[hedef_idx, "Yeni %"] = round(df.at[hedef_idx, "Yeni %"] + fark_yuzde, 2)

    # ADIM 3: İdare Ödeme Hesaplaması
    for index, row in df.iterrows():
        df.at[index, "İdare Ödeme (TL)"] = round((df.at[index, "Yeni %"] / 100.0) * yeni_toplam_tl, 2)

    # İdare Kuruş Düzeltmesi (Kullanıcı Seçimi)
    idare_fark_tl = round(yeni_toplam_tl - df["İdare Ödeme (TL)"].sum(), 2)
    if idare_fark_tl != 0:
        df.at[hedef_idx, "İdare Ödeme (TL)"] = round(df.at[hedef_idx, "İdare Ödeme (TL)"] + idare_fark_tl, 2)
        st.toast(f"Mühendislik Düzeltmesi: Toplamı eşitlemek için {idare_fark_tl:+.2f} TL seçtiğiniz kaleme ({secili_duzeltme_kalemi}) yedirildi.", icon="🛠️")

    # ==========================================
    # 5. SONUÇ EKRANI VE DIŞA AKTARIM
    # ==========================================
    st.divider()
    st.subheader("📊 Yeni Durum (Sistem Uyumlu Dağılım)")
    
    check_yuzde = round(df["Yeni %"].sum(), 2)
    check_idare_tl = round(df["İdare Ödeme (TL)"].sum(), 2)
    
    c_res1, c_res2, c_res3 = st.columns(3)
    c_res1.metric("Hedef Toplam TL", f"{yeni_toplam_tl:,.2f} TL")
    c_res2.metric("Nihai Sistem Ödemesi (TL)", f"{check_idare_tl:,.2f} TL", delta=f"{check_idare_tl - yeni_toplam_tl:,.2f} Fark")
    c_res3.metric("Dağıtılan Toplam %", f"% {check_yuzde:.2f}", delta=f"{check_yuzde - 100.00:.2f} Fark")
    
    formatli_df = df.copy()
    formatli_df["Eski TL"] = formatli_df["Eski TL"].apply(lambda x: f"{x:,.2f}")
    formatli_df["Yeni TL"] = formatli_df["Yeni TL"].apply(lambda x: f"{x:,.2f}")
    formatli_df["Yeni %"] = formatli_df["Yeni %"].apply(lambda x: f"% {x:.2f}")
    formatli_df["İdare Ödeme (TL)"] = formatli_df["İdare Ödeme (TL)"].apply(lambda x: f"{x:,.2f}")
    
    st.dataframe(formatli_df[["Kalem", "Durum", "Eski TL", "Eski %", "Yeni TL", "Yeni %", "İdare Ödeme (TL)"]], use_container_width=True)

    st.markdown("### 📥 Sonuçları Dışa Aktar")
    d_col1, d_col2 = st.columns(2)
    
    d_col1.download_button(
        label="📊 Google E-Tablo / Excel Olarak İndir",
        data=df_to_excel(df),
        file_name="revize_pursantaj_sonuc.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    json_data = df.to_json(orient="records", force_ascii=False, indent=4)
    d_col2.download_button(
        label="💾 JSON Olarak İndir",
        data=json_data,
        file_name="revize_pursantaj_sonuc.json",
        mime="application/json",
        use_container_width=True
    )
