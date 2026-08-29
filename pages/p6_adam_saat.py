import streamlit as st
import pandas as pd
import numpy as np
import re

st.markdown("<h1 style='color: #2c3e50;'>👷‍♂️ P6 Adam-Saat Analizi</h1>", unsafe_allow_html=True)
st.markdown("XER veritabanını tarayarak iş kalemlerinin kaynak dağılımlarını şelale yöntemiyle hesaplayın.")

def clean_dec_float(val):
    if pd.isna(val) or val is None: return 0.0
    s = str(val).strip().upper()
    s = re.sub(r'[^\d.,-]', '', s)
    if not s or s in ['-', '.', ',']: return 0.0
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
        else: s = s.replace(',', '')
    elif ',' in s:
        parts = s.split(',')
        if len(parts[-1]) <= 2: s = s.replace(',', '.')
        else: s = s.replace(',', '')
    try: return float(s)
    except: return 0.0

@st.cache_data
def parse_xer(file_bytes):
    text = file_bytes.decode('windows-1254', errors='ignore')
    lines = text.splitlines()
    tables = {}
    current_table = None
    columns = []
    for line in lines:
        if line.startswith('%T'):
            current_table = line.split('\t')[1].strip()
            tables[current_table] = []
        elif line.startswith('%F') and current_table:
            columns = [col.strip().lower() for col in line.split('\t')[1:]]
        elif line.startswith('%R') and current_table:
            values = line.split('\t')[1:]
            row_dict = {columns[i]: values[i].strip() if i < len(values) else "" for i in range(len(columns))}
            tables[current_table].append(row_dict)
    def clean_df(data_list):
        df = pd.DataFrame(data_list)
        if not df.empty:
            for col in df.columns: df[col] = df[col].astype(str).str.strip()
        return df
    return clean_df(tables.get('TASK', [])), clean_df(tables.get('TASKRSRC', [])), clean_df(tables.get('RSRC', [])), clean_df(tables.get('ACTVTYPE', [])), clean_df(tables.get('ACTVCODE', [])), clean_df(tables.get('TASKACTV', []))

@st.cache_data
def prepare_xer_data(df_task, df_taskrsrc, df_rsrc, df_actvtype, df_actvcode, df_taskactv, grup_secimi, code_type_id=None):
    if df_task.empty or df_taskrsrc.empty: return pd.DataFrame()
    df_t = df_task.drop_duplicates(subset=['task_id']).copy()
    df_t['task_id'] = df_t['task_id'].astype(str)
    df_tr = df_taskrsrc.copy()
    df_tr['task_id'] = df_tr['task_id'].astype(str)
    if 'rsrc_id' in df_tr.columns: df_tr['rsrc_id'] = df_tr['rsrc_id'].astype(str)
    
    df_merged = pd.merge(df_tr, df_t, on='task_id', how='inner')
    
    if not df_rsrc.empty and 'rsrc_id' in df_merged.columns:
        df_r = df_rsrc.drop_duplicates(subset=['rsrc_id']).copy()
        df_r['rsrc_id'] = df_r['rsrc_id'].astype(str)
        df_merged = pd.merge(df_merged, df_r, on='rsrc_id', how='left')
        if 'rsrc_short_name' not in df_merged.columns: df_merged['rsrc_short_name'] = 'EK-AS'
    else: df_merged['rsrc_short_name'] = 'EK-AS'
        
    if grup_secimi == "Sadece Kaynak Adına Göre":
        df_merged['Grup_Adi'] = df_merged.get('rsrc_short_name', 'EK-AS')
    else:
        if not df_actvcode.empty and not df_taskactv.empty and code_type_id:
            hedef_kodlar = df_actvcode[df_actvcode['actv_code_type_id'].astype(str) == str(code_type_id)].copy()
            desc_col = 'actv_code_name' if 'actv_code_name' in hedef_kodlar.columns else 'short_name' if 'short_name' in hedef_kodlar.columns else 'actv_code_id'
            df_ta = df_taskactv.copy()
            df_ta['actv_code_id'] = df_ta['actv_code_id'].astype(str)
            df_ta['task_id'] = df_ta['task_id'].astype(str)
            hedef_kodlar['actv_code_id'] = hedef_kodlar['actv_code_id'].astype(str)
            baglanti_df = pd.merge(df_ta, hedef_kodlar, on='actv_code_id', how='inner').drop_duplicates(subset=['task_id']) 
            df_merged = pd.merge(df_merged, baglanti_df[['task_id', desc_col]], on='task_id', how='left')
            df_merged['Grup_Adi'] = df_merged[desc_col].fillna("Atanmamış İşler")
        else: df_merged['Grup_Adi'] = "Atanmamış İşler"

    df_merged['Grup_Adi'] = df_merged['Grup_Adi'].replace(['nan', 'None', '', '<NA>'], 'Atanmamış İşler')
    df_merged['Gecerli_Baslangic'] = pd.NaT
    df_merged['Gecerli_Bitis'] = pd.NaT
    
    for col in ['target_start_date', 'early_start_date', 'act_start_date']:
        if col in df_merged.columns: df_merged['Gecerli_Baslangic'] = df_merged['Gecerli_Baslangic'].fillna(pd.to_datetime(df_merged[col], errors='coerce'))
    for col in ['target_end_date', 'early_end_date', 'act_end_date']:
        if col in df_merged.columns: df_merged['Gecerli_Bitis'] = df_merged['Gecerli_Bitis'].fillna(pd.to_datetime(df_merged[col], errors='coerce'))
            
    for col in ['target_qty', 'remain_qty', 'act_qty']:
        if col in df_merged.columns:
            df_merged[col] = df_merged[col].apply(clean_dec_float)
            df_merged[col] = np.where(df_merged[col] > 50000, 0.0, df_merged[col])
        else: df_merged[col] = 0.0

    beklenen = ['task_code', 'task_name', 'rsrc_short_name', 'Gecerli_Baslangic', 'Gecerli_Bitis', 'Grup_Adi', 'target_qty', 'remain_qty', 'act_qty']
    for col in beklenen:
        if col not in df_merged.columns: df_merged[col] = None

    return df_merged[beklenen].copy()

def spread_to_months(df_filtered, qty_column):
    df_valid = df_filtered.dropna(subset=['Gecerli_Baslangic', 'Gecerli_Bitis']).copy()
    df_valid = df_valid[df_valid[qty_column] > 0]
    if df_valid.empty: return pd.DataFrame()
    df_valid['toplam_gun'] = (df_valid['Gecerli_Bitis'] - df_valid['Gecerli_Baslangic']).dt.days + 1
    df_valid['toplam_gun'] = df_valid['toplam_gun'].clip(lower=1, upper=3650).astype(int) 
    df_valid['Gunluk_AS'] = df_valid[qty_column] / df_valid['toplam_gun']
    df_spread = df_valid.loc[df_valid.index.repeat(df_valid['toplam_gun'])].copy()
    days_to_add = pd.to_timedelta(df_spread.groupby(level=0).cumcount(), unit='D')
    df_spread['Tarih'] = df_spread['Gecerli_Baslangic'] + days_to_add
    df_spread['Ay'] = df_spread['Tarih'].dt.to_period('M')
    aylik_pivot = df_spread.groupby(['Ay', 'Grup_Adi'])['Gunluk_AS'].sum().reset_index()
    aylik_pivot.rename(columns={'Ay': 'Planlanan Ay', 'Gunluk_AS': 'Toplam Adam-Saat', 'Grup_Adi': 'Ekip/Kod'}, inplace=True)
    aylik_pivot['Planlanan Ay'] = aylik_pivot['Planlanan Ay'].astype(str)
    return aylik_pivot

uploaded_xer = st.file_uploader("📂 XER Dosyasını Yükle", type=['xer'])
if uploaded_xer:
    with st.spinner("P6 Veritabanı okunuyor..."):
        df_task, df_taskrsrc, df_rsrc, df_actvtype, df_actvcode, df_taskactv = parse_xer(uploaded_xer.getvalue())
            
    kod_secenekleri = ["Sadece Kaynak Adına Göre"]
    kod_sozlugu = {}
    if not df_actvtype.empty:
        for _, row in df_actvtype.iterrows():
            gosterim_adi = f"Aktivite Kodu: {row.get('actv_code_type', 'Bilinmeyen')}"
            kod_secenekleri.append(gosterim_adi)
            kod_sozlugu[gosterim_adi] = str(row.get('actv_code_type_id', '')).strip()
            
    col_ayarlar1, col_ayarlar2 = st.columns(2)
    grup_tercihi = col_ayarlar1.selectbox("📌 1. Veriler Neye Göre Gruplansın?", kod_secenekleri)
    aylik_kapasite = col_ayarlar2.number_input("1 İşçinin Aylık Kapasitesi (Saat)", value=208)

    secilen_kod_id = kod_sozlugu.get(grup_tercihi, None)
    
    with st.spinner("Kayıtlar eşleştiriliyor..."):
        ham_veri_df = prepare_xer_data(df_task, df_taskrsrc, df_rsrc, df_actvtype, df_actvcode, df_taskactv, grup_tercihi, secilen_kod_id)
    
    if not ham_veri_df.empty:
        projedeki_kaynaklar = sorted([str(k) for k in ham_veri_df['rsrc_short_name'].dropna().unique() if k != 'Tanımsız Kaynak'])
        col_k1, col_k2 = st.columns(2)
        secilen_kaynak = col_k1.multiselect("📌 2. Hangi Kaynaklar Analiz Edilsin?", projedeki_kaynaklar, default=projedeki_kaynaklar)
        qty_tipi = col_k2.radio("📌 3. Hangi Değer Tipi Hesaplansın?", ["Bütçelenen (Planlanan/Budgeted)", "Kalan (Remaining)", "Gerçekleşen (Actual)"])
        qty_sutunu = 'target_qty' if 'Bütçelenen' in qty_tipi else 'remain_qty' if 'Kalan' in qty_tipi else 'act_qty'

        if secilen_kaynak: ham_veri_df = ham_veri_df[ham_veri_df['rsrc_short_name'].isin(secilen_kaynak)]
        tum_gruplar = sorted([str(g) for g in ham_veri_df['Grup_Adi'].dropna().unique()])
        
        st.markdown("### 🔍 4. Kalem Seçimi (Filtre)")
        arama_kelimesi = st.text_input("Kelime ile Ara (Örn: seramik, duvar)")
        gosterilecek_gruplar = [g for g in tum_gruplar if arama_kelimesi.lower() in str(g).lower()] if arama_kelimesi else tum_gruplar
        tumunu_sec = st.checkbox(f"Arama sonucundaki {len(gosterilecek_gruplar)} kalemin TÜMÜNÜ SEÇ", value=False)
        varsayilan_secim = gosterilecek_gruplar if tumunu_sec else (gosterilecek_gruplar[:10] if not arama_kelimesi and len(gosterilecek_gruplar) > 10 else gosterilecek_gruplar)
        
        secilen_gruplar = st.multiselect("Analiz Edilecek Kalemler:", tum_gruplar, default=varsayilan_secim)
        
        if secilen_gruplar:
            filtrelenmis_df = ham_veri_df[ham_veri_df['Grup_Adi'].isin(secilen_gruplar)]
            toplam_as = filtrelenmis_df[qty_sutunu].sum()
            st.info(f"💡 Seçtiğiniz Kalemlerdeki Toplam **{qty_tipi.split(' ')[0]}** Adam-Saat: **{toplam_as:,.1f}**")
            
            if st.button("📊 Seçili Kalemleri Aylara Dağıt", type="primary"):
                with st.spinner("Şelale Analizi yapılıyor..."):
                    aylik_pivot = spread_to_months(filtrelenmis_df, qty_sutunu)
                    if aylik_pivot.empty: st.warning("Tarih bulunamadı.")
                    else:
                        aylik_pivot['Gerekli Ekip (Kişi)'] = np.ceil(aylik_pivot['Toplam Adam-Saat'] / aylik_kapasite).astype(int)
                        aylik_pivot['Toplam Adam-Saat'] = aylik_pivot['Toplam Adam-Saat'].apply(lambda x: f"{x:,.1f}")
                        c1, c2 = st.columns([1, 2])
                        with c1: st.dataframe(aylik_pivot, use_container_width=True)
                        with c2:
                            st.subheader("📈 İhtiyaç Yığılma Grafiği")
                            grafik_df = aylik_pivot.copy()
                            grafik_df['Gerekli Ekip (Kişi)'] = pd.to_numeric(grafik_df['Gerekli Ekip (Kişi)'])
                            st.bar_chart(grafik_df.pivot(index='Planlanan Ay', columns='Ekip/Kod', values='Gerekli Ekip (Kişi)').fillna(0))
