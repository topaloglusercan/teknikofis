import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

st.markdown("<h1 style='color: #2c3e50;'>📊 P6 Aktivite Kodu Analizi (Dashboard)</h1>", unsafe_allow_html=True)
st.markdown("Projenizdeki Adam-Saat veya miktarları, belirlediğiniz Kaynaklara ve Aktivite Kodlarına (Disiplin, Faz vb.) göre filtreleyip gruplayın.")

def clean_dec_float(val):
    try:
        if pd.isna(val) or val is None: return 0.0
        s = str(val).strip()
        if not s: return 0.0
        if '.' in s and ',' in s:
            if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
            else: s = s.replace(',', '')
        elif ',' in s: s = s.replace(',', '.')
        return float(s)
    except:
        return 0.0

def format_b(x, p=None):
    if x >= 1e9: return f"{x/1e9:.2f} Milyar"
    elif x >= 1e6: return f"{x/1e6:.2f} Milyon"
    elif x >= 1e3: return f"{x/1e3:.0f} Bin"
    else: return f"{int(x)}"

@st.cache_data
def parse_xer_for_activity_codes(file_bytes):
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

    return (
        clean_df(tables.get('TASK', [])), 
        clean_df(tables.get('TASKRSRC', [])),
        clean_df(tables.get('ACTVTYPE', [])), 
        clean_df(tables.get('ACTVCODE', [])), 
        clean_df(tables.get('TASKACTV', [])),
        clean_df(tables.get('RSRC', [])) # Kaynak isimleri için RSRC eklendi
    )

uploaded_xer = st.file_uploader("📂 XER Dosyasını Yükle (Aktivite Kodu Analizi İçin)", type=['xer'])

if uploaded_xer:
    with st.spinner("P6 Veritabanı, Kodlama ve Kaynak Yapısı Taranıyor..."):
        df_task, df_tr, df_at, df_ac, df_ta, df_rsrc = parse_xer_for_activity_codes(uploaded_xer.getvalue())
        
        if df_task.empty or df_tr.empty:
            st.error("⚠️ XER dosyasında Aktivite (TASK) veya Kaynak (TASKRSRC) verisi bulunamadı.")
        elif df_at.empty or df_ac.empty or df_ta.empty:
            st.warning("⚠️ Bu projede herhangi bir Aktivite Kodu (Activity Code) hiyerarşisi tanımlanmamış.")
        else:
            # 1. Kaynak Tablosunu Temizle ve İsimlendir
            df_tr_clean = df_tr.loc[:, ~df_tr.columns.duplicated()].copy()
            cost_col = 'target_qty' if 'target_qty' in df_tr_clean.columns else 'target_cost'
            
            if cost_col not in df_tr_clean.columns:
                st.error("Bütçe miktarı (Budgeted Units / Cost) sütunu bulunamadı.")
            else:
                df_tr_clean['task_id'] = df_tr_clean['task_id'].astype(str)
                df_tr_clean['Birim_Miktar'] = df_tr_clean[cost_col].apply(clean_dec_float)
                
                # Kaynak İsimlerini Eşleştir
                if 'rsrc_id' in df_tr_clean.columns:
                    df_tr_clean['rsrc_id'] = df_tr_clean['rsrc_id'].astype(str)
                    if not df_rsrc.empty and 'rsrc_id' in df_rsrc.columns:
                        df_r = df_rsrc.copy()
                        df_r = df_r.loc[:, ~df_r.columns.duplicated()].drop_duplicates(subset=['rsrc_id'])
                        df_r['rsrc_id'] = df_r['rsrc_id'].astype(str)
                        
                        name_col = next((c for c in ['rsrc_name', 'rsrc_short_name'] if c in df_r.columns), None)
                        if name_col:
                            isim_sozlugu = df_r.set_index('rsrc_id')[name_col].to_dict()
                            df_tr_clean['Kaynak_Adi'] = df_tr_clean['rsrc_id'].map(isim_sozlugu).fillna('İsimsiz').astype(str)
                            df_tr_clean['Kaynak_Gorunum'] = df_tr_clean['rsrc_id'] + " - " + df_tr_clean['Kaynak_Adi']
                        else:
                            df_tr_clean['Kaynak_Gorunum'] = df_tr_clean['rsrc_id']
                    else:
                        df_tr_clean['Kaynak_Gorunum'] = df_tr_clean['rsrc_id']
                else:
                    df_tr_clean['Kaynak_Gorunum'] = 'Tüm Kaynaklar'

                # 2. Aktivite Kodu İlişkilerini Kur (Merge)
                df_ta['task_id'] = df_ta['task_id'].astype(str)
                df_ta['actv_code_id'] = df_ta['actv_code_id'].astype(str)
                df_ac['actv_code_id'] = df_ac['actv_code_id'].astype(str)
                df_ac['actv_code_type_id'] = df_ac['actv_code_type_id'].astype(str)
                df_at['actv_code_type_id'] = df_at['actv_code_type_id'].astype(str)

                df_codes = pd.merge(df_ta[['task_id', 'actv_code_id']], df_ac[['actv_code_id', 'actv_code_type_id', 'short_name']], on='actv_code_id', how='inner')
                df_codes = pd.merge(df_codes, df_at[['actv_code_type_id', 'actv_code_type']], on='actv_code_type_id', how='inner')

                # 3. Kullanıcı Arayüzü (Kategori ve Kaynak Seçici)
                kategori_listesi = sorted(df_codes['actv_code_type'].unique())
                kaynak_listesi = sorted([str(x) for x in df_tr_clean['Kaynak_Gorunum'].unique()])
                
                if not kategori_listesi:
                    st.warning("Aktivitelere atanmış hiçbir kod bulunamadı.")
                else:
                    st.markdown("---")
                    c1, c2 = st.columns(2)
                    with c1:
                        secilen_kategori = st.selectbox("📌 1. Kırılım Kategorisini Seçin:", kategori_listesi)
                    with c2:
                        secilen_kaynaklar = st.multiselect("🔍 2. Analiz Edilecek Kaynakları Filtreleyin (Boşsa Tümü):", kaynak_listesi, default=[])
                    
                    # 4. Seçimlere Göre Veriyi Filtrele ve Hesapla
                    if secilen_kaynaklar:
                        df_tr_filt = df_tr_clean[df_tr_clean['Kaynak_Gorunum'].isin(secilen_kaynaklar)].copy()
                    else:
                        df_tr_filt = df_tr_clean.copy()

                    df_task_totals = df_tr_filt.groupby('task_id')['Birim_Miktar'].sum().reset_index()
                    toplam_proje_miktari = df_task_totals['Birim_Miktar'].sum()

                    df_secilen_kodlar = df_codes[df_codes['actv_code_type'] == secilen_kategori][['task_id', 'short_name']]
                    df_dashboard = pd.merge(df_task_totals, df_secilen_kodlar, on='task_id', how='left')
                    df_dashboard['short_name'] = df_dashboard['short_name'].fillna('Atanmamış (Tanımsız)')
                    
                    df_summary = df_dashboard.groupby('short_name')['Birim_Miktar'].sum().reset_index()
                    df_summary = df_summary.sort_values(by='Birim_Miktar', ascending=True)
                    df_summary = df_summary[df_summary['Birim_Miktar'] > 0]

                    # 5. Özet Metrikler
                    st.markdown("---")
                    c_met1, c_met2 = st.columns(2)
                    with c_met1:
                        st.metric("Toplam Budgeted Units (Seçilen Kaynaklar)", f"{toplam_proje_miktari:,.2f}")
                    with c_met2:
                        st.metric(f"'{secilen_kategori}' Kırılım Sayısı", len(df_summary))

                    # 6. Görselleştirme (Yatay Bar Chart)
                    st.markdown(f"### 📊 {secilen_kategori} Dağılım Grafiği")
                    plt.style.use('ggplot')
                    fig, ax = plt.subplots(figsize=(12, max(6, len(df_summary) * 0.5)))
                    
                    bars = ax.barh(df_summary['short_name'], df_summary['Birim_Miktar'], color='#4C72B0', edgecolor='black', alpha=0.8)
                    ax.set_xlabel('Budgeted Units (Seçilen Kaynaklar)', fontweight='bold')
                    ax.xaxis.set_major_formatter(FuncFormatter(format_b))
                    
                    # Sayısal değerleri yazdır
                    for bar in bars:
                        width = bar.get_width()
                        ax.annotate(f"{width:,.0f}",
                                    xy=(width, bar.get_y() + bar.get_height() / 2),
                                    xytext=(5, 0),
                                    textcoords="offset points",
                                    ha='left', va='center', fontweight='bold', fontsize=10)
                        
                    plt.tight_layout()
                    st.pyplot(fig)

                    # 7. Veri Tablosu
                    st.markdown("### 📋 Detaylı Veri Tablosu")
                    df_rapor = df_summary.sort_values(by='Birim_Miktar', ascending=False).copy()
                    df_rapor['Yüzde (%)'] = (df_rapor['Birim_Miktar'] / toplam_proje_miktari * 100).round(2) if toplam_proje_miktari > 0 else 0
                    df_rapor.rename(columns={'short_name': secilen_kategori, 'Birim_Miktar': 'Budgeted Units'}, inplace=True)
                    
                    st.dataframe(df_rapor, use_container_width=True)
