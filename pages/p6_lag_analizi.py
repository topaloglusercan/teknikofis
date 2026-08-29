import streamlit as st
import pandas as pd
import re

st.markdown("<h1 style='color: #2c3e50;'>⏳ P6 Lag (Bekleme Süresi) Analizi</h1>", unsafe_allow_html=True)
st.markdown("XER dosyanızı yükleyerek aktiviteler arasına gizlenmiş bekleme sürelerini anında listeleyin.")

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
def parse_xer_lag(file_bytes):
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

    return clean_df(tables.get('TASK', [])), clean_df(tables.get('TASKPRED', []))

uploaded_xer_lag = st.file_uploader("📂 XER Dosyasını Yükle (Lag Analizi İçin)", type=['xer'])

if uploaded_xer_lag:
    with st.spinner("İlişki Ağları (Network) Taranıyor..."):
        df_task, df_taskpred = parse_xer_lag(uploaded_xer_lag.getvalue())
        
    if not df_taskpred.empty and not df_task.empty:
        df_rels = df_taskpred.copy()
        df_tasks_dict = df_task[['task_id', 'task_code', 'task_name']].copy()
        
        lag_col = 'lag_hr_cnt' if 'lag_hr_cnt' in df_rels.columns else 'lag_cnt'
        if lag_col not in df_rels.columns: df_rels[lag_col] = 0.0
            
        df_rels[lag_col] = df_rels[lag_col].apply(clean_dec_float)
        lagged_rels = df_rels[df_rels[lag_col] != 0].copy()
        
        if lagged_rels.empty:
            st.success("✅ Bu projede hiç Lag (Bekleme Süresi) içeren ilişki bulunamadı.")
        else:
            p6_mesai_saati = st.number_input("P6 Takvimi Günlük Mesai (Saat)", value=8.0)
            
            rel_map = {'PR_FS': 'FS (Bitiş-Başla)', 'PR_SS': 'SS (Başla-Başla)', 'PR_FF': 'FF (Bitiş-Bitiş)', 'PR_SF': 'SF (Başla-Bitiş)'}
            lagged_rels['İlişki Tipi'] = lagged_rels['pred_type'].map(rel_map).fillna(lagged_rels.get('pred_type', 'Bilinmeyen'))
            
            merged = pd.merge(lagged_rels, df_tasks_dict, on='task_id', how='left')
            merged.rename(columns={'task_code': 'Ardıl ID', 'task_name': 'Ardıl Aktivite Adı'}, inplace=True)
            
            merged = pd.merge(merged, df_tasks_dict, left_on='pred_task_id', right_on='task_id', how='left', suffixes=('', '_pred_drop'))
            merged.rename(columns={'task_code': 'Öncül ID', 'task_name': 'Öncül Aktivite Adı'}, inplace=True)
            
            merged['Lag (Saat)'] = merged[lag_col]
            merged['Lag (Gün)'] = (merged[lag_col] / p6_mesai_saati).round(1)
            
            sonuc_tablosu = merged[['Öncül ID', 'Öncül Aktivite Adı', 'İlişki Tipi', 'Lag (Gün)', 'Lag (Saat)', 'Ardıl ID', 'Ardıl Aktivite Adı']]
            
            c_m1, c_m2 = st.columns(2)
            c_m1.info(f"📌 Lag İçeren Toplam İlişki Sayısı: **{len(sonuc_tablosu)}**")
            c_m2.error(f"⚠️ Tespit Edilen Maksimum Lag: **{sonuc_tablosu['Lag (Gün)'].max()} Gün**")

            st.dataframe(sonuc_tablosu, use_container_width=True)
            csv_lag = sonuc_tablosu.to_csv(index=False).encode('utf-8')
            st.download_button("💾 Raporu CSV İndir", data=csv_lag, file_name="P6_Lag_Analizi.csv", mime="text/csv")
    else:
        st.error("XER dosyasında İlişki (TASKPRED) tablosu bulunamadı.")