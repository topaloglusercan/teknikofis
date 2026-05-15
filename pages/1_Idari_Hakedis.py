import streamlit as st
import pandas as pd
import warnings
import json
from decimal import Decimal, ROUND_HALF_UP

warnings.filterwarnings("ignore")

st.set_page_config(page_title="İdari Hakediş Modülü", layout="wide", page_icon="📂")

# --- YARDIMCI FONKSİYONLAR ---
def parse_turkish_date(date_str):
    if pd.isna(date_str) or str(date_str).strip() == '': return pd.NaT
    date_str = str(date_str).strip().replace('.', ' ').lower()
    months = {'oca': '01', 'ocak': '01', 'şub': '02', 'şubat': '02', 'mar': '03', 'mart': '03',
              'nis': '04', 'nisan': '04', 'may': '05', 'mayıs': '05', 'haz': '06', 'haziran': '06',
              'tem': '07', 'temmuz': '07', 'ağu': '08', 'ağustos': '08', 'eyl': '09', 'eylül': '09',
              'eki': '10', 'ekim': '10', 'kas': '11', 'kasım': '11', 'ara': '12', 'aralık': '12'}
    parts = date_str.split()
    if len(parts) == 2:
        m_num = months.get(parts[0], '01')
        y_num = parts[1] if len(parts[1]) == 4 else f"20{parts[1]}"
        return f"{y_num}-{m_num}"
    return pd.NaT

def clean_decimal(val):
    if pd.isna(val) or str(val).strip() == '': return Decimal('0.0')
    try:
        clean_str = str(val).replace('.', '').replace(',', '.')
        return Decimal(clean_str)
    except:
        return Decimal('0.0')

def tr_format(val):
    if pd.isna(val) or val == "": return ""
    try:
        formatted = "{:,.2f}".format(float(val))
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return val

# --- ANA HESAPLAMA MOTORU ---
def hesapla(df_prog, df_endeks, df_alt, df_b):
    df_prog.columns = df_prog.columns.str.strip()
    
    end_col = 'AYLAR' if 'AYLAR' in df_endeks.columns else 'Aylar'
    df_endeks['AyKodu'] = pd.to_datetime(df_endeks[end_col].apply(parse_turkish_date)).dt.to_period('M')
    df_endeks = df_endeks.drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_b['AyKodu'] = pd.to_datetime(df_b['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_b = df_b.drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_prog['AyKodu'] = pd.to_datetime(df_prog['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')

    son_endeks_ayi = df_endeks.index.max()
    
    katsayilar = {str(row['Ağırlık']).strip().lower(): clean_decimal(row['Katsayı']) for _, row in df_alt.iterrows()}
    temel_endeksler = {str(row['Ağırlık']).strip().lower(): clean_decimal(row['Temel Endeks']) for _, row in df_alt.iterrows()}
    endeks_haritasi = {'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}

    prog_kum_col = df_prog.columns[1] 
    imalat_kum_col = df_prog.columns[2] 

    kovalar = []
    onceki_kum = Decimal('0.0')
    for _, row in df_prog.iterrows():
        kum = clean_decimal(row[prog_kum_col])
        kapasite = kum - onceki_kum
        kovalar.append({'ay': row['AyKodu'], 'kapasite': kapasite})
        onceki_kum = kum

    final_ff_listesi, matris_verileri = [], []
    onceki_imalat_kum, kümülatif_toplam_ff = Decimal('0.0'), Decimal('0.0')

    for _, row in df_prog.iterrows():
        uyg_ayi = row['AyKodu']
        guncel_imalat_kum = clean_decimal(row[imalat_kum_col])
        aylik_imalat = guncel_imalat_kum - onceki_imalat_kum
        
        if aylik_imalat <= Decimal('0.0'):
            final_ff_listesi.append(float(kümülatif_toplam_ff) if guncel_imalat_kum > Decimal('0.0') else 0.0)
            if guncel_imalat_kum > Decimal('0.0'): onceki_imalat_kum = guncel_imalat_kum
            continue
            
        b_val = df_b.loc[uyg_ayi, 'B'] if uyg_ayi in df_b.index else Decimal('1.0')
        b_kat = clean_decimal(b_val) if clean_decimal(b_val) > Decimal('0.0') else Decimal('1.0')
        
        gercek_endeks_ayi = min(uyg_ayi, son_endeks_ayi)
        endeks_uyg = df_endeks.loc[gercek_endeks_ayi] if gercek_endeks_ayi in df_endeks.index else pd.Series(dtype=float)
            
        toplam_ff_aylik, kalan_para = Decimal('0.0'), aylik_imalat
        
        for kova in kovalar:
            if kalan_para <= Decimal('0.0'): break 
            if kova['kapasite'] > Decimal('0.0'):
                kullanilan_tutar = min(kalan_para, kova['kapasite'])
                matris_verileri.append({
                    'Hakediş Ayı': str(uyg_ayi),
                    'İş Programı (Ödenek) Ayı': str(kova['ay']),
                    'Kullanılan Tutar': float(kullanilan_tutar)
                })
                
                gercek_prog_ayi = min(kova['ay'], son_endeks_ayi)
                endeks_prog = df_endeks.loc[gercek_prog_ayi] if kova['ay'] < uyg_ayi and gercek_prog_ayi in df_endeks.index else endeks_uyg
                
                pn = Decimal('0.0')
                for k, sutun in endeks_haritasi.items():
                    e_temel = temel_endeksler.get(k, Decimal('0.0'))
                    e_gecerli = min(clean_decimal(endeks_uyg.get(sutun, 0)), clean_decimal(endeks_prog.get(sutun, 0))) if kova['ay'] < uyg_ayi else clean_decimal(endeks_uyg.get(sutun, 0))
                    if e_temel > Decimal('0.0'):
                        pn += katsayilar.get(k, Decimal('0.0')) * (e_gecerli / e_temel)
                
                ff_dilim_yuvarlanmis = (kullanilan_tutar * b_kat * (pn - Decimal('1.0'))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                toplam_ff_aylik += ff_dilim_yuvarlanmis
                kova['kapasite'] -= kullanilan_tutar
                kalan_para -= kullanilan_tutar
        
        kümülatif_toplam_ff += toplam_ff_aylik
        final_ff_listesi.append(float(kümülatif_toplam_ff))
        onceki_imalat_kum = guncel_imalat_kum

    df_sonuc = df_prog.copy()
    df_sonuc['KÜMÜLATİF FİYAT FARKI'] = final_ff_listesi
    df_matris_ham = pd.DataFrame(matris_verileri)
    if not df_matris_ham.empty:
        df_pivot = df_matris_ham.pivot_table(index='Hakediş Ayı', columns='İş Programı (Ödenek) Ayı', values='Kullanılan Tutar', aggfunc='sum', fill_value=0)
        df_pivot['HAKEDİŞ TUTARI (Toplam)'] = df_pivot.sum(axis=1)
        df_pivot.loc['ÖDENEK MİKTARI (Kullanılan Toplam)'] = df_pivot.sum()
    else: df_pivot = pd.DataFrame()

    return df_sonuc, df_pivot

# --- ARAYÜZ (UI) ---
if 'prog_df' not in st.session_state:
    st.session_state.prog_df = pd.DataFrame({"AYLAR": ["Oca 22"], "İŞ PROGRAMI KÜMÜLATİF": [""], "İMALAT TUTARI KÜMÜLATİF": [""]})
    st.session_state.endeks_df = pd.DataFrame({"AYLAR": ["Oca 22"], "I o": [""], "Ç o": [""], "D o": [""], "Y o": [""], "K o": [""], "G o": [""], "M o": [""]})
    st.session_state.alt_df = pd.DataFrame({"Ağırlık": ["a", "b1", "b2", "b3", "b4", "b5", "c"], "Katsayı": ["0,00", "0,00", "0,00", "0,00", "0,00", "0,00", "0,00"], "Temel Endeks": ["", "", "", "", "", "", ""]})
    st.session_state.b_df = pd.DataFrame({"AYLAR": ["Oca 22"], "B": ["1,00"]})

st.title("📂 İdari Hakediş & Teyit Matrisi")

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Hakediş Projesi Yönetimi")
uploaded_file = st.sidebar.file_uploader("Önceki Projeyi Yükle (.json)", type=["json"])

if uploaded_file is not None:
    data = json.load(uploaded_file)
    st.session_state.prog_df = pd.DataFrame(data['prog'])
    st.session_state.endeks_df = pd.DataFrame(data['endeks'])
    st.session_state.alt_df = pd.DataFrame(data['alt'])
    st.session_state.b_df = pd.DataFrame(data['b'])
    st.sidebar.success("Proje başarıyla yüklendi!")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. İş Programı ve İmalatlar")
    edited_prog = st.data_editor(st.session_state.prog_df, num_rows="dynamic", use_container_width=True)
    st.subheader("3. Alt Endeks Ağırlıkları")
    edited_alt = st.data_editor(st.session_state.alt_df, num_rows="dynamic", use_container_width=True)
with col2:
    st.subheader("2. Endeks Tablosu")
    edited_endeks = st.data_editor(st.session_state.endeks_df, num_rows="dynamic", use_container_width=True)
    st.subheader("4. B Katsayısı Tablosu")
    edited_b = st.data_editor(st.session_state.b_df, num_rows="dynamic", use_container_width=True)

project_data = {
    'prog': edited_prog.to_dict(orient='records'),
    'endeks': edited_endeks.to_dict(orient='records'),
    'alt': edited_alt.to_dict(orient='records'),
    'b': edited_b.to_dict(orient='records')
}
st.sidebar.download_button(
    label="💾 Mevcut Veriyi Bilgisayara İndir",
    data=json.dumps(project_data, indent=4),
    file_name="hakedis_projem.json",
    mime="application/json",
    use_container_width=True
)

st.markdown("---")
if st.button("🚀 Hesapla ve Matrisi Çıkar", use_container_width=True):
    try:
        p = edited_prog[edited_prog.iloc[:,0].astype(str).str.strip() != '']
        e = edited_endeks[edited_endeks.iloc[:,0].astype(str).str.strip() != '']
        a = edited_alt[edited_alt.iloc[:,0].astype(str).str.strip() != '']
        b = edited_b[edited_b.iloc[:,0].astype(str).str.strip() != '']
        
        df_sonuc, df_pivot = hesapla(p, e, a, b)
        
        st.subheader("📊 Ödenek Dilimlerinin Hakedişlere Göre Kullanılması (Teyit Matrisi)")
        df_pivot_tr = df_pivot.map(tr_format)
        st.dataframe(df_pivot_tr.style.set_properties(subset=['HAKEDİŞ TUTARI (Toplam)'], **{'font-weight': 'bold', 'background-color': '#e6f2ff'}), use_container_width=True)
        
        st.subheader("📑 Kümülatif Fiyat Farkı Sonuçları")
        for col in df_sonuc.columns:
            if any(x in col.upper() for x in ['TUTAR', 'PROGRAM', 'FARKI']):
                df_sonuc[col] = df_sonuc[col].apply(tr_format)
        st.dataframe(df_sonuc, use_container_width=True)
    except Exception as ex:
        st.error(f"🚨 Hata: {ex}")