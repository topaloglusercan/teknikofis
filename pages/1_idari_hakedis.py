"""
İdari Hakediş & Teyit Matrisi - Web Modülü (Streamlit)
Orijinal Tkinter masaüstü uygulamasının birebir web uyarlamasıdır.
- Rerun (çift tıklama/odak kaybı) sorunu çözülmüştür.
- Matematiksel motor orijinal .pyc dosyasından alınmıştır.
"""

import streamlit as st
import pandas as pd
import json
import io
import warnings
from decimal import Decimal, ROUND_HALF_UP, getcontext

# Hassasiyet ve uyarı ayarları (Orijinal koddaki gibi)
getcontext().prec = 28 
warnings.filterwarnings("ignore")

st.set_page_config(page_title="İdari Hakediş Modülü", layout="wide", page_icon="📂")

# ==========================================
# 1. MATEMATİK VE YARDIMCI FONKSİYONLAR
# ==========================================
def parse_turkish_date(date_str):
    if pd.isna(date_str) or str(date_str).strip() == '': 
        return pd.NaT
    s = str(date_str).strip().replace('.', ' ').lower()
    if s in ['none', 'nan', 'nat', '<na>']: 
        return pd.NaT
    
    # Decompile sırasındaki mask_37 hatası düzeltildi
    months = {
        'oca': '01', 'ocak': '01', 'şub': '02', 'şubat': '02', 
        'mar': '03', 'mart': '03', 'nis': '04', 'nisan': '04', 
        'may': '05', 'mayıs': '05', 'haz': '06', 'haziran': '06', 
        'tem': '07', 'temmuz': '07', 'ağu': '08', 'ağustos': '08', 
        'eyl': '09', 'eylül': '09', 'eki': '10', 'ekim': '10', 
        'kas': '11', 'kasım': '11', 'ara': '12', 'aralık': '12'
    }
    
    parts = s.split()
    if len(parts) == 2:
        m_num = months.get(parts[0], '01')
        y_num = parts[1] if len(parts[1]) == 4 else f"20{parts[1]}"
        return f"{y_num}-{m_num}"
    return pd.NaT

def clean_decimal(val):
    if val is None or pd.isna(val): return Decimal('0.0')
    s = str(val).strip().replace('TL', '').replace('%', '').strip()
    if s.lower() in ['', 'none', 'nan', 'nat', '<na>']: return Decimal('0.0')
    
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    else:
        if ',' in s:
            s = s.replace(',', '.')
        elif s.count('.') > 1:
            s = s.replace('.', '')
            
    try:
        d = Decimal(s)
        if d.is_nan(): return Decimal('0.0')
        return d
    except:
        return Decimal('0.0')

def tr_format(val):
    if pd.isna(val) or val == "": return ""
    try:
        formatted = "{:,.2f}".format(float(val))
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(val)

def filter_empty_rows(df):
    if df.empty: return df
    mask = df.iloc[:, 0].astype(str).str.strip().str.lower().isin(['', 'none', 'nan', 'nat', '<na>'])
    return df[~mask]

# ==========================================
# 2. HESAPLAMA MOTORU (Orijinal Tkinter Core)
# ==========================================
def hesapla(df_prog, df_endeks, df_alt, df_b):
    df_prog = filter_empty_rows(df_prog.copy())
    df_endeks = filter_empty_rows(df_endeks.copy())
    df_alt = filter_empty_rows(df_alt.copy())
    df_b = filter_empty_rows(df_b.copy())
    
    if df_prog.empty or df_endeks.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_prog.columns = df_prog.columns.str.strip()
    
    end_col = 'AYLAR' if 'AYLAR' in df_endeks.columns else 'Aylar'
    df_endeks['AyKodu'] = pd.to_datetime(df_endeks[end_col].apply(parse_turkish_date)).dt.to_period('M')
    df_endeks = df_endeks.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_b['AyKodu'] = pd.to_datetime(df_b['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_b = df_b.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_prog['AyKodu'] = pd.to_datetime(df_prog['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_prog = df_prog.dropna(subset=['AyKodu'])

    if df_endeks.empty or df_prog.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    son_endeks_ayi = df_endeks.index.max()
    
    katsayilar = {str(r['Ağırlık']).strip().lower(): clean_decimal(r['Katsayı']) for _, r in df_alt.iterrows()}
    temel_endeksler = {str(r['Ağırlık']).strip().lower(): clean_decimal(r['Temel Endeks']) for _, r in df_alt.iterrows()}
    
    _default_harita = {'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}
    if 'Endeks Sütunu' in df_alt.columns:
        endeks_haritasi = {str(r['Ağırlık']).strip().lower(): str(r['Endeks Sütunu']).strip() for _, r in df_alt.iterrows() if str(r.get('Endeks Sütunu', '')).strip() not in ['', 'nan', 'None']}
        if not endeks_haritasi:
            endeks_haritasi = _default_harita
    else:
        endeks_haritasi = _default_harita

    prog_kum_col = df_prog.columns[1] 
    imalat_kum_col = df_prog.columns[2] 

    kovalar = []
    onceki_kum = Decimal('0.0')
    for _, row in df_prog.iterrows():
        kum = clean_decimal(row[prog_kum_col])
        cap = kum - onceki_kum
        kovalar.append({'ay': row['AyKodu'], 'kapasite': cap if cap > Decimal('0.0') else Decimal('0.0')})
        onceki_kum = kum

    final_ff, matris = [], []
    onceki_imalat_kum = Decimal('0.0')
    kumulatif_ff = Decimal('0.0')

    for _, row in df_prog.iterrows():
        uyg_ayi = row['AyKodu']
        guncel = clean_decimal(row[imalat_kum_col])
        aylik = guncel - onceki_imalat_kum
        
        if aylik <= Decimal('0.0'):
            final_ff.append(float(kumulatif_ff))
            if guncel > Decimal('0.0'): onceki_imalat_kum = guncel
            continue
            
        b_val = df_b.loc[uyg_ayi, 'B'] if uyg_ayi in df_b.index else Decimal('1.0')
        b_kat = clean_decimal(b_val)
        if b_kat <= Decimal('0.0'): b_kat = Decimal('1.0')
        
        gercek_end_ayi = min(uyg_ayi, son_endeks_ayi)
        if gercek_end_ayi in df_endeks.index:
            endeks_uyg = df_endeks.loc[gercek_end_ayi]
        else:
            endeks_uyg = df_endeks.iloc[-1]
            
        toplam_ff_aylik = Decimal('0.0')
        kalan = aylik
        
        for kova in kovalar:
            if kalan <= Decimal('0.0'): break 
            if kova['kapasite'] > Decimal('0.0'):
                kullanilan = min(kalan, kova['kapasite'])
                gercek_prog_ayi = min(kova['ay'], son_endeks_ayi)
                gecikme = kova['ay'] < uyg_ayi
                
                if gecikme:
                    comp_ayi = min(gercek_end_ayi, gercek_prog_ayi)
                    endeks_prog = df_endeks.loc[comp_ayi] if comp_ayi in df_endeks.index else endeks_uyg
                else:
                    endeks_prog = endeks_uyg
                
                pn = Decimal('0.0')
                for k, sutun in endeks_haritasi.items():
                    e_temel = temel_endeksler.get(k, Decimal('0.0'))
                    e_uyg = clean_decimal(endeks_uyg.get(sutun, 0))
                    e_prog = clean_decimal(endeks_prog.get(sutun, 0))
                    e_gecerli = min(e_uyg, e_prog) if gecikme else e_uyg
                    katsayi = katsayilar.get(k, Decimal('0.0'))
                    
                    if e_temel > Decimal('0.0'):
                        pn += katsayi * (e_gecerli / e_temel)
                    elif katsayi > Decimal('0.0'):
                        pn += katsayi
                
                ff_dilim = (kullanilan * b_kat * (pn - Decimal('1.0'))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                
                matris.append({
                    'Hakediş Ayı': str(uyg_ayi),
                    'İş Programı (Ödenek) Ayı': str(kova['ay']),
                    'Kullanılan Tutar': float(kullanilan),
                    'Uygulanan Pn (15 Hane)': float(pn),
                    'Fiyat Farkı Tutarı': float(ff_dilim)
                })
                
                toplam_ff_aylik += ff_dilim
                kova['kapasite'] -= kullanilan
                kalan -= kullanilan
        
        kumulatif_ff += toplam_ff_aylik
        final_ff.append(float(kumulatif_ff))
        onceki_imalat_kum = guncel

    df_sonuc = df_prog.copy()
    df_sonuc['KÜMÜLATİF FİYAT FARKI'] = final_ff
    df_detay = pd.DataFrame(matris)
    
    if not df_detay.empty:
        df_pivot = df_detay.pivot_table(index='Hakediş Ayı', columns='İş Programı (Ödenek) Ayı', values='Kullanılan Tutar', aggfunc='sum', fill_value=0)
        df_pivot['HAKEDİŞ TUTARI (Toplam)'] = df_pivot.sum(axis=1)
        df_pivot.loc['ÖDENEK MİKTARI'] = df_pivot.sum()
    else: 
        df_pivot = pd.DataFrame()

    return df_sonuc, df_pivot, df_detay

# ==========================================
# 3. VERİ YÜKLEME VE HAFIZA (STATE)
# ==========================================
def load_from_excel(file):
    xls = pd.ExcelFile(file)
    dfs = {}
    sheet_map = {'IsProgrami': 'prog_df', 'Endeks': 'endeks_df', 'AltEndeks': 'alt_df', 'B': 'b_df'}
    
    for sheet in xls.sheet_names:
        if sheet in sheet_map:
            df_temp = pd.read_excel(xls, sheet_name=sheet, nrows=5)
            skip = 0
            cols = [str(c).upper() for c in df_temp.columns]
            
            if 'AYLAR' not in cols and 'AĞIRLIK' not in cols:
                for i, row in df_temp.iterrows():
                    row_vals = [str(v).upper() for v in row.values]
                    if 'AYLAR' in row_vals or 'AĞIRLIK' in row_vals:
                        skip = i + 1; break
            
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=skip)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df = df.astype(str).replace(['nan', 'NaN', 'None', '<NA>'], '')
            dfs[sheet_map[sheet]] = df
    return dfs

def init_state():
    if 'load_count' not in st.session_state: st.session_state.load_count = 0
    if 'prog_df' not in st.session_state:
        st.session_state.prog_df = pd.DataFrame({"AYLAR": ["Oca 22"], "İŞ PROGRAMI KÜMÜLATİF": ["0,00"], "İMALAT TUTARI KÜMÜLATİF": ["0,00"]})
    if 'endeks_df' not in st.session_state:
        st.session_state.endeks_df = pd.DataFrame({"AYLAR": ["Oca 22"], "I o": ["0,00"], "Ç o": ["0,00"], "D o": ["0,00"], "Y o": ["0,00"], "K o": ["0,00"], "G o": ["0,00"], "M o": ["0,00"]})
    if 'alt_df' not in st.session_state:
        st.session_state.alt_df = pd.DataFrame({"Ağırlık": ["a", "b1", "b2", "b3", "b4", "b5", "c"], "Katsayı": ["0,00"] * 7, "Temel Endeks": ["0,00"] * 7, "Endeks Sütunu": ["I o", "Ç o", "D o", "Y o", "K o", "G o", "M o"]})
    if 'b_df' not in st.session_state:
        st.session_state.b_df = pd.DataFrame({"AYLAR": ["Oca 22"], "B": ["1,00"]})

init_state()

# ==========================================
# 4. ARAYÜZ TASARIMI (UI)
# ==========================================
st.title("📂 İdari Hakediş & Teyit Matrisi")

# -- YAN MENÜ (SIDEBAR) --
st.sidebar.markdown("### 📥 Excel ile Proje Yükle")
st.sidebar.caption("Excel şablonunuzu (.xlsx) buradan yükleyin.")
uploaded_excel = st.sidebar.file_uploader("Excel Dosyası Seç (.xlsx)", type=["xlsx"], key="excel_uploader")

if uploaded_excel is not None:
    try:
        dfs = load_from_excel(uploaded_excel)
        if 'prog_df' in dfs: st.session_state.prog_df = dfs['prog_df']
        if 'endeks_df' in dfs: st.session_state.endeks_df = dfs['endeks_df']
        if 'alt_df' in dfs: st.session_state.alt_df = dfs['alt_df']
        if 'b_df' in dfs: st.session_state.b_df = dfs['b_df']
        st.session_state.load_count += 1
        st.sidebar.success("Excel başarıyla yüklendi!")
    except Exception as e:
        st.sidebar.error(f"Hata: {e}")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📥 Önceki Projeyi Yükle (.json)")
st.sidebar.caption("Projeyi Yükle (.json)")
uploaded_json = st.sidebar.file_uploader("JSON Dosyası Seç", type=["json"], key="json_uploader")

if uploaded_json is not None:
    data = json.load(uploaded_json)
    if 'prog' in data: st.session_state.prog_df = pd.DataFrame(data['prog']).astype(str)
    if 'endeks' in data: st.session_state.endeks_df = pd.DataFrame(data['endeks']).astype(str)
    if 'alt' in data: st.session_state.alt_df = pd.DataFrame(data['alt']).astype(str)
    if 'b' in data: st.session_state.b_df = pd.DataFrame(data['b']).astype(str)
    st.session_state.load_count += 1
    st.sidebar.success("Proje başarıyla yüklendi!")

# -- ANA EKRAN TABLOLARI --
suffix = st.session_state.load_count

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. İş Programı ve İmalatlar")
    # Veriyi st.session_state'den alır, tablo üzerinde yapılan değişiklikler 'edited_prog' a aktarılır.
    # Bu sayede focus kaybı/Rerun döngüsü kırılmış olur.
    edited_prog = st.data_editor(st.session_state.prog_df, num_rows="dynamic", use_container_width=True, key=f"prog_ed_{suffix}")
    
    st.subheader("3. Alt Endeks Ağırlıkları")
    edited_alt = st.data_editor(st.session_state.alt_df, num_rows="dynamic", use_container_width=True, key=f"alt_ed_{suffix}")

with col2:
    st.subheader("2. Endeks Tablosu")
    edited_endeks = st.data_editor(st.session_state.endeks_df, num_rows="dynamic", use_container_width=True, key=f"end_ed_{suffix}")
    
    st.subheader("4. B Katsayısı Tablosu")
    edited_b = st.data_editor(st.session_state.b_df, num_rows="dynamic", use_container_width=True, key=f"b_ed_{suffix}")

# -- İNDİRME VE KAYDETME BUTONLARI --
st.sidebar.markdown("---")
st.sidebar.markdown("### 📤 Dışa Aktar")
project_data = {
    'prog': edited_prog.to_dict(orient='records'),
    'endeks': edited_endeks.to_dict(orient='records'),
    'alt': edited_alt.to_dict(orient='records'),
    'b': edited_b.to_dict(orient='records')
}
st.sidebar.download_button(
    label="💾 JSON Olarak Kaydet",
    data=json.dumps(project_data, indent=4),
    file_name="hakedis_projesi.json",
    mime="application/json",
    use_container_width=True
)

st.markdown("---")

# -- HESAPLAMA BUTONU --
if st.button("🚀 Hesapla ve Matrisi Çıkar", use_container_width=True, type="primary"):
    with st.spinner("Matematiksel motor çalışıyor..."):
        df_sonuc, df_pivot, df_detay = hesapla(edited_prog, edited_endeks, edited_alt, edited_b)
    
    if df_sonuc.empty:
        st.warning("⚠️ Lütfen tablolara geçerli veri giriniz (İş Programı ve Endeks boş olamaz).")
    else:
        st.success("✅ Hesaplama Başarılı!")
        
        tab1, tab2, tab3 = st.tabs(["🔍 Dilim Bazlı Detay", "📊 Teyit Matrisi", "📑 Kümülatif Sonuç"])
        
        with tab1:
            df_detay_gosterim = df_detay.copy()
            df_detay_gosterim['Kullanılan Tutar'] = df_detay_gosterim['Kullanılan Tutar'].apply(tr_format)
            df_detay_gosterim['Fiyat Farkı Tutarı'] = df_detay_gosterim['Fiyat Farkı Tutarı'].apply(tr_format)
            df_detay_gosterim['Uygulanan Pn (15 Hane)'] = df_detay_gosterim['Uygulanan Pn (15 Hane)'].apply(lambda x: "{:.15f}".format(x).rstrip('0').rstrip('.').replace('.', ','))
            st.dataframe(df_detay_gosterim, use_container_width=True)
            
        with tab2:
            if not df_pivot.empty:
                df_pivot_tr = df_pivot.map(tr_format)
                st.dataframe(df_pivot_tr.style.set_properties(subset=['HAKEDİŞ TUTARI (Toplam)'], **{'font-weight': 'bold', 'background-color': '#e6f2ff'}), use_container_width=True)
            else:
                st.info("Teyit matrisi oluşturulacak yeterli hakediş verisi bulunamadı.")
                
        with tab3:
            for col in df_sonuc.columns:
                if any(x in col.upper() for x in ['TUTAR', 'PROGRAM', 'FARKI']):
                    df_sonuc[col] = df_sonuc[col].apply(tr_format)
            st.dataframe(df_sonuc, use_container_width=True)
