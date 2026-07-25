"""
Hakediş Fiyat Farkı Simülatörü
================================
Mevcut "İdari Hakediş" motorunun BİREBİR AYNI hesaplama çekirdeği
üzerine kurulu; slider'larla anlık simülasyon ve senaryo karşılaştırma
katmanı eklenmiştir.

Çalıştırma:
    pip install streamlit pandas openpyxl plotly
    streamlit run hakedis_simulator.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import math, re, json, warnings, copy, io
from decimal import Decimal, ROUND_HALF_UP, getcontext
import plotly.graph_objects as go

getcontext().prec = 28
warnings.filterwarnings("ignore")

try:
    st.set_page_config(page_title="Hakediş Fiyat Farkı Simülatörü", page_icon="🏗️", layout="wide")
except Exception:
    pass

# ==========================================
# --- YARDIMCI FONKSİYONLAR ---
# ==========================================
def parse_turkish_date(date_str):
    if pd.isna(date_str) or str(date_str).strip() == '': return pd.NaT
    date_str = str(date_str).strip().replace('.', ' ').lower()
    if date_str in ['none', 'nan', 'nat', '<na>']: return pd.NaT
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
    if pd.isna(val): return Decimal('0.0')
    val_str = str(val).strip()
    if val_str.lower() in ['', 'none', 'nan', 'nat', '<na>']: return Decimal('0.0')
    
    val_str = val_str.replace('TL', '').replace('%', '').strip()
    
    if '.' in val_str and ',' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    else:
        if ',' in val_str:
            val_str = val_str.replace(',', '.')
        elif val_str.count('.') > 1:
            val_str = val_str.replace('.', '')
            
    try:
        d = Decimal(val_str)
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
        return val

def filter_empty_rows(df):
    if df.empty: return df
    mask = df.iloc[:, 0].astype(str).str.strip().str.lower().isin(['', 'none', 'nan', 'nat', '<na>'])
    return df[~mask]

# ==========================================
# --- EXCEL / JSON YÜKLEME ---
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
            
            if 'AYLAR' in cols or 'AĞIRLIK' in cols:
                skip = 0
            else:
                for i, row in df_temp.iterrows():
                    row_vals = [str(v).upper() for v in row.values]
                    if 'AYLAR' in row_vals or 'AĞIRLIK' in row_vals:
                        skip = i + 1
                        break
            
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=skip)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df = df.astype(str).replace(['nan', 'NaN', 'None', '<NA>'], '')
            dfs[sheet_map[sheet]] = df
    return dfs

def generate_excel_download(df_prog, df_endeks, df_alt, df_b):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_prog.to_excel(writer, sheet_name='IsProgrami', index=False)
        df_endeks.to_excel(writer, sheet_name='Endeks', index=False)
        df_alt.to_excel(writer, sheet_name='AltEndeks', index=False)
        df_b.to_excel(writer, sheet_name='B', index=False)
    return output.getvalue()

# Arayüz etiketleri için (Simülatör Gelişmiş Seçenekleri)
KOD_BILGI = {
    'a':  {'kolon': 'I o', 'resmi_kod': 'İn',    'kisa': 'İşçilik',      'ad': 'İşçilik (TÜFE bağlı)'},
    'b1': {'kolon': 'Ç o', 'resmi_kod': 'Çn-23', 'kisa': 'Çimento/Mineral', 'ad': 'Metalik Olmayan Diğer Mineral Ürünler (Çimento vb.)'},
    'b2': {'kolon': 'D o', 'resmi_kod': 'Dn-24', 'kisa': 'Demir-Çelik',  'ad': 'Ana Metaller (Demir-Çelik)'},
    'b3': {'kolon': 'Y o', 'resmi_kod': 'Ayn',   'kisa': 'Akaryakıt',    'ad': 'Akaryakıt Ürünleri'},
    'b4': {'kolon': 'K o', 'resmi_kod': 'Kn-16', 'kisa': 'Ağaç/Mantar',  'ad': 'Ağaç ve Mantar Ürünleri (mobilya hariç)'},
    'b5': {'kolon': 'G o', 'resmi_kod': 'Gn',    'kisa': 'Genel ÜFE',    'ad': 'Genel Yurt İçi ÜFE (Yİ-ÜFE)'},
    'c':  {'kolon': 'M o', 'resmi_kod': 'Mn-28', 'kisa': 'Makine/Ekipman', 'ad': 'Makine ve Ekipmanlar b.y.s.'},
}
KOD_ETIKET = {k: v['kisa'] for k, v in KOD_BILGI.items()}

# ==========================================
# --- ANA HESAPLAMA MOTORU (idari_hakedis.py ile BİREBİR AYNI) ---
# ==========================================
def hesapla(df_prog, df_endeks, df_alt, df_b):
    df_prog = filter_empty_rows(df_prog.copy())
    df_endeks = filter_empty_rows(df_endeks.copy())
    df_alt = filter_empty_rows(df_alt.copy())
    df_b = filter_empty_rows(df_b.copy())
    
    if df_prog.empty or df_endeks.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_prog.columns = df_prog.columns.str.strip()
    
    end_col = 'AYLAR' if 'AYLAR' in df_endeks.columns else 'Aylar'
    df_endeks['AyKodu'] = pd.to_datetime(df_endeks[end_col].apply(parse_turkish_date)).dt.to_period('M')
    df_endeks = df_endeks.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_b['AyKodu'] = pd.to_datetime(df_b['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_b = df_b.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_prog['AyKodu'] = pd.to_datetime(df_prog['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_prog = df_prog.dropna(subset=['AyKodu'])

    if df_endeks.empty or df_prog.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

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
        capacity = kum - onceki_kum
        kovalar.append({'ay': row['AyKodu'], 'kapasite': capacity if capacity > Decimal('0.0') else Decimal('0.0'), 'orig': capacity if capacity > Decimal('0.0') else Decimal('0.0')})
        onceki_kum = kum

    final_ff_listesi, matris_verileri, aylik_rows = [], [], []
    onceki_imalat_kum, kümülatif_toplam_ff = Decimal('0.0'), Decimal('0.0')

    for _, row in df_prog.iterrows():
        uyg_ayi = row['AyKodu']
        guncel_imalat_kum = clean_decimal(row[imalat_kum_col])
        aylik_imalat = guncel_imalat_kum - onceki_imalat_kum
        
        if aylik_imalat <= Decimal('0.0'):
            final_ff_listesi.append(float(kümülatif_toplam_ff))
            if guncel_imalat_kum > Decimal('0.0'): onceki_imalat_kum = guncel_imalat_kum
            continue
            
        b_val = df_b.loc[uyg_ayi, 'B'] if uyg_ayi in df_b.index else Decimal('1.0')
        b_kat = clean_decimal(b_val) if clean_decimal(b_val) > Decimal('0.0') else Decimal('1.0')
        
        gercek_endeks_ayi = min(uyg_ayi, son_endeks_ayi)
        if gercek_endeks_ayi in df_endeks.index:
            endeks_uyg = df_endeks.loc[gercek_endeks_ayi]
        else:
            endeks_uyg = df_endeks.iloc[-1]
            
        toplam_ff_aylik, kalan_para = Decimal('0.0'), aylik_imalat
        
        for kova in kovalar:
            if kalan_para <= Decimal('0.0'): break 
            if kova['kapasite'] > Decimal('0.0'):
                kullanilan_tutar = min(kalan_para, kova['kapasite'])
                
                gercek_prog_ayi = min(kova['ay'], son_endeks_ayi)
                gecikme = kova['ay'] < uyg_ayi
                
                if gecikme:
                    comp_ayi = min(gercek_endeks_ayi, gercek_prog_ayi)
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
                
                ff_dilim = kullanilan_tutar * b_kat * (pn - Decimal('1.0'))
                ff_dilim_yuvarlanmis = ff_dilim.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                
                matris_verileri.append({
                    'Hakediş Ayı': str(uyg_ayi),
                    'İş Programı (Ödenek) Ayı': str(kova['ay']),
                    'Kullanılan Tutar': float(kullanilan_tutar),
                    'Uygulanan Pn (Excel - 15 Hane)': float(pn),
                    'Fiyat Farkı Tutarı': float(ff_dilim_yuvarlanmis)
                })
                
                toplam_ff_aylik += ff_dilim_yuvarlanmis
                kova['kapasite'] -= kullanilan_tutar
                kalan_para -= kullanilan_tutar
        
        kümülatif_toplam_ff += toplam_ff_aylik
        final_ff_listesi.append(float(kümülatif_toplam_ff))
        
        # Grafik ve Simülatör Analizleri için aylık özet
        aylik_rows.append({'Ay': str(uyg_ayi), 'Aylık İmalat': float(aylik_imalat),
                            'B Katsayısı': float(b_kat),
                            'Aylık Fiyat Farkı': float(toplam_ff_aylik),
                            'Kümülatif Fiyat Farkı': float(kümülatif_toplam_ff)})
                            
        onceki_imalat_kum = guncel_imalat_kum

    df_sonuc = df_prog.copy()
    df_sonuc['KÜMÜLATİF FİYAT FARKI'] = final_ff_listesi
    df_detay = pd.DataFrame(matris_verileri)
    
    if not df_detay.empty:
        df_pivot = df_detay.pivot_table(index='Hakediş Ayı', columns='İş Programı (Ödenek) Ayı', values='Kullanılan Tutar', aggfunc='sum', fill_value=0)
        df_pivot['HAKEDİŞ TUTARI (Toplam)'] = df_pivot.sum(axis=1)
        df_pivot.loc['ÖDENEK MİKTARI (Kullanılan Toplam)'] = df_pivot.sum()
    else: 
        df_pivot = pd.DataFrame()

    df_aylik = pd.DataFrame(aylik_rows)
    return df_sonuc, df_pivot, df_detay, df_aylik

# ==========================================
# --- SİMÜLASYON DÖNÜŞÜMLERİ ---
# ==========================================
def endeks_uzat(df_end, artis, ek_ay=36):
    df = df_end.copy()
    end_col = 'AYLAR' if 'AYLAR' in df.columns else 'Aylar'
    df['_ay'] = pd.to_datetime(df[end_col].apply(parse_turkish_date)).dt.to_period('M')
    df = df.sort_values('_ay').reset_index(drop=True)
    son = df.iloc[-1]
    son_ay = df['_ay'].iloc[-1]
    cols = [c for c in df.columns if c not in (end_col, '_ay')]

    if isinstance(artis, dict):
        artis_map = {c: artis.get(c, 0.0) for c in cols}
    else:
        artis_map = {c: artis for c in cols}

    yeni = []
    for k in range(1, ek_ay + 1):
        yeni_ay = son_ay + k
        satir = {end_col: str(yeni_ay)}
        for c in cols:
            val_str = str(son[c]).replace(',', '.')
            satir[c] = f"{(float(val_str) * ((1 + artis_map[c] / 100) ** k)):.6f}".replace('.', ',')
        yeni.append(satir)

    df_ek = pd.DataFrame(yeni)
    return pd.concat([df.drop(columns='_ay'), df_ek], ignore_index=True)

def imalat_donustur(df_prog, hiz_carpani=1.0, gecikme_ay=0, tek_ay_index=None, tek_ay_kaydirma=0):
    df = df_prog.copy()
    imal_col = df.columns[2]
    kum = [clean_decimal(v) for v in df[imal_col]]
    n = len(kum)

    aylik = [kum[0]] + [kum[i] - kum[i - 1] for i in range(1, n)]
    aylik = [a * Decimal(str(hiz_carpani)) for a in aylik]

    if tek_ay_index is not None:
        j = tek_ay_index + tek_ay_kaydirma
        j = max(0, min(n - 1, j))
        deger = aylik[tek_ay_index]
        aylik[tek_ay_index] = Decimal('0.0')
        aylik[j] += deger
    elif gecikme_ay != 0:
        shifted = [Decimal('0.0')] * n
        for i, a in enumerate(aylik):
            j = i + gecikme_ay
            j = max(0, min(n - 1, j))
            shifted[j] += a
        aylik = shifted

    yeni_kum, running = [], Decimal('0.0')
    for a in aylik:
        running += a
        yeni_kum.append(str(running).replace('.', ','))

    df[imal_col] = yeni_kum
    return df

def b_override_uygula(df_b, b_deger):
    df = df_b.copy()
    df['B'] = str(b_deger).replace('.', ',')
    return df

def katsayi_override_uygula(df_alt, katsayi_dict):
    df = df_alt.copy()
    for i, row in df.iterrows():
        kod = str(row['Ağırlık']).strip().lower()
        if kod in katsayi_dict:
            df.at[i, 'Katsayı'] = str(katsayi_dict[kod]).replace('.', ',')
    return df

def senaryo_calistir(df_prog, df_end, df_alt, df_b, gecikme_ay, hiz_carpani, endeks_artis,
                      b_deger, tek_ay_index=None, tek_ay_kaydirma=0, katsayi_override=None):
    p2 = imalat_donustur(df_prog, hiz_carpani=hiz_carpani, gecikme_ay=gecikme_ay,
                          tek_ay_index=tek_ay_index, tek_ay_kaydirma=tek_ay_kaydirma)
    e2 = endeks_uzat(df_end, endeks_artis)
    b2 = b_override_uygula(df_b, b_deger) if b_deger is not None else df_b
    a2 = katsayi_override_uygula(df_alt, katsayi_override) if katsayi_override else df_alt
    return hesapla(p2, e2, a2, b2)


# ==========================================
# --- SESSION STATE VE ÖRNEK VERİ ---
# ==========================================
if 'load_count' not in st.session_state:
    st.session_state.load_count = 0
if 'prog_df' not in st.session_state:
    st.session_state.prog_df = pd.DataFrame({"AYLAR": ["Oca 22"], "İŞ PROGRAMI KÜMÜLATİF": ["0,00"], "İMALAT TUTARI KÜMÜLATİF": ["0,00"]})
if 'endeks_df' not in st.session_state:
    st.session_state.endeks_df = pd.DataFrame({"AYLAR": ["Oca 22"], "I o": ["0,00"], "Ç o": ["0,00"], "D o": ["0,00"], "Y o": ["0,00"], "K o": ["0,00"], "G o": ["0,00"], "M o": ["0,00"]})
if 'alt_df' not in st.session_state:
    st.session_state.alt_df = pd.DataFrame({"Ağırlık": ["a", "b1", "b2", "b3", "b4", "b5", "c"], "Katsayı": ["0,00", "0,00", "0,00", "0,00", "0,00", "0,00", "0,00"], "Temel Endeks": ["0,00", "0,00", "0,00", "0,00", "0,00", "0,00", "0,00"]})
if 'b_df' not in st.session_state:
    st.session_state.b_df = pd.DataFrame({"AYLAR": ["Oca 22"], "B": ["1,00"]})
if "senaryolar" not in st.session_state:
    st.session_state.senaryolar = {}

st.title("🏗️ Hakediş Fiyat Farkı Simülatörü")
st.caption("Kova sistemi · Gecikme matrisi · Slider'lı anlık simülasyon · Senaryo karşılaştırma")

tab1, tab2, tab3, tab4 = st.tabs(["📋 Veri Girişi", "📊 Baz Sonuç", "🎛️ Simülatör", "⚖️ Senaryo Karşılaştır"])

# ══════════════════════ TAB 1 — VERİ GİRİŞİ ══════════════════════
with tab1:
    st.info("Kendi projenizin verilerini buraya girin/yapıştırın, veya JSON/Excel dosyanızı sol menüden içe aktarın.")

    st.sidebar.markdown("---")
    st.sidebar.subheader("📥 Proje Verisi Yükle")
    st.sidebar.caption("JSON projenizi veya Excel şablonunuzu buradan yükleyin.")
    uploaded_file = st.sidebar.file_uploader("Dosya Seç (.json veya .xlsx)", type=["json", "xlsx"])

    if uploaded_file is not None:
        if uploaded_file.name.endswith('.json'):
            data = json.load(uploaded_file)
            st.session_state.prog_df = pd.DataFrame(data['prog']).map(str)
            st.session_state.endeks_df = pd.DataFrame(data['endeks']).map(str)
            st.session_state.alt_df = pd.DataFrame(data['alt']).map(str)
            st.session_state.b_df = pd.DataFrame(data['b']).map(str)
            st.session_state.load_count += 1
            st.sidebar.success("JSON projesi başarıyla yüklendi!")
        elif uploaded_file.name.endswith('.xlsx'):
            try:
                dfs = load_from_excel(uploaded_file)
                if 'prog_df' in dfs: st.session_state.prog_df = dfs['prog_df']
                if 'endeks_df' in dfs: st.session_state.endeks_df = dfs['endeks_df']
                if 'alt_df' in dfs: st.session_state.alt_df = dfs['alt_df']
                if 'b_df' in dfs: st.session_state.b_df = dfs['b_df']
                st.session_state.load_count += 1
                st.sidebar.success("Excel şablonu başarıyla okundu!")
            except Exception as e:
                st.sidebar.error(f"Excel okuma hatası. Detay: {e}")

    suffix = st.session_state.load_count

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1️⃣ İş Programı ve İmalatlar")
        ep = st.data_editor(st.session_state.prog_df, num_rows="dynamic", use_container_width=True, key=f"prog_ed_{suffix}")
        st.session_state.prog_edited = ep

        st.divider()
        st.subheader("3️⃣ Alt Endeks Ağırlıkları")
        ea = st.data_editor(st.session_state.alt_df, num_rows="dynamic", use_container_width=True, key=f"alt_ed_{suffix}")
        st.session_state.alt_edited = ea

    with c2:
        st.subheader("2️⃣ Aylık Endeks Tablosu")
        ee = st.data_editor(st.session_state.endeks_df, num_rows="dynamic", use_container_width=True, key=f"end_ed_{suffix}")
        st.session_state.end_edited = ee

        st.divider()
        st.subheader("4️⃣ B Katsayısı")
        eb = st.data_editor(st.session_state.b_df, num_rows="dynamic", use_container_width=True, key=f"b_ed_{suffix}")
        st.session_state.b_df_edited = eb

    project_data = {
        'prog': ep.to_dict(orient='records'),
        'endeks': ee.to_dict(orient='records'),
        'alt': ea.to_dict(orient='records'),
        'b': eb.to_dict(orient='records')
    }
    excel_data = generate_excel_download(ep, ee, ea, eb)

    st.sidebar.markdown("---")
    st.sidebar.subheader("📤 Mevcut Veriyi Dışa Aktar")
    st.sidebar.download_button("💾 JSON Olarak Kaydet", data=json.dumps(project_data, indent=4), file_name="hakedis_projem.json", mime="application/json", use_container_width=True)
    st.sidebar.download_button("📊 Excel Şablonu İndir", data=excel_data, file_name="idari_hakedis_sablonu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)


# ══════════════════════ TAB 2 — BAZ SONUÇ ══════════════════════
with tab2:
    try:
        df_sonuc, df_pivot, df_detay, df_aylik = hesapla(
            st.session_state.get("prog_edited", st.session_state.prog_df), 
            st.session_state.get("end_edited", st.session_state.endeks_df), 
            st.session_state.get("alt_edited", st.session_state.alt_df), 
            st.session_state.get("b_df_edited", st.session_state.b_df)
        )
        if not df_aylik.empty:
            toplam_ff = df_aylik['Aylık Fiyat Farkı'].sum()
            st.metric("Toplam Fiyat Farkı (Baz Senaryo)", f"{tr_format(toplam_ff)} TL")

            st.subheader("Aylık / Kümülatif Fiyat Farkı")
            st.dataframe(df_aylik.style.format({
                'Aylık İmalat': lambda x: tr_format(x), 'B Katsayısı': '{:.4f}',
                'Aylık Fiyat Farkı': lambda x: tr_format(x), 'Kümülatif Fiyat Farkı': lambda x: tr_format(x),
            }), use_container_width=True)

            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df_aylik['Ay'], y=df_aylik['Kümülatif Fiyat Farkı'], mode='lines+markers', name='Kümülatif FF'))
            fig.update_layout(title="Kümülatif Fiyat Farkı Gelişimi", xaxis_title="Ay", yaxis_title="TL", height=350)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Lütfen Tab 1'den veri giriniz.")
    except Exception as e:
        st.error(f"🚨 Hata: Lütfen giriş verilerinin eksiksiz olduğundan emin olun.")
        import traceback
        with st.expander("Teknik hata detayı"):
            st.code(traceback.format_exc())

# ══════════════════════ TAB 3 — SİMÜLATÖR ══════════════════════
with tab3:
    st.markdown("Sliderları hareket ettirin — sonuçlar **anlık** yeniden hesaplanır.")

    mod = st.radio("Kaydırma Modu", ["Genel (tüm aylar)", "Belirli Ay"], horizontal=True)

    prog_aktif = st.session_state.get("prog_edited", st.session_state.prog_df)
    ay_listesi = list(prog_aktif[prog_aktif.columns[0]])
    tek_ay_index, tek_ay_kaydirma, gecikme_ay = None, 0, 0

    s1, s2, s3, s4 = st.columns(4)
    with s1:
        if mod == "Genel (tüm aylar)":
            gecikme_ay = st.slider("Gecikme / Hızlanma (ay)", -6, 6, 0)
        else:
            secim_ay = st.selectbox("Hangi ay kaydırılsın?", ay_listesi)
            tek_ay_index = ay_listesi.index(secim_ay)
            tek_ay_kaydirma = st.slider(f"'{secim_ay}' kaç ay kaydırılsın", -6, 6, 0)
    with s2:
        hiz_carpani = st.slider("Aylık İmalat Hızı Çarpanı", 0.3, 2.0, 1.0, 0.05)
    with s3:
        endeks_artis_genel = st.slider("Gelecek Aylar İçin Endeks Artışı (%/ay)", 0.0, 10.0, 0.0, 0.1)
    with s4:
        b_override_aktif = st.checkbox("B Katsayısını Simüle Et")
        if b_override_aktif:
            b_ovr = st.slider("Sabit B Katsayısı (Tüm aylar)", 0.0, 2.0, 1.0, 0.01)
        else:
            b_ovr = None

    alt_bazinda = st.checkbox("Alt endeks bazında ayrı artış oranı ayarla (gelişmiş)")
    endeks_artis = endeks_artis_genel
    if alt_bazinda:
        with st.expander("Alt Endeks Bazlı Artış Oranları (%/ay)", expanded=True):
            ac = st.columns(7)
            endeks_artis_dict = {}
            for i, kod in enumerate(KOD_ETIKET):
                bilgi = KOD_BILGI[kod]
                kol = bilgi['kolon']
                with ac[i]:
                    endeks_artis_dict[kol] = st.slider(bilgi['kisa'], 0.0, 10.0, endeks_artis_genel, 0.1, key=f"artis_{kol}")
        endeks_artis = endeks_artis_dict

    katsayi_bazinda = st.checkbox("Alt endeks ağırlıklarını (Katsayı) simüle et (gelişmiş)")
    katsayi_override = None
    if katsayi_bazinda:
        alt_aktif = st.session_state.get("alt_edited", st.session_state.alt_df)
        orijinal_kat = {str(r['Ağırlık']).strip().lower(): float(clean_decimal(r['Katsayı'])) for _, r in alt_aktif.iterrows()}

        if st.session_state.get("kat_pending") is not None:
            for kod, val in st.session_state["kat_pending"].items():
                st.session_state[f"kat_{kod}"] = val
            st.session_state["kat_pending"] = None

        for kod in KOD_ETIKET:
            key = f"kat_{kod}"
            if key not in st.session_state:
                st.session_state[key] = orijinal_kat.get(kod, 0.0)

        with st.expander("Alt Endeks Ağırlıkları (Katsayı) — Toplamı 1.000 Olmalı", expanded=True):
            pcols = st.columns(7)
            for i, kod in enumerate(KOD_ETIKET):
                with pcols[i]:
                    bilgi = KOD_BILGI[kod]
                    st.slider(KOD_ETIKET[kod], 0.0, 1.0, step=0.01, key=f"kat_{kod}",
                               help=f"Resmi kod: {bilgi['resmi_kod']} — {bilgi['ad']}")

            toplam_kat = sum(st.session_state[f"kat_{kod}"] for kod in KOD_ETIKET)
            if abs(toplam_kat - 1.0) < 0.001:
                st.success(f"✅ Toplam: {toplam_kat:.3f} (Doğru)")
            else:
                st.warning(f"⚠️ Toplam: {toplam_kat:.3f} — 1.000 olmalı! (Sonuçlar yine hesaplanır ama gerçek dışı olur)")

            pb1, pb2, pb3, pb4, pb5 = st.columns(5)
            with pb1:
                if st.button("Sadece İşçilik"):
                    st.session_state["kat_pending"] = {kod: (1.0 if kod == 'a' else 0.0) for kod in KOD_ETIKET}
                    st.rerun()
            with pb2:
                if st.button("Sadece ÜFE"):
                    st.session_state["kat_pending"] = {kod: (1.0 if kod == 'b5' else 0.0) for kod in KOD_ETIKET}
                    st.rerun()
            with pb3:
                if st.button("50 İşçilik 50 ÜFE"):
                    st.session_state["kat_pending"] = {kod: (0.5 if kod in ('a', 'b5') else 0.0) for kod in KOD_ETIKET}
                    st.rerun()
            with pb4:
                if st.button("Normalize Et"):
                    s = sum(st.session_state[f"kat_{kod}"] for kod in KOD_ETIKET)
                    if s > 0:
                        st.session_state["kat_pending"] = {kod: st.session_state[f"kat_{kod}"] / s for kod in KOD_ETIKET}
                        st.rerun()
            with pb5:
                if st.button("Orijinale Dön"):
                    st.session_state["kat_pending"] = {kod: orijinal_kat.get(kod, 0.0) for kod in KOD_ETIKET}
                    st.rerun()

        katsayi_override = {kod: st.session_state[f"kat_{kod}"] for kod in KOD_ETIKET}

    try:
        end_aktif = st.session_state.get("end_edited", st.session_state.endeks_df)
        alt_aktif = st.session_state.get("alt_edited", st.session_state.alt_df)
        b_aktif = st.session_state.get("b_df_edited", st.session_state.b_df)

        base_sonuc, base_pivot, _, base_aylik = hesapla(prog_aktif, end_aktif, alt_aktif, b_aktif)
        sim_sonuc, sim_pivot, _, sim_aylik = senaryo_calistir(
            prog_aktif, end_aktif, alt_aktif, b_aktif,
            gecikme_ay, hiz_carpani, endeks_artis, b_ovr,
            tek_ay_index=tek_ay_index, tek_ay_kaydirma=tek_ay_kaydirma,
            katsayi_override=katsayi_override
        )

        toplam_base = base_aylik['Aylık Fiyat Farkı'].sum() if not base_aylik.empty else 0
        toplam_sim = sim_aylik['Aylık Fiyat Farkı'].sum() if not sim_aylik.empty else 0
        fark = toplam_sim - toplam_base
        fark_pct = (fark / toplam_base * 100) if toplam_base != 0 else 0

        m1, m2, m3 = st.columns(3)
        m1.metric("Baz Senaryo Toplam FF", f"{tr_format(toplam_base)} TL")
        m2.metric("Simülasyon Toplam FF", f"{tr_format(toplam_sim)} TL", delta=f"{tr_format(fark)} TL")
        m3.metric("Fark (%)", f"{fark_pct:+.1f}%")

        fig2 = go.Figure()
        if not base_aylik.empty:
            fig2.add_trace(go.Scatter(x=base_aylik['Ay'], y=base_aylik['Kümülatif Fiyat Farkı'], mode='lines+markers', name='Baz Senaryo', line=dict(dash='dot')))
        if not sim_aylik.empty:
            fig2.add_trace(go.Scatter(x=sim_aylik['Ay'], y=sim_aylik['Kümülatif Fiyat Farkı'], mode='lines+markers', name='Simülasyon'))
        fig2.update_layout(title="Kümülatif Fiyat Farkı — Baz vs Simülasyon", xaxis_title="Ay", yaxis_title="TL", height=380)
        st.plotly_chart(fig2, use_container_width=True)

        # ════════════════════════════════════════════════════════════════
        # YENİ EKLENEN BÖLÜM: TEORİK KIYASLAMA (İş Programına Tam Uyum Senaryosu)
        # ════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.subheader("💡 Teorik Kıyaslama: İş Programına Tam Uyum Senaryosunun Getirisi")
        st.info("Eğer imalatlar, İş Programı ile **birebir aynı tutarda ve zamanda** gerçekleşseydi, orijinal şartlara göre seçtiğiniz *yeni senaryonun* (örneğin İşçiliğin 100 olması) size ekstra getirisi ne olurdu?")

        # 1. Hacim Sınırını Belirliyoruz (Tablo 1'deki orijinal ilerleyişten)
        imal_col = prog_aktif.columns[2]
        prog_col = prog_aktif.columns[1]
        if not prog_aktif.empty:
            toplam_sim_imalat = max([clean_decimal(val) for val in prog_aktif[imal_col]])
        else:
            toplam_sim_imalat = Decimal('0.0')

        # 2. Kusursuz İş Programı (Teorik) Verisini Hazırlıyoruz
        df_teorik = prog_aktif.copy()
        yeni_imalat_kum = []
        for planlanan in df_teorik[prog_col]:
            p_val = clean_decimal(planlanan)
            if p_val > toplam_sim_imalat:
                yeni_imalat_kum.append(f"{float(toplam_sim_imalat):.2f}".replace('.', ','))
            else:
                yeni_imalat_kum.append(f"{float(p_val):.2f}".replace('.', ','))
        df_teorik[imal_col] = yeni_imalat_kum
        
        # 3. Teorik BAZ Durumu (Hiçbir slidera dokunmadan önceki standart halin)
        _, _, _, teorik_aylik_baz = hesapla(df_teorik, end_aktif, alt_aktif, b_aktif)
        toplam_teorik_baz = teorik_aylik_baz['Aylık Fiyat Farkı'].sum() if not teorik_aylik_baz.empty else 0
        
        # 4. Teorik SENARYO Durumu (Senin slider kısıtların uygulandığında)
        e_sim = endeks_uzat(end_aktif, endeks_artis)
        b_sim = b_override_uygula(b_aktif, b_ovr) if b_ovr is not None else b_aktif
        a_sim = katsayi_override_uygula(alt_aktif, katsayi_override) if katsayi_override else alt_aktif
        
        _, _, _, teorik_aylik_senaryo = hesapla(df_teorik, e_sim, a_sim, b_sim)
        toplam_teorik_senaryo = teorik_aylik_senaryo['Aylık Fiyat Farkı'].sum() if not teorik_aylik_senaryo.empty else 0
        
        # 5. Getiri Farkı
        fark_teorik = float(toplam_teorik_senaryo) - float(toplam_teorik_baz)
        
        # 6. Sonuçları Ekrana Basıyoruz
        t_col1, t_col2, t_col3 = st.columns(3)
        t_col1.metric("Teorik FF (Orijinal Şartlar)", f"{tr_format(toplam_teorik_baz)} TL")
        t_col2.metric("Teorik FF (Yeni Senaryonuz)", f"{tr_format(toplam_teorik_senaryo)} TL")
        
        if fark_teorik > 0:
            t_col3.metric("Fark (Yeni Senaryonun Getirisi)", f"+{tr_format(fark_teorik)} TL", delta_color="normal")
            st.success("✅ **Analiz:** Seçtiğiniz yeni ekonomik şartlar, iş programına tam uyduğunuz kusursuz bir senaryoda size DAHA FAZLA fiyat farkı getirisi sağlıyor.")
        elif fark_teorik < 0:
            t_col3.metric("Fark (Yeni Senaryonun Getirisi)", f"{tr_format(fark_teorik)} TL", delta_color="inverse")
            st.error("⚠️ **Analiz:** Seçtiğiniz yeni ekonomik şartlar, iş programına tam uyduğunuz senaryoda fiyat farkı getirisini DÜŞÜRÜYOR.")
        # Eski Hali:
        # else:
        #    t_col3.metric("Fark", "0,00 TL")
        #    st.info("Senaryonuz, iş programına tam uyumlu kusursuz durumda herhangi bir getiri farkı yaratmadı.")

        # Yeni Hali:
        else:
            t_col3.metric("Fark", "0,00 TL")
            st.caption("Orijinal sözleşme katsayıları devrede.")

        st.divider()
        senaryo_adi = st.text_input("Bu ayarları senaryo olarak kaydet:", placeholder="örn. 'Senaryo A'")
        if st.button("💾 Senaryoyu Kaydet"):
            st.session_state.senaryolar[senaryo_adi.strip()] = {
                'mod': mod, 'gecikme_ay': gecikme_ay, 'toplam_ff': float(toplam_sim), 'aylik': sim_aylik.to_dict('records'),
            }
            st.success("Kaydedildi. Karşılaştırmak için 'Senaryo Karşılaştır' sekmesine geçin.")

    except Exception as e:
        st.error(f"🚨 Hata oluştu. Detay: {e}")

# ══════════════════════ TAB 4 — SENARYO KARŞILAŞTIR ══════════════════════
with tab4:
    if not st.session_state.senaryolar:
        st.info("Henüz kaydedilmiş senaryo yok.")
    else:
        secilenler = st.multiselect("Karşılaştırılacak senaryolar:", list(st.session_state.senaryolar.keys()), default=list(st.session_state.senaryolar.keys()))
        if secilenler:
            fig3 = go.Figure()
            for ad in secilenler:
                s = st.session_state.senaryolar[ad]
                aylik = pd.DataFrame(s['aylik'])
                if not aylik.empty:
                    fig3.add_trace(go.Scatter(x=aylik['Ay'], y=aylik['Kümülatif Fiyat Farkı'], mode='lines+markers', name=ad))
            fig3.update_layout(title="Senaryolar Arası Karşılaştırma", xaxis_title="Ay", yaxis_title="TL", height=400)
            st.plotly_chart(fig3, use_container_width=True)
        if st.button("🗑️ Tüm senaryoları temizle"):
            st.session_state.senaryolar = {}
            st.rerun()
