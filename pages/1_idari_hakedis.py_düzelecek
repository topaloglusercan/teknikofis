"""
İdari Hakediş & Teyit Matrisi - Web Modülü
- Excel Şablon İndirme motoru geri eklendi.
- Görsel temizlik ve kilitlenme karşıtı 'TextColumn' aktif.
- Koyu temalı analiz grafiği aktif.
"""

import streamlit as st
import pandas as pd
import json
import io
import datetime
import warnings
from decimal import Decimal, ROUND_HALF_UP, getcontext
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

getcontext().prec = 28 
warnings.filterwarnings("ignore")

st.set_page_config(page_title="İdari Hakediş Modülü", layout="wide", page_icon="📂")

# ==========================================
# 1. GÖRSEL TEMİZLİK VE FORMATLAMA
# ==========================================
def clean_df_for_ui(df):
    """Excel/JSON verilerini ekrana basmadan önce insan okunabilir Türk formatına çevirir."""
    df_clean = df.copy()
    months = ['', 'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık']
    
    for col in df_clean.columns:
        is_numeric = str(col).upper() not in ['AYLAR', 'AĞIRLIK', 'ENDEKS SÜTUNU']
        new_vals = []
        for val in df_clean[col]:
            if pd.isna(val) or val is None:
                new_vals.append("")
                continue
                
            if isinstance(val, (pd.Timestamp, datetime.datetime)):
                new_vals.append(f"{months[val.month]} {str(val.year)[-2:]}")
                continue
            
            s = str(val).strip()
            if s.lower() in ['nan', 'none', 'nat', '<na>', '']:
                new_vals.append("")
                continue
                
            if s.endswith(" 00:00:00"):
                s = s.replace(" 00:00:00", "")
                try: 
                    dt = pd.to_datetime(s)
                    s = f"{months[dt.month]} {str(dt.year)[-2:]}"
                except: pass
            else:
                if is_numeric and "," not in s:
                    try:
                        fval = float(s)
                        if fval != 0 or "." in s: 
                            s = "{:,.2f}".format(fval).replace(",", "X").replace(".", ",").replace("X", ".")
                    except: pass
            new_vals.append(s)
        df_clean[col] = new_vals
    return df_clean.astype(str)

def get_text_config(df):
    return {col: st.column_config.TextColumn(col) for col in df.columns}

# ==========================================
# 2. MATEMATİK VE YARDIMCI FONKSİYONLAR
# ==========================================
def parse_turkish_date(date_str):
    if pd.isna(date_str) or str(date_str).strip() == '': return pd.NaT
    s = str(date_str).strip().replace('.', ' ').lower()
    if s in ['none', 'nan', 'nat', '<na>']: return pd.NaT
    try: return pd.to_datetime(s).strftime('%Y-%m')
    except: pass
    
    months = {'oca': '01', 'ocak': '01', 'şub': '02', 'şubat': '02', 'mar': '03', 'mart': '03', 
              'nis': '04', 'nisan': '04', 'may': '05', 'mayıs': '05', 'haz': '06', 'haziran': '06', 
              'tem': '07', 'temmuz': '07', 'ağu': '08', 'ağustos': '08', 'eyl': '09', 'eylül': '09', 
              'eki': '10', 'ekim': '10', 'kas': '11', 'kasım': '11', 'ara': '12', 'aralık': '12'}
    
    parts = s.split()
    if len(parts) >= 2:
        m_num = months.get(parts[0], '01')
        y_num = parts[1] if len(parts[1]) == 4 else f"20{parts[1]}"
        return f"{y_num}-{m_num}"
    return pd.NaT

def clean_decimal(val):
    if val is None or pd.isna(val): return Decimal('0.0')
    s = str(val).strip().replace('TL', '').replace('%', '').strip()
    if s.lower() in ['', 'none', 'nan', 'nat', '<na>']: return Decimal('0.0')
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
        else: s = s.replace(',', '')
    else:
        if ',' in s: s = s.replace(',', '.')
        elif s.count('.') > 1: s = s.replace('.', '')
    try:
        d = Decimal(s)
        if d.is_nan(): return Decimal('0.0')
        return d
    except:
        return Decimal('0.0')

def tr_format(val):
    if pd.isna(val) or val == "": return ""
    try:
        return "{:,.2f}".format(float(val)).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(val)

def filter_empty_rows(df):
    if df.empty: return df
    mask = df.iloc[:, 0].astype(str).str.strip().str.lower().isin(['', 'none', 'nan', 'nat', '<na>'])
    return df[~mask]

# ==========================================
# 3. HESAPLAMA MOTORU 
# ==========================================
def hesapla(df_prog, df_endeks, df_alt, df_b):
    df_prog = filter_empty_rows(df_prog.copy())
    df_endeks = filter_empty_rows(df_endeks.copy())
    df_alt = filter_empty_rows(df_alt.copy())
    df_b = filter_empty_rows(df_b.copy())
    
    if df_prog.empty or df_endeks.empty: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df_prog.columns = df_prog.columns.str.strip()
    
    end_col = 'AYLAR' if 'AYLAR' in df_endeks.columns else 'Aylar'
    df_endeks['AyKodu'] = pd.to_datetime(df_endeks[end_col].apply(parse_turkish_date)).dt.to_period('M')
    df_endeks = df_endeks.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_b['AyKodu'] = pd.to_datetime(df_b['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_b = df_b.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_prog['AyKodu'] = pd.to_datetime(df_prog['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_prog = df_prog.dropna(subset=['AyKodu'])

    if df_endeks.empty or df_prog.empty: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    son_endeks_ayi = df_endeks.index.max()
    katsayilar = {str(r['Ağırlık']).strip().lower(): clean_decimal(r['Katsayı']) for _, r in df_alt.iterrows()}
    temel_endeksler = {str(r['Ağırlık']).strip().lower(): clean_decimal(r['Temel Endeks']) for _, r in df_alt.iterrows()}
    
    _default_harita = {'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}
    if 'Endeks Sütunu' in df_alt.columns:
        endeks_haritasi = {str(r['Ağırlık']).strip().lower(): str(r['Endeks Sütunu']).strip() for _, r in df_alt.iterrows() if str(r.get('Endeks Sütunu', '')).strip() not in ['', 'nan', 'None']}
        if not endeks_haritasi: endeks_haritasi = _default_harita
    else: endeks_haritasi = _default_harita

    prog_kum_col = df_prog.columns[1]; imalat_kum_col = df_prog.columns[2] 
    kovalar = []
    onceki_kum = Decimal('0.0')
    
    for _, row in df_prog.iterrows():
        kum = clean_decimal(row[prog_kum_col])
        cap = kum - onceki_kum
        kovalar.append({'ay': row['AyKodu'], 'kapasite': cap if cap > Decimal('0.0') else Decimal('0.0')})
        onceki_kum = kum

    final_ff, matris = [], []
    onceki_imalat_kum = Decimal('0.0'); kumulatif_ff = Decimal('0.0')

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
        endeks_uyg = df_endeks.loc[gercek_end_ayi] if gercek_end_ayi in df_endeks.index else df_endeks.iloc[-1]
            
        toplam_ff_aylik = Decimal('0.0'); kalan = aylik
        
        for kova in kovalar:
            if kalan <= Decimal('0.0'): break 
            if kova['kapasite'] > Decimal('0.0'):
                kullanilan = min(kalan, kova['kapasite'])
                gercek_prog_ayi = min(kova['ay'], son_endeks_ayi)
                gecikme = kova['ay'] < uyg_ayi
                
                if gecikme:
                    comp_ayi = min(gercek_end_ayi, gercek_prog_ayi)
                    endeks_prog = df_endeks.loc[comp_ayi] if comp_ayi in df_endeks.index else endeks_uyg
                else: endeks_prog = endeks_uyg
                
                pn = Decimal('0.0')
                for k, sutun in endeks_haritasi.items():
                    e_temel = temel_endeksler.get(k, Decimal('0.0'))
                    e_uyg = clean_decimal(endeks_uyg.get(sutun, 0))
                    e_prog = clean_decimal(endeks_prog.get(sutun, 0))
                    e_gecerli = min(e_uyg, e_prog) if gecikme else e_uyg
                    katsayi = katsayilar.get(k, Decimal('0.0'))
                    
                    if e_temel > Decimal('0.0'): pn += katsayi * (e_gecerli / e_temel)
                    elif katsayi > Decimal('0.0'): pn += katsayi
                
                ff_dilim = (kullanilan * b_kat * (pn - Decimal('1.0'))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                matris.append({
                    'Hakediş Ayı': str(uyg_ayi), 'İş Programı (Ödenek) Ayı': str(kova['ay']),
                    'Kullanılan Tutar': float(kullanilan), 'Uygulanan Pn (15 Hane)': float(pn), 'Fiyat Farkı Tutarı': float(ff_dilim)
                })
                toplam_ff_aylik += ff_dilim; kova['kapasite'] -= kullanilan; kalan -= kullanilan
        
        kumulatif_ff += toplam_ff_aylik
        final_ff.append(float(kumulatif_ff)); onceki_imalat_kum = guncel

    df_sonuc = df_prog.copy()
    df_sonuc['KÜMÜLATİF FİYAT FARKI'] = final_ff
    df_detay = pd.DataFrame(matris)
    
    if not df_detay.empty:
        df_pivot = df_detay.pivot_table(index='Hakediş Ayı', columns='İş Programı (Ödenek) Ayı', values='Kullanılan Tutar', aggfunc='sum', fill_value=0)
        df_pivot['HAKEDİŞ TUTARI (Toplam)'] = df_pivot.sum(axis=1)
        df_pivot.loc['ÖDENEK MİKTARI'] = df_pivot.sum()
    else: df_pivot = pd.DataFrame()
    return df_sonuc, df_pivot, df_detay

# ==========================================
# 4. EXCEL ŞABLON MOTORU (GERİ EKLENDİ)
# ==========================================
def generate_excel_download(df_prog, df_endeks, df_alt, df_b):
    """ Ekranda bulunan tabloların verilerini kullanarak renkli ve kilitli profesyonel bir şablon üretir. """
    wb = Workbook()
    thin = Side(style='thin', color='B0C4DE')
    def brd(): return Border(top=thin, left=thin, right=thin, bottom=thin)
    def fill(c): return PatternFill('solid', fgColor=c)
    def fnt(c='1A1A2E', bold=False, sz=10): return Font(color=c, bold=bold, size=sz, name='Calibri')
    def aln(h='left', wrap=False): return Alignment(horizontal=h, vertical='center', wrap_text=wrap)
    
    HEADER = '2C5F8A'; WHITE = 'FFFFFF'; YELLOW = 'FFFDE7'; LIGHT = 'EAF4FB'; NOTE = 'FFF9C4'
    
    def make_header(ws, title, note=''):
        ws.sheet_view.showGridLines = False
        ws.merge_cells('A1:Z1')
        ws['A1'] = title; ws['A1'].fill = fill(HEADER); ws['A1'].font = fnt(WHITE, True, 12); ws['A1'].alignment = aln('center')
        ws.row_dimensions[1].height = 22
        if note:
            ws.merge_cells('A2:Z2')
            ws['A2'] = note; ws['A2'].fill = fill(NOTE); ws['A2'].font = fnt('7B5800', False, 9); ws['A2'].alignment = aln('center')
            ws.row_dimensions[2].height = 14
            
    def col_header(ws, row, headers, widths):
        for ci, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row, ci, h)
            c.fill = fill(HEADER); c.font = fnt(WHITE, True, 9); c.alignment = aln('center'); c.border = brd()
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[row].height = 16
        
    def write_df(ws, df, start_row, yellow_cols):
        for ri, row in enumerate(df.values, start_row):
            bg = YELLOW if ri % 2 == 0 else 'FFFEF0'
            for ci, val in enumerate(row, 1):
                c = ws.cell(ri, ci, str(val) if pd.notna(val) and str(val)!='nan' else '')
                c.fill = fill(YELLOW if ci in yellow_cols else bg)
                c.font = fnt('1A1A2E', False, 9); c.alignment = aln('right' if ci > 1 else 'left'); c.border = brd()
        
        # Gelecek aylar için 10 adet boş sarı giriş satırı ekle
        for ri in range(start_row + len(df), start_row + len(df) + 10):
            bg = YELLOW if ri % 2 == 0 else 'FFFEF0'
            for ci in range(1, len(df.columns) + 1):
                c = ws.cell(ri, ci, '')
                c.fill = fill(YELLOW if ci in yellow_cols else bg)
                c.font = fnt('1A1A2E', False, 9); c.alignment = aln('right' if ci > 1 else 'left'); c.border = brd()

    ws1 = wb.active; ws1.title = 'IsProgrami'
    make_header(ws1, '1. İŞ PROGRAMI VE İMALATLAR', 'Sarı hücrelere verileri girin. Ay formatı: Oca 22, Şub 22, ... / Tüm tutarlar TL')
    col_header(ws1, 3, list(df_prog.columns), [16, 30, 30])
    write_df(ws1, df_prog, 4, yellow_cols=[1, 2, 3])
    
    ws2 = wb.create_sheet('Endeks')
    make_header(ws2, '2. ENDEKS TABLOSU', 'Kaynak: TÜİK / Çevre ve Şehircilik Bakanlığı yayınları')
    col_header(ws2, 3, list(df_endeks.columns), [16] + [12]*(len(df_endeks.columns)-1))
    write_df(ws2, df_endeks, 4, yellow_cols=list(range(1, len(df_endeks.columns)+1)))
    
    ws3 = wb.create_sheet('AltEndeks')
    make_header(ws3, '3. ALT ENDEKS AĞIRLIKLARI', 'Katsayılar toplamı = 1.00 olmalıdır. Temel Endeks = sözleşme tarihindeki endeks değeri.')
    col_header(ws3, 3, list(df_alt.columns), [12, 14, 16, 16][:len(df_alt.columns)])
    write_df(ws3, df_alt, 4, yellow_cols=list(range(1, len(df_alt.columns)+1)))
    
    ws4 = wb.create_sheet('B')
    make_header(ws4, '4. B KATSAYISI TABLOSU', 'Müteahhit fiyat farkı katsayısı (genellikle 1.00).')
    col_header(ws4, 3, list(df_b.columns), [16, 14])
    write_df(ws4, df_b, 4, yellow_cols=[1, 2])
    
    ws5 = wb.create_sheet('Kilavuz')
    ws5.sheet_view.showGridLines = False
    ws5.column_dimensions['A'].width = 90
    ws5['A1'] = 'İDARİ HAKEDİŞ ŞABLONU — KULLANIM KILAVUZU'
    ws5['A1'].fill = fill(HEADER); ws5['A1'].font = fnt(WHITE, True, 14); ws5['A1'].alignment = aln('center')
    ws5.row_dimensions[1].height = 28
    
    help_text = [
        "",
        "💡 EXCEL ŞABLONU İLE VERİ GİRİŞİ",
        "1. Bu Excel dosyasındaki sarı alanları kendi projenizin değerleriyle doldurun.",
        "2. Programın sol menüsündeki 'Excel Dosyası Seç (.xlsx)' bölümünden bu dosyayı yükleyin.",
        "3. Verileriniz anında tabloya aktarılacaktır. Ardından 'Hesapla' butonuna basabilirsiniz."
    ]
    for i, text in enumerate(help_text, 2):
        c = ws5.cell(i, 1, text)
        if text.startswith("💡") or text.startswith("⚠️"):
            c.font = fnt('1A3A6E', True, 11); c.fill = fill('D6E4F0')
        else:
            c.font = fnt('1A1A2E', False, 10)
        ws5.row_dimensions[i].height = 18

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def load_from_excel(file):
    xls = pd.ExcelFile(file)
    dfs = {}
    sheet_map = {'IsProgrami': 'prog_df', 'Endeks': 'endeks_df', 'AltEndeks': 'alt_df', 'B': 'b_df'}
    for sheet in xls.sheet_names:
        if sheet in sheet_map:
            df_temp = pd.read_excel(xls, sheet_name=sheet, nrows=5)
            skip = 0; cols = [str(c).upper() for c in df_temp.columns]
            if 'AYLAR' not in cols and 'AĞIRLIK' not in cols:
                for i, row in df_temp.iterrows():
                    if 'AYLAR' in [str(v).upper() for v in row.values] or 'AĞIRLIK' in [str(v).upper() for v in row.values]:
                        skip = i + 1; break
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=skip)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            
            if sheet == 'AltEndeks' and 'Endeks Sütunu' not in df.columns:
                df['Endeks Sütunu'] = df['Ağırlık'].astype(str).str.strip().str.lower().map({'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}).fillna('')
            
            dfs[sheet_map[sheet]] = clean_df_for_ui(df)
    return dfs

# ==========================================
# 5. STATE & ARAYÜZ (UI)
# ==========================================
if 'load_count' not in st.session_state: st.session_state.load_count = 0
if 'prog_df' not in st.session_state:
    st.session_state.prog_df = pd.DataFrame({"AYLAR": ["Oca 22"], "İŞ PROGRAMI KÜMÜLATİF": ["0,00"], "İMALAT TUTARI KÜMÜLATİF": ["0,00"]})
if 'endeks_df' not in st.session_state:
    st.session_state.endeks_df = pd.DataFrame({"AYLAR": ["Oca 22"], "I o": ["0,00"], "Ç o": ["0,00"], "D o": ["0,00"], "Y o": ["0,00"], "K o": ["0,00"], "G o": ["0,00"], "M o": ["0,00"]})
if 'alt_df' not in st.session_state:
    st.session_state.alt_df = pd.DataFrame({"Ağırlık": ["a", "b1", "b2", "b3", "b4", "b5", "c"], "Katsayı": ["0,00"] * 7, "Temel Endeks": ["0,00"] * 7, "Endeks Sütunu": ["I o", "Ç o", "D o", "Y o", "K o", "G o", "M o"]})
if 'b_df' not in st.session_state:
    st.session_state.b_df = pd.DataFrame({"AYLAR": ["Oca 22"], "B": ["1,00"]})

st.title("📂 İdari Hakediş & Teyit Matrisi")

# -- YAN MENÜ (SIDEBAR) --
st.sidebar.markdown("### 📥 Excel ile Proje Yükle")
uploaded_excel = st.sidebar.file_uploader("Excel Dosyası Seç (.xlsx)", type=["xlsx"])
if uploaded_excel is not None:
    try:
        dfs = load_from_excel(uploaded_excel)
        if 'prog_df' in dfs: st.session_state.prog_df = dfs['prog_df']
        if 'endeks_df' in dfs: st.session_state.endeks_df = dfs['endeks_df']
        if 'alt_df' in dfs: st.session_state.alt_df = dfs['alt_df']
        if 'b_df' in dfs: st.session_state.b_df = dfs['b_df']
        st.session_state.load_count += 1
        st.sidebar.success("✅ Veriler yüklendi ve formatlandı!")
    except Exception as e: st.sidebar.error(f"Hata: {e}")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📥 Önceki Projeyi Yükle (.json)")
uploaded_json = st.sidebar.file_uploader("JSON Dosyası Seç", type=["json"])
if uploaded_json is not None:
    data = json.load(uploaded_json)
    if 'prog' in data: st.session_state.prog_df = clean_df_for_ui(pd.DataFrame(data['prog']))
    if 'endeks' in data: st.session_state.endeks_df = clean_df_for_ui(pd.DataFrame(data['endeks']))
    if 'alt' in data: st.session_state.alt_df = clean_df_for_ui(pd.DataFrame(data['alt']))
    if 'b' in data: st.session_state.b_df = clean_df_for_ui(pd.DataFrame(data['b']))
    st.session_state.load_count += 1
    st.sidebar.success("✅ Proje yüklendi!")

# -- ANA EKRAN TABLOLARI --
suffix = st.session_state.load_count
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. İş Programı ve İmalatlar")
    edited_prog = st.data_editor(st.session_state.prog_df, column_config=get_text_config(st.session_state.prog_df), num_rows="dynamic", use_container_width=True, key=f"prog_ed_{suffix}")
    st.subheader("3. Alt Endeks Ağırlıkları")
    edited_alt = st.data_editor(st.session_state.alt_df, column_config=get_text_config(st.session_state.alt_df), num_rows="dynamic", use_container_width=True, key=f"alt_ed_{suffix}")

with col2:
    st.subheader("2. Endeks Tablosu")
    edited_endeks = st.data_editor(st.session_state.endeks_df, column_config=get_text_config(st.session_state.endeks_df), num_rows="dynamic", use_container_width=True, key=f"end_ed_{suffix}")
    st.subheader("4. B Katsayısı Tablosu")
    edited_b = st.data_editor(st.session_state.b_df, column_config=get_text_config(st.session_state.b_df), num_rows="dynamic", use_container_width=True, key=f"b_ed_{suffix}")

# -- İNDİRME BUTONLARI (JSON & EXCEL) --
st.sidebar.markdown("---")
st.sidebar.markdown("### 📤 Projeyi Dışa Aktar / Şablon Al")

col_s1, col_s2 = st.sidebar.columns(2)
with col_s1:
    st.download_button(
        label="💾 JSON",
        data=json.dumps({'prog': edited_prog.to_dict(orient='records'), 'endeks': edited_endeks.to_dict(orient='records'), 'alt': edited_alt.to_dict(orient='records'), 'b': edited_b.to_dict(orient='records')}, indent=4),
        file_name="hakedis_projem.json", mime="application/json", use_container_width=True
    )
with col_s2:
    st.download_button(
        label="📊 EXCEL",
        data=generate_excel_download(edited_prog, edited_endeks, edited_alt, edited_b),
        file_name="idari_hakedis_sablonu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("---")

# ==========================================
# 6. HESAPLAMA VE GRAFİK EKRANI
# ==========================================
if st.button("🚀 Hesapla ve Sonuçları Göster", use_container_width=True, type="primary"):
    with st.spinner("Matematiksel motor çalışıyor..."):
        df_sonuc, df_pivot, df_detay = hesapla(edited_prog, edited_endeks, edited_alt, edited_b)
    
    if df_sonuc.empty:
        st.warning("⚠️ Lütfen tablolara geçerli veri giriniz (İş Programı ve Endeks boş olamaz).")
    else:
        st.success("✅ Hesaplama Başarılı!")
        tab1, tab2, tab3, tab4 = st.tabs(["🔍 Dilim Detay", "📊 Teyit Matrisi", "📑 Kümülatif Sonuç", "📈 Analiz Grafiği"])
        
        with tab1:
            df_detay_gosterim = df_detay.copy()
            df_detay_gosterim['Kullanılan Tutar'] = df_detay_gosterim['Kullanılan Tutar'].apply(tr_format)
            df_detay_gosterim['Fiyat Farkı Tutarı'] = df_detay_gosterim['Fiyat Farkı Tutarı'].apply(tr_format)
            df_detay_gosterim['Uygulanan Pn (15 Hane)'] = df_detay_gosterim['Uygulanan Pn (15 Hane)'].apply(lambda x: "{:.15f}".format(x).rstrip('0').rstrip('.').replace('.', ','))
            st.dataframe(df_detay_gosterim, use_container_width=True)
            
        with tab2:
            if not df_pivot.empty:
                st.dataframe(df_pivot.map(tr_format).style.set_properties(subset=['HAKEDİŞ TUTARI (Toplam)'], **{'font-weight': 'bold', 'background-color': '#e6f2ff'}), use_container_width=True)
                
        with tab3:
            for col in df_sonuc.columns:
                if any(x in col.upper() for x in ['TUTAR', 'PROGRAM', 'FARKI']):
                    df_sonuc[col] = df_sonuc[col].apply(tr_format)
            st.dataframe(df_sonuc, use_container_width=True)
            
        with tab4:
            try:
                df_prog_raw = filter_empty_rows(edited_prog.copy())
                df_prog_raw['AyKodu'] = pd.to_datetime(df_prog_raw['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
                df_prog_raw = df_prog_raw.dropna(subset=['AyKodu'])
                
                aylar = [str(a) for a in df_prog_raw['AyKodu']]
                prog_vals = [float(clean_decimal(v)) for v in df_prog_raw.iloc[:, 1]]
                iml_vals = [float(clean_decimal(v)) for v in df_prog_raw.iloc[:, 2]]
                
                ff_col = next((c for c in df_sonuc.columns if 'FARKI' in str(c).upper() or 'FIYAT' in str(c).upper()), df_sonuc.columns[-1])
                ff_vals = [float(clean_decimal(v)) for v in df_sonuc[ff_col]]
                
                x = list(range(len(aylar)))
                fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 7), sharex=True, gridspec_kw={'height_ratios': [3, 1], 'hspace': 0.1})
                fig.patch.set_facecolor('#0d1117')
                
                for ax in [ax1, ax2]:
                    ax.set_facecolor('#161b22')
                    ax.tick_params(colors='#8b949e', labelsize=9)
                    for spine in ax.spines.values(): spine.set_edgecolor('#30363d')
                    ax.grid(color='#21262d', linewidth=0.7, linestyle='--')
                    
                ax1.plot(x, prog_vals, color='#58a6ff', linewidth=2.5, marker='o', markersize=5, label='İş Programı (Kümülatif)')
                ax1.plot(x, iml_vals, color='#3fb950', linewidth=2.5, marker='s', markersize=5, linestyle='--', label='İmalat Tutarı (Kümülatif)')
                ax1.fill_between(x, prog_vals, iml_vals, where=[p >= i for p, i in zip(prog_vals, iml_vals)], alpha=0.15, color='#f85149', label='Gecikme')
                ax1.fill_between(x, prog_vals, iml_vals, where=[p < i for p, i in zip(prog_vals, iml_vals)], alpha=0.15, color='#3fb950', label='Öne Geçme')
                
                ax1.set_ylabel('Tutar (TL)', color='#8b949e', fontsize=10)
                ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: tr_format(v)))
                ax1.legend(fontsize=9, facecolor='#21262d', edgecolor='#30363d', labelcolor='#c9d1d9', loc='upper left')
                ax1.set_title('Kümülatif İş Programı & İmalat Takibi', color='#e6edf3', fontsize=13, pad=12)
                
                colors_ff = ['#3fb950' if v >= 0 else '#f85149' for v in ff_vals]
                ax2.bar(x, ff_vals, color=colors_ff, alpha=0.85, width=0.6)
                ax2.axhline(0, color='#30363d', linewidth=1)
                ax2.set_ylabel('Fiyat Farkı (TL)', color='#8b949e', fontsize=10)
                ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: tr_format(v)))
                ax2.set_title('Kümülatif Fiyat Farkı', color='#c9d1d9', fontsize=11, pad=6)
                ax2.set_xticks(x)
                ax2.set_xticklabels(aylar, rotation=45, ha='right', fontsize=9, color='#8b949e')
                
                st.pyplot(fig)
            except Exception as e:
                st.error(f"Grafik oluşturulamadı: {e}")
