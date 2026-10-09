"""
Fiyat Farkı Tahmin & Stokastik Endeks Simülatörü (5. Sekme Backtesting & Resmi Raporlama Sürümü)
Özellikler:
- Okul Projesi 19 Ay (Eylül 25 - Mart 27, 433.444.000 TL) tam veri seti eksiksiz yüklü.
- "5. En İyi Karma (Alt Endekslere Göre Dinamik)" modeli aktif.
- Mevcut duruma kadar 1_idari_hakedis.py ile %100 birebir aynı hakediş hesabı yapar.
- Kalan aylar iş programına paralel tamamlandığında ~550 Milyon TL toplam FF üretir.
- 5. Sekmede Kapsamlı Backtesting (MAPE & RMSE) Tablosu ve Doğrulama Grafiği.
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import io
import re
import datetime
import warnings
import requests
from decimal import Decimal, ROUND_HALF_UP, getcontext
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

getcontext().prec = 28
warnings.filterwarnings("ignore")

try:
    st.set_page_config(page_title="Fiyat Farkı Tahmin & Endeks Simülatörü", page_icon="📈", layout="wide")
except Exception:
    pass

# ==========================================
# 1. RESMİ ENDEKS KODLARI VE YARDIMCI FONKSİYONLAR
# ==========================================
KOD_BILGI = {
    'a':  {'kolon': 'I o', 'resmi_kod': 'İn',    'kisa': 'İşçilik',         'ad': 'İşçilik (Asgari Ücret / TÜFE Bağlı)'},
    'b1': {'kolon': 'Ç o', 'resmi_kod': 'Çn-23', 'kisa': 'Çimento/Mineral', 'ad': 'Metalik Olmayan Diğer Mineral Ürünler (Çimento, Hazır Beton)'},
    'b2': {'kolon': 'D o', 'resmi_kod': 'Dn-24', 'kisa': 'Demir-Çelik',     'ad': 'Ana Metaller (Demir-Çelik)'},
    'b3': {'kolon': 'Y o', 'resmi_kod': 'Ayn',   'kisa': 'Akaryakıt',       'ad': 'Akaryakıt Ürünleri (Motorin)'},
    'b4': {'kolon': 'K o', 'resmi_kod': 'Kn-16', 'kisa': 'Ağaç/Kereste',    'ad': 'Ağaç ve Mantar Ürünleri (Kereste vb.)'},
    'b5': {'kolon': 'G o', 'resmi_kod': 'Gn',    'kisa': 'Genel Yİ-ÜFE',    'ad': 'Genel Yurt İçi Üretici Fiyat Endeksi (Yİ-ÜFE)'},
    'c':  {'kolon': 'M o', 'resmi_kod': 'Mn-28', 'kisa': 'Makine/Ekipman',  'ad': 'Makine ve Ekipmanlar (Amortisman)'},
}

AYLAR_LIST = ['', 'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık']

MONTHS_MAP = {
    'oca': '01', 'ocak': '01', 'şub': '02', 'sub': '02', 'şubat': '02', 'subat': '02',
    'mar': '03', 'mart': '03', 'nis': '04', 'nisan': '04', 'may': '05', 'mayıs': '05', 'mayis': '05',
    'haz': '06', 'haziran': '06', 'tem': '07', 'temmuz': '07', 'ağu': '08', 'agu': '08', 'ağustos': '08', 'agustos': '08',
    'eyl': '09', 'eylül': '09', 'eylul': '09', 'eki': '10', 'ekim': '10', 'kas': '11', 'kasım': '11', 'kasim': '11',
    'ara': '12', 'aralık': '12', 'aralik': '12'
}

def parse_turkish_date(date_str):
    if pd.isna(date_str) or str(date_str).strip() == '': return pd.NaT
    s = str(date_str).strip().replace('.', ' ').replace('\n', ' ').replace('\r', ' ').lower()
    if s in ['none', 'nan', 'nat', '<na>']: return pd.NaT
    try: return pd.to_datetime(s).strftime('%Y-%m')
    except Exception: pass
    parts = [p for p in s.split() if p]
    if len(parts) >= 2:
        if parts[0] in MONTHS_MAP and parts[1].isdigit():
            m_num = MONTHS_MAP[parts[0]]
            y_num = parts[1] if len(parts[1]) == 4 else f"20{parts[1]}"
            return f"{y_num}-{m_num}"
        if parts[0].isdigit() and parts[1] in MONTHS_MAP:
            m_num = MONTHS_MAP[parts[1]]
            y_num = parts[0] if len(parts[0]) == 4 else f"20{parts[0]}"
            return f"{y_num}-{m_num}"
    return pd.NaT

def period_to_tr_full_str(period_val, short_year=False):
    yr = str(period_val.year)[-2:] if short_year else str(period_val.year)
    return f"{AYLAR_LIST[period_val.month]} {yr}"

def clean_decimal(val):
    if val is None or pd.isna(val): return Decimal('0.0')
    if isinstance(val, Decimal): return Decimal('0.0') if val.is_nan() else val
    if isinstance(val, (int, float, np.number)):
        if np.isnan(val): return Decimal('0.0')
        return Decimal(str(val))
    s = str(val).strip().replace('TL', '').replace('₺', '').replace('%', '').strip()
    if s.lower() in ['', 'none', 'nan', 'nat', '<na>']: return Decimal('0.0')
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
        else: s = s.replace(',', '')
    elif ',' in s: s = s.replace(',', '.')
    elif s.count('.') > 1: s = s.replace('.', '')
    try:
        d = Decimal(s)
        return Decimal('0.0') if d.is_nan() else d
    except Exception:
        return Decimal('0.0')

def tr_format(val, decimals=2):
    if pd.isna(val) or val == "": return ""
    try:
        fmt = f"{{:,.{decimals}f}}"
        return fmt.format(float(val)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)

def filter_empty_rows(df):
    if df is None or df.empty: return pd.DataFrame()
    mask = df.iloc[:, 0].astype(str).str.strip().str.lower().isin(['', 'none', 'nan', 'nat', '<na>'])
    return df[~mask]

def clean_df_for_ui(df):
    df_clean = df.copy()
    for col in df_clean.columns:
        is_numeric = str(col).upper() not in ['AYLAR', 'AĞIRLIK', 'ENDEKS SÜTUNU']
        new_vals = []
        for val in df_clean[col]:
            if pd.isna(val) or val is None:
                new_vals.append("")
                continue
            if isinstance(val, (pd.Timestamp, datetime.datetime)):
                new_vals.append(f"{AYLAR_LIST[val.month]} {str(val.year)[-2:]}")
                continue
            s = str(val).strip()
            if s.lower() in ['nan', 'none', 'nat', '<na>', '']:
                new_vals.append("")
                continue
            if s.endswith(" 00:00:00"):
                s = s.replace(" 00:00:00", "")
                try:
                    dt = pd.to_datetime(s)
                    s = f"{AYLAR_LIST[dt.month]} {str(dt.year)[-2:]}"
                except Exception: pass
            else:
                if is_numeric and "," not in s:
                    try:
                        dval = Decimal(s).normalize()
                        s_val = f"{dval:f}"
                        if "." in s_val:
                            int_part, dec_part = s_val.split(".")
                            int_part = "{:,}".format(int(int_part)).replace(",", ".")
                            s = f"{int_part},{dec_part}"
                        else:
                            s = "{:,}".format(int(s_val)).replace(",", ".")
                    except Exception: pass
            new_vals.append(s)
        df_clean[col] = new_vals
    return df_clean.astype(str).astype("string")

def ensure_text_df(df):
    if df is None or df.empty: return df
    return clean_df_for_ui(df)

def get_text_config(df):
    return {col: st.column_config.TextColumn(col, width="medium") for col in df.columns}

def ayristir_gercek_endeks(df_end):
    df_c = filter_empty_rows(df_end.copy())
    if df_c.empty: return ensure_text_df(df_end)
    end_col = 'AYLAR' if 'AYLAR' in df_c.columns else df_c.columns[0]
    num_cols = [c for c in df_c.columns if c != end_col]
    gecerli_satirlar = []
    for idx, row in df_c.iterrows():
        vals = [float(clean_decimal(row[c])) for c in num_cols]
        if all(v > 0 for v in vals): gecerli_satirlar.append(idx)
    if gecerli_satirlar:
        df_c = df_c.loc[gecerli_satirlar].reset_index(drop=True)
    if len(df_c) <= 5: return ensure_text_df(df_c)
    def satir_tahmin_mi(row):
        uzun_kesir = 0
        for c in num_cols:
            val_s = str(row[c]).strip()
            if ',' in val_s:
                dec_part = val_s.split(',')[-1]
                if len(dec_part) == 6 and not dec_part.endswith('0000'): uzun_kesir += 1
        return uzun_kesir >= 4
    flags = [satir_tahmin_mi(r) for _, r in df_c.iterrows()]
    if not flags[0] and any(flags):
        for i in range(len(flags)):
            if all(flags[i:]):
                if i >= 3: return ensure_text_df(df_c.iloc[:i].reset_index(drop=True))
                break
    return ensure_text_df(df_c)

# ==========================================
# 2. VARSAYILAN VERİ SETİ (OKUL PROJESİ - 19 AY TAM LİSTE)
# ==========================================
DEFAULT_PROG = [
    {"AYLAR": "Eylül 25",   "İŞ PROGRAMI KÜMÜLATİF": "2.165.869,79",   "İMALAT TUTARI KÜMÜLATİF": "0"},
    {"AYLAR": "Ekim 25",    "İŞ PROGRAMI KÜMÜLATİF": "6.470.938,48",   "İMALAT TUTARI KÜMÜLATİF": "0"},
    {"AYLAR": "Kasım 25",   "İŞ PROGRAMI KÜMÜLATİF": "17.293.937,21",  "İMALAT TUTARI KÜMÜLATİF": "0"},
    {"AYLAR": "Aralık 25",  "İŞ PROGRAMI KÜMÜLATİF": "30.287.885,90",  "İMALAT TUTARI KÜMÜLATİF": "0"},
    {"AYLAR": "Ocak 26",    "İŞ PROGRAMI KÜMÜLATİF": "55.011.822,59",  "İMALAT TUTARI KÜMÜLATİF": "54.223.844,40"},
    {"AYLAR": "Şubat 26",   "İŞ PROGRAMI KÜMÜLATİF": "84.518.453,41",  "İMALAT TUTARI KÜMÜLATİF": "54.223.844,40"},
    {"AYLAR": "Mart 26",    "İŞ PROGRAMI KÜMÜLATİF": "116.193.494,10", "İMALAT TUTARI KÜMÜLATİF": "86.775.488,80"},
    {"AYLAR": "Nisan 26",   "İŞ PROGRAMI KÜMÜLATİF": "148.697.481,54", "İMALAT TUTARI KÜMÜLATİF": "128.039.357,60"},
    {"AYLAR": "Mayıs 26",   "İŞ PROGRAMI KÜMÜLATİF": "180.326.778,21", "İMALAT TUTARI KÜMÜLATİF": "162.021.367,20"},
    {"AYLAR": "Haziran 26", "İŞ PROGRAMI KÜMÜLATİF": "210.665.272,86", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Temmuz 26",  "İŞ PROGRAMI KÜMÜLATİF": "240.061.456,36", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Ağustos 26", "İŞ PROGRAMI KÜMÜLATİF": "268.254.235,09", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Eylül 26",   "İŞ PROGRAMI KÜMÜLATİF": "294.265.183,12", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Ekim 26",    "İŞ PROGRAMI KÜMÜLATİF": "320.300.266,54", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Kasım 26",   "İŞ PROGRAMI KÜMÜLATİF": "346.314.783,44", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Aralık 26",  "İŞ PROGRAMI KÜMÜLATİF": "370.147.808,63", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Ocak 27",    "İŞ PROGRAMI KÜMÜLATİF": "396.149.133,62", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Şubat 27",   "İŞ PROGRAMI KÜMÜLATİF": "417.840.015,98", "İMALAT TUTARI KÜMÜLATİF": ""},
    {"AYLAR": "Mart 27",    "İŞ PROGRAMI KÜMÜLATİF": "433.444.000,00", "İMALAT TUTARI KÜMÜLATİF": ""}
]

DEFAULT_ENDEKS = [
    {"AYLAR": "Temmuz 2025", "I o": "100,421925", "Ç o": "4972,280000", "D o": "6034,240000", "Y o": "44,717708", "K o": "3481,270000", "G o": "4409,730000", "M o": "3218,000000"},
    {"AYLAR": "Eylül 2025",  "I o": "105,780006", "Ç o": "5110,090000", "D o": "6187,730000", "Y o": "45,110900", "K o": "3572,840000", "G o": "4632,890000", "M o": "3319,760000"},
    {"AYLAR": "Ekim 2025",   "I o": "108,477581", "Ç o": "5170,070000", "D o": "6350,020000", "Y o": "44,919783", "K o": "3577,460000", "G o": "4708,200000", "M o": "3368,810000"},
    {"AYLAR": "Kasım 2025",  "I o": "109,415936", "Ç o": "5216,990000", "D o": "6453,420000", "Y o": "47,560158", "K o": "3606,930000", "G o": "4747,630000", "M o": "3419,550000"},
    {"AYLAR": "Aralık 2025", "I o": "110,386963", "Ç o": "5277,720000", "D o": "6568,100000", "Y o": "44,895592", "K o": "3667,480000", "G o": "4783,040000", "M o": "3475,040000"},
    {"AYLAR": "Ocak 2026",   "I o": "115,730000", "Ç o": "5454,970000", "D o": "6710,150000", "Y o": "46,162817", "K o": "3786,330000", "G o": "4910,530000", "M o": "3624,850000"},
    {"AYLAR": "Şubat 2026",  "I o": "119,160000", "Ç o": "5591,460000", "D o": "6805,630000", "Y o": "48,474967", "K o": "3895,970000", "G o": "5029,760000", "M o": "3708,220000"},
    {"AYLAR": "Mart 2026",   "I o": "121,470000", "Ç o": "5683,710000", "D o": "6960,900000", "Y o": "56,330433", "K o": "3968,370000", "G o": "5145,360000", "M o": "3726,740000"},
    {"AYLAR": "Nisan 2026",  "I o": "126,550000", "Ç o": "5841,560000", "D o": "7209,790000", "Y o": "61,810350", "K o": "4068,800000", "G o": "5308,460000", "M o": "3813,110000"},
    {"AYLAR": "Mayıs 2026",  "I o": "128,720000", "Ç o": "5972,290000", "D o": "7364,500000", "Y o": "56,946300", "K o": "4233,670000", "G o": "5454,580000", "M o": "3857,200000"}
]

DEFAULT_ALT = [
    {"Ağırlık": "a",  "Katsayı": "0,15", "Temel Endeks": "100,421925",  "Endeks Sütunu": "I o"},
    {"Ağırlık": "b1", "Katsayı": "0,20", "Temel Endeks": "4972,280000", "Endeks Sütunu": "Ç o"},
    {"Ağırlık": "b2", "Katsayı": "0,20", "Temel Endeks": "6034,240000", "Endeks Sütunu": "D o"},
    {"Ağırlık": "b3", "Katsayı": "0,15", "Temel Endeks": "44,717708",   "Endeks Sütunu": "Y o"},
    {"Ağırlık": "b4", "Katsayı": "0,05", "Temel Endeks": "3481,270000", "Endeks Sütunu": "K o"},
    {"Ağırlık": "b5", "Katsayı": "0,10", "Temel Endeks": "4409,730000", "Endeks Sütunu": "G o"},
    {"Ağırlık": "c",  "Katsayı": "0,15", "Temel Endeks": "3218,000000", "Endeks Sütunu": "M o"}
]

DEFAULT_B = [
    {"AYLAR": "Eylül 25",   "B": "0,90"},
    {"AYLAR": "Ekim 25",    "B": "0,90"},
    {"AYLAR": "Kasım 25",   "B": "0,90"},
    {"AYLAR": "Aralık 25",  "B": "0,90"},
    {"AYLAR": "Ocak 26",    "B": "0,90"},
    {"AYLAR": "Şubat 26",   "B": "0,90"},
    {"AYLAR": "Mart 26",    "B": "0,90"},
    {"AYLAR": "Nisan 26",   "B": "0,90"},
    {"AYLAR": "Mayıs 26",   "B": "0,90"},
    {"AYLAR": "Haziran 26", "B": "0,90"},
    {"AYLAR": "Temmuz 26",  "B": "0,90"},
    {"AYLAR": "Ağustos 26", "B": "0,90"},
    {"AYLAR": "Eylül 26",   "B": "0,90"},
    {"AYLAR": "Ekim 26",    "B": "0,90"},
    {"AYLAR": "Kasım 26",   "B": "0,90"},
    {"AYLAR": "Aralık 26",  "B": "0,90"},
    {"AYLAR": "Ocak 27",    "B": "0,90"},
    {"AYLAR": "Şubat 27",   "B": "0,90"},
    {"AYLAR": "Mart 27",    "B": "0,90"}
]

# ==========================================
# 3. WEB SCRAPING & TEYİT MOTORU
# ==========================================
@st.cache_data(ttl=3600)
def webden_endeks_cek_ve_guncelle(df_mevcut_base):
    url = "https://www.hakedis.org/endeksler/yapim-isleri-fiyat-farki-endeksleri-kasim-2013-sonrasi"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    hedef_kolonlar = ["I o", "Ç o", "D o", "Y o", "K o", "G o", "M o"]
    try:
        r = requests.get(url, headers=headers, timeout=10)
        r.raise_for_status()
        dfs = pd.read_html(io.StringIO(r.text), thousands='.', decimal=',')
        for df in dfs:
            if len(df.columns) >= 8:
                df_raw = df.iloc[:, :8].copy()
                df_raw.columns = ["AYLAR"] + hedef_kolonlar
                df_raw['_per'] = pd.to_datetime(df_raw["AYLAR"].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M')
                df_valid = df_raw.dropna(subset=['_per']).copy()
                if df_valid.empty: continue
                pozitif_mask = []
                for _, row in df_valid.iterrows():
                    vals = [float(clean_decimal(row[c])) for c in hedef_kolonlar]
                    pozitif_mask.append(all(v > 1.0 for v in vals))
                df_valid = df_valid[pozitif_mask].drop_duplicates(subset=['_per']).sort_values('_per').reset_index(drop=True)
                
                df_valid_formatli = df_valid.copy()
                df_valid_formatli["AYLAR"] = [period_to_tr_full_str(p) for p in df_valid_formatli['_per']]
                for col in hedef_kolonlar:
                    df_valid_formatli[col] = [tr_format(float(clean_decimal(v)), 2) for v in df_valid_formatli[col]]
                df_ham_web = clean_df_for_ui(df_valid_formatli.drop(columns=['_per']))
                
                if len(df_valid) >= 3:
                    df_m = filter_empty_rows(df_mevcut_base.copy())
                    if not df_m.empty:
                        m_col = 'AYLAR' if 'AYLAR' in df_m.columns else df_m.columns[0]
                        df_m['_per'] = pd.to_datetime(df_m[m_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M')
                        df_m = df_m.dropna(subset=['_per']).sort_values('_per').reset_index(drop=True)
                        max_m_per = df_m['_per'].max()
                        df_yeni_aylar = df_valid[df_valid['_per'] > max_m_per].copy()
                        if df_yeni_aylar.empty:
                            return clean_df_for_ui(df_mevcut_base), True, f"Mevcut projenin endeksleri ({period_to_tr_full_str(max_m_per)}) Hakedis.org'daki son açıklanan aydan güncel veya eşit. Tablo korundu.", df_ham_web
                        
                        for col in hedef_kolonlar:
                            son_m_val = float(clean_decimal(df_m[col].iloc[-1]))
                            y_ilk_val = float(clean_decimal(df_yeni_aylar[col].iloc[0]))
                            if son_m_val > 0 and y_ilk_val > 0 and ((son_m_val / y_ilk_val) > 3.0 or (y_ilk_val / son_m_val) > 3.0):
                                ortak = df_valid[df_valid['_per'] == max_m_per]
                                ref_web = float(clean_decimal(ortak[col].iloc[0])) if not ortak.empty else y_ilk_val
                                carpan = (son_m_val / ref_web) if ref_web > 0 else 1.0
                                df_yeni_aylar[col] = [tr_format(float(clean_decimal(v)) * carpan, 6) for v in df_yeni_aylar[col]]
                            else:
                                df_yeni_aylar[col] = [tr_format(float(clean_decimal(v)), 6) for v in df_yeni_aylar[col]]
                        df_yeni_aylar["AYLAR"] = [period_to_tr_full_str(p) for p in df_yeni_aylar['_per']]
                        df_birlesik = pd.concat([df_m.drop(columns=['_per']), df_yeni_aylar.drop(columns=['_per'])], ignore_index=True)
                        return clean_df_for_ui(df_birlesik), True, f"Hakedis.org'dan {len(df_yeni_aylar)} yeni açıklanan ay eklendi.", df_ham_web
                    else:
                        return df_ham_web, True, f"Hakedis.org üzerinden {len(df_valid)} aylık güncel TÜİK endeksi başarıyla çekildi.", df_ham_web
        return clean_df_for_ui(df_mevcut_base), False, "Hakedis.org tablosunda yeni açıklanmış pozitif endeks satırı bulunamadı; mevcut proje endeksleri korundu.", None
    except Exception as e:
        return clean_df_for_ui(df_mevcut_base), False, f"Hakedis.org bağlantı/güvenlik engeli ({type(e).__name__}); mevcut proje endeksleri korundu.", None

def hazirla_prog_ve_turetilen_ay(df_prog_raw, df_b_raw, tamamlama_aktif=True):
    df_p = filter_empty_rows(df_prog_raw.copy())
    df_b = filter_empty_rows(df_b_raw.copy())
    if df_p.empty or not tamamlama_aktif:
        return ensure_text_df(df_p), ensure_text_df(df_b), None
    ay_col = df_p.columns[0]
    prog_col = df_p.columns[1]
    iml_col = df_p.columns[2]

    # Tablodaki boş hücreleri kontrol et
    bos_indisler = [i for i in range(len(df_p)) if str(df_p.iloc[i, 2]).strip() in ['', 'none', 'nan', '<na>']]
    
    if len(bos_indisler) > 0:
        # DURUM A (Okul Projesi gibi): Boş bırakılan gelecek ayları iş programına paralel doldur
        for idx in bos_indisler:
            df_p.iloc[idx, 2] = df_p.iloc[idx, 1]
        ilk_bos = bos_indisler[0]
        baz_iml = clean_decimal(df_p.iloc[ilk_bos - 1, 2]) if ilk_bos > 0 else Decimal('0.0')
        kalan_tutar = clean_decimal(df_p.iloc[-1, 1]) - baz_iml
        return ensure_text_df(df_p), ensure_text_df(df_b), {
            'tip': 'mevcut_bos_aylar',
            'kalan_tutar': float(kalan_tutar),
            'aciklama': f"Tablodaki boş bırakılan **{len(bos_indisler)} gelecek ay** iş programına paralel olarak tamamlandı (Kalan İmalat: **{tr_format(kalan_tutar)} TL**)."
        }
    else:
        # DURUM B (Konut 42 gibi): Boş ay yoksa ve imalat toplamı iş programını karşılamıyorsa +1 ay türet
        son_prog = clean_decimal(df_p.iloc[-1, 1])
        son_iml = clean_decimal(df_p.iloc[-1, 2])
        kalan_tutar = son_prog - son_iml
        if kalan_tutar > Decimal('0.01'):
            son_ay_str = str(df_p.iloc[-1, 0])
            son_per = pd.to_datetime(parse_turkish_date(son_ay_str), format='%Y-%m', errors='coerce').to_period('M')
            if pd.notna(son_per):
                yeni_per = son_per + 1
                yeni_ay_adi = period_to_tr_full_str(yeni_per, short_year=True)
                yeni_prog_row = pd.DataFrame([{ay_col: yeni_ay_adi, prog_col: tr_format(son_prog), iml_col: tr_format(son_prog)}])
                df_p = pd.concat([df_p, yeni_prog_row], ignore_index=True)
                if not df_b.empty:
                    b_ay_col = df_b.columns[0]
                    b_val_col = df_b.columns[1]
                    son_b_val = str(df_b.iloc[-1, 1])
                    yeni_b_row = pd.DataFrame([{b_ay_col: yeni_ay_adi, b_val_col: son_b_val}])
                    df_b = pd.concat([df_b, yeni_b_row], ignore_index=True)
                return ensure_text_df(df_p), ensure_text_df(df_b), {
                    'tip': 'turetilen_ay',
                    'yeni_ay': yeni_ay_adi,
                    'yeni_per': yeni_per,
                    'kalan_tutar': float(kalan_tutar),
                    'aciklama': f"Kalan boş ay olmadığı ve son imalat toplam iş programını karşılamadığı için **{yeni_ay_adi}** ayı otomatik türetildi ve kalan **{tr_format(kalan_tutar)} TL** imalat bu aya gömüldü."
                }
    return ensure_text_df(df_p), ensure_text_df(df_b), None

def hesapla_gereken_ufuk(df_prog, df_base_end):
    try:
        df_p = filter_empty_rows(df_prog)
        df_e = filter_empty_rows(df_base_end)
        p_col = 'AYLAR' if 'AYLAR' in df_p.columns else df_p.columns[0]
        e_col = 'AYLAR' if 'AYLAR' in df_e.columns else df_e.columns[0]
        p_max = pd.to_datetime(df_p[p_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M').max()
        e_max = pd.to_datetime(df_e[e_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M').max()
        son_prog = clean_decimal(df_p.iloc[-1, 1])
        son_iml_str = str(df_p.iloc[-1, 2]).strip()
        son_iml = clean_decimal(son_iml_str) if son_iml_str != '' else Decimal('0.0')
        ekstra_turetilen = 1 if (son_iml_str != '' and (son_prog - son_iml) > Decimal('0.01')) else 0
        if pd.notna(p_max) and pd.notna(e_max) and (p_max + ekstra_turetilen) > e_max:
            hedef_max = p_max + ekstra_turetilen
            diff = (hedef_max.year - e_max.year) * 12 + (hedef_max.month - e_max.month)
            return int(max(1, min(diff, 36)))
    except Exception: pass
    return 10

# ==========================================
# 4. İLERİ DÜZEY ZAMAN SERİSİ, STOKASTİK TAHMİN VE BACKTESTING MOTORU
# ==========================================
def tek_seri_tahmin_uret(y_arr, ek_ay, model_tipi, alpha=0.65, beta=0.35, phi=0.90, mc_quantile=80, n_sims=1000):
    y_raw = [float(v) for v in y_arr if float(v) > 0]
    if len(y_raw) == 0: return [1.0] * max(ek_ay, 0)
    y = np.array(y_raw, dtype=float)
    n = len(y)
    if n < 2 or ek_ay <= 0: return [float(y[-1])] * max(ek_ay, 0)
        
    preds = []
    if "Holt" in model_tipi:
        L, T = y[0], y[1] - y[0]
        for i in range(1, n):
            L_prev = L
            L = alpha * y[i] + (1.0 - alpha) * (L + T)
            T = beta * (L - L_prev) + (1.0 - beta) * T
        T = max(T, y[-1] * 0.005)
        for h in range(1, ek_ay + 1): preds.append(L + h * T)
    elif "Sönümlü" in model_tipi or "Damped" in model_tipi:
        L, T = y[0], y[1] - y[0]
        for i in range(1, n):
            L_prev = L
            L = alpha * y[i] + (1.0 - alpha) * (L_prev + phi * T)
            T = beta * (L - L_prev) + (1.0 - beta) * phi * T
        T = max(T, y[-1] * 0.005)
        for h in range(1, ek_ay + 1):
            phi_sum = sum(phi ** j for j in range(1, h + 1))
            preds.append(L + phi_sum * T)
    elif "Monte Carlo" in model_tipi or "GBM" in model_tipi:
        np.random.seed(42)
        log_returns = np.diff(np.log(y))
        mu = max(float(np.mean(log_returns)), 0.005)
        sigma = max(float(np.std(log_returns)), 0.005)
        sim_paths = np.zeros((n_sims, ek_ay))
        for s in range(n_sims):
            curr = y[-1]
            for h in range(ek_ay):
                z = np.random.normal(0, 1)
                curr = curr * np.exp((mu - 0.5 * (sigma ** 2)) + sigma * z)
                sim_paths[s, h] = curr
        preds = np.percentile(sim_paths, mc_quantile, axis=0).tolist()
    else:
        pct_returns = np.diff(y) / y[:-1]
        weights = np.exp(np.linspace(-1.0, 0.0, len(pct_returns)))
        weights /= weights.sum()
        ewma_r = max(float(np.sum(pct_returns * weights)), 0.005)
        last_r = max(float(pct_returns[-1]), 0.005)
        ar_rho = 0.65
        curr = y[-1]
        curr_r = last_r
        for h in range(1, ek_ay + 1):
            curr_r = ar_rho * curr_r + (1.0 - ar_rho) * ewma_r
            curr = curr * (1.0 + max(curr_r, 0.003))
            preds.append(curr)
            
    safe_preds = []
    prev_val = y[-1]
    for p_val in preds:
        safe_val = max(float(p_val), prev_val * 1.002)
        prev_val = safe_val
        safe_preds.append(safe_val)
    return safe_preds

def backtest_modelleri_hesapla(df_base_endeks, df_alt, alpha=0.65, beta=0.35, phi=0.90):
    df_clean = ayristir_gercek_endeks(df_base_endeks.copy())
    end_col = 'AYLAR' if 'AYLAR' in df_clean.columns else df_clean.columns[0]
    endeks_sutunlari = [c for c in df_clean.columns if c != end_col]
    
    katsayilar = {}
    for _, r in df_alt.iterrows():
        k_kod = str(r['Ağırlık']).strip().lower()
        sut = str(r.get('Endeks Sütunu', KOD_BILGI.get(k_kod, {}).get('kolon', ''))).strip()
        katsayilar[sut] = float(clean_decimal(r['Katsayı']))

    modeller = [("1. Holt's Çift Üstel", "Holt"), ("2. Sönümlü Trend (Damped)", "Sönümlü"), ("3. Monte Carlo GBM (P50 Medyan)", "Monte Carlo"), ("4. Otoregresif (AR-1 & EWMA)", "Otoregresif")]
    satirlar = []
    karma_mape = {etiket: 0.0 for etiket, _ in modeller}
    detay_gecmis = {}
    en_iyi_modeller_dict = {}
    
    n_total = len(df_clean)
    baslangic_idx = max(2, n_total - 6)
    test_aylari = df_clean[end_col].iloc[baslangic_idx:].tolist()
    
    for col in endeks_sutunlari:
        y_all = np.array([max(float(clean_decimal(v)), 1e-6) for v in df_clean[col]], dtype=float)
        agirlik = katsayilar.get(col, 0.0)
        row_dict = {"Alt Endeks": col, "Sözleşme Ağırlığı": f"%{agirlik*100:.0f}"}
        en_iyi_mape = 999.0
        en_iyi_model_isim = ""
        en_iyi_model_kod = ""
        detay_gecmis[col] = {"Aylar": test_aylari, "Gerçekleşen": y_all[baslangic_idx:].tolist()}
        
        for etiket, m_kod in modeller:
            gercekler, tahminler = [], []
            for t in range(baslangic_idx, n_total):
                p_val = tek_seri_tahmin_uret(y_all[:t], 1, m_kod, alpha=alpha, beta=beta, phi=phi, mc_quantile=50)[0]
                gercekler.append(max(y_all[t], 1e-6))
                tahminler.append(p_val)
                
            if len(gercekler) > 0:
                gercekler_arr = np.array(gercekler)
                tahminler_arr = np.array(tahminler)
                mape = float(np.mean(np.abs((gercekler_arr - tahminler_arr) / np.maximum(gercekler_arr, 1e-6))) * 100.0)
                rmse = float(np.sqrt(np.mean((gercekler_arr - tahminler_arr) ** 2)))
            else:
                mape, rmse = 0.0, 0.0
            
            row_dict[f"{etiket} MAPE"] = f"%{mape:.2f} (RMSE: {tr_format(rmse, 2)})"
            karma_mape[etiket] += agirlik * mape
            detay_gecmis[col][etiket] = tahminler
            if mape < en_iyi_mape:
                en_iyi_mape = mape
                en_iyi_model_isim = etiket.split(".")[1].strip()
                en_iyi_model_kod = m_kod
                
        row_dict["En Düşük Hata (Önerilen)"] = f"🏆 {en_iyi_model_isim} (%{en_iyi_mape:.2f})"
        en_iyi_modeller_dict[col] = en_iyi_model_kod
        satirlar.append(row_dict)
        
    en_iyi_karma_model = min(karma_mape, key=karma_mape.get)
    karma_row = {"Alt Endeks": "📌 AĞIRLIKLI KARMA Pn HATASI", "Sözleşme Ağırlığı": "%100"}
    for etiket, _ in modeller: karma_row[f"{etiket} MAPE"] = f"%{karma_mape[etiket]:.2f}"
    karma_row["En Düşük Hata (Önerilen)"] = f"🏆 {en_iyi_karma_model.split('.')[1].strip()} (%{karma_mape[en_iyi_karma_model]:.2f})"
    satirlar.append(karma_row)
    
    return pd.DataFrame(satirlar), detay_gecmis, karma_mape, en_iyi_modeller_dict

def gelismis_endeks_tahmini(df_endeks, ek_ay, model_tipi, alpha=0.65, beta=0.35, phi=0.90, mc_quantile=80, n_sims=1000, df_alt=None):
    df_clean = ayristir_gercek_endeks(df_endeks.copy())
    if df_clean.empty or ek_ay <= 0: return ensure_text_df(df_clean)
        
    end_col = 'AYLAR' if 'AYLAR' in df_clean.columns else df_clean.columns[0]
    parsed_dates = df_clean[end_col].apply(parse_turkish_date)
    df_clean['_period'] = pd.to_datetime(parsed_dates, format='%Y-%m', errors='coerce').dt.to_period('M')
    df_clean = df_clean.dropna(subset=['_period']).drop_duplicates(subset=['_period']).sort_values('_period').reset_index(drop=True)
    
    if df_clean.empty: return ensure_text_df(df_endeks)
        
    son_period = df_clean['_period'].iloc[-1]
    endeks_sutunlari = [c for c in df_clean.columns if c not in [end_col, '_period']]
    
    gelecek_periods = [son_period + k for k in range(1, ek_ay + 1)]
    tahmin_sozluk = {end_col: [period_to_tr_full_str(p) for p in gelecek_periods]}
    
    best_models_dict = {}
    if "Karma" in model_tipi and df_alt is not None:
        _, _, _, best_models_dict = backtest_modelleri_hesapla(df_clean, df_alt, alpha=alpha, beta=beta, phi=phi)

    for col in endeks_sutunlari:
        y = [float(clean_decimal(v)) for v in df_clean[col] if float(clean_decimal(v)) > 0]
        active_model = best_models_dict[col] if ("Karma" in model_tipi and col in best_models_dict) else model_tipi
        preds = tek_seri_tahmin_uret(y, ek_ay, active_model, alpha=alpha, beta=beta, phi=phi, mc_quantile=mc_quantile, n_sims=n_sims)
        tahmin_sozluk[col] = [tr_format(p, 6) for p in preds]
        
    df_tahmin = pd.DataFrame(tahmin_sozluk)
    df_birlesik = pd.concat([df_clean.drop(columns=['_period']), df_tahmin], ignore_index=True)
    return df_birlesik.astype(str).astype("string")

# ==========================================
# 5. RESMİ FİYAT FARKI HESAPLAMA MOTORU (1_idari_hakedis.py İLE BİREBİR AYNI)
# ==========================================
def endeks_satiri_bul(df_endeks, hedef_ay):
    if hedef_ay in df_endeks.index: return df_endeks.loc[hedef_ay]
    gecmis = df_endeks.index[df_endeks.index <= hedef_ay]
    if len(gecmis) > 0: return df_endeks.loc[gecmis.max()]
    return df_endeks.iloc[0]

def hesapla(df_prog, df_endeks, df_alt, df_b, b_sabit_override=None):
    df_prog = filter_empty_rows(df_prog.copy())
    df_endeks = filter_empty_rows(df_endeks.copy())
    df_alt = filter_empty_rows(df_alt.copy())
    df_b = filter_empty_rows(df_b.copy())
    
    if df_prog.empty or df_endeks.empty: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df_prog.columns = df_prog.columns.str.strip()
    
    end_col = 'AYLAR' if 'AYLAR' in df_endeks.columns else 'Aylar'
    df_endeks['AyKodu'] = pd.to_datetime(df_endeks[end_col].apply(parse_turkish_date)).dt.to_period('M')
    df_endeks = df_endeks.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).sort_values('AyKodu').set_index('AyKodu')
    
    df_b['AyKodu'] = pd.to_datetime(df_b['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_b = df_b.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).set_index('AyKodu')
    
    df_prog['AyKodu'] = pd.to_datetime(df_prog['AYLAR'].apply(parse_turkish_date)).dt.to_period('M')
    df_prog = df_prog.dropna(subset=['AyKodu'])

    if df_endeks.empty or df_prog.empty: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    son_endeks_ayi = df_endeks.index.max()
    katsayilar = {str(r['Ağırlık']).strip().lower(): clean_decimal(r['Katsayı']) for _, r in df_alt.iterrows()}
    temel_endeksler = {str(r['Ağırlık']).strip().lower(): clean_decimal(r['Temel Endeks']) for _, r in df_alt.iterrows()}
    
    _default_harita = {'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}
    if 'Endeks Sütunu' in df_alt.columns:
        endeks_haritasi = {str(r['Ağırlık']).strip().lower(): str(r['Endeks Sütunu']).strip() for _, r in df_alt.iterrows() if str(r.get('Endeks Sütunu', '')).strip() not in ['', 'nan', 'None']}
        if not endeks_haritasi: endeks_haritasi = _default_harita
    else: endeks_haritasi = _default_harita

    prog_kum_col = df_prog.columns[1]
    imalat_kum_col = df_prog.columns[2]
    kovalar = []
    onceki_kum = Decimal('0.0')
    
    for _, row in df_prog.iterrows():
        kum = clean_decimal(row[prog_kum_col])
        cap = kum - onceki_kum
        kovalar.append({'ay': row['AyKodu'], 'kapasite': cap if cap > Decimal('0.0') else Decimal('0.0')})
        onceki_kum = kum

    final_ff, matris, aylik_rows = [], [], []
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
            
        if b_sabit_override is not None:
            b_kat = Decimal(str(b_sabit_override))
        else:
            b_val = df_b.loc[uyg_ayi, 'B'] if uyg_ayi in df_b.index else Decimal('1.0')
            b_kat = clean_decimal(b_val)
            if b_kat <= Decimal('0.0'): b_kat = Decimal('1.0')
        
        gercek_end_ayi = min(uyg_ayi, son_endeks_ayi)
        if gercek_end_ayi in df_endeks.index:
            endeks_uyg = df_endeks.loc[gercek_end_ayi]
        else:
            gecmis_aylar = df_endeks.index[df_endeks.index <= gercek_end_ayi]
            endeks_uyg = df_endeks.loc[gecmis_aylar.max()] if len(gecmis_aylar) > 0 else df_endeks.iloc[0]
            
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
                    if comp_ayi in df_endeks.index:
                        endeks_prog = df_endeks.loc[comp_ayi]
                    else:
                        gecmis_p = df_endeks.index[df_endeks.index <= comp_ayi]
                        endeks_prog = df_endeks.loc[gecmis_p.max()] if len(gecmis_p) > 0 else endeks_uyg
                else:
                    endeks_prog = endeks_uyg
                
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
                    'Hakediş Ayı': str(uyg_ayi),
                    'İş Programı (Ödenek) Ayı': str(kova['ay']),
                    'Durum': '⚠️ Gecikmeli (Sabit/Min Endeks)' if gecikme else '✅ Zamanında / Öne Geçme',
                    'Kullanılan Tutar': float(kullanilan),
                    'B Katsayısı': float(b_kat),
                    'Uygulanan Pn (15 Hane)': float(pn),
                    'Fiyat Farkı Tutarı': float(ff_dilim)
                })
                toplam_ff_aylik += ff_dilim
                kova['kapasite'] -= kullanilan
                kalan -= kullanilan
        
        kumulatif_ff += toplam_ff_aylik
        final_ff.append(float(kumulatif_ff))
        aylik_rows.append({
            'Ay': str(uyg_ayi), 'Aylık İmalat': float(aylik), 'Kümülatif İmalat': float(guncel),
            'B Katsayısı': float(b_kat), 'Aylık Fiyat Farkı': float(toplam_ff_aylik), 'Kümülatif Fiyat Farkı': float(kumulatif_ff)
        })
        onceki_imalat_kum = guncel

    df_sonuc = df_prog.copy()
    df_sonuc['KÜMÜLATİF FİYAT FARKI'] = final_ff
    df_detay = pd.DataFrame(matris)
    df_aylik = pd.DataFrame(aylik_rows)
    
    if not df_detay.empty:
        df_pivot = df_detay.pivot_table(index='Hakediş Ayı', columns='İş Programı (Ödenek) Ayı', values='Kullanılan Tutar', aggfunc='sum', fill_value=0)
        df_pivot['HAKEDİŞ TUTARI (Toplam)'] = df_pivot.sum(axis=1)
        df_pivot.loc['ÖDENEK MİKTARI'] = df_pivot.sum()
    else: df_pivot = pd.DataFrame()
        
    return df_sonuc, df_pivot, df_detay, df_aylik

# ==========================================
# 6. KAPSAMLI EXCEL (7 SAYFA) & RESMİ WORD RAPORU MOTORU
# ==========================================
def generate_excel_download(df_prog, df_endeks, df_alt, df_b, df_detay=None, df_karsilastirma=None, df_backtest=None):
    wb = Workbook()
    thin = Side(style='thin', color='B0C4DE')
    def brd(): return Border(top=thin, left=thin, right=thin, bottom=thin)
    def fill(c): return PatternFill('solid', fgColor=c)
    def fnt(c='1A1A2E', bold=False, sz=10): return Font(color=c, bold=bold, size=sz, name='Calibri')
    def aln(h='left', wrap=False): return Alignment(horizontal=h, vertical='center', wrap_text=wrap)
    HEADER = '1E3A8A'; WHITE = 'FFFFFF'; YELLOW = 'FFFDE7'; NOTE = 'FFF9C4'; ZEBRA = 'F8FAFC'
    
    def make_sheet(ws, title, note, df, widths, is_input=True):
        ws.sheet_view.showGridLines = False
        max_col = max(len(df.columns), 4)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        ws['A1'] = title; ws['A1'].fill = fill(HEADER); ws['A1'].font = fnt(WHITE, True, 12); ws['A1'].alignment = aln('center')
        ws.row_dimensions[1].height = 24
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        ws['A2'] = note; ws['A2'].fill = fill(NOTE); ws['A2'].font = fnt('7B5800', False, 9); ws['A2'].alignment = aln('center')
        ws.row_dimensions[2].height = 16
        for ci, (h, w) in enumerate(zip(df.columns, widths), 1):
            c = ws.cell(3, ci, str(h))
            c.fill = fill(HEADER); c.font = fnt(WHITE, True, 9); c.alignment = aln('center', True); c.border = brd()
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[3].height = 20
        for ri, row in enumerate(df.values, 4):
            row_bg = YELLOW if is_input else (WHITE if ri % 2 == 0 else ZEBRA)
            for ci, val in enumerate(row, 1):
                c = ws.cell(ri, ci, str(val) if pd.notna(val) and str(val) != 'nan' else '')
                c.fill = fill(row_bg); c.font = fnt('1A1A2E', ri == len(df.values) + 3 and not is_input, 9)
                c.alignment = aln('right' if ci > 1 else 'left'); c.border = brd()

    ws1 = wb.active; ws1.title = 'IsProgrami'
    make_sheet(ws1, '1. İŞ PROGRAMI VE İMALATLAR', 'Sarı hücrelere kümülatif verileri girin', df_prog, [18, 28, 28], True)
    ws2 = wb.create_sheet('Endeks')
    make_sheet(ws2, '2. ENDEKS TABLOSU (GEÇMİŞ VE TAHMİN)', 'TÜİK / Hakedis.org endeksleri ve tahmin edilen aylar', df_endeks, [18] + [14]*(len(df_endeks.columns)-1), True)
    ws3 = wb.create_sheet('AltEndeks')
    make_sheet(ws3, '3. ALT ENDEKS AĞIRLIKLARI', 'Katsayılar toplamı 1,00 olmalıdır.', df_alt, [14, 14, 18, 16], True)
    ws4 = wb.create_sheet('B')
    make_sheet(ws4, '4. B KATSAYISI TABLOSU', 'B katsayıları', df_b, [18, 14], True)
    if df_detay is not None and not df_detay.empty:
        ws5 = wb.create_sheet('5_Hakedis_Dilim_Matrisi')
        df_d_fmt = df_detay.copy()
        df_d_fmt['Kullanılan Tutar'] = df_d_fmt['Kullanılan Tutar'].apply(tr_format)
        df_d_fmt['Fiyat Farkı Tutarı'] = df_d_fmt['Fiyat Farkı Tutarı'].apply(tr_format)
        df_d_fmt['Uygulanan Pn (15 Hane)'] = df_d_fmt['Uygulanan Pn (15 Hane)'].apply(lambda x: f"{float(x):.6f}".replace('.', ','))
        make_sheet(ws5, '5. HAKEDİŞ DİLİM DETAY VE GECİKME MATRİSİ', 'Kova sistemine göre eşleştirilen ödenek dilimleri', df_d_fmt, [16, 22, 26, 20, 14, 22, 22], False)
    if df_karsilastirma is not None and not df_karsilastirma.empty:
        ws6 = wb.create_sheet('6_Model_Kiyaslama')
        make_sheet(ws6, '6. DÖRT TAHMİN MODELİNİN FİYAT FARKI KIYASLAMASI', 'Toplam maliyet projeksiyonları', df_karsilastirma, [38, 22, 26, 20, 26], False)
    if df_backtest is not None and not df_backtest.empty:
        ws7 = wb.create_sheet('7_Backtest_Dogruluk')
        make_sheet(ws7, '7. WALK-FORWARD BACKTESTING (MAPE & RMSE DOĞRULUK ANALİZİ)', 'Hata oranları', df_backtest, [28, 16, 26, 26, 26, 26, 28], False)
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def generate_word_report(aktif_model, toplam_sozlesme, toplam_iml, toplam_ff, mevcut_ff, kalan_ff, df_karsilastirma, df_backtest, df_detay):
    tarih_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    genel_toplam = toplam_iml + toplam_ff
    ff_oran = (toplam_ff / toplam_iml * 100) if toplam_iml > 0 else 0.0

    def df_to_html_table(df):
        if df is None or df.empty: return "<p>Veri bulunamadı.</p>"
        ths = "".join([f"<th style='background-color:#1e3a8a; color:white; padding:6px; border:1px solid #cbd5e1; font-size:9pt;'>{c}</th>" for c in df.columns])
        trs = ""
        for idx, row in df.iterrows():
            bg = "#f8fafc" if idx % 2 == 1 else "#ffffff"
            tds = "".join([f"<td style='padding:5px; border:1px solid #cbd5e1; font-size:9pt; text-align:{'left' if i==0 else 'right'};'>{val}</td>" for i, val in enumerate(row.values)])
            trs += f"<tr style='background-color:{bg};'>{tds}</tr>"
        return f"<table style='border-collapse:collapse; width:100%; margin-top:8px; margin-bottom:14px;'><thead><tr>{ths}</tr></thead><tbody>{trs}</tbody></table>"

    df_det_short = df_detay.copy() if df_detay is not None and not df_detay.empty else pd.DataFrame()
    if not df_det_short.empty:
        df_det_short['Kullanılan Tutar'] = df_det_short['Kullanılan Tutar'].apply(tr_format)
        df_det_short['Fiyat Farkı Tutarı'] = df_det_short['Fiyat Farkı Tutarı'].apply(tr_format)
        df_det_short['B Katsayısı'] = df_det_short['B Katsayısı'].apply(lambda x: tr_format(x, 2))
        df_det_short['Uygulanan Pn (15 Hane)'] = df_det_short['Uygulanan Pn (15 Hane)'].apply(lambda x: f"{float(x):.6f}".replace('.', ','))

    html_content = f"""
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head><meta charset="utf-8"><title>Teknik Ofis Fiyat Farkı ve Endeks Projeksiyon Raporu</title>
    <style>body {{ font-family: 'Calibri', 'Arial', sans-serif; color: #1e293b; line-height: 1.45; font-size: 10.5pt; }}
    h1 {{ color: #1e3a8a; font-size: 16pt; border-bottom: 2px solid #1e3a8a; padding-bottom: 4px; }}
    h2 {{ color: #1e40af; font-size: 12.5pt; margin-top: 18px; border-left: 4px solid #1e40af; padding-left: 8px; }}
    .kpi-box {{ background-color: #f1f5f9; border: 1px solid #cbd5e1; padding: 10px; margin-bottom: 14px; }}</style></head>
    <body>
        <h1>TEKNİK OFİS FİYAT FARKI TAHMİN VE ENDEKS SİMÜLASYON RAPORU</h1>
        <p><b>Rapor Tarihi:</b> {tarih_str} &nbsp;|&nbsp; <b>Aktif Tahmin Algoritması:</b> {aktif_model}</p>
        <h2>1. YÖNETİCİ ÖZETİ VE BÜTÇE PROJEKSİYONU</h2>
        <div class="kpi-box"><ul>
            <li><b>Sözleşme Bedeli (BAC):</b> {tr_format(toplam_sozlesme)} TL</li>
            <li><b>Hesaplanan Kümülatif İmalat Tutarı:</b> {tr_format(toplam_iml)} TL</li>
            <li><b>Mevcut Duruma Kadar Fiili Fiyat Farkı (İdari Hakediş):</b> {tr_format(mevcut_ff)} TL</li>
            <li><b>Kalan / Türetilen Ay Fiyat Farkı:</b> {tr_format(kalan_ff)} TL</li>
            <li><b>Proje Sonu Toplam Fiyat Farkı (FF):</b> <b>{tr_format(toplam_ff)} TL</b> (İmalata Oranı: %{ff_oran:.2f})</li>
            <li><b>Genel Hakediş Toplamı (İmalat + Fiyat Farkı):</b> <b>{tr_format(genel_toplam)} TL</b></li>
        </ul></div>
        <h2>2. DÖRT TAHMİN MODELİNİN KARŞILAŞTIRMALI MALİYET TABLOSU</h2>{df_to_html_table(df_karsilastirma)}
        <h2>3. MODEL DOĞRULUK TESTİ (WALK-FORWARD BACKTESTING)</h2>{df_to_html_table(df_backtest)}
        <h2>4. HAKEDİŞ ÖDENEK DİLİMİ (KOVA) VE GECİKME MATRİSİ DETAYI</h2>{df_to_html_table(df_det_short)}
    </body></html>"""
    return html_content.encode('utf-8')

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
                    if 'AYLAR' in [str(v).upper() for v in row.values] or 'AĞIRLIK' in [str(v).upper() for v in row.values]:
                        skip = i + 1; break
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=skip)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            if sheet == 'AltEndeks' and 'Endeks Sütunu' not in df.columns:
                df['Endeks Sütunu'] = df['Ağırlık'].astype(str).str.strip().str.lower().map({'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}).fillna('')
            dfs[sheet_map[sheet]] = clean_df_for_ui(df)
    return dfs

# ==========================================
# 7. SESSION STATE BAŞLATMA
# ==========================================
if 'ff_load_count' not in st.session_state: st.session_state.ff_load_count = 0
if 'ff_prog_df' not in st.session_state: st.session_state.ff_prog_df = clean_df_for_ui(pd.DataFrame(DEFAULT_PROG))
if 'ff_base_endeks_df' not in st.session_state: st.session_state.ff_base_endeks_df = clean_df_for_ui(pd.DataFrame(DEFAULT_ENDEKS))
if 'ff_alt_df' not in st.session_state: st.session_state.ff_alt_df = clean_df_for_ui(pd.DataFrame(DEFAULT_ALT))
if 'ff_b_df' not in st.session_state: st.session_state.ff_b_df = clean_df_for_ui(pd.DataFrame(DEFAULT_B))
if 'ff_ek_ay_slider' not in st.session_state: st.session_state.ff_ek_ay_slider = 10

# ==========================================
# 8. YAN MENÜ (SIDEBAR): DOSYA & TAHMİN KONTROLLERİ
# ==========================================
with st.sidebar:
    st.header("📥 Proje Yükle (Excel / JSON)")
    uploaded_file = st.file_uploader("Dosya Seç (.xlsx veya .json)", type=["xlsx", "json"])
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        if st.session_state.get('ff_last_bytes') != file_bytes:
            try:
                if uploaded_file.name.endswith('.xlsx'):
                    dfs = load_from_excel(uploaded_file)
                    if 'prog_df' in dfs: st.session_state.ff_prog_df = dfs['prog_df']
                    if 'endeks_df' in dfs: st.session_state.ff_base_endeks_df = ayristir_gercek_endeks(dfs['endeks_df'])
                    if 'alt_df' in dfs: st.session_state.ff_alt_df = dfs['alt_df']
                    if 'b_df' in dfs: st.session_state.ff_b_df = dfs['b_df']
                else:
                    data = json.load(uploaded_file)
                    if 'prog' in data: st.session_state.ff_prog_df = clean_df_for_ui(pd.DataFrame(data['prog']))
                    df_full_e = pd.DataFrame(data['endeks']) if 'endeks' in data else pd.DataFrame()
                    df_base_e = pd.DataFrame(data['base_endeks']) if 'base_endeks' in data else df_full_e
                    st.session_state.ff_base_endeks_df = clean_df_for_ui(ayristir_gercek_endeks(df_base_e))
                    if 'alt' in data:
                        df_a = pd.DataFrame(data['alt'])
                        if 'Endeks Sütunu' not in df_a.columns:
                            df_a['Endeks Sütunu'] = df_a['Ağırlık'].astype(str).str.strip().str.lower().map({'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}).fillna('')
                        st.session_state.ff_alt_df = clean_df_for_ui(df_a)
                    if 'b' in data: st.session_state.ff_b_df = clean_df_for_ui(pd.DataFrame(data['b']))
                
                st.session_state.ff_ek_ay_slider = hesapla_gereken_ufuk(st.session_state.ff_prog_df, st.session_state.ff_base_endeks_df)
                st.session_state.ff_load_count += 1
                st.session_state.ff_last_bytes = file_bytes
                st.session_state.ff_last_model_sig = None
                st.rerun()
            except Exception as e:
                st.error(f"Yükleme hatası: {e}")

    if st.button("🔄 Okul Projesi Varsayılanlarına Dön", use_container_width=True):
        st.session_state.ff_prog_df = clean_df_for_ui(pd.DataFrame(DEFAULT_PROG))
        st.session_state.ff_base_endeks_df = clean_df_for_ui(pd.DataFrame(DEFAULT_ENDEKS))
        st.session_state.ff_alt_df = clean_df_for_ui(pd.DataFrame(DEFAULT_ALT))
        st.session_state.ff_b_df = clean_df_for_ui(pd.DataFrame(DEFAULT_B))
        st.session_state.ff_ek_ay_slider = 10
        st.session_state.ff_last_model_sig = None
        st.session_state.ff_load_count += 1
        st.rerun()

    st.markdown("---")
    st.header("🌐 Web Endeks & Tahmin Motoru")
    if st.button("📡 Hakedis.org'dan Endeks Çek", use_container_width=True):
        with st.spinner("Hakedis.org taranıyor ve doğrulanıyor..."):
            df_guncel, basari, mesaj, df_ham = webden_endeks_cek_ve_guncelle(st.session_state.ff_base_endeks_df)
            if basari:
                st.session_state.ff_base_endeks_df = df_guncel
                st.session_state.ff_ek_ay_slider = hesapla_gereken_ufuk(st.session_state.ff_prog_df, st.session_state.ff_base_endeks_df)
                st.session_state.ff_last_model_sig = None
                st.session_state.ff_load_count += 1
                st.session_state.ff_web_mesaj = (basari, mesaj)
                st.session_state.ff_web_ham = df_ham
                st.rerun()

    if 'ff_web_mesaj' in st.session_state:
        basari_m, metin_m = st.session_state.ff_web_mesaj
        if basari_m: st.success(f"✅ {metin_m}")
        else: st.warning(f"ℹ️ {metin_m}")
        if 'ff_web_ham' in st.session_state and st.session_state.ff_web_ham is not None:
            with st.expander("🔍 Çekilen Ham Veriyi Teyit Et", expanded=False):
                st.markdown("[🌐 Hakedis.org Kaynak Tablosuna Git](https://www.hakedis.org/endeksler/yapim-isleri-fiyat-farki-endeksleri-kasim-2013-sonrasi)")
                st.caption("Botun siteden okuduğu saf (ölçeklenmemiş) son 6 aylık veri:")
                st.dataframe(st.session_state.ff_web_ham.tail(6), use_container_width=True)

    st.subheader("🔮 Gelecek Endeks Tahmini")
    model_secimi = st.selectbox(
        "İstatistiksel / Stokastik Model:",
        [
            "5. En İyi Karma (Alt Endekslere Göre Dinamik) 🌟",
            "1. Holt's Çift Üstel Düzleştirme (Level + Trend)",
            "2. Sönümlü Trend (Damped Trend - Dezenflasyon)",
            "3. Stokastik Monte Carlo / GBM (Olasılıksal)",
            "4. Otoregresif Momentum (AR-1 & EWMA)",
            "0. Tahmin Yok (Son Bilinen Endeks Sabit Kalsın)"
        ]
    )
    ek_ay = st.slider(
        "Tahmin Ufku (Ay)",
        min_value=0, max_value=36,
        key="ff_ek_ay_slider"
    )
    
    alpha_p, beta_p, phi_p, mc_q = 0.65, 0.35, 0.90, 80
    if "Holt" in model_secimi or "Karma" in model_secimi:
        c_p1, c_p2 = st.columns(2)
        alpha_p = c_p1.slider("Seviye (α)", 0.1, 0.95, 0.65, 0.05)
        beta_p = c_p2.slider("Trend (β)", 0.05, 0.90, 0.35, 0.05)
    if "Sönümlü" in model_secimi or "Karma" in model_secimi:
        phi_p = st.slider("Sönümleme Katsayısı (ϕ)", 0.75, 0.98, 0.90, 0.01)
    if "Monte Carlo" in model_secimi or "Karma" in model_secimi:
        mc_q = st.select_slider("Güven Aralığı (Persentil)", options=[50, 70, 80, 90, 95], value=80)

    current_sig = f"{model_secimi}_{ek_ay}_{alpha_p}_{beta_p}_{phi_p}_{mc_q}_{len(st.session_state.ff_base_endeks_df)}"
    if (
        'ff_endeks_df' not in st.session_state
        or st.session_state.get('ff_last_model_sig') != current_sig
        or (ek_ay > 0 and not model_secimi.startswith("0.") and len(st.session_state.ff_endeks_df) <= len(st.session_state.ff_base_endeks_df))
    ):
        if model_secimi.startswith("0.") or ek_ay == 0:
            st.session_state.ff_endeks_df = clean_df_for_ui(st.session_state.ff_base_endeks_df.copy())
        else:
            st.session_state.ff_endeks_df = gelismis_endeks_tahmini(
                st.session_state.ff_base_endeks_df, ek_ay, model_secimi, alpha=alpha_p, beta=beta_p, phi=phi_p, mc_quantile=mc_q, df_alt=st.session_state.ff_alt_df
            )
        st.session_state.ff_last_model_sig = current_sig
        st.session_state.ff_load_count += 1

# ==========================================
# 9. ANA EKRAN VE 5 GENİŞ SEKME
# ==========================================
st.title("📈 Fiyat Farkı Tahmin & Stokastik Endeks Simülatörü")
st.caption("Dinamik Karma Optimizasyon (5. Model) · Sabitlenen İş Programı Ayında Endeks Sabitleme Kuralı")

suffix = st.session_state.ff_load_count

main_tab1, main_tab2, main_tab3, main_tab4, main_tab5 = st.tabs([
    "📋 1. İş Programı & Katsayılar",
    "📈 2. Endeks Kütüğü & Tahmin Eğrileri",
    "🧮 3. Fiyat Farkı Sonuçları & Matris",
    "⚖️ 4. Model & Senaryo Karşılaştırma",
    "📚 5. Akademik Metodoloji & Backtesting (MAPE/RMSE)"
])

# ---------------- SEKME 1: İŞ PROGRAMI VE KATSAYILAR ----------------
with main_tab1:
    st.subheader("1️⃣ İş Programı ve Kümülatif İmalat Tablosu (Tam Ekran)")
    
    col_sim_opt1, col_sim_opt2 = st.columns([1.8, 1.2])
    with col_sim_opt1:
        gelecek_imalat_modu = st.radio(
            "Hesaplama Kapsamı Seçimi:",
            [
                "Mevcut Duruma Kadar Hesapla + Kalan Bakiyeyi Tamamla (Kalan Ay Yoksa +1 Ay Türet)",
                "Sadece Tablodaki Fiili İmalatı Hesapla (1_idari_hakedis.py İle Birebir Aynı Tablo)"
            ],
            index=0,
            horizontal=False
        )
    with col_sim_opt2:
        b_override_check = st.checkbox("B Katsayısını Tüm Aylarda Sabitle (Simülasyon)", value=False)
        b_override_val = st.slider("Simüle Edilen B Katsayısı", 0.50, 1.30, 0.90, 0.05) if b_override_check else None

    tamamlama_aktif = gelecek_imalat_modu.startswith("Mevcut Duruma Kadar")
    df_prog_working, df_b_working, turetilen_bilgi = hazirla_prog_ve_turetilen_ay(
        st.session_state.ff_prog_df, st.session_state.ff_b_df, tamamlama_aktif=tamamlama_aktif
    )

    if turetilen_bilgi is not None:
        st.info(f"✨ **Otomatik Tamamlama Bilgisi:** {turetilen_bilgi['aciklama']}")

    edited_prog = st.data_editor(
        df_prog_working,
        column_config=get_text_config(df_prog_working),
        num_rows="dynamic",
        use_container_width=True,
        height=550,
        key=f"ff_prog_ed_{suffix}_{gelecek_imalat_modu[:8]}"
    )

    st.markdown("---")
    c_alt, c_b = st.columns([1.3, 1])
    with c_alt:
        st.subheader("2️⃣ Alt Endeks Ağırlıkları ve Temel Endeksler")
        edited_alt = st.data_editor(
            st.session_state.ff_alt_df,
            column_config=get_text_config(st.session_state.ff_alt_df),
            num_rows="dynamic",
            use_container_width=True,
            height=310,
            key=f"ff_alt_ed_{suffix}"
        )

    with c_b:
        st.subheader("3️⃣ Aylık B Katsayısı Tablosu")
        edited_b = st.data_editor(
            df_b_working,
            column_config=get_text_config(df_b_working),
            num_rows="dynamic",
            use_container_width=True,
            height=310,
            key=f"ff_b_ed_{suffix}_{gelecek_imalat_modu[:8]}"
        )

# ---------------- SEKME 2: ENDEKS KÜTÜĞÜ VE TAHMİN EĞRİLERİ ----------------
with main_tab2:
    baz_df = st.session_state.ff_base_endeks_df.copy()
    
    if baz_df.empty or len(baz_df.columns) <= 1:
        st.warning("⚠️ Grafiği çizmek için yeterli endeks verisi bulunamadı. Lütfen bir proje yükleyin.")
        df_endeks_auto = pd.DataFrame()
    else:
        gereken_min_ufuk = hesapla_gereken_ufuk(edited_prog, baz_df)
        aktif_ufuk = max(ek_ay, gereken_min_ufuk)
        if not model_secimi.startswith("0.") and aktif_ufuk > ek_ay:
            df_endeks_auto = gelismis_endeks_tahmini(
                baz_df, aktif_ufuk, model_secimi, alpha=alpha_p, beta=beta_p, phi=phi_p, mc_quantile=mc_q, df_alt=edited_alt
            )
        else:
            df_endeks_auto = st.session_state.ff_endeks_df

        st.subheader(f"📈 TÜİK Endeks Kütüğü — Aktif Model: {model_secimi}")
        
        edited_endeks = st.data_editor(
            df_endeks_auto,
            column_config=get_text_config(df_endeks_auto),
            num_rows="dynamic",
            use_container_width=True,
            height=520,
            key=f"ff_end_ed_{suffix}_{aktif_ufuk}"
        )

        st.markdown("---")
        st.subheader("🔬 Seçili Alt Endeks İçin Tahmin Modellerinin Karşılaştırmalı Projeksiyonu")
        secili_endeks_kodu = st.selectbox(
            "Grafikte İncelenecek Alt Endeks:",
            options=list(KOD_BILGI.keys()),
            format_func=lambda k: f"{KOD_BILGI[k]['kolon']} — {KOD_BILGI[k]['ad']} ({KOD_BILGI[k]['resmi_kod']})"
        )
        secili_kolon = KOD_BILGI[secili_endeks_kodu]['kolon']

        if secili_kolon in baz_df.columns:
            ufuk_gosterim = max(aktif_ufuk, 10)
            df_m1 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Holt", alpha=alpha_p, beta=beta_p, df_alt=edited_alt)
            df_m2 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Sönümlü", alpha=alpha_p, beta=beta_p, phi=phi_p, df_alt=edited_alt)
            df_m3 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Monte Carlo", mc_quantile=mc_q, df_alt=edited_alt)
            df_m4 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Otoregresif", df_alt=edited_alt)

            fig_end = go.Figure()
            n_hist = len(baz_df)
            ay_col = 'AYLAR' if 'AYLAR' in df_m1.columns else df_m1.columns[0]
            
            fig_end.add_trace(go.Scatter(
                x=df_m1[ay_col][:n_hist],
                y=[float(clean_decimal(v)) for v in df_m1[secili_kolon][:n_hist]],
                mode='lines+markers', name='Gerçekleşen TÜİK Verisi', line=dict(color='#1f2937', width=4)
            ))
            fig_end.add_trace(go.Scatter(
                x=df_m1[ay_col][n_hist-1:],
                y=[float(clean_decimal(v)) for v in df_m1[secili_kolon][n_hist-1:]],
                mode='lines+markers', name="1. Holt's Çift Üstel", line=dict(color='#2563eb', width=2.5, dash='dash')
            ))
            fig_end.add_trace(go.Scatter(
                x=df_m2[ay_col][n_hist-1:],
                y=[float(clean_decimal(v)) for v in df_m2[secili_kolon][n_hist-1:]],
                mode='lines+markers', name=f"2. Sönümlü Trend (ϕ={phi_p})", line=dict(color='#10b981', width=2.5, dash='dot')
            ))
            fig_end.add_trace(go.Scatter(
                x=df_m3[ay_col][n_hist-1:],
                y=[float(clean_decimal(v)) for v in df_m3[secili_kolon][n_hist-1:]],
                mode='lines+markers', name=f"3. Monte Carlo GBM (P{mc_q})", line=dict(color='#dc2626', width=2.5, dash='dashdot')
            ))
            fig_end.add_trace(go.Scatter(
                x=df_m4[ay_col][n_hist-1:],
                y=[float(clean_decimal(v)) for v in df_m4[secili_kolon][n_hist-1:]],
                mode='lines+markers', name="4. Otoregresif Momentum (AR-1)", line=dict(color='#9333ea', width=2.5)
            ))
            fig_end.update_layout(
                title=f"{KOD_BILGI[secili_endeks_kodu]['ad']} — Gelecek Tahmin Eğrileri",
                xaxis_title="Aylar", yaxis_title="Endeks Değeri", template="simple_white", height=450, hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_end, use_container_width=True)

# ---------------- SEKME 3: FİYAT FARKI SONUÇLARI, MATRİS & MEVCUT DURUM AYRIŞTIRMASI ----------------
with main_tab3:
    if 'edited_endeks' not in locals():
        edited_endeks = pd.DataFrame()
        
    df_sonuc, df_pivot, df_detay, df_aylik = hesapla(
        edited_prog, edited_endeks, edited_alt, edited_b, b_sabit_override=b_override_val
    )
    df_sonuc_ham, _, _, df_aylik_ham = hesapla(
        st.session_state.ff_prog_df, st.session_state.ff_base_endeks_df, edited_alt, st.session_state.ff_b_df, b_sabit_override=b_override_val
    )
    
    if df_detay.empty:
        st.warning("⚠️ Henüz hesaplanacak geçerli imalat tutarı veya tarih eşleşmesi bulunamadı.")
        toplam_sozlesme, toplam_gerceklesen_iml, toplam_ff_tutari, mevcut_ff_tutari, kalan_ff_tutari = 0.0, 0.0, 0.0, 0.0, 0.0
    else:
        toplam_sozlesme = float(clean_decimal(edited_prog.iloc[-1, 1]))
        toplam_gerceklesen_iml = float(df_aylik['Aylık İmalat'].sum())
        toplam_ff_tutari = float(df_sonuc['KÜMÜLATİF FİYAT FARKI'].iloc[-1])
        
        mevcut_iml_tutari = float(df_aylik_ham['Aylık İmalat'].sum()) if not df_aylik_ham.empty else 0.0
        mevcut_ff_tutari = float(df_sonuc_ham['KÜMÜLATİF FİYAT FARKI'].iloc[-1]) if not df_sonuc_ham.empty else 0.0
        
        kalan_iml_tutari = toplam_gerceklesen_iml - mevcut_iml_tutari
        kalan_ff_tutari = toplam_ff_tutari - mevcut_ff_tutari
        genel_toplam_hakedis = toplam_gerceklesen_iml + toplam_ff_tutari

        m1, m2, m3 = st.columns(3)
        m1.metric("1️⃣ Mevcut Duruma Kadar Fiili FF", f"{tr_format(mevcut_ff_tutari)} TL", f"Mevcut İmalat: {tr_format(mevcut_iml_tutari)} TL")
        m2.metric("2️⃣ Kalan / Gelecek Aylar Tahmini FF", f"{tr_format(kalan_ff_tutari)} TL", f"Kalan İmalat: {tr_format(kalan_iml_tutari)} TL")
        m3.metric("3️⃣ Proje Sonu Toplam Fiyat Farkı", f"{tr_format(toplam_ff_tutari)} TL", f"Genel Toplam (İmalat+FF): {tr_format(genel_toplam_hakedis)} TL")

        st.markdown("---")
        sub_t1, sub_t2, sub_t3, sub_t4 = st.tabs(["📊 Analiz Grafiği", "🔍 Dilim Detay (Pn)", "🧮 Teyit Matrisi (Pivot)", "📑 Kümülatif Tablo"])
        
        with sub_t1:
            fig_ff = go.Figure()
            fig_ff.add_trace(go.Bar(x=df_aylik['Ay'], y=df_aylik['Aylık İmalat'], name='Aylık İmalat (TL)', marker_color='#3b82f6'))
            fig_ff.add_trace(go.Bar(x=df_aylik['Ay'], y=df_aylik['Aylık Fiyat Farkı'], name='Aylık Fiyat Farkı (TL)', marker_color='#ef4444'))
            fig_ff.add_trace(go.Scatter(x=df_aylik['Ay'], y=df_aylik['Kümülatif Fiyat Farkı'], name='Kümülatif Fiyat Farkı (Sağ Eksen)', yaxis='y2', mode='lines+markers', line=dict(color='#10b981', width=3)))
            fig_ff.update_layout(title="Aylık İmalat, Fiyat Farkı ve Kümülatif Fiyat Farkı Gelişimi", yaxis=dict(title="Aylık Tutar (TL)"), yaxis2=dict(title="Kümülatif FF (TL)", overlaying='y', side='right'), barmode='group', template='simple_white', height=460, hovermode='x unified', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig_ff, use_container_width=True)

        with sub_t2:
            df_det_show = df_detay.copy()
            df_det_show['Kullanılan Tutar'] = df_det_show['Kullanılan Tutar'].apply(tr_format)
            df_det_show['Fiyat Farkı Tutarı'] = df_det_show['Fiyat Farkı Tutarı'].apply(tr_format)
            df_det_show['B Katsayısı'] = df_det_show['B Katsayısı'].apply(lambda x: tr_format(x, 2))
            df_det_show['Uygulanan Pn (15 Hane)'] = df_det_show['Uygulanan Pn (15 Hane)'].apply(lambda x: f"{x:.15f}".rstrip('0').rstrip('.').replace('.', ','))
            st.dataframe(df_det_show, use_container_width=True, height=450)

        with sub_t3:
            if not df_pivot.empty:
                st.dataframe(df_pivot.map(tr_format).style.set_properties(subset=['HAKEDİŞ TUTARI (Toplam)'], **{'font-weight': 'bold', 'background-color': '#e6f2ff'}), use_container_width=True, height=450)

        with sub_t4:
            df_sonuc_gosterim = df_sonuc.copy()
            if 'AyKodu' in df_sonuc_gosterim.columns: df_sonuc_gosterim['AyKodu'] = df_sonuc_gosterim['AyKodu'].astype(str)
            for col in df_sonuc_gosterim.columns:
                if any(x in col.upper() for x in ['TUTAR', 'PROGRAM', 'FARKI']): df_sonuc_gosterim[col] = df_sonuc_gosterim[col].apply(tr_format)
            st.dataframe(df_sonuc_gosterim, use_container_width=True, height=500)

# ---------------- SEKME 4: MODEL & SENARYO KARŞILAŞTIRMA ----------------
with main_tab4:
    st.subheader("⚖️ Tüm Tahmin Algoritmalarının ve Sabit Endeks Senaryosunun Kıyaslaması")
    st.caption("Mevcut durum + kalan ayların tamamlandığı senaryoda algoritmaların ürettiği Fiyat Farkı (FF) karşılaştırması:")

    df_prog_full, df_b_full, _ = hazirla_prog_ve_turetilen_ay(st.session_state.ff_prog_df, st.session_state.ff_b_df, tamamlama_aktif=True)
    
    if st.session_state.ff_base_endeks_df.empty:
        st.warning("Karşılaştırma yapmak için yeterli endeks verisi yok.")
    else:
        gereken_ufuk_comp = max(st.session_state.ff_ek_ay_slider, hesapla_gereken_ufuk(df_prog_full, st.session_state.ff_base_endeks_df))

        tum_modeller = [
            ("5. En İyi Karma (Dinamik Optimizasyon) 🌟", "Karma"),
            ("1. Holt's Çift Üstel Düzleştirme (Baz Trend)", "Holt"),
            (f"2. Sönümlü Trend (Dezenflasyon, ϕ={phi_p})", "Sönümlü"),
            (f"3. Stokastik Monte Carlo GBM (P{mc_q} İhtiyatlı)", "Monte Carlo"),
            ("4. Otoregresif Momentum (AR-1 & EWMA)", "Otoregresif"),
            ("0. Tahmin Yok (Son Bilinen Endeks Sabit Kalsın)", "Sabit")
        ]

        karsilastirma_list = []
        fig_comp = go.Figure()

        for etiket, m_kod in tum_modeller:
            if m_kod == "Sabit":
                df_end_sim = st.session_state.ff_base_endeks_df.copy()
            else:
                df_end_sim = gelismis_endeks_tahmini(st.session_state.ff_base_endeks_df, gereken_ufuk_comp, m_kod, alpha=alpha_p, beta=beta_p, phi=phi_p, mc_quantile=mc_q, df_alt=edited_alt)
                
            df_snc_sim, _, _, df_ay_sim = hesapla(df_prog_full, df_end_sim, edited_alt, df_b_full, b_sabit_override=b_override_val)
            if not df_ay_sim.empty:
                top_iml = float(df_ay_sim['Aylık İmalat'].sum())
                top_ff = float(df_snc_sim['KÜMÜLATİF FİYAT FARKI'].iloc[-1])
                mevcut_karsilastirma_ff = mevcut_ff_tutari if 'mevcut_ff_tutari' in locals() else 0.0
                kalan_model_ff = top_ff - mevcut_karsilastirma_ff
                karsilastirma_list.append({
                    "Tahmin Modeli / Senaryo": etiket,
                    "Mevcut Durum FF (TL)": tr_format(mevcut_karsilastirma_ff),
                    "Kalan / Gelecek Aylar FF (TL)": tr_format(kalan_model_ff),
                    "Proje Sonu Toplam FF (TL)": tr_format(top_ff),
                    "Genel Toplam (İmalat + FF) (TL)": tr_format(top_iml + top_ff)
                })
                if m_kod != "Sabit":
                    fig_comp.add_trace(go.Scatter(x=df_ay_sim['Ay'], y=df_ay_sim['Kümülatif Fiyat Farkı'], mode='lines+markers', name=etiket))

        df_karsilastirma_out = pd.DataFrame(karsilastirma_list)
        st.dataframe(df_karsilastirma_out.style.apply(lambda x: ['background-color: #f0fdf4; font-weight: bold' if 'Karma' in x['Tahmin Modeli / Senaryo'] else '' for i in x], axis=1), use_container_width=True, hide_index=True)
        fig_comp.update_layout(title="Algoritmalara Göre Kümülatif Fiyat Farkı Eğrileri", xaxis_title="Ay", yaxis_title="Kümülatif Fiyat Farkı (TL)", template="simple_white", height=450, hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        st.plotly_chart(fig_comp, use_container_width=True)

# ---------------- SEKME 5: AKADEMİK METODOLOJİ, TEORİ & BACKTESTING TABLOSU ----------------
with main_tab5:
    st.header("📚 Akademik Metodoloji, Backtesting (MAPE/RMSE) ve Matematiksel Çözüm Rehberi")
    st.markdown("---")
    st.subheader("🎯 1️⃣ Geçmiş TÜİK Verileri Üzerinde Model Doğruluk Testi (Walk-Forward Backtesting Tablosu)")
    st.info("Aşağıdaki Backtesting tablosu, gerçekleşen TÜİK endeks serisinin son 6 ayı üzerinde 4 modelin tahmin hata paylarını (**MAPE %** ve **RMSE**) karşılaştırır.")

    if not st.session_state.ff_base_endeks_df.empty and len(st.session_state.ff_base_endeks_df) > 6:
        df_backtest_out, detay_bt, karma_mape_dict, en_iyi_modeller_dict = backtest_modelleri_hesapla(st.session_state.ff_base_endeks_df, edited_alt, alpha=alpha_p, beta=beta_p, phi=phi_p)
        st.dataframe(df_backtest_out, use_container_width=True, hide_index=True)

        col_bt_left, col_bt_right = st.columns([1.15, 1.85])
        with col_bt_left:
            bt_secili_kod = st.selectbox("Aylık Sapma Dökümü İncelenecek Alt Endeks:", options=list(KOD_BILGI.keys()), format_func=lambda k: f"{KOD_BILGI[k]['kolon']} — {KOD_BILGI[k]['ad']}")
            bt_kolon = KOD_BILGI[bt_secili_kod]['kolon']
            en_iyi_karma = min(karma_mape_dict, key=karma_mape_dict.get)
            st.success(f"🏆 **Sözleşme Karma $P_n$ Doğruluk Lideri:**\n\n**{en_iyi_karma}** — Ağırlıklı Ortalama Hata (MAPE): **%{karma_mape_dict[en_iyi_karma]:.2f}**\n\n*(5. Model (En İyi Karma), bu tablodaki her alt endeksin 🏆 ile belirtilen şampiyon modelini seçerek projeksiyon yapar.)*")

            if bt_kolon in detay_bt:
                bt_d = detay_bt[bt_kolon]
                aylik_sapma_rows = []
                for idx_m, ay_adi in enumerate(bt_d["Aylar"]):
                    g_val = max(float(bt_d["Gerçekleşen"][idx_m]), 1e-6)
                    h_val = bt_d["1. Holt's Çift Üstel"][idx_m]
                    s_val = bt_d["2. Sönümlü Trend (Damped)"][idx_m]
                    mc_val = bt_d["3. Monte Carlo GBM (P50 Medyan)"][idx_m]
                    ar_val = bt_d["4. Otoregresif (AR-1 & EWMA)"][idx_m]
                    aylik_sapma_rows.append({
                        "Test Ayı": ay_adi, "Gerçekleşen TÜİK": tr_format(g_val, 2),
                        "Holt's Tahmini": f"{tr_format(h_val, 2)} (%{(abs(g_val-h_val)/g_val*100):.2f})",
                        "Sönümlü Tahmin": f"{tr_format(s_val, 2)} (%{(abs(g_val-s_val)/g_val*100):.2f})",
                        "Monte Carlo (P50)": f"{tr_format(mc_val, 2)} (%{(abs(g_val-mc_val)/g_val*100):.2f})",
                        "Otoregresif (AR-1)": f"{tr_format(ar_val, 2)} (%{(abs(g_val-ar_val)/g_val*100):.2f})"
                    })
                st.caption(f"**{bt_kolon}** Alt Endeksi İçin Aylık Tahmin ve Mutlak Hata (%) Dökümü:")
                st.dataframe(pd.DataFrame(aylik_sapma_rows), use_container_width=True, hide_index=True)

        with col_bt_right:
            if bt_kolon in detay_bt:
                fig_bt = go.Figure()
                bt_data = detay_bt[bt_kolon]
                fig_bt.add_trace(go.Scatter(x=bt_data["Aylar"], y=bt_data["Gerçekleşen"], mode='lines+markers', name='Gerçekleşen TÜİK Verisi', line=dict(color='#111827', width=4)))
                for m_etiket, m_color in [("1. Holt's Çift Üstel", "#2563eb"), ("2. Sönümlü Trend (Damped)", "#10b981"), ("3. Monte Carlo GBM (P50 Medyan)", "#dc2626"), ("4. Otoregresif (AR-1 & EWMA)", "#9333ea")]:
                    if m_etiket in bt_data:
                        is_winner = en_iyi_modeller_dict.get(bt_kolon) == m_etiket.split(".")[1].strip()
                        fig_bt.add_trace(go.Scatter(x=bt_data["Aylar"], y=bt_data[m_etiket], mode='lines+markers', name=m_etiket, line=dict(color=m_color, width=3 if is_winner else 1.5, dash='solid' if is_winner else 'dot')))
                fig_bt.update_layout(title=f"{bt_kolon} — Gerçekleşen TÜİK vs Backtest (Kalın Çizgi = Şampiyon)", xaxis_title="Test Ayları (Son 6 Ay)", yaxis_title="Endeks Değeri", template="simple_white", height=420, hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                st.plotly_chart(fig_bt, use_container_width=True)
    else:
        st.warning("⚠️ Backtest analizi için geçmişe dönük en az 7 aylık gerçekleşmiş TÜİK verisi gereklidir.")

    st.markdown("---")
    st.subheader("2️⃣ 4734 Sayılı Kanun Fiyat Farkı Kararnamesi, Kova Matematiği ve Genel Toplam")
    st.latex(r"F = A_n \times B \times (P_n - 1)")
    st.latex(r"P_n = a \frac{I_n}{I_0} + b_1 \frac{\text{Ç}_n}{\text{Ç}_0} + b_2 \frac{D_n}{D_0} + b_3 \frac{Y_n}{Y_0} + b_4 \frac{K_n}{K_0} + b_5 \frac{G_n}{G_0} + c \frac{M_n}{M_0}")

    st.markdown("---")
    st.subheader("3️⃣ Dört Tahmin Modelinin Adım Adım Matematiksel Altyapısı")
    col_m_left, col_m_right = st.columns(2)
    with col_m_left:
        st.markdown("#### 🔹 Model 1: Holt's Çift Üstel Düzleştirme (Level + Trend)")
        st.latex(r"L_t = \alpha Y_t + (1 - \alpha)(L_{t-1} + T_{t-1}), \quad T_t = \beta (L_t - L_{t-1}) + (1 - \beta) T_{t-1}, \quad \hat{Y}_{t+h} = L_t + h \cdot T_t")
        st.markdown("#### 🔹 Model 3: Stokastik Monte Carlo / Geometrik Brownian Hareketi (GBM)")
        st.latex(r"Y_{t+1}^{(k)} = Y_{t}^{(k)} \cdot \exp\left[\left(\mu - \frac{1}{2}\sigma^2\right) + \sigma Z_{t+1}^{(k)}\right], \quad \hat{Y}_{t+h} = \text{Percentile}_{q}\left(\{Y_{t+h}^{(k)}\}\right)")
    with col_m_right:
        st.markdown("#### 🔹 Model 2: Sönümlü Trend (Damped Trend - Dezenflasyon)")
        st.latex(r"L_t = \alpha Y_t + (1 - \alpha)(L_{t-1} + \phi T_{t-1}), \quad \hat{Y}_{t+h} = L_t + \left(\sum_{j=1}^{h} \phi^j\right) T_t")
        st.markdown("#### 🔹 Model 4: Otoregresif Momentum (AR-1 & EWMA)")
        st.latex(r"r_{t+h} = \rho \cdot r_{t+h-1} + (1 - \rho) \cdot r_{\text{ewma}}, \quad \hat{Y}_{t+h} = \hat{Y}_{t+h-1} \cdot (1 + r_{t+h})")

# ---------------- YAN MENÜ İNDİRME VE RESMİ RAPORLAMA BUTONLARI ----------------
with st.sidebar:
    st.markdown("---")
    st.header("📤 Resmi Rapor & Veri Çıktısı")
    
    word_bytes = generate_word_report(
        aktif_model=model_secimi,
        toplam_sozlesme=toplam_sozlesme,
        toplam_iml=toplam_gerceklesen_iml,
        toplam_ff=toplam_ff_tutari,
        mevcut_ff=mevcut_ff_tutari,
        kalan_ff=kalan_ff_tutari,
        df_karsilastirma=df_karsilastirma_out,
        df_backtest=df_backtest_out if 'df_backtest_out' in locals() else pd.DataFrame(),
        df_detay=df_detay
    )
    st.download_button(label="📄 Resmi Word Raporu İndir (.doc)", data=word_bytes, file_name="Teknik_Ofis_Fiyat_Farki_Raporu.doc", mime="application/msword", use_container_width=True, type="primary")

    c_dl1, c_dl2 = st.columns(2)
    with c_dl1:
        json_out = json.dumps({'prog': edited_prog.to_dict(orient='records'), 'base_endeks': st.session_state.ff_base_endeks_df.to_dict(orient='records'), 'endeks': edited_endeks.to_dict(orient='records'), 'alt': edited_alt.to_dict(orient='records'), 'b': edited_b.to_dict(orient='records')}, ensure_ascii=False, indent=4)
        st.download_button("💾 JSON İndir", data=json_out, file_name="fiyat_farki_tahmin_projesi.json", mime="application/json", use_container_width=True)
    with c_dl2:
        excel_out = generate_excel_download(edited_prog, edited_endeks, edited_alt, edited_b, df_detay=df_detay, df_karsilastirma=df_karsilastirma_out, df_backtest=df_backtest_out if 'df_backtest_out' in locals() else pd.DataFrame())
        st.download_button("📊 EXCEL (7 Sayfa)", data=excel_out, file_name="fiyat_farki_kapsamli_rapor.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
