import streamlit as st
import pandas as pd
import warnings
import io
import json
from decimal import Decimal, ROUND_HALF_UP, getcontext

getcontext().prec = 28
warnings.filterwarnings("ignore")

st.set_page_config(page_title="İdari Hakediş Modülü", layout="wide", page_icon="📂")

# --- YARDIMCI FONKSİYONLAR ---
def parse_turkish_date(date_str):
    if pd.isna(date_str) or str(date_str).strip() == '':
        return pd.NaT
    date_str = str(date_str).strip().replace('.', ' ').lower()
    if date_str in ['none', 'nan', 'nat', '<na>']:
        return pd.NaT
    months = {
        'oca': '01', 'ocak': '01', 'şub': '02', 'şubat': '02', 'mar': '03', 'mart': '03',
        'nis': '04', 'nisan': '04', 'may': '05', 'mayıs': '05', 'haz': '06', 'haziran': '06',
        'tem': '07', 'temmuz': '07', 'ağu': '08', 'ağustos': '08', 'eyl': '09', 'eylül': '09',
        'eki': '10', 'ekim': '10', 'kas': '11', 'kasım': '11', 'ara': '12', 'aralık': '12'
    }
    parts = date_str.split()
    if len(parts) == 2:
        m_num = months.get(parts[0], '01')
        y_num = parts[1] if len(parts[1]) == 4 else f"20{parts[1]}"
        return f"{y_num}-{m_num}"
    return pd.NaT


def clean_decimal(val):
    if pd.isna(val):
        return Decimal('0.0')
    val_str = str(val).strip()
    if val_str.lower() in ['', 'none', 'nan', 'nat', '<na>']:
        return Decimal('0.0')

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
        if d.is_nan():
            return Decimal('0.0')
        return d
    except Exception:
        return Decimal('0.0')


def tr_format(val):
    if pd.isna(val) or val == "":
        return ""
    try:
        formatted = "{:,.2f}".format(float(val))
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return val


def filter_empty_rows(df):
    if df.empty:
        return df
    mask = df.iloc[:, 0].astype(str).str.strip().str.lower().isin(['', 'none', 'nan', 'nat', '<na>'])
    return df[~mask]


# --- EXCEL YÜKLEME / İNDİRME ---
def load_from_excel(file):
    """Read an uploaded Excel and map sheets flexibly to expected dataframes."""
    xls = pd.ExcelFile(file)
    dfs = {}
    for sheet in xls.sheet_names:
        name_norm = sheet.lower()
        try:
            if 'ispro' in name_norm or 'program' in name_norm or 'iş pro' in name_norm or 'işprogram' in name_norm:
                key = 'prog_df'
            elif 'endeks' in name_norm:
                key = 'endeks_df'
            elif 'alt' in name_norm or 'ağırlık' in name_norm:
                key = 'alt_df'
            elif sheet.strip().upper() == 'B':
                key = 'b_df'
            else:
                continue

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
            dfs[key] = df
        except Exception:
            continue
    return dfs


def generate_excel_download(df_prog, df_endeks, df_alt, df_b):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_prog.to_excel(writer, sheet_name='IsProgrami', index=False)
        df_endeks.to_excel(writer, sheet_name='Endeks', index=False)
        df_alt.to_excel(writer, sheet_name='AltEndeks', index=False)
        df_b.to_excel(writer, sheet_name='B', index=False)
    return output.getvalue()


# --- HESAPLAMA MOTORU ---
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

    # alt_df sütunlarını güvenle al
    try:
        ag_col = [c for c in df_alt.columns if c.lower().strip() == 'ağırlık'][0]
    except Exception:
        ag_col = df_alt.columns[0]
    try:
        kats_col = [c for c in df_alt.columns if c.lower().strip() == 'katsayı'][0]
    except Exception:
        kats_col = df_alt.columns[1] if len(df_alt.columns) > 1 else df_alt.columns[0]
    try:
        temel_col = [c for c in df_alt.columns if c.lower().strip() == 'temel endeks'][0]
    except Exception:
        temel_col = df_alt.columns[2] if len(df_alt.columns) > 2 else df_alt.columns[-1]

    katsayilar = {str(row[ag_col]).strip().lower(): clean_decimal(row[kats_col]) for _, row in df_alt.iterrows()}
    temel_endeksler = {str(row[ag_col]).strip().lower(): clean_decimal(row[temel_col]) for _, row in df_alt.iterrows()}
    endeks_haritasi = {'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}

    if df_prog.shape[1] < 3:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    prog_kum_col = df_prog.columns[1]
    imalat_kum_col = df_prog.columns[2]

    kovalar = []
    onceki_kum = Decimal('0.0')
    for _, row in df_prog.iterrows():
        kum = clean_decimal(row[prog_kum_col])
        capacity = kum - onceki_kum
        kovalar.append({'ay': row['AyKodu'], 'kapasite': capacity if capacity > Decimal('0.0') else Decimal('0.0')})
        onceki_kum = kum

    final_ff_listesi, matris_verileri = [], []
    onceki_imalat_kum, kümülatif_toplam_ff = Decimal('0.0'), Decimal('0.0')

    for _, row in df_prog.iterrows():
        uyg_ayi = row['AyKodu']
        guncel_imalat_kum = clean_decimal(row[imalat_kum_col])
        aylik_imalat = guncel_imalat_kum - onceki_imalat_kum

        if aylik_imalat <= Decimal('0.0'):
            final_ff_listesi.append(float(kümülatif_toplam_ff))
            if guncel_imalat_kum > Decimal('0.0'):
                onceki_imalat_kum = guncel_imalat_kum
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
            if kalan_para <= Decimal('0.0'):
                break
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

    return df_sonuc, df_pivot, df_detay


# --- ARAYÜZ VE HAFIZA YÖNETİMİ ---
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

st.title("📂 İdari Hakediş & Teyit Matrisi")

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Excel ile Proje Yükle")
st.sidebar.caption("💡 Excel şablonunuzu (`.xlsx`) buradan yükleyin.")
uploaded_xlsx = st.sidebar.file_uploader("Excel Dosyası Seç (.xlsx)", type=["xlsx"] , key='xlsx_upload')

if uploaded_xlsx is not None:
    try:
        dfs = load_from_excel(uploaded_xlsx)
        if 'prog_df' in dfs:
            st.session_state.prog_df = dfs['prog_df']
        if 'endeks_df' in dfs:
            st.session_state.endeks_df = dfs['endeks_df']
        if 'alt_df' in dfs:
            st.session_state.alt_df = dfs['alt_df']
        if 'b_df' in dfs:
            st.session_state.b_df = dfs['b_df']
        st.session_state.load_count += 1
        st.sidebar.success("Excel başarıyla okundu!")
    except Exception as e:
        st.sidebar.error(f"Dosya okuma hatası: Lütfen dosya yapısının doğru olduğundan emin olun. Detay: {e}")

# JSON proje yükleme / indirme (eski akış korunuyor)
st.sidebar.subheader("📥 Önceki Projeyi Yükle (.json)")
uploaded_json = st.sidebar.file_uploader("Projeyi Yükle (.json)", type=["json"], key='json_upload')
if uploaded_json is not None:
    try:
        data = json.load(uploaded_json)
        st.session_state.prog_df = pd.DataFrame(data.get('prog', [])).astype(str)
        st.session_state.endeks_df = pd.DataFrame(data.get('endeks', [])).astype(str)
        st.session_state.alt_df = pd.DataFrame(data.get('alt', [])).astype(str)
        st.session_state.b_df = pd.DataFrame(data.get('b', [])).astype(str)
        st.session_state.load_count += 1
        st.sidebar.success("Proje başarıyla yüklendi!")
    except Exception as e:
        st.sidebar.error(f"JSON yükleme hatası: {e}")

suffix = st.session_state.load_count

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. İş Programı ve İmalatlar")
    edited_prog = st.data_editor(st.session_state.prog_df.astype(str), num_rows="dynamic", use_container_width=True, key=f"prog_ed_{suffix}")
    st.subheader("3. Alt Endeks Ağırlıkları")
    edited_alt = st.data_editor(st.session_state.alt_df.astype(str), num_rows="dynamic", use_container_width=True, key=f"alt_ed_{suffix}")
with col2:
    st.subheader("2. Endeks Tablosu")
    edited_endeks = st.data_editor(st.session_state.endeks_df.astype(str), num_rows="dynamic", use_container_width=True, key=f"end_ed_{suffix}")
    st.subheader("4. B Katsayısı Tablosu")
    edited_b = st.data_editor(st.session_state.b_df.astype(str), num_rows="dynamic", use_container_width=True, key=f"b_ed_{suffix}")

# İndirme butonları: JSON ve Excel
project_data = {
    'prog': edited_prog.to_dict(orient='records'),
    'endeks': edited_endeks.to_dict(orient='records'),
    'alt': edited_alt.to_dict(orient='records'),
    'b': edited_b.to_dict(orient='records')
}

st.sidebar.markdown("---")
st.sidebar.subheader("📤 Mevcut Veriyi İndir")
st.sidebar.download_button(
    label="💾 Mevcut Veriyi JSON olarak indir",
    data=json.dumps(project_data, indent=2, ensure_ascii=False),
    file_name="hakedis_projem.json",
    mime="application/json",
    use_container_width=True
)

# Excel indirme
try:
    excel_data = generate_excel_download(edited_prog, edited_endeks, edited_alt, edited_b)
    st.sidebar.download_button(
        label="💾 Mevcut Veriyi Excel olarak indir (.xlsx)",
        data=excel_data,
        file_name="idari_hakedis_projem.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
except Exception:
    st.sidebar.info("Excel oluşturulurken bir sorun oldu. openpyxl yüklü olduğundan emin olun.")

st.markdown("---")
if st.button("🚀 Hesapla ve Matrisi Çıkar", use_container_width=True):
    try:
        p = edited_prog.copy()
        e = edited_endeks.copy()
        a = edited_alt.copy()
        b = edited_b.copy()

        df_sonuc, df_pivot, df_detay = hesapla(p, e, a, b)

        if df_sonuc.empty:
            st.warning("⚠️ Lütfen tablolara geçerli hakediş ve endeks verilerini giriniz.")
        else:
            st.subheader("🔍 Detaylı Fiyat Farkı Analizi (Dilim Bazlı)")
            df_detay_gosterim = df_detay.copy()
            if 'Kullanılan Tutar' in df_detay_gosterim.columns:
                df_detay_gosterim['Kullanılan Tutar'] = df_detay_gosterim['Kullanılan Tutar'].apply(tr_format)
            if 'Fiyat Farkı Tutarı' in df_detay_gosterim.columns:
                df_detay_gosterim['Fiyat Farkı Tutarı'] = df_detay_gosterim['Fiyat Farkı Tutarı'].apply(tr_format)
            if 'Uygulanan Pn (Excel - 15 Hane)' in df_detay_gosterim.columns:
                try:
                    df_detay_gosterim['Uygulanan Pn (Excel - 15 Hane)'] = df_detay_gosterim['Uygulanan Pn (Excel - 15 Hane)'].apply(
                        lambda x: "{:.15f}".format(x).rstrip('0').rstrip('.').replace('.', ',')
                    )
                except Exception:
                    pass
            st.dataframe(df_detay_gosterim, use_container_width=True)

            st.subheader("📊 Ödenek Dilimlerinin Hakedişlere Göre Kullanılması (Teyit Matrisi)")
            if not df_pivot.empty:
                df_pivot_tr = df_pivot.map(tr_format)
                try:
                    st.dataframe(df_pivot_tr.style.set_properties(subset=['HAKEDİŞ TUTARI (Toplam)'], **{'font-weight': 'bold', 'background-color': '#e6f2ff'}), use_container_width=True)
                except Exception:
                    st.dataframe(df_pivot_tr, use_container_width=True)
            else:
                st.info("Teyit matrisi için detay verisi bulunamadı.")

            st.subheader("📑 Kümülatif Fiyat Farkı Sonuçları")
            for col in df_sonuc.columns:
                if any(x in col.upper() for x in ['TUTAR', 'PROGRAM', 'FARKI']):
                    df_sonuc[col] = df_sonuc[col].apply(tr_format)
            st.dataframe(df_sonuc, use_container_width=True)

    except Exception as ex:
        st.error(f"🚨 Hesaplama Sırasında Hata Oluştu: {ex}")
