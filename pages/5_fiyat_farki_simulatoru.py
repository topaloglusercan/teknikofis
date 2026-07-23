"""
Hakediş Fiyat Farkı Simülatörü
================================
Mevcut "Hakediş Fiyat Farkı Hesaplayıcı" motorunun (kova sistemi + gecikme
matrisi + B katsayısı + endeks tablosu) BİREBİR AYNI hesaplama çekirdeği
üzerine kurulu; slider'larla anlık simülasyon ve senaryo karşılaştırma
katmanı eklenmiştir.

Çalıştırma:
    pip install streamlit pandas openpyxl plotly
    streamlit run hakedis_simulator.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import math, re, json, warnings, copy
from decimal import Decimal, ROUND_HALF_UP
import plotly.graph_objects as go

warnings.filterwarnings("ignore")

try:
    st.set_page_config(page_title="Hakediş Fiyat Farkı Simülatörü", page_icon="🏗️", layout="wide")
except Exception:
    pass  # Çoklu sayfa (pages/) uygulamasında ana dosya zaten set_page_config çağırmış olabilir

# ─────────────────────────────────────────────────────────
# YARDIMCI FONKSİYONLAR  (orijinal motorla birebir aynı)
# ─────────────────────────────────────────────────────────
MONTHS = {
    "oca": "01", "ocak": "01", "şub": "02", "sub": "02", "şubat": "02", "subat": "02",
    "mar": "03", "mart": "03", "nis": "04", "nisan": "04",
    "may": "05", "mayıs": "05", "mayis": "05",
    "haz": "06", "haziran": "06", "tem": "07", "temmuz": "07",
    "ağu": "08", "agu": "08", "ağustos": "08", "agustos": "08",
    "eyl": "09", "eylül": "09", "eylul": "09",
    "eki": "10", "ekim": "10", "kas": "11", "kasım": "11", "kasim": "11",
    "ara": "12", "aralık": "12", "aralik": "12",
}

def parse_tarih(s):
    s = str(s).strip()
    if re.match(r'^\d{4}-\d{2}(-\d{2})?$', s):
        return s[:7] + "-01"
    s = s.lower()
    parts = s.replace(".", " ").replace("/", " ").split()
    if len(parts) < 2:
        return None
    ay_str, yil = parts[0], parts[1]
    ay = MONTHS.get(ay_str[:3], MONTHS.get(ay_str))
    if ay is None:
        return None
    if len(yil) == 2:
        yil = "20" + yil
    return f"{yil}-{ay}-01"

def dec(x):
    try:
        if x is None:
            return Decimal('0')
        if isinstance(x, (int, float)):
            if isinstance(x, float) and math.isnan(x):
                return Decimal('0')
            return Decimal(str(x))
        s = str(x).strip()
        if s == '' or s.lower() == 'nan':
            return Decimal('0')
        return Decimal(s.replace('.', '').replace(',', '.')) if ',' in s else Decimal(s)
    except Exception:
        return Decimal('0')

def tr(val, d=2):
    try:
        if val is None or (isinstance(val, float) and math.isnan(val)) or float(val) == 0:
            return "-"
        s = f"{{:,.{d}f}}".format(float(val))
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)

EMAP = {'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}
ENDEKS_KOLONLARI = list(EMAP.values())

# Resmi "Fiyat Farkı Listesi / Temel Endeks" tablosuna göre kod eşlemesi
# (İn / Çn-23 / Dn-24 / Ayn / Kn-16 / Gn (Yİ-ÜFE) / Mn-28)
KOD_BILGI = {
    'a':  {'kolon': 'I o', 'resmi_kod': 'İn',    'kisa': 'İşçilik',      'ad': 'İşçilik (TÜFE bağlı)'},
    'b1': {'kolon': 'Ç o', 'resmi_kod': 'Çn-23', 'kisa': 'Çimento/Mineral', 'ad': 'Metalik Olmayan Diğer Mineral Ürünler (Çimento vb.)'},
    'b2': {'kolon': 'D o', 'resmi_kod': 'Dn-24', 'kisa': 'Demir-Çelik',  'ad': 'Ana Metaller (Demir-Çelik)'},
    'b3': {'kolon': 'Y o', 'resmi_kod': 'Ayn',   'kisa': 'Akaryakıt',    'ad': 'Akaryakıt Ürünleri'},
    'b4': {'kolon': 'K o', 'resmi_kod': 'Kn-16', 'kisa': 'Ağaç/Mantar',  'ad': 'Ağaç ve Mantar Ürünleri (mobilya hariç)'},
    'b5': {'kolon': 'G o', 'resmi_kod': 'Gn',    'kisa': 'Genel ÜFE',   'ad': 'Genel Yurt İçi ÜFE (Yİ-ÜFE)'},
    'c':  {'kolon': 'M o', 'resmi_kod': 'Mn-28', 'kisa': 'Makine/Ekipman', 'ad': 'Makine ve Ekipmanlar b.y.s.'},
}
KOD_ETIKET = {k: v['kisa'] for k, v in KOD_BILGI.items()}
KOLON_BILGI = {v['kolon']: v for v in KOD_BILGI.values()}

# ─────────────────────────────────────────────────────────
# JSON İÇE / DIŞA AKTARMA  (mevcut motorunuzun dosya formatıyla uyumlu)
# ─────────────────────────────────────────────────────────
def _num(x):
    return float(dec(x))

def json_to_dataframes(data):
    """Motorunuzun JSON formatını (prog/endeks/alt/b anahtarları, Türkçe
    sayı biçimi '1.234,56') iç DataFrame yapısına çevirir."""
    prog_rows = []
    for r in data.get('prog', []):
        prog_rows.append({
            'AYLAR': r['AYLAR'],
            'İŞ PROGRAMI KÜMÜLATİF': _num(r.get('İŞ PROGRAMI KÜMÜLATİF', '0')),
            'İMALAT TUTARI KÜMÜLATİF': _num(r.get('İMALAT TUTARI KÜMÜLATİF', '0')),
        })
    df_prog = pd.DataFrame(prog_rows)

    end_rows = []
    for r in data.get('endeks', []):
        row = {'AYLAR': r['AYLAR']}
        for k in ENDEKS_KOLONLARI:
            row[k] = _num(r.get(k, '0'))
        end_rows.append(row)
    df_end = pd.DataFrame(end_rows)

    alt_rows = []
    for r in data.get('alt', []):
        alt_rows.append({
            'Ağırlık': r['Ağırlık'],
            'Katsayı': _num(r.get('Katsayı', '0')),
            'Temel Endeks': _num(r.get('Temel Endeks', '0')),
        })
    df_alt = pd.DataFrame(alt_rows)

    b_rows = []
    for r in data.get('b', []):
        b_rows.append({'AYLAR': r['AYLAR'], 'B': _num(r.get('B', '1'))})
    df_b = pd.DataFrame(b_rows)

    return df_prog, df_end, df_alt, df_b


def _tl_str(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return tr(x, 2)

def _plain_str(x, d=6):
    try:
        return f"{float(x):.{d}f}".replace('.', ',')
    except Exception:
        return str(x)

def dataframes_to_json(df_prog, df_end, df_alt, df_b):
    """İç DataFrame yapısını, motorunuzun okuyabileceği JSON formatına
    (Türkçe sayı biçimiyle) geri çevirir."""
    data = {'prog': [], 'endeks': [], 'alt': [], 'b': []}
    for _, r in df_prog.iterrows():
        imal = r['İMALAT TUTARI KÜMÜLATİF']
        data['prog'].append({
            'AYLAR': r['AYLAR'],
            'İŞ PROGRAMI KÜMÜLATİF': _tl_str(r['İŞ PROGRAMI KÜMÜLATİF']),
            'İMALAT TUTARI KÜMÜLATİF': _tl_str(imal) if not (imal == 0 or imal == '') else "",
        })
    for _, r in df_end.iterrows():
        row = {'AYLAR': r['AYLAR']}
        for k in ENDEKS_KOLONLARI:
            row[k] = _plain_str(r[k])
        data['endeks'].append(row)
    for _, r in df_alt.iterrows():
        data['alt'].append({
            'Ağırlık': r['Ağırlık'],
            'Katsayı': _plain_str(r['Katsayı'], d=2),
            'Temel Endeks': _plain_str(r['Temel Endeks']),
        })
    for _, r in df_b.iterrows():
        data['b'].append({'AYLAR': r['AYLAR'], 'B': _plain_str(r['B'], d=1)})
    return data

# ─────────────────────────────────────────────────────────
# HESAPLAMA MOTORU  (orijinal ile birebir aynı mantık)
# ─────────────────────────────────────────────────────────
def hesapla(df_prog, df_end, df_alt, df_b):
    df_prog = df_prog.copy(); df_end = df_end.copy()
    df_alt  = df_alt.copy();  df_b   = df_b.copy()
    df_prog.columns = df_prog.columns.str.strip()

    df_end['_ay'] = pd.to_datetime(df_end['AYLAR'].apply(parse_tarih)).dt.to_period('M')
    df_end = df_end.drop_duplicates('_ay').set_index('_ay')

    df_b['_ay'] = pd.to_datetime(df_b['AYLAR'].apply(parse_tarih)).dt.to_period('M')
    df_b = df_b.drop_duplicates('_ay').set_index('_ay')

    df_prog['_ay'] = pd.to_datetime(df_prog['AYLAR'].apply(parse_tarih)).dt.to_period('M')
    son_end = df_end.index.max()

    kat = {str(r['Ağırlık']).strip(): dec(r['Katsayı'])      for _, r in df_alt.iterrows()}
    tbl = {str(r['Ağırlık']).strip(): dec(r['Temel Endeks']) for _, r in df_alt.iterrows()}

    pkol = 'İŞ PROGRAMI KÜMÜLATİF'
    ikol = 'İMALAT TUTARI KÜMÜLATİF'

    # ── 1. Kovalar ──
    kovalar, prev = [], Decimal('0')
    for _, r in df_prog.iterrows():
        kum = dec(r[pkol])
        kovalar.append({'ay': r['_ay'], 'kap': kum - prev, 'orig': kum - prev})
        prev = kum

    # ── 2. Hesaplama döngüsü ──
    ff_list, matris, aylik_rows = [], [], []
    prev_imal, kum_ff = Decimal('0'), Decimal('0')

    for _, r in df_prog.iterrows():
        ay = r['_ay']
        kum_imal = dec(r[ikol])
        ayl_imal = kum_imal - prev_imal

        if ayl_imal <= 0:
            ff_list.append(kum_ff if kum_imal > 0 else Decimal('0'))
            if kum_imal > 0:
                prev_imal = kum_imal
            continue

        b_val = dec(df_b.loc[ay, 'B']) if ay in df_b.index else Decimal('1')
        if b_val == 0:
            b_val = Decimal('1')

        real_end = min(ay, son_end)
        e_uyg = df_end.loc[real_end]
        ayl_ff = Decimal('0')
        kalan = ayl_imal

        for kova in kovalar:
            if kalan <= 0:
                break
            if kova['kap'] <= 0:
                continue

            kullan = min(kalan, kova['kap'])
            prog_ay = kova['ay']
            gecikme = prog_ay < ay
            real_prog = min(prog_ay, son_end)
            e_prog = df_end.loc[real_prog] if gecikme else e_uyg

            matris.append({'Kova Ayı': str(prog_ay), 'Hakediş Ayı': str(ay),
                            'Tutar': float(kullan), 'Gecikme': gecikme})

            pn = Decimal('0')
            for k, sut in EMAP.items():
                et = tbl.get(k, Decimal('0'))
                eu = dec(e_uyg[sut]); ep_ = dec(e_prog[sut])
                eg = min(eu, ep_) if gecikme else eu
                if et > 0:
                    pn += kat.get(k, Decimal('0')) * (eg / et)

            dilim = (kullan * b_val * (pn - 1)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            ayl_ff += dilim
            kova['kap'] -= kullan
            kalan -= kullan

        kum_ff += ayl_ff
        ff_list.append(kum_ff)
        aylik_rows.append({'Ay': str(ay), 'Aylık İmalat': float(ayl_imal),
                            'B Katsayısı': float(b_val),
                            'Aylık Fiyat Farkı': float(ayl_ff),
                            'Kümülatif Fiyat Farkı': float(kum_ff)})
        prev_imal = kum_imal

    df_sonuc = df_prog.drop(columns=['_ay'], errors='ignore').copy()
    df_sonuc['KÜMÜLATİF FF (TL)'] = [float(x) for x in ff_list]

    df_mat = pd.DataFrame(matris)
    df_pivot = pd.DataFrame()
    if not df_mat.empty:
        df_pivot = df_mat.pivot_table(index='Kova Ayı', columns='Hakediş Ayı',
                                       values='Tutar', aggfunc='sum', fill_value=0)
        kap_map = {str(k['ay']): float(k['orig']) for k in kovalar}
        df_pivot.insert(0, 'Kova Kapasitesi', pd.Series(kap_map))
        df_pivot.loc['TOPLAM'] = df_pivot.sum()

    df_kov = pd.DataFrame([{
        'Kova Ayı': str(k['ay']),
        'Başlangıç Kapasitesi': float(k['orig']),
        'Kalan Kapasite': float(k['kap']),
        'Kullanılan': float(k['orig'] - k['kap']),
        'Doluluk %': round(float((k['orig'] - k['kap']) / k['orig'] * 100), 1) if k['orig'] > 0 else 0,
    } for k in kovalar])

    df_aylik = pd.DataFrame(aylik_rows)
    return df_sonuc, df_pivot, df_kov, df_aylik, son_end


# ─────────────────────────────────────────────────────────
# SİMÜLASYON DÖNÜŞÜMLERİ  (yeni katman — motora dokunmaz)
# ─────────────────────────────────────────────────────────
def endeks_uzat(df_end, artis, ek_ay=36):
    """Endeks tablosunu, son bilinen aydan itibaren aylık bileşik artış
    oranıyla ileriye doğru uzatır (gelecekteki enflasyon/artış senaryosu).
    artis: tek bir float (%/ay, tüm kolonlara aynı oran uygulanır) VEYA
    {'I o': pct, 'Ç o': pct, ...} şeklinde kolon bazında sözlük."""
    df = df_end.copy()
    df['_ay'] = pd.to_datetime(df['AYLAR'].apply(parse_tarih)).dt.to_period('M')
    df = df.sort_values('_ay').reset_index(drop=True)
    son = df.iloc[-1]
    son_ay = df['_ay'].iloc[-1]
    cols = [c for c in df.columns if c not in ('AYLAR', '_ay')]

    if isinstance(artis, dict):
        artis_map = {c: artis.get(c, 0.0) for c in cols}
    else:
        artis_map = {c: artis for c in cols}

    yeni = []
    for k in range(1, ek_ay + 1):
        yeni_ay = son_ay + k
        satir = {'AYLAR': str(yeni_ay)}
        for c in cols:
            satir[c] = float(son[c]) * ((1 + artis_map[c] / 100) ** k)
        yeni.append(satir)

    df_ek = pd.DataFrame(yeni)
    return pd.concat([df.drop(columns='_ay'), df_ek], ignore_index=True)


def imalat_donustur(df_prog, hiz_carpani=1.0, gecikme_ay=0, tek_ay_index=None, tek_ay_kaydirma=0):
    """İmalat kümülatif eğrisini dönüştürür:
    - hiz_carpani: tüm ayların imalat artışını ölçekler.
    - gecikme_ay: TÜM ayları birlikte zaman ekseninde kaydırır (genel mod).
      Pozitif = iş geride kalır, negatif = iş önde gider.
    - tek_ay_index verilirse (Belirli Ay modu): sadece o aya ait imalat
      artışı, tek_ay_kaydirma kadar başka bir aya taşınır; diğer aylar
      olduğu gibi kalır. Bu durumda gecikme_ay dikkate alınmaz."""
    df = df_prog.copy()
    imal_col = 'İMALAT TUTARI KÜMÜLATİF'
    kum = [dec(v) for v in df[imal_col]]
    n = len(kum)

    aylik = [kum[0]] + [kum[i] - kum[i - 1] for i in range(1, n)]
    aylik = [a * Decimal(str(hiz_carpani)) for a in aylik]

    if tek_ay_index is not None:
        j = tek_ay_index + tek_ay_kaydirma
        j = max(0, min(n - 1, j))
        deger = aylik[tek_ay_index]
        aylik[tek_ay_index] = Decimal('0')
        aylik[j] += deger
    elif gecikme_ay != 0:
        shifted = [Decimal('0')] * n
        for i, a in enumerate(aylik):
            j = i + gecikme_ay
            j = max(0, min(n - 1, j))
            shifted[j] += a
        aylik = shifted

    yeni_kum, running = [], Decimal('0')
    for a in aylik:
        running += a
        yeni_kum.append(float(running))

    df[imal_col] = yeni_kum
    return df


def b_override_uygula(df_b, b_deger):
    df = df_b.copy()
    df['B'] = b_deger
    return df


def katsayi_override_uygula(df_alt, katsayi_dict):
    """Alt endeks ağırlıklarını (Katsayı sütunu) verilen sözlükle değiştirir.
    katsayi_dict: {'a': 0.5, 'b1': 0.0, ...} — kod bazında."""
    df = df_alt.copy()
    for i, row in df.iterrows():
        kod = str(row['Ağırlık']).strip()
        if kod in katsayi_dict:
            df.at[i, 'Katsayı'] = katsayi_dict[kod]
    return df


def senaryo_calistir(df_prog, df_end, df_alt, df_b, gecikme_ay, hiz_carpani, endeks_artis,
                      b_deger, tek_ay_index=None, tek_ay_kaydirma=0, katsayi_override=None):
    """endeks_artis: tek bir float (%/ay, tüm alt endekslere uygulanır) VEYA
    {'I o': pct, 'Ç o': pct, ...} şeklinde alt-endeks bazında sözlük olabilir.
    katsayi_override: {'a': 0.5, 'b3': 0.5, ...} verilirse alt endeks ağırlıkları
    (Katsayı) bu değerlerle değiştirilir."""
    p2 = imalat_donustur(df_prog, hiz_carpani=hiz_carpani, gecikme_ay=gecikme_ay,
                          tek_ay_index=tek_ay_index, tek_ay_kaydirma=tek_ay_kaydirma)
    e2 = endeks_uzat(df_end, endeks_artis)
    b2 = b_override_uygula(df_b, b_deger) if b_deger is not None else df_b
    a2 = katsayi_override_uygula(df_alt, katsayi_override) if katsayi_override else df_alt
    return hesapla(p2, e2, a2, b2)


# ─────────────────────────────────────────────────────────
# ÖRNEK VERİ  (gösterim amaçlı — kendi verinizle değiştirin)
# ─────────────────────────────────────────────────────────
def ornek_veri():
    aylar = ["Oca 2025", "Şub 2025", "Mar 2025", "Nis 2025", "May 2025", "Haz 2025"]

    prog = pd.DataFrame({
        "AYLAR": aylar,
        "İŞ PROGRAMI KÜMÜLATİF": [2_000_000, 5_000_000, 9_000_000, 14_000_000, 20_000_000, 27_000_000],
        "İMALAT TUTARI KÜMÜLATİF": [1_800_000, 4_500_000, 8_000_000, 12_500_000, 18_000_000, 25_500_000],
    })

    base = {"I o": 3000.0, "Ç o": 4800.0, "D o": 5900.0, "Y o": 44.7, "K o": 3400.0, "G o": 4300.0, "M o": 3100.0}
    end_rows = []
    for i, ay in enumerate(aylar):
        row = {"AYLAR": ay}
        for k, v in base.items():
            row[k] = round(v * (1.015 ** i), 2)
        end_rows.append(row)
    end = pd.DataFrame(end_rows)

    alt = pd.DataFrame({
        "Ağırlık": ["a", "b1", "b2", "b3", "b4", "b5", "c"],
        "Katsayı": [0.15, 0.20, 0.20, 0.15, 0.05, 0.10, 0.15],
        "Temel Endeks": [base["I o"], base["Ç o"], base["D o"], base["Y o"], base["K o"], base["G o"], base["M o"]],
    })

    b = pd.DataFrame({"AYLAR": aylar, "B": [0.9] * len(aylar)})
    return prog, end, alt, b


# ─────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────
if "prog" not in st.session_state:
    p, e, a, b = ornek_veri()
    st.session_state.prog, st.session_state.end = p, e
    st.session_state.alt, st.session_state.b_df = a, b
if "senaryolar" not in st.session_state:
    st.session_state.senaryolar = {}
if "rc" not in st.session_state:
    st.session_state.rc = 0

st.title("🏗️ Hakediş Fiyat Farkı Simülatörü")
st.caption("Kova sistemi · Gecikme matrisi · Slider'lı anlık simülasyon · Senaryo karşılaştırma")

tab1, tab2, tab3, tab4 = st.tabs(["📋 Veri Girişi", "📊 Baz Sonuç", "🎛️ Simülatör", "⚖️ Senaryo Karşılaştır"])

# ══════════════════════ TAB 1 — VERİ GİRİŞİ ══════════════════════
with tab1:
    st.info("Kendi projenizin verilerini buraya girin/yapıştırın, veya JSON dosyanızı içe aktarın. Aşağıdaki değerler örnek veridir.", icon="💡")

    with st.expander("📂 JSON İçe / Dışa Aktar (motorunuzun dosya formatı)", expanded=False):
        jc1, jc2 = st.columns(2)
        with jc1:
            st.markdown("**İçe Aktar**")
            yuklenen = st.file_uploader("JSON dosyası seçin", type=['json'], key='json_uploader')
            if yuklenen is not None:
                if st.button("✅ Bu JSON ile verileri değiştir"):
                    try:
                        veri = json.load(yuklenen)
                        p2, e2, a2, b2 = json_to_dataframes(veri)
                        st.session_state.prog, st.session_state.end = p2, e2
                        st.session_state.alt, st.session_state.b_df = a2, b2
                        st.session_state.rc += 1
                        st.success("JSON verisi içe aktarıldı.")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"JSON okunamadı: {ex}")
        with jc2:
            st.markdown("**Dışa Aktar**")
            try:
                json_data = dataframes_to_json(st.session_state.prog, st.session_state.end,
                                                st.session_state.alt, st.session_state.b_df)
                json_str = json.dumps(json_data, ensure_ascii=False, indent=2)
                st.download_button("💾 JSON olarak indir", data=json_str,
                                    file_name="hakedis_verisi.json", mime="application/json")
            except Exception as ex:
                st.error(f"JSON oluşturulamadı: {ex}")

    rc = st.session_state.rc
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("1️⃣ İş Programı ve İmalatlar")
        ep = st.data_editor(st.session_state.prog, num_rows="dynamic", use_container_width=True, key=f"ep_{rc}")
        st.session_state.prog = ep

        st.divider()
        st.subheader("3️⃣ Alt Endeks Ağırlıkları")
        ea = st.data_editor(st.session_state.alt, num_rows="dynamic", use_container_width=True, key=f"ea_{rc}")
        st.session_state.alt = ea
        try:
            ks = pd.to_numeric(ea['Katsayı'], errors='coerce').sum()
            ok = abs(ks - 1) < 0.001
            st.caption(f"{'✅' if ok else '⚠️'} Katsayı toplamı: **{ks:.4f}** {'(Doğru)' if ok else '→ 1.0000 olmalı!'}")
        except Exception:
            pass

        with st.expander("ℹ️ Hangi kod (Ağırlık) neye karşılık geliyor?"):
            st.caption("Resmi 'Fiyat Farkı Listesi / Temel Endeks' tablosuna göre eşleme:")
            ref = pd.DataFrame([
                {'Ağırlık Kodu': kod, 'Endeks Kolonu': b['kolon'],
                 'Resmi Kod': b['resmi_kod'], 'Açıklama': b['ad']}
                for kod, b in KOD_BILGI.items()
            ])
            st.dataframe(ref, use_container_width=True, hide_index=True)

    with c2:
        st.subheader("2️⃣ Aylık Endeks Tablosu")
        ee = st.data_editor(st.session_state.end, num_rows="dynamic", use_container_width=True, key=f"ee_{rc}")
        st.session_state.end = ee

        st.divider()
        st.subheader("4️⃣ B Katsayısı")
        eb = st.data_editor(st.session_state.b_df, num_rows="dynamic", use_container_width=True, key=f"eb_{rc}")
        st.session_state.b_df = eb

# ══════════════════════ TAB 2 — BAZ SONUÇ ══════════════════════
with tab2:
    try:
        df_sonuc, df_pivot, df_kov, df_aylik, son_end = hesapla(
            st.session_state.prog, st.session_state.end, st.session_state.alt, st.session_state.b_df
        )
        toplam_ff = df_aylik['Aylık Fiyat Farkı'].sum() if not df_aylik.empty else 0
        st.metric("Toplam Fiyat Farkı (Baz Senaryo)", f"{tr(toplam_ff)} TL")

        colA, colB = st.columns(2)
        with colA:
            st.subheader("Aylık / Kümülatif Fiyat Farkı")
            st.dataframe(df_aylik.style.format({
                'Aylık İmalat': lambda x: tr(x), 'B Katsayısı': '{:.4f}',
                'Aylık Fiyat Farkı': lambda x: tr(x), 'Kümülatif Fiyat Farkı': lambda x: tr(x),
            }), use_container_width=True)
        with colB:
            st.subheader("Kova Durumu")
            st.dataframe(df_kov, use_container_width=True)

        st.subheader("Kova Matrisi (Kova Ayı × Hakediş Ayı)")
        st.caption("🟩 Diyagonal = zamanında yapılan iş · 🟧 Üst-diyagonal = gecikmeli tüketim (düşük endeks cezası)")
        if not df_pivot.empty:
            st.dataframe(df_pivot.style.format(lambda x: tr(x) if isinstance(x, (int, float)) else x), use_container_width=True)

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df_aylik['Ay'], y=df_aylik['Kümülatif Fiyat Farkı'], mode='lines+markers', name='Kümülatif FF'))
        fig.update_layout(title="Kümülatif Fiyat Farkı Gelişimi", xaxis_title="Ay", yaxis_title="TL", height=350)
        st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"🚨 {e}")
        import traceback
        with st.expander("Teknik hata detayı"):
            st.code(traceback.format_exc())

# ══════════════════════ TAB 3 — SİMÜLATÖR ══════════════════════
with tab3:
    st.markdown("Sliderları hareket ettirin — sonuçlar **anlık** yeniden hesaplanır.")

    mod = st.radio("Kaydırma Modu", ["Genel (tüm aylar)", "Belirli Ay"], horizontal=True,
                    help="Genel: tüm ayları birlikte kaydırır. Belirli Ay: sadece seçtiğiniz tek bir ayın imalatını taşır, diğer aylar olduğu gibi kalır.")

    ay_listesi = list(st.session_state.prog['AYLAR'])
    tek_ay_index, tek_ay_kaydirma, gecikme_ay = None, 0, 0

    s1, s2, s3, s4 = st.columns(4)
    with s1:
        if mod == "Genel (tüm aylar)":
            gecikme_ay = st.slider("Gecikme / Hızlanma (ay)", -6, 6, 0,
                                    help="Pozitif: iş programın gerisinde kalır (üretim ileri aylara kayar). Negatif: iş programın önüne geçer.")
        else:
            secim_ay = st.selectbox("Hangi ay kaydırılsın?", ay_listesi)
            tek_ay_index = ay_listesi.index(secim_ay)
            tek_ay_kaydirma = st.slider(f"'{secim_ay}' kaç ay kaydırılsın", -6, 6, 0,
                                         help="Pozitif: bu ayın imalatı ileri bir aya taşınır (o ay gecikmeli yapılmış gibi). Negatif: geriye taşınır.")
    with s2:
        hiz_carpani = st.slider("Aylık İmalat Hızı Çarpanı", 0.3, 2.0, 1.0, 0.05,
                                 help="Her ayın imalat artışını bu katsayıyla ölçekler.")
    with s3:
        endeks_artis_genel = st.slider("Gelecek Aylar İçin Varsayılan Endeks Artışı (%/ay)", 0.0, 10.0, 1.5, 0.1,
                                        help="Endeks tablosunda veri olmayan aylar için bileşik aylık artış varsayımı.")
    with s4:
        b_ovr = st.slider("B Katsayısı (override)", 0.0, 2.0, 0.9, 0.01)

    with st.expander("ℹ️ 'Gecikme/Hızlanma' ve 'İmalat Hızı Çarpanı' ne yapar? (örnekli açıklama)"):
        st.markdown("""
### 🤔 Basit bir benzetme

Şantiyenizi bir **koşu bandı** gibi düşünün:

- 🔵 **İmalat Hızı Çarpanı** = bandın **hızı**. Ne kadar hızlı koşuyorsunuz
  (ayda ne kadar iş bitiyor)?
- 🔴 **Gecikme / Hızlanma** = koştuğunuz mesafenin **hangi takvim ayına**
  yazıldığı. Aynı mesafeyi, farklı bir ayda koşmuş gibi göstermek.

Kısacası: **Hız Çarpanı → "ne kadar iş"**, **Gecikme/Hızlanma → "ne zaman"**.
İkisi birbirinden bağımsızdır ve genelde birlikte kullanılır.
""")

        ornek_ay = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran"]
        ornek_orijinal = [10, 15, 20, 25, 30, 40]

        def _ornek_kaydir(vals, kaydirma):
            n = len(vals)
            yeni = [0.0] * n
            for i, v in enumerate(vals):
                j = max(0, min(n - 1, i + kaydirma))
                yeni[j] += v
            return yeni

        st.markdown("#### 🔴① En basit hâli: sadece TEK bir ayı kaydırmak (\"Belirli Ay\" modu)")
        st.markdown(
            "Örnek: **Mart** ayının işini (20 milyon TL) **+2** kaydırıyoruz. '+2' demek "
            "**2 takvim ayı ileri** demektir: Mart → (+1) → Nisan → (+1) → **Mayıs**. "
            "Nisan sadece 'yol üstünde', kendisi hiç değişmiyor — iş doğrudan Mayıs'a düşüyor:"
        )
        tek_ay_ornek = ornek_orijinal.copy()
        tek_ay_ornek[2] = 0  # Mart boşalıyor
        tek_ay_ornek[4] += ornek_orijinal[2]  # Mayıs'a ekleniyor
        fig_tek = go.Figure()
        fig_tek.add_trace(go.Bar(x=ornek_ay, y=ornek_orijinal, name='Orijinal', marker_color='#94a3b8'))
        fig_tek.add_trace(go.Bar(x=ornek_ay, y=tek_ay_ornek, name="Mart'ı +2 kaydır", marker_color='#f97316'))
        fig_tek.add_annotation(
            x="Mayıs", y=53, ax="Mart", ay=53, xref="x", yref="y", axref="x", ayref="y",
            showarrow=True, arrowhead=3, arrowsize=1.3, arrowwidth=2, arrowcolor="#ef4444",
            text="tam 2 ay ileri (Nisan atlanıyor)", font=dict(color="#ef4444", size=12), yshift=14
        )
        fig_tek.update_layout(barmode='group', title="Sadece Mart'ın İşi Mayıs'a Taşınıyor, Gerisi Aynı",
                               yaxis_title='Milyon TL (örnek)', height=300, margin=dict(t=40, b=10))
        st.plotly_chart(fig_tek, use_container_width=True)
        st.caption(
            "Mart'ın 20 milyon TL'lik işi artık Mayıs ayında görünüyor, Mart'ta '0' kalıyor. Nisan "
            "**hiç etkilenmedi** (sadece aradan geçilen bir ay, kendi işi hâlâ 25). Ocak, Şubat ve "
            "Haziran da hiç değişmedi. Program hâlâ bu işi Mart ayına yazdığı için, Mayıs'a kayan bu "
            "iş **gecikmeli** sayılır ve Mart'ın (daha düşük) endeksiyle sabitlenir."
        )

        st.markdown("#### 🔴② Daha kapsamlı hâli: TÜM ayları birlikte kaydırmak (\"Genel\" modu)")
        st.markdown(
            "Genel modda tek bir ay değil, **her ayın işi kendi + N ay sonrasına** taşınır — "
            "yani hepsi aynı anda, aynı miktarda kayar. Aşağıda tüm aylara +2 uygulanmış hâli:"
        )
        gecikmeli = _ornek_kaydir(ornek_orijinal, 2)
        fig_gec = go.Figure()
        fig_gec.add_trace(go.Bar(x=ornek_ay, y=ornek_orijinal, name='Orijinal (gecikme=0)', marker_color='#94a3b8'))
        fig_gec.add_trace(go.Bar(x=ornek_ay, y=gecikmeli, name='Gecikme = +2 (tüm aylar)', marker_color='#ef4444'))
        fig_gec.update_layout(barmode='group', title='Her Ayın İşi Kendi +2 Ayına Taşınıyor (Toplam TL Aynı)',
                               yaxis_title='Milyon TL (örnek)', height=300, margin=dict(t=40, b=10))
        st.plotly_chart(fig_gec, use_container_width=True)

        esleme = []
        n = len(ornek_ay)
        for i, ay in enumerate(ornek_ay):
            j = max(0, min(n - 1, i + 2))
            not_str = ""
            if i + 2 > n - 1:
                not_str = " (program burada bittiği için son aya yığıldı)"
            esleme.append({'Orijinal Ay': ay, 'Tutar (milyon TL)': ornek_orijinal[i],
                            '→ Yeni Ay': ornek_ay[j] + not_str})
        st.dataframe(pd.DataFrame(esleme), use_container_width=True, hide_index=True)

        st.caption(
            "Her satırda görüldüğü gibi Ocak'ın işi Mart'a, Şubat'ınki Nisan'a, Mart'ınki Mayıs'a "
            "taşınıyor — hepsi kendi +2 ayına. Ama Nisan, Mayıs ve Haziran'ın işi +2 kaydırılınca "
            "programın son ayı olan Haziran'ı aşıyor; taşacak yer olmadığı için hepsi **Haziran'da "
            "birikiyor** (bu yüzden Haziran çubuğu bu kadar yüksek). Toplam iş miktarı "
            f"({sum(ornek_orijinal)} milyon TL) hiç değişmiyor, sadece hangi ayda göründüğü değişiyor. "
            "Program hâlâ eski aylarını beklediği için, kaymış her iş **gecikmeli** sayılır ve kendi "
            "(daha düşük) programlanmış ayının endeksiyle sabitlenir — sonuç olarak fiyat farkı "
            "genelde **düşer**."
        )


        st.markdown("#### 🔵 Örnek: İmalat Hızı Çarpanı = 1.3 uygulanırsa")
        hizli = [round(v * 1.3, 1) for v in ornek_orijinal]
        fig_hiz = go.Figure()
        fig_hiz.add_trace(go.Bar(x=ornek_ay, y=ornek_orijinal, name='Orijinal (çarpan=1.0)', marker_color='#94a3b8'))
        fig_hiz.add_trace(go.Bar(x=ornek_ay, y=hizli, name='Çarpan = 1.3', marker_color='#3b82f6'))
        fig_hiz.update_layout(barmode='group', title='Her Ayın Kendi İş Miktarı Değişiyor (Zaman Aynı)',
                               yaxis_title='Milyon TL (örnek)', height=300, margin=dict(t=40, b=10))
        st.plotly_chart(fig_hiz, use_container_width=True)
        st.caption(
            "Ocak yine Ocak, Şubat yine Şubat — hiçbir ay yer değiştirmiyor. Sadece her ayda yapılan iş "
            "%30 artıyor (10→13, 15→19,5 vb.). Daha fazla iş, o ayın kovasını daha hızlı doldurur; taşan "
            "kısım bir sonraki kovaya sızar."
        )

        st.markdown("""
#### 📋 Özet

| | Neyi değiştirir? | Neyi değiştirmez? | Fiyat farkına genel etkisi |
|---|---|---|---|
| 🔴 **Gecikme/Hızlanma** | **Zaman** — hangi ayda yapıldığı | Toplam TL tutarı | (+) gecikme → **düşer**, (−) hızlanma → **artar** |
| 🔵 **İmalat Hızı Çarpanı** | **Miktar** — ne kadar iş yapıldığı | Zaman ekseni (aylar) | >1 → kovalar hızlı dolar, <1 → programın gerisinde kalınır |

**İkisini birlikte kullanmak:** "İş 2 ay gecikmeli ama aynı zamanda normalden
%20 daha hızlı ilerliyor" senaryosunu görmek için Gecikme=+2 **ve**
Hız Çarpanı=1.2 birlikte ayarlanabilir.
        """)

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
                    endeks_artis_dict[kol] = st.slider(bilgi['kisa'], 0.0, 10.0, endeks_artis_genel, 0.1,
                                                        key=f"artis_{kol}",
                                                        help=f"Resmi kod: {bilgi['resmi_kod']} — {bilgi['ad']}")
        endeks_artis = endeks_artis_dict

    # ── Alt endeks ağırlıkları (Katsayı) simülasyonu ──
    katsayi_bazinda = st.checkbox("Alt endeks ağırlıklarını (Katsayı) simüle et (gelişmiş)",
                                   help="Örn: 'sadece işçilik %100 olsa' veya 'işçilik %50 + akaryakıt %50 olsa' ne olurdu?")
    katsayi_override = None
    if katsayi_bazinda:
        orijinal_kat = {str(r['Ağırlık']).strip(): float(r['Katsayı']) for _, r in st.session_state.alt.iterrows()}

        # Widget'lar oluşturulmadan ÖNCE, bekleyen (buton kaynaklı) bir güncelleme
        # varsa uygula. Streamlit, bir widget'ın key'ine, o widget bu run içinde
        # oluşturulduktan SONRA session_state üzerinden yazılmasına izin vermez;
        # bu yüzden güncelleme her zaman bir önceki rerun'da (widget'lar henüz
        # oluşmadan) uygulanır.
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
        base_sonuc, base_pivot, base_kov, base_aylik, _ = hesapla(
            st.session_state.prog, st.session_state.end, st.session_state.alt, st.session_state.b_df
        )
        sim_sonuc, sim_pivot, sim_kov, sim_aylik, _ = senaryo_calistir(
            st.session_state.prog, st.session_state.end, st.session_state.alt, st.session_state.b_df,
            gecikme_ay, hiz_carpani, endeks_artis, b_ovr,
            tek_ay_index=tek_ay_index, tek_ay_kaydirma=tek_ay_kaydirma,
            katsayi_override=katsayi_override
        )

        toplam_base = base_aylik['Aylık Fiyat Farkı'].sum() if not base_aylik.empty else 0
        toplam_sim = sim_aylik['Aylık Fiyat Farkı'].sum() if not sim_aylik.empty else 0
        fark = toplam_sim - toplam_base
        fark_pct = (fark / toplam_base * 100) if toplam_base != 0 else 0

        m1, m2, m3 = st.columns(3)
        m1.metric("Baz Senaryo Toplam FF", f"{tr(toplam_base)} TL")
        m2.metric("Simülasyon Toplam FF", f"{tr(toplam_sim)} TL", delta=f"{tr(fark)} TL")
        m3.metric("Fark (%)", f"{fark_pct:+.1f}%")

        fig2 = go.Figure()
        if not base_aylik.empty:
            fig2.add_trace(go.Scatter(x=base_aylik['Ay'], y=base_aylik['Kümülatif Fiyat Farkı'],
                                       mode='lines+markers', name='Baz Senaryo', line=dict(dash='dot')))
        if not sim_aylik.empty:
            fig2.add_trace(go.Scatter(x=sim_aylik['Ay'], y=sim_aylik['Kümülatif Fiyat Farkı'],
                                       mode='lines+markers', name='Simülasyon'))
        fig2.update_layout(title="Kümülatif Fiyat Farkı — Baz vs Simülasyon", xaxis_title="Ay", yaxis_title="TL", height=380)
        st.plotly_chart(fig2, use_container_width=True)

        with st.expander("📊 Aylık detay tablosu (simülasyon)"):
            st.dataframe(sim_aylik.style.format({
                'Aylık İmalat': lambda x: tr(x), 'B Katsayısı': '{:.4f}',
                'Aylık Fiyat Farkı': lambda x: tr(x), 'Kümülatif Fiyat Farkı': lambda x: tr(x),
            }), use_container_width=True)

        st.divider()
        senaryo_adi = st.text_input("Bu ayarları senaryo olarak kaydet (isim ver):", placeholder="örn. 'Senaryo A - 3 ay gecikme'")
        if st.button("💾 Senaryoyu Kaydet"):
            if senaryo_adi.strip():
                st.session_state.senaryolar[senaryo_adi.strip()] = {
                    'mod': mod,
                    'gecikme_ay': gecikme_ay,
                    'tek_ay': ay_listesi[tek_ay_index] if tek_ay_index is not None else None,
                    'tek_ay_kaydirma': tek_ay_kaydirma,
                    'hiz_carpani': hiz_carpani,
                    'endeks_artis': endeks_artis if not isinstance(endeks_artis, dict) else "alt endeks bazlı",
                    'b_ovr': b_ovr,
                    'katsayi': katsayi_override if katsayi_override else "orijinal",
                    'toplam_ff': toplam_sim, 'aylik': sim_aylik.to_dict('records'),
                }
                st.success(f"'{senaryo_adi}' kaydedildi. Karşılaştırmak için '⚖️ Senaryo Karşılaştır' sekmesine geçin.")
            else:
                st.warning("Lütfen bir senaryo ismi girin.")

    except Exception as e:
        st.error(f"🚨 {e}")
        import traceback
        with st.expander("Teknik hata detayı"):
            st.code(traceback.format_exc())

# ══════════════════════ TAB 4 — SENARYO KARŞILAŞTIR ══════════════════════
with tab4:
    st.markdown("Simülatör sekmesinde kaydettiğiniz senaryoları burada yan yana karşılaştırın.")

    if not st.session_state.senaryolar:
        st.info("👈 Henüz kaydedilmiş senaryo yok. '🎛️ Simülatör' sekmesinde ayarları belirleyip **Senaryoyu Kaydet** butonuna basın.")
    else:
        secilenler = st.multiselect("Karşılaştırılacak senaryolar:", list(st.session_state.senaryolar.keys()),
                                     default=list(st.session_state.senaryolar.keys()),
                                     help="İstediğiniz kadar senaryoyu aynı grafikte üst üste görebilirsiniz.")

        if secilenler:
            ozet = []
            for ad in secilenler:
                s = st.session_state.senaryolar[ad]
                ozet.append({
                    'Senaryo': ad,
                    'Mod': s.get('mod', 'Genel (tüm aylar)'),
                    'Gecikme (ay)': s['gecikme_ay'],
                    'Kaydırılan Ay': s.get('tek_ay') or '-',
                    'Ay Kaydırma': s.get('tek_ay_kaydirma', 0),
                    'İmalat Hız Çarpanı': s['hiz_carpani'],
                    'Endeks Artışı (%/ay)': s['endeks_artis'],
                    'B Katsayısı': s['b_ovr'],
                    'Toplam Fiyat Farkı (TL)': s['toplam_ff'],
                })
            df_ozet = pd.DataFrame(ozet)
            st.dataframe(df_ozet.style.format({'Toplam Fiyat Farkı (TL)': lambda x: tr(x)}), use_container_width=True)

            fig3 = go.Figure()
            for ad in secilenler:
                s = st.session_state.senaryolar[ad]
                aylik = pd.DataFrame(s['aylik'])
                if not aylik.empty:
                    fig3.add_trace(go.Scatter(x=aylik['Ay'], y=aylik['Kümülatif Fiyat Farkı'], mode='lines+markers', name=ad))
            fig3.update_layout(title="Senaryolar Arası Kümülatif Fiyat Farkı Karşılaştırması",
                                xaxis_title="Ay", yaxis_title="TL", height=400)
            st.plotly_chart(fig3, use_container_width=True)

            if len(secilenler) == 2:
                a, b_ = secilenler
                fark = st.session_state.senaryolar[b_]['toplam_ff'] - st.session_state.senaryolar[a]['toplam_ff']
                st.metric(f"'{b_}' − '{a}' Farkı", f"{tr(fark)} TL")

        if st.button("🗑️ Tüm senaryoları temizle"):
            st.session_state.senaryolar = {}
            st.rerun()
