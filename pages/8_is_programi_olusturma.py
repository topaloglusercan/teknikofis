"""
Yapay Zeka Destekli İş Programı & Risk Simülatörü - Sürüm 60.19 (NİHAİ TAM SÜRÜM)
Özellikler: Excel Şablon Butonu Düzeltildi, Tez Grafikleri (P50/P90, WBS Riskleri), Parametre Kütüphanesi.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import scipy.stats as st_stats
import re
import io
import json
import requests
from datetime import datetime, timedelta
import folium
from streamlit_folium import st_folium

warnings = __import__('warnings')
warnings.filterwarnings("ignore")

# ==========================================
# 1. YARDIMCI VE GÜVENLİK FONKSİYONLARI
# ==========================================
def safe_float(val, default=0.0):
    if pd.isna(val) or val is None or str(val).strip().lower() in ["", "none", "nan", "-"]: return default
    try: return float(str(val).replace(',', ''))
    except: return default

def safe_str(val):
    if pd.isna(val) or val is None or str(val).strip().lower() in ["", "none", "nan", "-"]: return ""
    return str(val).strip()

def parse_single_pred(p):
    p = str(p).upper().strip()
    if not p: return "", "FS", 0.0
    match = re.search(r"(FS|SS|FF|SF)", p)
    if match:
        link_type = match.group(1)
        parts = p.split(link_type)
        pred_id = parts[0].strip()
        lag_str = parts[1].replace(" ", "") if len(parts) > 1 and parts[1].strip() else "0"
        try: lag = float(lag_str)
        except: lag = 0.0
        return pred_id, link_type, lag
    else:
        match = re.search(r"([+-]\s*\d+(?:\.\d+)?)", p)
        if match:
            lag_str = match.group(1).replace(" ", "")
            pred_id = p.replace(match.group(1), "").strip()
            try: lag = float(lag_str)
            except: lag = 0.0
            return pred_id, "FS", lag
        else:
            return p.strip(), "FS", 0.0

def fetch_weather_data(lat, lon, start_date, end_date):
    url = "https://archive-api.open-meteo.com/v1/archive"
    params = {"latitude": lat, "longitude": lon, "start_date": start_date, "end_date": end_date, "daily": ["precipitation_sum", "wind_speed_10m_max", "temperature_2m_min"], "timezone": "auto"}
    response = requests.get(url, params=params)
    response.raise_for_status()
    data = response.json()
    df = pd.DataFrame({"Tarih": pd.to_datetime(data["daily"]["time"]), "Yagis_mm": data["daily"].get("precipitation_sum", []), "Ruzgar_kmh": data["daily"].get("wind_speed_10m_max", []), "Sicaklik_min": data["daily"].get("temperature_2m_min", [])})
    df['Yagis_mm'] = df['Yagis_mm'].fillna(0)
    df['Ruzgar_kmh'] = df['Ruzgar_kmh'].fillna(0)
    df['Sicaklik_min'] = df['Sicaklik_min'].fillna(20) 
    return df

def perform_ks_test(df_hist):
    results = []
    if 'Aktivite Kodu' not in df_hist.columns or 'Planlanan_Sure' not in df_hist.columns: return pd.DataFrame()
    for akod, group in df_hist.groupby('Aktivite Kodu'):
        p_arr = pd.to_numeric(group['Planlanan_Sure'], errors='coerce')
        g_arr = pd.to_numeric(group['Gercek_Sure'], errors='coerce')
        mask = (p_arr > 0) & (~g_arr.isna())
        ratios = (g_arr[mask] / p_arr[mask]).values
        if len(ratios) < 3: continue
        best_dist, best_p = "Triangular", 0
        try:
            p_norm = st_stats.kstest(ratios, 'norm', args=st_stats.norm.fit(ratios))[1]
            if p_norm > best_p: best_p, best_dist = p_norm, "Normal"
        except: pass
        try:
            p_uni = st_stats.kstest(ratios, 'uniform', args=st_stats.uniform.fit(ratios))[1]
            if p_uni > best_p: best_p, best_dist = p_uni, "Uniform"
        except: pass
        try:
            p_log = st_stats.kstest(ratios, 'lognorm', args=st_stats.lognorm.fit(ratios))[1]
            if p_log > best_p: best_p, best_dist = p_log, "Lognormal"
        except: pass
        if best_p < 0.05: best_dist = "Triangular"
        results.append({"Aktivite Kodu": akod, "Dağılım Tipi": best_dist, "Min_Ratio": np.min(ratios), "Med_Ratio": np.median(ratios), "Max_Ratio": np.max(ratios), "Veri_Sayisi": len(ratios)})
    return pd.DataFrame(results)

def generate_excel_template():
    df_tasks = pd.DataFrame({"ID": [f"T{i:02d}" for i in range(1, 16)], "WBS": ["1. Kaba İşler"]*4 + ["2. İnce İşler"]*4 + ["3. Elektrik"]*3 + ["4. Mekanik"]*2 + ["5. Cephe", "6. Peyzaj"], "Aktivite Kodu": ["HAF-01", "BET-01", "BET-01", "BET-01", "DUV-01", "SVA-01", "SVA-01", "SER-01", "ELK-01", "ELK-01", "ELK-01", "MEK-01", "MEK-01", "CPH-01", "PEY-01"], "Görev": ["Hafriyat", "Temel", "Kolon/Perde", "Döşeme", "Duvar", "Sıva", "Boya", "Seramik", "Kablo", "Pano", "Aydınlatma", "Sıhhi Tesisat", "Havalandırma", "Dış Cephe", "Çim"], "Alt Yüklenici": ["Kazıcı A.Ş.", "Beton A.Ş.", "Beton A.Ş.", "Beton A.Ş.", "İnce Yapı", "İnce Yapı", "Boya Ltd.", "İnce Yapı", "Elektrik", "Elektrik", "Elektrik", "Mekanik", "Mekanik", "Cephe", "Peyzaj"], "Bütçe Tutarı (TL)": [150000, 400000, 600000, 450000, 200000, 180000, 120000, 300000, 150000, 250000, 100000, 280000, 220000, 750000, 80000], "Planlanan Süre": [6, 7, 10, 7, 9, 7, 6, 10, 7, 5, 6, 9, 7, 14, 6], "İlerleme (%)": [0]*15, "Gerçekleşen Süre": [0]*15, "Duruş (Gün)": [0]*15, "Dağılım Tipi": ["BetaPERT", "Triangular", "Triangular", "Uniform", "Triangular", "Normal", "Lognormal", "BetaPERT", "Triangular", "Uniform", "Triangular", "Triangular", "Normal", "BetaPERT", "Uniform"], "İyimser": [4, 5, 8, 5, 6, 5, 4, 7, 5, 3, 4, 6, 5, 10, 4], "Olası": [6, 7, 10, 7, 9, 7, 6, 10, 7, 5, 6, 9, 7, 14, 6], "Kötümser": [10, 12, 15, 11, 15, 11, 9, 15, 11, 8, 9, 14, 10, 20, 10], "Öncüller": ["", "T01FS", "T02FS", "T03FS", "T04FS", "T05FS", "T06FS", "T07FS, T11FS", "T05FS", "T09FS", "T10FS", "T04FS", "T12FS", "T04FS", "T14FS"], "Risk / Fırsat ID": ["W-01", "W-01", "R-01", "", "", "", "", "O-01", "", "", "", "", "", "W-02", "W-01"]})
    df_risks = pd.DataFrame({"ID": ["W-01", "W-02", "R-01", "O-01"], "Grup": ["Hava (API)", "Hava (API)", "Operasyonel", "Operasyonel"], "Tip": ["Risk", "Risk", "Risk", "Fırsat"], "Tanım": ["Yağış", "Rüzgar", "Vinç Arızası", "Ekip Ekleme"], "İhtimal (%)": [15, 8, 15, 20], "Etki (Gün)": [1.5, 1.0, 5.0, 4.0]})
    df_hist = pd.DataFrame({"Aktivite Kodu": ["BET-01", "BET-01", "BET-01", "BET-01", "BET-01", "SVA-01"], "Planlanan_Sure": [6.0, 7.0, 5.0, 7.0, 6.0, 8.0], "Gercek_Sure": [6.5, 7.7, 5.8, 8.0, 7.5, 9.5], "Aciklama": ["1. Kat", "2. Kat", "3. Kat", "4. Kat", "Zemin", "Örnek"]})
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_tasks.to_excel(writer, index=False, sheet_name='Görevler')
        df_risks.to_excel(writer, index=False, sheet_name='Riskler')
        df_hist.to_excel(writer, index=False, sheet_name='Tablo_C')
    return output.getvalue()

# ==========================================
# 2. VERİ TABANI BAŞLATMA VE ARŞİV
# ==========================================
if 'risk_df' not in st.session_state:
    st.session_state.risk_df = pd.DataFrame({"ID": ["R-01", "O-01"], "Grup": ["Operasyonel", "Operasyonel"], "Tip": ["Risk", "Fırsat"], "Tanım": ["Kule Vinç Arızası", "Ekstra Ekiplerin Gelmesi"], "İhtimal (%)": [25.0, 20.0], "Etki (Gün)": [5.0, 3.0]})

if 'tasks_df' not in st.session_state:
    st.session_state.tasks_df = pd.DataFrame({"ID": [f"T{i:02d}" for i in range(1, 16)], "WBS": ["1. Kaba İşler"]*4 + ["2. İnce İşler"]*4 + ["3. Elektrik İşleri"]*3 + ["4. Mekanik İşler"]*2 + ["5. Cephe İşleri", "6. Peyzaj İşleri"], "Aktivite Kodu": ["HAF-01", "BET-01", "BET-01", "BET-01", "DUV-01", "SVA-01", "SVA-01", "SER-01", "ELK-01", "ELK-01", "ELK-01", "MEK-01", "MEK-01", "CPH-01", "PEY-01"], "Görev": ["Hafriyat", "Temel Kalıp", "Kolon/Perde", "Döşeme", "Duvar", "Kaba Sıva", "İnce Sıva", "Seramik", "Kablo", "Pano", "Aydınlatma", "Borulama", "Kanal", "İskele", "Peyzaj"], "Alt Yüklenici": ["Hafriyatçı A.Ş.", "Beton A.Ş.", "Beton A.Ş.", "Beton A.Ş.", "İnce İşler Ltd.", "İnce İşler Ltd.", "Boya A.Ş.", "İnce İşler Ltd.", "Elektrik Ltd.", "Elektrik Ltd.", "Elektrik Ltd.", "Mekanik A.Ş.", "Mekanik A.Ş.", "Cephe Ltd.", "Peyzaj A.Ş."], "Bütçe Tutarı (TL)": [150000, 400000, 600000, 450000, 200000, 180000, 120000, 300000, 150000, 250000, 100000, 280000, 220000, 750000, 80000], "Planlanan Süre": [6, 7, 10, 7, 9, 7, 6, 10, 7, 5, 6, 9, 7, 14, 6], "İlerleme (%)": [0]*15, "Gerçekleşen Süre": [0]*15, "Duruş (Gün)": [0]*15, "Dağılım Tipi": ["BetaPERT", "Triangular", "Triangular", "Uniform", "Triangular", "Normal", "Lognormal", "BetaPERT", "Triangular", "Uniform", "Triangular", "Triangular", "Normal", "BetaPERT", "Uniform"], "İyimser": [4, 5, 8, 5, 6, 5, 4, 7, 5, 3, 4, 6, 5, 10, 4], "Olası": [6, 7, 10, 7, 9, 7, 6, 10, 7, 5, 6, 9, 7, 14, 6], "Kötümser": [10, 12, 15, 11, 15, 11, 9, 15, 11, 8, 9, 14, 10, 20, 10], "Öncüller": ["", "T01FS", "T02FS", "T03FS", "T04FS", "T05FS", "T06FS", "T07FS, T11FS", "T05FS", "T09FS", "T10FS", "T04FS", "T12FS", "T04FS", "T14FS"], "Risk / Fırsat ID": ["", "", "R-01", "", "", "", "", "O-01", "", "", "", "", "", "", ""]})

if 'scenario_archive' not in st.session_state: st.session_state.scenario_archive = {}
if 'baseline_data' not in st.session_state: st.session_state.baseline_data = None
if 'param_library' not in st.session_state: st.session_state.param_library = None

st.title("🎓 Akademik Şantiye Simülatörü & Raporlama Modülü (v60.19)")
st.divider()

# ==========================================
# YAN MENÜ: AYARLAR VE ÖZELLİKLER
# ==========================================
with st.sidebar:
    st.header("🏢 Proje Kimliği")
    p_adi = st.text_input("Proje Adı", placeholder="Örn: USÛL TEKNİK A.Ş. Merkez")
    baslangic_tarihi = st.date_input("Planlanan Başlangıç", datetime.today())
    n_simulations = st.slider("İterasyon Sayısı", 100, 5000, 1000, step=100)
    
    st.markdown("---")
    st.header("📉 EVM Performans Ayarı")
    apply_spi = st.checkbox("🔮 EVM Gelecek Tahmini", value=True, help="Aktifse: Mevcut düşük performans, henüz başlanmayan işlerin de süresini uzatır. Pasifse: Başlanmayan işler planlandığı gibi biter varsayılır.")

    st.markdown("---")
    st.header("🛡️ Senaryo 1: Risk İptali")
    risk_list = ["Uygulanmasın"] + list(st.session_state.risk_df["ID"].dropna().unique())
    secili_iptal_risk = st.selectbox("Etkisi Sıfırlanacak Risk", risk_list)

    st.markdown("---")
    st.header("🎯 Senaryo 2: Taşeron Performansı")
    alt_yukleniciler = ["Uygulanmasın"] + list(st.session_state.tasks_df["Alt Yüklenici"].dropna().unique())
    secili_tasaronlar_raw = st.multiselect("Simüle Edilecek Alt Yüklenici", alt_yukleniciler, default=["Uygulanmasın"])
    secili_tasaronlar = [t for t in secili_tasaronlar_raw if t != "Uygulanmasın"]
    perf_iyilesme = st.slider("Performans Değişimi (%)", min_value=-50, max_value=50, value=0, step=5) if secili_tasaronlar else 0

    st.markdown("---")
    st.header("📈 Aşama 2: K-S Testi (Parametre Kütüphanesi)")
    ks_file = st.file_uploader("Tablo C (Geçmiş Veriler) Yükle", type=["xlsx"])
    if ks_file:
        try:
            df_hist = pd.read_excel(ks_file, sheet_name="Tablo_C")
            if "Aktivite Kodu" in df_hist.columns and "Gercek_Sure" in df_hist.columns and "Planlanan_Sure" in df_hist.columns:
                if st.button("K-S Testini Çalıştır ve Ana Tabloyu Güncelle", use_container_width=True):
                    with st.spinner("İstatistiksel sapma testleri uygulanıyor..."):
                        new_params = perform_ks_test(df_hist)
                        if not new_params.empty:
                            st.session_state.param_library = new_params.copy() # AS2-AS5 İÇİN KÜTÜPHANE KAYDI
                            for _, r in new_params.iterrows():
                                idx_list = st.session_state.tasks_df.index[st.session_state.tasks_df['Aktivite Kodu'] == r['Aktivite Kodu']].tolist()
                                for i in idx_list:
                                    plan_sure = safe_float(st.session_state.tasks_df.at[i, 'Planlanan Süre'], 1.0)
                                    st.session_state.tasks_df.at[i, 'Dağılım Tipi'] = r['Dağılım Tipi']
                                    st.session_state.tasks_df.at[i, 'İyimser'] = round(plan_sure * r['Min_Ratio'], 1)
                                    st.session_state.tasks_df.at[i, 'Olası'] = round(plan_sure * r['Med_Ratio'], 1)
                                    st.session_state.tasks_df.at[i, 'Kötümser'] = round(plan_sure * r['Max_Ratio'], 1)
                            st.success(f"Başarılı! {len(new_params)} Havuz için kütüphane oluşturuldu ve süreler hesaplandı.")
        except Exception as e: st.error(f"Tablo okunamadı: {e}")

    st.markdown("---")
    st.header("⛈️ Dinamik İklim Kütüğü (API)")
    m = folium.Map(location=[39.0, 35.0], zoom_start=5)
    m.add_child(folium.LatLngPopup())
    map_data = st_folium(m, height=250, use_container_width=True)
    lat, lon = (map_data["last_clicked"]["lat"], map_data["last_clicked"]["lng"]) if map_data and map_data.get("last_clicked") else (41.0082, 28.9784)
    
    c1, c2, c3 = st.columns(3)
    with c1: c_yagmur = st.checkbox("🌧️ Yağ", value=True)
    with c2: c_ruzgar = st.checkbox("💨 Rüz")
    with c3: c_don = st.checkbox("❄️ Don")
    
    ce1, ce2, ce3 = st.columns(3)
    with ce1: yagis_esik = st.number_input("Yağ.(mm)", min_value=0.1, value=3.0, step=0.5)
    with ce2: ruzgar_esik = st.number_input("Rüz.(km)", min_value=10.0, value=45.0, step=5.0)
    with ce3: sicaklik_esik = st.number_input("Sıc.(°C)", max_value=15.0, value=0.0, step=1.0)
    
    ay_isimleri = {1:"Ocak", 2:"Şubat", 3:"Mart", 4:"Nisan", 5:"Mayıs", 6:"Haziran", 7:"Temmuz", 8:"Ağustos", 9:"Eylül", 10:"Ekim", 11:"Kasım", 12:"Aralık"}
    secilen_aylar = st.multiselect("Ayları Seçin", options=list(ay_isimleri.keys()), default=[1, 2, 3], format_func=lambda x: ay_isimleri[x], label_visibility="collapsed")

    if st.button("📡 API'den İklim Verisini Çek", type="primary", use_container_width=True):
        if not (c_yagmur or c_ruzgar or c_don): 
            st.error("Seçim yapınız!")
        elif lat and lon and secilen_aylar:
            with st.spinner("Geçmiş 5 yıl taranıyor..."):
                try:
                    df_weather = fetch_weather_data(lat, lon, (datetime.today().date() - timedelta(days=365*5)).strftime("%Y-%m-%d"), (datetime.today().date() - timedelta(days=5)).strftime("%Y-%m-%d"))
                    df_filtered = df_weather[df_weather['Tarih'].dt.month.isin(secilen_aylar)]
                    tdays = max(len(df_filtered), 1)
                    aylar_str = "-".join([ay_isimleri[m] for m in secilen_aylar])
                    
                    w_ids = [int(str(x).split('-')[1]) for x in st.session_state.risk_df['ID'].dropna() if str(x).startswith('W-') and len(str(x).split('-'))==2]
                    nxt = max(w_ids) + 1 if w_ids else 1
                    
                    new_rows = []
                    if c_yagmur: 
                        new_rows.append({"ID": f"W-{nxt:02d}", "Grup": "Hava (API)", "Tip": "Risk", "Tanım": f"Aşırı Yağış ({aylar_str})", "İhtimal (%)": round((len(df_filtered[df_filtered['Yagis_mm'] >= yagis_esik]) / tdays)*100, 1), "Etki (Gün)": 1.0})
                        nxt += 1
                    if c_ruzgar: 
                        new_rows.append({"ID": f"W-{nxt:02d}", "Grup": "Hava (API)", "Tip": "Risk", "Tanım": f"Rüzgar ({aylar_str})", "İhtimal (%)": round((len(df_filtered[df_filtered['Ruzgar_kmh'] >= ruzgar_esik]) / tdays)*100, 1), "Etki (Gün)": 1.0})
                        nxt += 1
                    if c_don: 
                        new_rows.append({"ID": f"W-{nxt:02d}", "Grup": "Hava (API)", "Tip": "Risk", "Tanım": f"Don/Buzlanma ({aylar_str})", "İhtimal (%)": round((len(df_filtered[df_filtered['Sicaklik_min'] <= sicaklik_esik]) / tdays)*100, 1), "Etki (Gün)": 1.0})
                    
                    if new_rows: 
                        st.session_state.risk_df = pd.concat([st.session_state.risk_df, pd.DataFrame(new_rows)], ignore_index=True)
                        st.rerun()
                except Exception as e: 
                    st.error(f"Hata: {e}")

    # --- JSON KAYIT VE GERİ YÜKLEME SİSTEMİ ---
    st.markdown("---")
    st.header("💾 Proje Kayıt İşlemleri")
    
    def export_to_json():
        export_data = {
            "tasks": st.session_state.tasks_df.to_dict(orient="records"),
            "risks": st.session_state.risk_df.to_dict(orient="records")
        }
        return json.dumps(export_data, ensure_ascii=False, indent=4)
        
    dosya_ismi = f"{p_adi.replace(' ', '_') if p_adi else 'santiye_projesi'}.json"
    st.download_button(label="📥 Projeyi Kaydet (JSON)", data=export_to_json(), file_name=dosya_ismi, mime="application/json", use_container_width=True)
    
    st.markdown("👇 **Kayıtlı Projeyi Yükle**")
    uploaded_json = st.file_uploader("Yüklemek için JSON dosyası seçin", type=["json"], label_visibility="collapsed")
    if uploaded_json is not None:
        if st.button("Yükle ve Verileri Güncelle", use_container_width=True, type="primary"):
            try:
                loaded_data = json.load(uploaded_json)
                st.session_state.tasks_df = pd.DataFrame(loaded_data["tasks"])
                st.session_state.risk_df = pd.DataFrame(loaded_data["risks"])
                st.session_state.scenario_archive = {}
                st.session_state.baseline_data = None
                if 'total_base' in st.session_state: del st.session_state['total_base']
                st.success("Proje başarıyla yüklendi!")
                st.rerun()
            except Exception as e:
                st.error(f"Dosya yüklenirken bir hata oluştu: {e}")

    # EKSİK OLAN EXCEL ŞABLON İNDİRME BUTONU BURAYA EKLENDİ
    st.markdown("---")
    st.download_button("📥 Boş Excel Şablonu İndir", data=generate_excel_template(), file_name="Tez_Gantt_Sablon.xlsx", use_container_width=True)

# ==========================================
# 3. VERİ GİRİŞ TABLOLARI 
# ==========================================
st.subheader("🚨 Tablo D: Dinamik Risk, Fırsat ve Hava Durumu Kütüğü")
st.session_state.risk_df = st.data_editor(st.session_state.risk_df, num_rows="dynamic", use_container_width=True, height=200)

st.write("")
st.subheader("📋 Tablo A-B: Görevler, SPI Performans Girişi ve Belirsizlik Dağılımları")
st.session_state.tasks_df = st.data_editor(st.session_state.tasks_df, num_rows="dynamic", height=400, use_container_width=True, column_config={
    "Dağılım Tipi": st.column_config.SelectboxColumn("Dağılım Tipi", options=["Triangular", "BetaPERT", "Uniform", "Normal", "Lognormal"], required=True),
    "Bütçe Tutarı (TL)": st.column_config.NumberColumn("Bütçe Tutarı (TL)", min_value=0, format="%d ₺"),
    "İlerleme (%)": st.column_config.NumberColumn("İlerleme (%)", min_value=0, max_value=100),
    "Gerçekleşen Süre": st.column_config.NumberColumn("Gerçekleşen Süre", min_value=0.0),
    "Duruş (Gün)": st.column_config.NumberColumn("Dış Etken/Duruş (Gün)", min_value=0.0)
})

# ==========================================
# SİMÜLASYON VE AĞ ANALİZİ MOTORU
# ==========================================
def calculate_ccpm_network(tasks_df, durations_dict):
    p50_durations = {tid: np.percentile(d_arr, 50) for tid, d_arr in durations_dict.items() if len(d_arr) > 0}
    p80_durations = {tid: np.percentile(d_arr, 80) for tid, d_arr in durations_dict.items() if len(d_arr) > 0}
    es, ef = {tid: 0.0 for tid in p50_durations.keys()}, {tid: 0.0 for tid in p50_durations.keys()}
    successors, preds_map = {tid: [] for tid in p50_durations.keys()}, {}
    for _, row in tasks_df.iterrows():
        tid = str(row['ID']).upper()
        if tid not in p50_durations: continue
        preds = [p.strip().upper() for p in str(row['Öncüller']).split(',') if p.strip()]
        preds_cleaned = []
        for p in preds:
            p_id = re.sub(r'[+-].*', '', p.replace('FS', '').replace('SS', '').replace('FF', '').replace('SF', '')).strip()
            if p_id in p50_durations:
                preds_cleaned.append(p_id)
                successors[p_id].append(tid)
        preds_map[tid] = preds_cleaned

    def get_es(task, path_stack=None):
        if path_stack is None: path_stack = set()
        if task in path_stack: return 0.0 
        path_stack.add(task)
        val = max([get_es(p, path_stack) + p50_durations[p] for p in preds_map[task]]) if preds_map.get(task) else 0.0
        path_stack.remove(task)
        return val

    for tid in p50_durations.keys(): es[tid], ef[tid] = get_es(tid), get_es(tid) + p50_durations[tid]
    project_duration = max(ef.values()) if ef else 0.0
    ls, lf = {tid: 0.0 for tid in p50_durations.keys()}, {tid: project_duration for tid in p50_durations.keys()}
    
    def get_lf(task, path_stack=None):
        if path_stack is None: path_stack = set()
        if task in path_stack: return project_duration
        path_stack.add(task)
        val = min([get_lf(s, path_stack) - p50_durations[s] for s in successors[task]]) if successors.get(task) else project_duration
        path_stack.remove(task)
        return val

    for tid in p50_durations.keys(): lf[tid], ls[tid] = get_lf(tid), get_lf(tid) - p50_durations[tid]
    slack = {tid: round(ls[tid] - es[tid], 2) for tid in p50_durations.keys()}
    is_critical = {tid: (slack[tid] <= 0.1) for tid in p50_durations.keys()}
    
    feeding_buffers = []
    for tid in p50_durations.keys():
        if not is_critical[tid] and not str(tid).startswith("FB_"): 
            for succ in successors[tid]:
                if is_critical[succ]:  
                    fb_size = max(0, round(p80_durations[tid] - p50_durations[tid], 1))
                    if fb_size > 0: feeding_buffers.append({"Yan Yol (Kritik Olmayan)": tid, "Kavşak Noktası (Kritik)": succ, "Bolluk (Gün)": slack[tid], "Besleme Tamponu (FB) Önerisi": fb_size})
    
    network_results = [{"ID": tid, "Aktivite": tasks_df[tasks_df['ID'] == tid]['Görev'].values[0] if not tasks_df[tasks_df['ID'] == tid].empty else tid, "P50 Süre": round(p50_durations[tid], 1), "Erken Bitiş (EF)": round(ef[tid], 1), "Geç Bitiş (LF)": round(lf[tid], 1), "Bolluk (Slack)": slack[tid], "Kritik Yol": "🚨 EVET" if is_critical[tid] else "Hayır"} for tid in p50_durations.keys()]
    return pd.DataFrame(network_results), pd.DataFrame(feeding_buffers), is_critical

def run_monte_carlo(t_df, r_df, n_sim, apply_spi_flag=True, scenario_sub=None, scenario_boost=0, cancelled_risk=None):
    risk_dict = {safe_str(r.get("ID")).upper(): {"prob": safe_float(r.get("İhtimal (%)"))/100.0, "impact": safe_float(r.get("Etki (Gün)")), "mult": -1 if "fırsat" in safe_str(r.get("Tip")).lower() else 1} for _, r in r_df.iterrows() if safe_str(r.get("ID")) != "-"}
    trade_ev, trade_ad = {}, {}
    for _, row in t_df.iterrows():
        akod = safe_str(row.get('Aktivite Kodu'))
        ilerleme, gercek, durus, plan = safe_float(row.get('İlerleme (%)')), safe_float(row.get('Gerçekleşen Süre')), safe_float(row.get('Duruş (Gün)')), safe_float(row.get('Planlanan Süre'))
        if ilerleme > 0 or gercek > 0:
            trade_ev[akod] = trade_ev.get(akod, 0) + (plan * (ilerleme / 100.0))
            trade_ad[akod] = trade_ad.get(akod, 0) + max(gercek - durus, 0.1)
            
    glob_spi = sum(trade_ev.values()) / sum(trade_ad.values()) if sum(trade_ad.values()) > 0 else 1.0
    trade_spi = {akod: trade_ev[akod] / trade_ad[akod] if trade_ad[akod] > 0 else 1.0 for akod in trade_ev}
    durations, preds_dict = {}, {}
    c_risk_id = safe_str(cancelled_risk).upper() if cancelled_risk and cancelled_risk != "Uygulanmasın" else None

    for _, row in t_df.iterrows():
        tid = safe_str(row.get('ID')).upper()
        if not tid or tid == "-": continue
        ilerleme, gercek, akod = safe_float(row.get('İlerleme (%)')), safe_float(row.get('Gerçekleşen Süre')), safe_str(row.get('Aktivite Kodu'))
        
        if ilerleme >= 100:
            durations[tid] = np.full(n_sim, gercek)
            preds_dict[tid] = safe_str(row.get('Öncüller', ''))
            continue
            
        if apply_spi_flag:
            spi_k = max(trade_spi.get(akod, glob_spi), 0.1)
        else:
            spi_k = max(trade_spi.get(akod, glob_spi), 0.1) if ilerleme > 0 else 1.0

        rem_ratio = (100.0 - ilerleme) / 100.0
        iyimser_base = safe_float(row.get('İyimser'))
        olasi_base = max(iyimser_base, safe_float(row.get('Olası')))
        kotumser_base = max(olasi_base, safe_float(row.get('Kötümser')))
        
        if scenario_sub and safe_str(row.get('Alt Yüklenici')) in scenario_sub:
            factor = 1.0 - (scenario_boost / 100.0)
            olasi_base, kotumser_base = max(iyimser_base, olasi_base * factor), max(olasi_base, kotumser_base * factor)
            
        r_iy, r_ol, r_ko = (iyimser_base * rem_ratio) / spi_k, (olasi_base * rem_ratio) / spi_k, (kotumser_base * rem_ratio) / spi_k
        dagilim_tipi = safe_str(row.get('Dağılım Tipi'))
        if dagilim_tipi == "Uniform": base_dur = np.random.uniform(r_iy, r_ko, n_sim)
        elif dagilim_tipi == "BetaPERT":
            if r_ko == r_iy: base_dur = np.full(n_sim, r_ol)
            else:
                alpha, beta_param = 1 + 4 * (r_ol - r_iy) / (r_ko - r_iy), 1 + 4 * (r_ko - r_ol) / (r_ko - r_iy)
                base_dur = r_iy + np.random.beta(alpha, beta_param, n_sim) * (r_ko - r_iy)
        elif dagilim_tipi == "Normal":
            base_dur = np.clip(np.random.normal((r_iy + 4 * r_ol + r_ko) / 6, (r_ko - r_iy) / 6 if r_ko > r_iy else 0.1, n_sim), r_iy, r_ko)
        elif dagilim_tipi == "Lognormal":
            mu, sigma = (r_iy + 4 * r_ol + r_ko) / 6, (r_ko - r_iy) / 6 if r_ko > r_iy else 0.1
            if mu > 0:
                sigma2 = np.log(1 + (sigma/mu)**2)
                base_dur = np.clip(np.random.lognormal(np.log(mu) - sigma2/2, np.sqrt(sigma2), n_sim), r_iy, r_ko)
            else: base_dur = np.full(n_sim, r_ol)
        else: base_dur = np.random.triangular(r_iy, r_ol, r_ko, n_sim)
            
        for r_id in [r.strip().upper() for r in safe_str(row.get('Risk / Fırsat ID', '')).split(',') if r.strip()]:
            if r_id in risk_dict and r_id != c_risk_id:
                base_dur += (np.random.binomial(1, risk_dict[r_id]["prob"], n_sim) * risk_dict[r_id]["impact"] * risk_dict[r_id]["mult"])
                
        durations[tid] = gercek + np.maximum(base_dur, 0.1)
        preds_dict[tid] = safe_str(row.get('Öncüller', ''))

    start_times, end_times = {}, {}
    def calc_times(task_id, path_stack=None):
        if path_stack is None: path_stack = set()
        if task_id in path_stack: return np.zeros(n_sim), np.zeros(n_sim)
        if task_id in start_times: return start_times[task_id], end_times[task_id]
        path_stack.add(task_id) 
        p_str = preds_dict.get(task_id, "")
        task_dur = durations.get(task_id, np.ones(n_sim))
        if not p_str or p_str == "-": st_time = np.zeros(n_sim)
        else:
            p_reqs = []
            for p in [p.strip() for p in p_str.split(",") if p.strip()]:
                pred_id, l_type, p_lag = parse_single_pred(p)
                if pred_id in durations:
                    p_st, p_en = calc_times(pred_id, path_stack)
                    if l_type == "FS": p_reqs.append(p_en + p_lag)
                    elif l_type == "SS": p_reqs.append(p_st + p_lag)
                    elif l_type == "FF": p_reqs.append(p_en + p_lag - task_dur)
                    elif l_type == "SF": p_reqs.append(p_st + p_lag - task_dur)
            st_time = np.maximum.reduce(p_reqs) if p_reqs else np.zeros(n_sim)
        st_time = np.maximum(st_time, 0) 
        start_times[task_id], end_times[task_id] = st_time, st_time + task_dur
        path_stack.remove(task_id)
        return start_times[task_id], end_times[task_id]

    for t in durations.keys(): calc_times(t)
    return start_times, end_times, durations

if not 'total_base' in st.session_state:
    st.info("👆 Lütfen verilerinizi ayarladıktan sonra yukarıdaki **'Modeli Çalıştır'** butonuna basın.")

if st.button("🎲 Karar Destek Modelini Çalıştır", type="primary", use_container_width=True):
    with st.spinner("Modeller işleniyor, Baseline & İlerleme verileri çekiliyor..."):
        st_base, et_base, dur_base = run_monte_carlo(st.session_state.tasks_df, st.session_state.risk_df, n_simulations, apply_spi_flag=apply_spi)
        total_base = np.maximum.reduce(list(et_base.values())) if et_base else np.zeros(n_simulations)
        
        st_active, et_active, dur_active = st_base, et_base, dur_base 
        
        if secili_tasaronlar or secili_iptal_risk != "Uygulanmasın":
            st_scen, et_scen, dur_scen = run_monte_carlo(st.session_state.tasks_df, st.session_state.risk_df, n_simulations, apply_spi_flag=apply_spi, scenario_sub=secili_tasaronlar, scenario_boost=perf_iyilesme, cancelled_risk=secili_iptal_risk)
            total_scen = np.maximum.reduce(list(et_scen.values())) if et_scen else np.zeros(n_simulations)
            st_active, et_active, dur_active = st_scen, et_scen, dur_scen
            s_parts = []
            if secili_tasaronlar: s_parts.append(f"Hız/Yavaş: {', '.join(secili_tasaronlar)} ({perf_iyilesme:+}%)")
            if secili_iptal_risk != "Uygulanmasın": s_parts.append(f"-Risk: {secili_iptal_risk}")
            st.session_state.scen_name = " | ".join(s_parts)
        else:
            total_scen, st.session_state.scen_name = total_base, "Değişiklik Yapılmadı"
            
        df_net, df_fb, is_critical_map = calculate_ccpm_network(st.session_state.tasks_df, dur_active)
        st.session_state.df_network, st.session_state.df_fb = df_net, df_fb
        
        gantt_data, cashflow_series_list = [], []
        b_date = pd.to_datetime(baslangic_tarihi)
        
        for _, row in st.session_state.tasks_df.iterrows():
            tid = safe_str(row.get('ID')).upper()
            ilerleme = safe_float(row.get('İlerleme (%)'))
            if tid in st_active:
                a_s, a_e = float(np.mean(st_active[tid])), float(np.mean(et_active[tid])) 
                s_d, e_d = b_date + timedelta(days=a_s), b_date + timedelta(days=a_e)
                if s_d == e_d: e_d += timedelta(hours=12)
                is_crit = is_critical_map.get(tid, False)
                görev_isim = f"{'🔴 ' if is_crit else ''}{'✅ ' if ilerleme >= 100 else (f'🔄(%{int(ilerleme)}) ' if ilerleme > 0 else '')}{tid} - {safe_str(row.get('Görev'))}"
                gantt_data.append({"ID": tid, "WBS Grubu": safe_str(row.get('WBS')), "Görev İsim": safe_str(row.get('Görev')), "Başlangıç": s_d, "Bitiş": e_d, "Süre_Num": round(a_e-a_s, 1), "Görev Gösterim": görev_isim, "Kritik": is_crit})
                
                budget = safe_float(row.get('Bütçe Tutarı (TL)'))
                if budget > 0:
                    delta_days = (e_d.date() - s_d.date()).days
                    if delta_days > 0:
                        daily_cost = budget / delta_days
                        cashflow_series_list.append(pd.Series(daily_cost, index=pd.date_range(start=s_d.date(), periods=delta_days, freq='D')))
        
        df_g = pd.DataFrame(gantt_data).sort_values(by=["Başlangıç", "ID"])
        p50, p80, p90 = int(np.percentile(total_base, 50)), int(np.percentile(total_base, 80)), int(np.percentile(total_base, 90))
        p50_scen, p80_scen, p90_scen = int(np.percentile(total_scen, 50)), int(np.percentile(total_scen, 80)), int(np.percentile(total_scen, 90))
        buf = p80_scen - p50_scen
        
        if buf > 0 and not df_g.empty:
            max_d = df_g["Bitiş"].max()
            df_g.loc[len(df_g)] = {"ID": "TAMPON", "WBS Grubu": "7. Tamponlar", "Görev İsim": "Proje Tamponu", "Başlangıç": max_d, "Bitiş": max_d + timedelta(days=float(buf)), "Süre_Num": float(buf), "Görev Gösterim": f"TAMPON - Proje Tamponu ({buf} Gün)", "Kritik": False}
        
        st.session_state.df_gantt, st.session_state.total_base, st.session_state.total_scen = df_g, total_base, total_scen
        st.session_state.dur_base = dur_active 
        st.session_state.kpi_base = {"p50": p50, "p80": p80, "p90": p90}
        st.session_state.kpi_scen = {"p50": p50_scen, "p80": p80_scen, "p90": p90_scen}
        
        if cashflow_series_list:
            df_cf_daily = pd.DataFrame({"Tutar (TL)": pd.concat(cashflow_series_list).groupby(level=0).sum()})
            df_cf_daily.index.name = 'Tarih' 
            df_cf_monthly = df_cf_daily.resample('ME').sum().reset_index()
            df_cf_monthly['Ay'] = df_cf_monthly['Tarih'].dt.strftime('%Y-%m')
            df_cf_monthly = df_cf_monthly[['Ay', 'Tutar (TL)']].sort_values('Ay')
            df_cf_monthly['Kümülatif Hakediş (TL)'] = df_cf_monthly['Tutar (TL)'].cumsum()
            st.session_state.df_monthly_cf = df_cf_monthly
        else: st.session_state.df_monthly_cf = pd.DataFrame()

# ==========================================
# 5. EKRAN ÇIKTILARI VE SEKMELER
# ==========================================
if 'total_base' in st.session_state and 'df_gantt' in st.session_state:
    st.divider()

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📊 Proje Planı", "🔗 Ağ Analizi (FB)", "🌪️ Pareto & Tornado", 
        "🔄 Senaryo (What-If)", "🎓 Tez Çıktıları & Kütüphane", "💰 Nakit Akışı", "📚 Kaynakça", "📍 Baseline Takibi"
    ])
    
    with tab1:
        kl = st.session_state.df_gantt["Görev Gösterim"].tolist()
        wbs_sorted = sorted(list(st.session_state.df_gantt['WBS Grubu'].unique()))
        c_map = {wbs: "#9333EA" if "7. Tamponlar" in wbs else px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)] for i, wbs in enumerate(wbs_sorted)}
        fig_g = px.timeline(st.session_state.df_gantt, x_start="Başlangıç", x_end="Bitiş", y="Görev Gösterim", color="WBS Grubu", text="Süre_Num", color_discrete_map=c_map, category_orders={"WBS Grubu": wbs_sorted})
        df_crit = st.session_state.df_gantt[st.session_state.df_gantt["Kritik"] == True].sort_values("Başlangıç")
        if not df_crit.empty:
            path_x, path_y, prev_row = [], [], None
            for _, c_row in df_crit.iterrows():
                if prev_row is not None:
                    path_x.extend([prev_row["Bitiş"], prev_row["Bitiş"], c_row["Başlangıç"], None])
                    path_y.extend([prev_row["Görev Gösterim"], c_row["Görev Gösterim"], c_row["Görev Gösterim"], None])
                prev_row = c_row
            fig_g.add_trace(go.Scatter(x=path_x, y=path_y, mode="lines", line=dict(color="red", width=2, dash="dot"), name="Kritik Akış", hoverinfo="skip"))
        fig_g.update_layout(plot_bgcolor='white', paper_bgcolor='white', margin=dict(l=10, r=100, t=40, b=10), yaxis=dict(autorange="reversed", categoryorder="array", categoryarray=kl), height=max(500, len(st.session_state.df_gantt)*40))
        st.plotly_chart(fig_g, use_container_width=True)

    with tab2:
        st.markdown("### Kritik Yollar ve Bolluk Süreleri")
        st.dataframe(st.session_state.df_network, use_container_width=True)
        st.markdown("### 🛡️ Önerilen Besleme Tamponları")
        if not st.session_state.df_fb.empty:
            df_fb_edit = st.session_state.df_fb.copy()
            if "Uygula" not in df_fb_edit.columns: df_fb_edit.insert(0, "Uygula", False)
            edited_fb = st.data_editor(df_fb_edit, column_config={"Uygula": st.column_config.CheckboxColumn("Seç", default=False)}, disabled=["Yan Yol (Kritik Olmayan)", "Kavşak Noktası (Kritik)", "Bolluk (Gün)", "Besleme Tamponu (FB) Önerisi"], use_container_width=True, hide_index=True)
            if st.button("✅ Seçili Tamponları Plana İşle", type="primary"):
                to_apply = edited_fb[edited_fb["Uygula"] == True]
                if not to_apply.empty:
                    new_tasks = []
                    for _, row in to_apply.iterrows():
                        y_yol, k_yol, fb_val = row["Yan Yol (Kritik Olmayan)"], row["Kavşak Noktası (Kritik)"], row["Besleme Tamponu (FB) Önerisi"]
                        fb_id = f"FB_{y_yol}"
                        if fb_id not in st.session_state.tasks_df['ID'].values:
                            new_tasks.append({"ID": fb_id, "WBS": "7. Tamponlar", "Aktivite Kodu": "TAMPON", "Görev": f"Tampon ({y_yol}->{k_yol})", "Alt Yüklenici": "Sistem", "Bütçe Tutarı (TL)": 0, "Planlanan Süre": fb_val, "İlerleme (%)": 0, "Gerçekleşen Süre": 0, "Duruş (Gün)": 0, "Dağılım Tipi": "Uniform", "İyimser": fb_val, "Olası": fb_val, "Kötümser": fb_val, "Öncüller": f"{y_yol}FS", "Risk / Fırsat ID": ""})
                            idx = st.session_state.tasks_df.index[st.session_state.tasks_df['ID'] == k_yol].tolist()[0]
                            st.session_state.tasks_df.at[idx, 'Öncüller'] = ", ".join([p.replace(y_yol, fb_id, 1) if p.startswith(y_yol) else p for p in str(st.session_state.tasks_df.at[idx, 'Öncüller']).split(',')])
                    if new_tasks:
                        st.session_state.tasks_df = pd.concat([st.session_state.tasks_df, pd.DataFrame(new_tasks)], ignore_index=True)
                        if 'total_base' in st.session_state: del st.session_state['total_base']
                        st.rerun()
                else: st.warning("Lütfen seçim yapınız.")
        else: st.success("Korunması gereken spesifik bir yan yol kavşağı bulunamadı.")

    with tab3:
        st.markdown("### 📊 Aktivite Tipine (WBS) Göre Gecikme Nedenleri (AS1)")
        wbs_risks = []
        for _, t_row in st.session_state.tasks_df.iterrows():
            r_ids = [x.strip().upper() for x in safe_str(t_row.get('Risk / Fırsat ID', '')).split(',') if x.strip()]
            for r_id in r_ids:
                if r_id != "-":
                    r_match = st.session_state.risk_df[st.session_state.risk_df['ID'].str.upper() == r_id]
                    if not r_match.empty:
                        r_row = r_match.iloc[0]
                        prob = safe_float(r_row.get("İhtimal (%)")) / 100.0
                        impact = safe_float(r_row.get("Etki (Gün)"))
                        r_name = safe_str(r_row.get("Tanım"))
                        if prob * impact > 0:
                            wbs_risks.append({"WBS Grubu": safe_str(t_row.get("WBS")), "Risk": r_name, "Beklenen Gecikme (Gün)": prob * impact})
        if wbs_risks:
            df_wbs_risk = pd.DataFrame(wbs_risks).groupby(['WBS Grubu', 'Risk']).sum().reset_index()
            fig_wbs = px.bar(df_wbs_risk, x="WBS Grubu", y="Beklenen Gecikme (Gün)", color="Risk", title="Hangi Aşama Hangi Riskten Ne Kadar Etkileniyor?", barmode='stack')
            fig_wbs.update_layout(plot_bgcolor='white'); st.plotly_chart(fig_wbs, use_container_width=True)
        else:
            st.info("Aktivitelere atanmış herhangi bir risk veya fırsat bulunamadı.")

        st.markdown("---")
        c_left, c_right = st.columns(2)
        with c_left:
            st.markdown("### 🌪️ Tornado Grafiği (Duyarlılık)")
            corrs_base = [{"Etken": f"{tid} - {st.session_state.tasks_df[st.session_state.tasks_df['ID']==tid]['Görev'].values[0] if not st.session_state.tasks_df[st.session_state.tasks_df['ID']==tid].empty else tid}", "Korelasyon (r)": round(np.corrcoef(d_arr, st.session_state.total_scen)[0, 1], 3)} for tid, d_arr in st.session_state.dur_base.items() if np.std(d_arr) > 0 and np.std(st.session_state.total_scen) > 0]
            df_c_base = pd.DataFrame(corrs_base).sort_values(by="Korelasyon (r)") if corrs_base else pd.DataFrame()
            if not df_c_base.empty:
                fig_t = px.bar(df_c_base, x="Korelasyon (r)", y="Etken", orientation='h', color="Korelasyon (r)", color_continuous_scale="Reds")
                fig_t.update_layout(plot_bgcolor='white', showlegend=False); st.plotly_chart(fig_t, use_container_width=True)
        with c_right:
            st.markdown("### 📊 Genel Gecikme Nedenleri (Pareto)")
            risk_impacts = []
            for _, r_row in st.session_state.risk_df.iterrows():
                r_id = safe_str(r_row.get('ID')).upper()
                if r_id == "-" or not r_id: continue
                prob, impact, r_name = safe_float(r_row.get("İhtimal (%)")) / 100.0, safe_float(r_row.get("Etki (Gün)")), safe_str(r_row.get("Tanım"))
                task_count = sum(1 for _, t_row in st.session_state.tasks_df.iterrows() if r_id in [x.strip().upper() for x in safe_str(t_row.get('Risk / Fırsat ID', '')).split(',') if x.strip()])
                if (prob * impact * task_count) > 0: risk_impacts.append({"Risk ID": r_id, "Nedeni": r_name, "Beklenen Gecikme (Gün)": prob * impact * task_count})
            if risk_impacts:
                df_r = pd.DataFrame(risk_impacts).sort_values(by="Beklenen Gecikme (Gün)", ascending=False)
                df_r["Kümülat %"] = 100 * df_r["Beklenen Gecikme (Gün)"].cumsum() / df_r["Beklenen Gecikme (Gün)"].sum()
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_r["Nedeni"], y=df_r["Beklenen Gecikme (Gün)"], name="Gecikme", marker_color='indianred'))
                fig.add_trace(go.Scatter(x=df_r["Nedeni"], y=df_r["Kümülat %"], name="Kümülatif %", yaxis='y2', mode='lines+markers', line=dict(color='blue', width=2)))
                fig.update_layout(yaxis2=dict(overlaying='y', side='right', range=[0, 105]), plot_bgcolor='white', hovermode="x unified"); st.plotly_chart(fig, use_container_width=True)

    with tab4:
        st.header("🎯 Senaryo Karşılaştırması ve Baseline")
        col_scen1, col_scen2, col_scen3 = st.columns([2, 1, 1])
        with col_scen1: scen_save_name = st.text_input("Senaryo/Baseline İsmi:", value=st.session_state.scen_name)
        with col_scen2:
            st.write(""); 
            if st.button("📥 Arşive Ekle", use_container_width=True):
                if st.session_state.scen_name != "Değişiklik Yapılmadı" or scen_save_name != "Değişiklik Yapılmadı":
                    st.session_state.scenario_archive[scen_save_name] = {"totals": st.session_state.total_scen, "kpi": st.session_state.kpi_scen}
                    st.success("Eklendi!")
        with col_scen3:
            st.write(""); 
            if st.button("📍 BASELINE Ata", type="primary", use_container_width=True):
                st.session_state.baseline_data = {"name": scen_save_name, "df_gantt": st.session_state.df_gantt.copy(), "df_cf": st.session_state.df_monthly_cf.copy(), "kpi": st.session_state.kpi_scen.copy() if st.session_state.scen_name != "Değişiklik Yapılmadı" else st.session_state.kpi_base.copy()}
                st.success("Donduruldu!")

        secilen_senaryolar = []
        if st.session_state.scenario_archive:
            secilen_senaryolar = st.multiselect("Karşılaştır:", options=list(st.session_state.scenario_archive.keys()))
        
        # TEZ EKLENTİSİ: P50, P80, P90 Çizgileri
        fig_h = go.Figure()
        fig_h.add_trace(go.Histogram(x=st.session_state.total_base, name='Baz Durum', opacity=0.6, marker_color='grey'))
        fig_h.add_vline(x=st.session_state.kpi_base['p50'], line_dash="dot", line_color="green", annotation_text="Baz P50")
        fig_h.add_vline(x=st.session_state.kpi_base['p80'], line_dash="dash", line_color="black", annotation_text="Baz P80")
        fig_h.add_vline(x=st.session_state.kpi_base['p90'], line_dash="solid", line_color="red", annotation_text="Baz P90")

        if st.session_state.scen_name != "Değişiklik Yapılmadı" and st.session_state.scen_name not in secilen_senaryolar:
            fig_h.add_trace(go.Histogram(x=st.session_state.total_scen, name="Aktif: " + st.session_state.scen_name, opacity=0.7, marker_color='blue'))
            fig_h.add_vline(x=st.session_state.kpi_scen['p50'], line_dash="dot", line_color="lightgreen", annotation_text="Aktif P50")
            fig_h.add_vline(x=st.session_state.kpi_scen['p80'], line_dash="dash", line_color="darkblue", annotation_text="Aktif P80")
            fig_h.add_vline(x=st.session_state.kpi_scen['p90'], line_dash="solid", line_color="orange", annotation_text="Aktif P90")

        colors = px.colors.qualitative.Set2
        for i, s_name in enumerate(secilen_senaryolar):
            s_data = st.session_state.scenario_archive[s_name]
            fig_h.add_trace(go.Histogram(x=s_data["totals"], name=s_name, opacity=0.7, marker_color=colors[i % len(colors)]))
            fig_h.add_vline(x=s_data["kpi"]['p80'], line_dash="dash", line_color=colors[i % len(colors)], annotation_text=f"{s_name} P80")
            
        fig_h.update_layout(barmode='overlay', title="Dağılım Kayması (P50, P80, P90 Sınırları)", plot_bgcolor='white', xaxis_title="Proje Süresi (Gün)"); st.plotly_chart(fig_h, use_container_width=True)

    with tab5:
        st.header("🎓 Tez Çıktıları & Parametre Kütüphanesi")
        st.info("**AS2 & AS5 Çıktısı:** Kurumsal Parametre Kütüphanesi (Geçmiş veriler ile istatistiksel dağılım tespiti)")
        
        if st.session_state.param_library is not None and not st.session_state.param_library.empty:
            st.success("Tarihsel geçmiş verilere (Tablo C) uygulanan Kolmogorov-Smirnov (K-S) testleri sonucunda, aktivite bazlı optimum dağılım tipleri ve sapma endeksleri aşağıdaki kütüphaneye kaydedilmiştir.")
            st.dataframe(st.session_state.param_library.style.format({
                "Min_Ratio": "{:.2f}", "Med_Ratio": "{:.2f}", "Max_Ratio": "{:.2f}"
            }), use_container_width=True)
        else:
            st.warning("⚠️ Sol menüden (Aşama 2) K-S Testi çalıştırıldığında Kurumsal Parametre Kütüphanesi burada oluşacaktır.")

    with tab6:
        st.markdown("### 📈 Aylık Nakit Akışı")
        if 'df_monthly_cf' in st.session_state and not st.session_state.df_monthly_cf.empty:
            df_m = st.session_state.df_monthly_cf
            c1, c2, c3 = st.columns(3)
            with c1: st.metric("Sözleşme Bütçesi", f"{df_m['Tutar (TL)'].sum():,.0f} ₺")
            with c2: st.metric("Bitiş Ayı", f"{df_m['Ay'].max()}")
            with c3: st.metric("Zirve Hakediş Ayı", f"{df_m.loc[df_m['Tutar (TL)'].idxmax(), 'Ay']}")
            fig_cf = go.Figure()
            fig_cf.add_trace(go.Bar(x=df_m['Ay'], y=df_m['Tutar (TL)'], name='Aylık Hakediş', marker_color='steelblue', yaxis='y1'))
            fig_cf.add_trace(go.Scatter(x=df_m['Ay'], y=df_m['Kümülatif Hakediş (TL)'], name='Kümülatif', mode='lines+markers', line=dict(color='darkred', width=3), yaxis='y2'))
            if st.session_state.baseline_data is not None and not st.session_state.baseline_data['df_cf'].empty:
                df_b_cf = st.session_state.baseline_data['df_cf']
                fig_cf.add_trace(go.Scatter(x=df_b_cf['Ay'], y=df_b_cf['Kümülatif Hakediş (TL)'], name='BASELINE (Hedef)', mode='lines', line=dict(color='grey', width=3, dash='dot'), yaxis='y2'))
            fig_cf.update_layout(plot_bgcolor='white', hovermode="x unified", yaxis=dict(title='Aylık Tutar', side='left', showgrid=False), yaxis2=dict(title='Kümülatif Toplam', side='right', overlaying='y', showgrid=False), legend=dict(x=0.01, y=0.99, bgcolor='rgba(255,255,255,0.8)'))
            st.plotly_chart(fig_cf, use_container_width=True)
            st.dataframe(df_m.style.format({"Tutar (TL)": "{:,.0f} ₺", "Kümülatif Hakediş (TL)": "{:,.0f} ₺"}), use_container_width=True)

    with tab7:
        st.header("📚 Kullanılan Kaynaklar (Literatür)")
        st.markdown("""
        **1. K-S Testi, Uzman Sistem ve Monte Carlo:** 
        * Koulinas, G. K., vd. (2020). *Schedule Delay Risk Analysis in Construction Projects with a Simulation-Based Expert System.*
        
        **2. Dağılım Modellemeleri (BetaPERT, Triangular, Uniform):** 
        * Davis, R. (2006). *Stochastic Project Duration Analysis Using PERT-beta Distributions.*
        * Deng, J., & Jian, W. (2022). *Estimating Construction Project Duration and Costs upon Completion Using Monte Carlo Simulations.*
        * Hassi, H., El Mkhalet, M., & Lamdouar, N. (2025). *Scheduling Optimization and Risk Analysis Using Monte Carlo Simulation for a Construction Project in Morocco.*
        
        **3. CCPM (Kritik Zincir) ve Besleme Tamponları:** 
        * Tamalika, T., vd. (2024). *Implementation of the Critical Chain Project Management (CCPM) Model for Improving Time and Cost in a Project for House Type 36.*
        * Tenera, A. B. (2008). *Critical chain buffer sizing: a comparative study.*
        
        **4. Riskten Arındırılmış Dinamik EVM:** 
        * Colin, J., & Vanhoucke, M. (2014). *Setting tolerance limits for statistical project control using Earned Value Management.*
        * PMI (2019). *Practice Standard for Earned Value Management.*
        """)
        
    with tab8:
        st.header("📍 Baseline (Hedef Çizgisi) ve İlerleme Takibi")
        if st.session_state.baseline_data is not None:
            col_b1, col_b2 = st.columns([4, 1])
            with col_b1:
                st.success(f"**Aktif Baseline:** {st.session_state.baseline_data['name']}")
            with col_b2:
                if st.button("🗑️ Baseline'ı Sil", use_container_width=True):
                    st.session_state.baseline_data = None
                    st.rerun()
            
            b_kpi = st.session_state.baseline_data['kpi']
            a_kpi = st.session_state.kpi_scen if st.session_state.scen_name != "Değişiklik Yapılmadı" else st.session_state.kpi_base
            fark_p80 = a_kpi['p80'] - b_kpi['p80']
            
            # TEZ EKLENTİSİ: Hedefe Yetişme Olasılığı (AS3)
            prob_meeting = np.sum(st.session_state.total_scen <= b_kpi['p80']) / len(st.session_state.total_scen) * 100
            
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.metric("Baseline P80 Hedefi", f"{b_kpi['p80']} Gün")
            with c2: st.metric("Mevcut P80 Tahmini", f"{a_kpi['p80']} Gün", delta=f"{abs(fark_p80)} Gün {'Gecikme' if fark_p80>0 else 'Kazanç'}", delta_color="inverse" if fark_p80>0 else "normal")
            with c3: st.metric("Hedefe Yetişme Olasılığı", f"%{prob_meeting:.1f}", help="Mevcut şartlar altında Baseline P80 teslim tarihine yetişme olasılığınız.")
            with c4: st.info("Gri çubuklar Sözleşme hedefini, renkli çubuklar aktif durumu gösterir.")
            
            st.markdown("### 📊 Kıyaslamalı Gantt Şeması")
            df_b_gantt = st.session_state.baseline_data['df_gantt']
            df_a_gantt = st.session_state.df_gantt
            b_dict = {row['ID']: {'start': row['Başlangıç'], 'end': row['Bitiş']} for _, row in df_b_gantt.iterrows()}
            fig_comp = go.Figure()
            y_labels = []
            
            for _, a_row in df_a_gantt.iloc[::-1].iterrows():
                tid, g_isim = a_row['ID'], a_row['Görev Gösterim']
                y_labels.append(g_isim)
                if tid in b_dict:
                    fig_comp.add_trace(go.Scatter(x=[b_dict[tid]['start'], b_dict[tid]['end']], y=[g_isim, g_isim], mode='lines', line=dict(color='lightgray', width=24), showlegend=False, hoverinfo='x+text', text=f"Hedef: {b_dict[tid]['start'].strftime('%d %b %Y')} - {b_dict[tid]['end'].strftime('%d %b %Y')}"))
                delay = (a_row['Bitiş'] - b_dict[tid]['end']).days if tid in b_dict else 0
                bar_color = "indianred" if delay > 0 else ("mediumseagreen" if delay < 0 else "steelblue")
                fig_comp.add_trace(go.Scatter(x=[a_row['Başlangıç'], a_row['Bitiş']], y=[g_isim, g_isim], mode='lines', line=dict(color=bar_color, width=10), showlegend=False, hoverinfo='x+text', text=f"Aktif: {a_row['Başlangıç'].strftime('%d %b %Y')} - {a_row['Bitiş'].strftime('%d %b %Y')} (Sapma: {delay} Gün)"))
            
            fig_comp.add_trace(go.Scatter(x=[None], y=[None], mode='lines', line=dict(color='lightgray', width=10), name='Baseline (Hedef)'))
            fig_comp.add_trace(go.Scatter(x=[None], y=[None], mode='lines', line=dict(color='indianred', width=10), name='Gecikmeli (Sapan)'))
            fig_comp.add_trace(go.Scatter(x=[None], y=[None], mode='lines', line=dict(color='steelblue', width=10), name='Hedefe Uygun'))
            fig_comp.update_layout(plot_bgcolor='white', height=max(500, len(y_labels)*35), margin=dict(l=10, r=20, t=40, b=10), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig_comp, use_container_width=True)

            st.markdown("### 📋 Varyans (Sapma) Tablosu")
            df_b_gantt_sub = df_b_gantt[['ID', 'Bitiş']].rename(columns={'Bitiş': 'Baseline_Bitis'})
            df_a_gantt_sub = df_a_gantt[['ID', 'Görev İsim', 'Bitiş']].rename(columns={'Bitiş': 'Aktif_Bitis'})
            df_var = pd.merge(df_a_gantt_sub, df_b_gantt_sub, on='ID', how='left')
            df_var['Gecikme (Gün)'] = (df_var['Aktif_Bitis'] - df_var['Baseline_Bitis']).dt.days
            df_var['Baseline_Bitis'] = df_var['Baseline_Bitis'].dt.strftime('%Y-%m-%d')
            df_var['Aktif_Bitis'] = df_var['Aktif_Bitis'].dt.strftime('%Y-%m-%d')
            st.dataframe(df_var[['ID', 'Görev İsim', 'Baseline_Bitis', 'Aktif_Bitis', 'Gecikme (Gün)']].style.map(lambda val: 'color: red; font-weight: bold' if val > 0 else ('color: green; font-weight: bold' if val < 0 else 'color: black'), subset=['Gecikme (Gün)']), use_container_width=True)
        else:
            st.warning("⚠️ Henüz bir Baseline (Hedef Çizgisi) atanmamış.")
