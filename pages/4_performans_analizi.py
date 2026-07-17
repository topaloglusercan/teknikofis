import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import json

# --- YARDIMCI FONKSİYONLAR ---
def format_tr(val):
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def parse_tr_number(val):
    if pd.isna(val) or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val = str(val).strip()
    if '.' in val and ',' in val:
        val = val.replace('.', '').replace(',', '.')
    elif ',' in val:
        val = val.replace(',', '.')
    try:
        return float(val)
    except:
        return 0.0

st.set_page_config(page_title="Performans Analizi", layout="wide")

st.title("Proje Performans Analizi")
st.markdown("Kazanılmış Takvim (ESA) ve Kazanılmış Değer (EVA) Metodolojileri Kıyaslaması")

# --- SESSION STATE (VERİ KALICILIĞI) ---
if 'bac' not in st.session_state:
    st.session_state.bac = 433444000.00
    st.session_state.sac = 19.0
    st.session_state.at = 10.0
    st.session_state.ev = 200121094.80
    st.session_state.ac = 200121094.80
    st.session_state.pv_df = pd.DataFrame({
        "Ay": list(range(1, 20)),
        "Aylik_PV_TL": [
            "2165869,79", "4305068,69", "10822998,73", "12993948,69", "24723936,69", 
            "29506630,82", "31675040,69", "32503987,44", "31629296,67", "30338494,65",
            "29396183,50", "28192778,73", "26010948,03", "26035083,42", "26014516,90",
            "23833025,19", "26001324,99", "21690882,36", "15603984,02"
        ]
    })

def load_data():
    if st.session_state.json_uploader is not None:
        try:
            data = json.load(st.session_state.json_uploader)
            st.session_state.bac = float(data["Parametreler"]["BAC"])
            st.session_state.sac = float(data["Parametreler"]["SAC"])
            st.session_state.at = float(data["Parametreler"]["AT"])
            st.session_state.ev = float(data["Parametreler"]["EV"])
            st.session_state.ac = float(data["Parametreler"]["AC"])
            loaded_df = pd.DataFrame(data["Planlanan_Degerler"])
            st.session_state.pv_df = loaded_df[["Ay", "Aylik_PV_TL"]]
        except Exception as e:
            st.sidebar.error("Dosya okuma hatası! Geçerli bir JSON yüklediğinizden emin olun.")

# --- SIDEBAR & INPUTS ---
st.sidebar.header("💾 Veri Kaydet ve Yükle")
st.sidebar.file_uploader("Önceki Çalışmayı Yükle (JSON)", type=["json"], key="json_uploader", on_change=load_data)

st.sidebar.markdown("---")
st.sidebar.header("Proje Parametreleri")

bac = st.sidebar.number_input("Toplam Bütçe (BAC - TL)", value=st.session_state.bac, step=100000.0, format="%.2f", key="bac")
sac = st.sidebar.number_input("Planlanan Süre (SAC / PD - Ay)", value=st.session_state.sac, step=1.0, key="sac")
at = st.sidebar.number_input("Gerçekleşen Zaman (AT - Ay)", value=st.session_state.at, step=1.0, key="at")
ev = st.sidebar.number_input("Kazanılmış Değer (EV - TL)", value=st.session_state.ev, step=100000.0, format="%.2f", key="ev")
ac = st.sidebar.number_input("Gerçekleşen Maliyet (AC - TL)", value=st.session_state.ac, step=100000.0, format="%.2f", key="ac")

st.sidebar.markdown("---")
st.sidebar.subheader("Aylık Planlanan Değerler (PV)")
st.sidebar.markdown("Yeni satır eklemek için tablonun en altını kullanabilirsiniz.")

edited_pv = st.sidebar.data_editor(
    st.session_state.pv_df, 
    num_rows="dynamic", 
    use_container_width=True,
    column_config={
        "Ay": st.column_config.NumberColumn("Ay", format="%d"),
        "Aylik_PV_TL": st.column_config.TextColumn("Aylık PV (TL)")
    }
)

# PV Hesaplamaları
edited_pv["Aylik_PV_Num"] = edited_pv["Aylik_PV_TL"].apply(parse_tr_number)
edited_pv = edited_pv.sort_values(by="Ay").reset_index(drop=True)
edited_pv["Kumulatif_PV"] = edited_pv["Aylik_PV_Num"].cumsum()
pv_values = [0.0] + edited_pv["Kumulatif_PV"].tolist()

# JSON Dışa Aktarma Butonu
export_data = {
    "Parametreler": {"BAC": bac, "SAC": sac, "AT": at, "EV": ev, "AC": ac},
    "Planlanan_Degerler": edited_pv.to_dict(orient="records")
}
json_string = json.dumps(export_data, indent=4, ensure_ascii=False)
st.sidebar.download_button("Çalışmayı Kaydet (JSON İndir)", file_name="proje_verileri.json", mime="application/json", data=json_string)

# --- TEMEL HESAPLAMALAR ---
if int(at) < len(pv_values):
    pv_at = pv_values[int(at)]
else:
    pv_at = pv_values[-1]
    
cv = ev - ac
sv = ev - pv_at
cpi = ev / ac if ac > 0 else 0
spi = ev / pv_at if pv_at > 0 else 0
eac_cost = bac / cpi if cpi > 0 else 0

par = bac / sac if sac > 0 else 0
es_lineer = ev / par if par > 0 else 0
tpi = es_lineer / at if at > 0 else 0
teac = sac / tpi if tpi > 0 else 0

c = 0
for i in range(len(pv_values)-1):
    if pv_values[i] <= ev < pv_values[i+1]:
        c = i
        break
    elif ev >= pv_values[-1]:
        c = len(pv_values) - 1
        break

if c < len(pv_values) - 1:
    pv_c = pv_values[c]
    pv_c1 = pv_values[c+1]
    i_val = (ev - pv_c) / (pv_c1 - pv_c) if (pv_c1 - pv_c) > 0 else 0
else:
    pv_c = pv_values[-1]
    pv_c1 = pv_values[-1]
    i_val = 0
    
es_scurve = c + i_val
spi_t = es_scurve / at if at > 0 else 0
ieac_t = sac / spi_t if spi_t > 0 else 0

# --- YÖNETİCİ ÖZETİ (MİNİMAL) ---
st.markdown("### Performans Özeti")
col_k1, col_k2, col_k3, col_k4 = st.columns(4)

col_k1.metric(label="Maliyet Endeksi (CPI)", value=f"{cpi:.2f}", delta=f"{cpi - 1.0:.2f}")
col_k2.metric(label="Maliyet Sapması (CV)", value=f"{format_tr(cv)} TL", delta="Negatif: Aşım" if cv < 0 else "Pozitif: Tasarruf", delta_color="off")
col_k3.metric(label="Zaman Endeksi (SPI-t)", value=f"{spi_t:.2f}", delta=f"{spi_t - 1.0:.2f}")
col_k4.metric(label="Zaman Sapması (ESA)", value=f"{format_tr(ieac_t - sac)} Ay", delta="Pozitif: Gecikme" if (ieac_t - sac) > 0 else "Negatif: Erken", delta_color="inverse")

st.markdown("---")

# --- TABS ---
tab1, tab2, tab3, tab4 = st.tabs(["ESA (Zaman Analizi)", "EVA (Maliyet Analizi)", "Teori ve Hesaplama Adımları", "Eğitim Merkezi: Uçak Analojisi"])

with tab1:
    st.markdown("#### Bitiş Süresi Projeksiyonları")
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Planlanan Bitiş (SAC)", f"{format_tr(sac)} Ay")
    col_m2.metric("Doğrusal Tahmin (TEAC)", f"{format_tr(teac)} Ay")
    col_m3.metric("S-Eğrisi Tahmini (IEAC-t)", f"{format_tr(ieac_t)} Ay")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**1. Doğrusal Yaklaşım Verileri**")
        st.write(f"- Planlanan Başarı Oranı (PAR): {format_tr(par)} TL/Ay")
        st.write(f"- Kazanılmış Takvim (ES): {format_tr(es_lineer)} Ay")
        st.write(f"- Zaman Performans Endeksi (TPI): {format_tr(tpi)}")
        
    with col2:
        st.markdown("**2. S-Eğrisi Yaklaşım Verileri**")
        st.write(f"- Tamamlanan Ay (C): {c}")
        st.write(f"- Kesirli Kısım (I): {format_tr(i_val)}")
        st.write(f"- Kazanılmış Takvim (ES): {format_tr(es_scurve)} Ay")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### İlerleme ve Projeksiyon Grafiği")
    
    fig1 = go.Figure()
    fig1.add_trace(go.Scatter(x=list(range(len(pv_values))), y=pv_values, mode='lines', name='PV (Planlanan)', line=dict(color='#2c3e50', width=2)))
    linear_pv = [par * i for i in range(len(pv_values))]
    fig1.add_trace(go.Scatter(x=list(range(len(pv_values))), y=linear_pv, mode='lines', name='Doğrusal Referans', line=dict(color='#95a5a6', width=1, dash='dash')))
    fig1.add_trace(go.Scatter(x=[at], y=[ev], mode='markers', name='Mevcut Durum (EV)', marker=dict(color='#e74c3c', size=10)))
    
    if teac > 0:
        fig1.add_trace(go.Scatter(x=[at, teac], y=[ev, bac], mode='lines', name='Doğrusal Tahmin', line=dict(color='#f39c12', width=1, dash='dot')))
    if ieac_t > 0:
        fig1.add_trace(go.Scatter(x=[at, ieac_t], y=[ev, bac], mode='lines', name='S-Eğrisi Tahmini', line=dict(color='#8e44ad', width=2, dash='dot')))

    max_x = max(len(pv_values)-1, teac, ieac_t)
    fig1.update_layout(
        template="simple_white",
        xaxis=dict(range=[0, max_x * 1.05], title="Zaman (Ay)"),
        yaxis=dict(title="Maliyet (TL)"),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    fig1.add_hline(y=bac, line_dash="dash", line_color="#7f8c8d", annotation_text="BAC", annotation_position="bottom right")

    st.plotly_chart(fig1, use_container_width=True)

with tab2:
    st.markdown("#### Maliyet ve Geleneksel Takvim Metrikleri")
    
    col3, col4 = st.columns(2)
    with col3:
        st.markdown("**Maliyet Parametreleri**")
        st.write(f"- Gerçekleşen Maliyet (AC): {format_tr(ac)} TL")
        st.write(f"- Maliyet Sapması (CV): {format_tr(cv)} TL")
        st.write(f"- Tamamlanma Tahmini Maliyeti (EAC): {format_tr(eac_cost)} TL")
        
    with col4:
        st.markdown("**Geleneksel Takvim Parametreleri (EVA)**")
        st.write(f"- Planlanan Değer (PV) - {int(at)}. Ay: {format_tr(pv_at)} TL")
        st.write(f"- Takvim Sapması (SV): {format_tr(sv)} TL")
        st.write(f"- Takvim Performans Endeksi (SPI): {format_tr(spi)}")
        st.caption("Not: Geleneksel SPI, proje bittiğinde her zaman 1.0 sonucunu verir.")
        
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### Durum Kıyaslaması (PV vs EV vs AC)")
    
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=["Planlanan (PV)", "Kazanılan (EV)", "Gerçekleşen (AC)"], y=[pv_at, ev, ac], marker_color=['#34495e', '#27ae60', '#c0392b'], text=[f"{format_tr(pv_at)}", f"{format_tr(ev)}", f"{format_tr(ac)}"], textposition='outside'))
    fig2.update_layout(template="simple_white", yaxis_title="Tutar (TL)")
    st.plotly_chart(fig2, use_container_width=True)

with tab3:
    st.markdown("#### 1. EVA (Kazanılmış Değer Analizi) Çözüm Adımları")
    st.markdown("Mevcut parametrelerle geleneksel maliyet ve zaman sapmalarının matematiksel çözümü:")
    
    st.latex(r"CV = EV - AC = " + format_tr(ev) + " - " + format_tr(ac) + " = " + format_tr(cv))
    st.latex(r"CPI = \frac{EV}{AC} = \frac{" + format_tr(ev) + "}{" + format_tr(ac) + "} = " + f"{cpi:.2f}")
    
    st.latex(r"SV = EV - PV = " + format_tr(ev) + " - " + format_tr(pv_at) + " = " + format_tr(sv))
    st.latex(r"SPI = \frac{EV}{PV} = \frac{" + format_tr(ev) + "}{" + format_tr(pv_at) + "} = " + f"{spi:.2f}")

    st.markdown("---")
    st.markdown("#### 2. ESA (Kazanılmış Takvim Analizi) Çözüm Adımları")
    st.markdown("**A. Doğrusal (Lineer) Yaklaşım** 📄*[^1]*")
    
    st.latex(r"PAR = \frac{BAC}{SAC} = \frac{" + format_tr(bac) + "}{" + format_tr(sac) + "} = " + format_tr(par))
    st.latex(r"ES = \frac{EV}{PAR} = \frac{" + format_tr(ev) + "}{" + format_tr(par) + "} = " + f"{es_lineer:.2f}")
    st.latex(r"TPI = \frac{ES}{AT} = \frac{" + f"{es_lineer:.2f}" + "}{" + f"{at:.2f}" + "} = " + f"{tpi:.2f}")
    st.latex(r"TEAC = \frac{SAC}{TPI} = \frac{" + format_tr(sac) + "}{" + f"{tpi:.2f}" + "} = " + f"{teac:.2f}")

    st.markdown("**B. S-Eğrisi Yaklaşımı** 📄*[^2]*")
    st.markdown(f"Kazanılmış değerimiz ({format_tr(ev)} TL), planlanan değerler tablosunda kümülatif olarak **{c}. aya** (PV_C = {format_tr(pv_c)} TL) karşılık gelmektedir. Ancak {c+1}. ayın hedefine (PV_C+1 = {format_tr(pv_c1)} TL) ulaşamamıştır. Bu nedenle kesirli kısım (I) hesaplanır:")
    
    if (pv_c1 - pv_c) > 0:
        st.latex(r"I = \frac{EV - PV_C}{PV_{C+1} - PV_C} = \frac{" + format_tr(ev) + " - " + format_tr(pv_c) + "}{" + format_tr(pv_c1) + " - " + format_tr(pv_c) + "} = " + f"{i_val:.2f}")
    else:
        st.latex(r"I = 0 \quad (\text{Bölen 0 olduğu için})")
        
    st.latex(r"ES = C + I = " + str(c) + " + " + f"{i_val:.2f}" + " = " + f"{es_scurve:.2f}")
    st.latex(r"SPI(t) = \frac{ES}{AT} = \frac{" + f"{es_scurve:.2f}" + "}{" + f"{at:.2f}" + "} = " + f"{spi_t:.2f}")
    st.latex(r"IEAC(t) = \frac{SAC}{SPI(t)} = \frac{" + format_tr(sac) + "}{" + f"{spi_t:.2f}" + "} = " + f"{ieac_t:.2f}")

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("### 📚 Kaynakça ve Akademik Referanslar")
    st.caption("[^1] **Anbari, F. T. (2003).** *Earned value project management method and extensions.* Project Management Journal, 34(4), 12-23. (Doğrusal/Lineer yaklaşım)
    st.caption("[^2] **Lipke, W. (2003).** *Schedule is different.* The Measurable News, 31(4), 31-34. (S-Eğrisi interpolasyonu)


with tab4:
    st.markdown("### ✈️ Uçak ve Yakıt Analojisi ile Performans Analizi")
    st.markdown("Kazanılmış Değer Analizi (EVA) ve Kazanılmış Takvim Analizini (ESA) anlamanın en kalıcı yolu **A noktasından B noktasına yapılan bir uçuş senaryosudur.**")
    
    st.info("**Uçuş Planı (Temel Kavramlar)** \n"
            "• **Planlanan Süre (SAC):** 10 Saat \n"
            "• **Toplam Mesafe (Proje Kapsamı):** 10.000 Km \n"
            "• **Planlanan Toplam Yakıt (Bütçe - BAC):** 100 Ton *(Her saat 10 ton yakıt harcanarak saatte 1.000 km yol katedilecek).*")

    # Eğitim için görsel grafik
    fig_edu = go.Figure()
    fig_edu.add_trace(go.Scatter(x=[0, 10], y=[0, 100], mode='lines', name='Planlanan Uçuş (PV)', line=dict(color='#2980b9', width=4, dash='dash')))
    fig_edu.add_trace(go.Scatter(x=[5], y=[40], mode='markers+text', name='Kazanılan Yol (EV)', marker=dict(color='#27ae60', size=16, symbol='star-triangle-up'), text=['Gerçekleşen Konum (EV: 40 Ton)'], textposition='bottom right'))
    fig_edu.add_trace(go.Scatter(x=[5], y=[60], mode='markers+text', name='Harcanan Yakıt (AC)', marker=dict(color='#c0392b', size=14, symbol='x'), text=['Depodan Eksilen (AC: 60 Ton)'], textposition='top left'))
    fig_edu.update_layout(title="A-B Arası Uçuş Senaryosu", xaxis_title="Uçuş Süresi (Saat)", yaxis_title="Kazanılan/Harcanan Yakıt Değeri (Ton)", template="simple_white", hovermode=False)
    st.plotly_chart(fig_edu, use_container_width=True)

    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.markdown('''
        **Uçak havalandı. Tam 5. saatin sonunda (AT = 5), 4.000 km yol alındığını ve 60 ton yakıt sarfiyatı yapıldığını varsayarak göstergelere bakıyoruz:**

        **1. Planlanan Değer (PV):** *Şu ana kadar nerede olmalıydık?*
        * 5 saat geçtiğine göre planımız yakıtın yarısını harcamaktı. Karşılığı: **50 Ton yakıt.** (PV = 50)

        **2. Kazanılmış Değer (EV):** *Gerçekte ne kadar yol geldik?*
        * Sadece 4.000 km yol gelebilmişiz. Yani 10.000 km'lik toplam hedefin %40'ı tamamlanmış. Bu ilerlemenin planımızdaki bütçe karşılığı nedir? (100 Ton x %40) = **40 Ton yakıt.** (EV = 40)

        **3. Gerçekleşen Maliyet (AC):** *Bu yolu gelirken depodan gerçekten ne kadar yaktık?*
        * Motorlar rüzgarla mücadele ederken çok zorlanmış. Göstergeye bakıyoruz, depodan **60 Ton yakıt** eksilmiş. (AC = 60)
        ''')
        
    with col_e2:
        st.markdown('''
        #### Geleneksel EVA'nın Kusuru
        EVA sisteminde takvim sapması (SV = EV - PV) hesaplandığında sonuç **-10 Ton** çıkar. Zamanın gerisinde kalmayı yakıt (para) cinsinden ölçmek bazı durumlarda yetersiz kalır.

        Eğer bu uçak hedefe 10 saatte değil de gecikerek 15 saatte varırsa; proje bittiği için 100 tonluk yol (10.000 km) tamamlanmış olur. Formül SV = 100 - 100 = 0 sonucunu verir. Yani proje 5 saat gecikmiş olsa bile, bittiği gün EVA metrikleri projeyi "tam zamanında bitmiş" gibi gösterir.

        #### Kazanılmış Takvim (ESA) ile Çözüm *[1]*
        ESA, zamanı para ile değil, gerçek zaman birimiyle ölçer.
        * **Kazanılmış Takvim (ES):** *"Bulunduğumuz yere plana göre ne kadar sürede gelmeliydik?"* 40 tonluk (4.000 km) yolu **4 saatte** gelmeliydik.
        * **Zaman Sapması:** ES - AT = 4 - 5 = **-1 Saat.** (Uçuşta tam 1 saat gerideyiz). 
        ''')

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### 📚 Analoji Kaynakçası")
    st.caption("*[1] Bu analoji, inşaat yönetimi literatüründe maliyet/zaman kavramlarını ayırmak için kullanılan klasik **Walt Lipke (2003)** metodolojisinin basitleştirilmiş ve ölçeklendirilmiş bir versiyonudur.*")
