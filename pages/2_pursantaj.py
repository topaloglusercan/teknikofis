import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# --- 1. ARAYÜZ VE TEMA AYARLARI ---
# DİKKAT: st.set_page_config() SATIRI BURADAN SİLİNDİ! SADECE AŞAĞIDAKİ MARKDOWN KALACAK.

st.markdown("""
<style>
    .info-box { background-color: #0F2044; color: white; padding: 18px; border-radius: 6px; margin-bottom: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .info-label { font-size: 0.75rem; color: #94a3b8; text-transform: uppercase; letter-spacing: 1.2px; font-weight: 600;}
    .info-val { font-size: 1.5rem; font-weight: 700; margin-top: 4px; color: #f8fafc; }
    .filter-container { background-color: #f8fafc; padding: 15px; border-radius: 6px; border: 1px solid #e2e8f0; margin-bottom: 15px; }
    
    .selected-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border-left: 4px solid #3b82f6;
        border-radius: 8px;
        padding: 15px 20px;
        color: white;
        margin-bottom: 15px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .sc-title { font-size: 0.85rem; color: #94a3b8; font-weight: 600; text-transform: uppercase; margin-bottom: 5px; }
    .sc-name { font-size: 1.3rem; font-weight: 700; color: #f8fafc; margin-bottom: 2px; }
    .sc-wbs { font-size: 0.9rem; color: #60a5fa; font-family: monospace; }
    .sc-metrics { display: flex; justify-content: space-between; margin-top: 12px; padding-top: 12px; border-top: 1px solid rgba(255,255,255,0.1); }
    .sc-metric-box { text-align: left; }
    .sc-metric-lbl { font-size: 0.75rem; color: #94a3b8; }
    .sc-metric-val { font-size: 1.1rem; font-weight: 600; color: #10b981; }
</style>
""", unsafe_allow_html=True)

if "bedel" not in st.session_state: st.session_state.bedel = 1_000_000.0
if "kesilen_wbs_listesi" not in st.session_state: st.session_state.kesilen_wbs_listesi = []

# --- 2. VERİ MODELİ VE ALFANÜMERİK SIRALAMA ---
def wbs_sirala(df_to_sort):
    def sort_key(wbs_str):
        parts = str(wbs_str).strip().split('.')
        key_parts = []
        for p in parts:
            if p.isdigit():
                key_parts.append((0, int(p)))
            else:
                match = re.match(r"([a-zA-Z]+)(\d+)", p)
                if match:
                    key_parts.append((1, match.group(1), int(match.group(2))))
                else:
                    key_parts.append((2, p))
        return tuple(key_parts)
        
    df_to_sort["_sort_key"] = df_to_sort["WBS"].apply(sort_key)
    df_to_sort = df_to_sort.sort_values(by="_sort_key").drop(columns=["_sort_key"])
    return df_to_sort.reset_index(drop=True)

if "df" not in st.session_state:
    data = [
        {"WBS": "A", "Ad": "A Blok", "Yerel_%": 100.0, "Kilit": True, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"},
        {"WBS": "A.1", "Ad": "Üst Yapı", "Yerel_%": 60.0, "Kilit": True, "İlerleme_%": 0.0, "Grup 1": "İnşaat", "Grup 2": "Ana İşveren", "Grup 3": "Tüm Katlar"},
        {"WBS": "A.1.1", "Ad": "Kaba İnşaat", "Yerel_%": 40.0, "Kilit": False, "İlerleme_%": 25.0, "Grup 1": "Kaba", "Grup 2": "Taşeron A", "Grup 3": "Zemin Kat"},
        {"WBS": "A.1.2", "Ad": "İnce İşler", "Yerel_%": 60.0, "Kilit": False, "İlerleme_%": 10.0, "Grup 1": "İnce", "Grup 2": "Taşeron B", "Grup 3": "Normal Katlar"},
        {"WBS": "A.2", "Ad": "Alt Yapı", "Yerel_%": 40.0, "Kilit": False, "İlerleme_%": 0.0, "Grup 1": "Altyapı", "Grup 2": "Taşeron C", "Grup 3": "Zemin Altı"},
    ]
    st.session_state.df = wbs_sirala(pd.DataFrame(data))

if "secili_wbs_listesi" not in st.session_state:
    st.session_state.secili_wbs_listesi = ["A"]

# --- 3. ALFANÜMERİK WBS İŞLEMLERİ ---
def parent_bul(wbs):
    return str(wbs).rsplit(".", 1)[0] if "." in str(wbs) else "KÖK"

def sonraki_wbs_uret(parent_wbs, current_df):
    if parent_wbs == "KÖK":
        sayilar = [int(w) for w in current_df["WBS"].astype(str) if "." not in w and w.isdigit()]
        if sayilar: return str(max(sayilar) + 1)
        return "YENİ_KÖK" 
    else:
        prefix = str(parent_wbs) + "."
        cocuklar = [w for w in current_df["WBS"].astype(str) if w.startswith(prefix) and len(w.split(".")) == len(str(parent_wbs).split(".")) + 1]
        
        max_num = 0
        for cocuk in cocuklar:
            son_kisim = cocuk.split(".")[-1]
            if son_kisim.isdigit():
                max_num = max(max_num, int(son_kisim))
            else:
                match = re.search(r'\d+$', son_kisim)
                if match: max_num = max(max_num, int(match.group()))
                
        return f"{prefix}{max_num + 1}"

def wbs_degistir(eski_kok, yeni_kok, dataframe):
    mask = (dataframe["WBS"] == eski_kok) | (dataframe["WBS"].astype(str).str.startswith(eski_kok + "."))
    dataframe.loc[mask, "WBS"] = dataframe.loc[mask, "WBS"].apply(lambda x: str(yeni_kok) + str(x)[len(str(eski_kok)):])
    return dataframe

def wbs_hedefe_tasi_coklu(kaynak_wbs_listesi, hedef_parent_wbs):
    df_temp = st.session_state.df.copy()
    gecerli_kaynaklar = []
    
    kokler = df_temp[~df_temp["WBS"].astype(str).str.contains(r'\.')]["WBS"].tolist()

    for k_wbs in kaynak_wbs_listesi:
        if k_wbs in kokler: continue 
        if k_wbs == hedef_parent_wbs: continue
        if str(hedef_parent_wbs).startswith(str(k_wbs) + "."): continue
        gecerli_kaynaklar.append(k_wbs)
        
    if not gecerli_kaynaklar: return df_temp

    maskeler = [ (df_temp["WBS"] == k) | (df_temp["WBS"].astype(str).str.startswith(k + ".")) for k in gecerli_kaynaklar ]
    tam_maske = pd.concat(maskeler, axis=1).any(axis=1) if maskeler else pd.Series([False]*len(df_temp))
    df_temp.loc[tam_maske, "Kilit"] = False

    for k_wbs in gecerli_kaynaklar:
        yeni_wbs = sonraki_wbs_uret(hedef_parent_wbs, df_temp)
        df_temp = wbs_degistir(k_wbs, yeni_wbs, df_temp)
    return df_temp

def wbs_girinti_artir(kaynak_wbs):
    parent = parent_bul(kaynak_wbs)
    if parent == "KÖK": return 
    kardesler = st.session_state.df[st.session_state.df["WBS"].apply(parent_bul) == parent]["WBS"].tolist()
    kardesler_df = pd.DataFrame({"WBS": kardesler})
    kardesler = wbs_sirala(kardesler_df)["WBS"].tolist()
    
    if kaynak_wbs not in kardesler: return
    idx = kardesler.index(kaynak_wbs)
    if idx > 0: 
        onceki_kardes = kardesler[idx - 1]
        yeni_wbs = sonraki_wbs_uret(onceki_kardes, st.session_state.df)
        st.session_state.df = wbs_degistir(kaynak_wbs, yeni_wbs, st.session_state.df)
        st.session_state.df = wbs_sirala(st.session_state.df)
        st.session_state.secili_wbs_listesi = [yeni_wbs]

def wbs_girinti_azalt(kaynak_wbs):
    if not "." in str(kaynak_wbs): return 
    parent = parent_bul(kaynak_wbs)
    if parent == "KÖK": return 
    buyuk_baba = parent_bul(parent)
    if buyuk_baba == "KÖK": return 
    
    yeni_wbs = sonraki_wbs_uret(buyuk_baba, st.session_state.df)
    st.session_state.df = wbs_degistir(kaynak_wbs, yeni_wbs, st.session_state.df)
    st.session_state.df = wbs_sirala(st.session_state.df)
    st.session_state.secili_wbs_listesi = [yeni_wbs]

def wbs_tasi(kaynak_wbs, yon):
    parent = parent_bul(kaynak_wbs)
    kardesler = st.session_state.df[st.session_state.df["WBS"].apply(parent_bul) == parent]["WBS"].tolist()
    kardesler_df = pd.DataFrame({"WBS": kardesler})
    kardesler = wbs_sirala(kardesler_df)["WBS"].tolist()
    
    if kaynak_wbs not in kardesler: return
    idx = kardesler.index(kaynak_wbs)
    hedef_wbs = None
    if yon == "yukari" and idx > 0: hedef_wbs = kardesler[idx - 1]
    elif yon == "asagi" and idx < len(kardesler) - 1: hedef_wbs = kardesler[idx + 1]
    
    if hedef_wbs:
        temp_wbs = "TEMP_" + kaynak_wbs
        df_temp = wbs_degistir(kaynak_wbs, temp_wbs, st.session_state.df.copy())
        df_temp = wbs_degistir(hedef_wbs, kaynak_wbs, df_temp)
        df_temp = wbs_degistir(temp_wbs, hedef_wbs, df_temp)
        st.session_state.df = wbs_sirala(df_temp)
        st.session_state.secili_wbs_listesi = [hedef_wbs] 

def wbs_agaci_kopyala(kaynak_wbs, kopya_sayisi):
    mask = (st.session_state.df["WBS"] == kaynak_wbs) | (st.session_state.df["WBS"].astype(str).str.startswith(kaynak_wbs + "."))
    kaynak_agac = st.session_state.df[mask].copy()
    if kaynak_agac.empty: return
    parent = parent_bul(kaynak_wbs)
    
    for i in range(kopya_sayisi):
        yeni_kok_wbs = sonraki_wbs_uret(parent, st.session_state.df)
        kopya_agac_i = kaynak_agac.copy()
        kopya_agac_i["WBS"] = kopya_agac_i["WBS"].apply(lambda x: str(yeni_kok_wbs) + str(x)[len(str(kaynak_wbs)):])
        kopya_agac_i.loc[kopya_agac_i["WBS"] == yeni_kok_wbs, "Ad"] += f" (Kopya {i+1})"
        kopya_agac_i["İlerleme_%"] = 0.0
        kopya_agac_i["Kilit"] = False
        st.session_state.df = pd.concat([st.session_state.df, kopya_agac_i], ignore_index=True)
        
    st.session_state.df = wbs_sirala(st.session_state.df)
    st.session_state.df = akilli_dengele(st.session_state.df, st.session_state.df)

def akilli_dengele(old_df, new_df):
    temp_df = new_df.copy()
    temp_df["Parent"] = temp_df["WBS"].astype(str).apply(parent_bul)
    parents = temp_df["Parent"].unique()
    
    for p in parents:
        mask = temp_df["Parent"] == p
        if len(old_df) != len(temp_df): continue
            
        degisen_satirlar = mask & ~np.isclose(old_df["Yerel_%"].astype(float), temp_df["Yerel_%"].astype(float), atol=0.001)
        if degisen_satirlar.sum() > 0:
            temp_df.loc[degisen_satirlar, "Kilit"] = True
            
        locked_mask = mask & temp_df["Kilit"]
        unlocked_mask = mask & ~temp_df["Kilit"]
        locked_sum = temp_df.loc[locked_mask, "Yerel_%"].sum()
        
        if unlocked_mask.sum() == 0 and not np.isclose(locked_sum, 100.0, atol=0.01):
            acilacak_satirlar = mask & ~degisen_satirlar
            if acilacak_satirlar.sum() > 0:
                temp_df.loc[acilacak_satirlar, "Kilit"] = False
                locked_mask = mask & temp_df["Kilit"]
                unlocked_mask = mask & ~temp_df["Kilit"]
                locked_sum = temp_df.loc[locked_mask, "Yerel_%"].sum()
                
        if unlocked_mask.sum() > 0:
            remaining = max(0.0, 100.0 - locked_sum)
            unlocked_current_sum = temp_df.loc[unlocked_mask, "Yerel_%"].sum()
            if unlocked_current_sum > 0:
                temp_df.loc[unlocked_mask, "Yerel_%"] = temp_df.loc[unlocked_mask, "Yerel_%"] * (remaining / unlocked_current_sum)
            else:
                temp_df.loc[unlocked_mask, "Yerel_%"] = remaining / unlocked_mask.sum()
                
    return temp_df.drop(columns=["Parent"])

# --- 4. HESAPLAMALAR VE GÖRSEL HAZIRLIK ---
df = st.session_state.df.copy()
df["Derinlik"] = df["WBS"].astype(str).apply(lambda x: len(x.split(".")) - 1)
df["Parent"] = df["WBS"].apply(parent_bul)
df["Is_Leaf"] = ~df["WBS"].isin(df["Parent"])

def gorsel_isim_olustur(row):
    derinlik = row["Derinlik"]
    if derinlik == 0: icon = "🏗️"
    elif derinlik == 1: icon = "🏢"
    elif derinlik == 2: icon = "🗂️"
    elif derinlik == 3: icon = "📂"
    elif derinlik == 4: icon = "📑"
    else: icon = "📄"
    prefix = " " * derinlik + ("└─ " if derinlik > 0 else "")
    return f"{prefix}{icon} {row['Ad']}"

df["Ağaç Görünümü"] = df.apply(gorsel_isim_olustur, axis=1)

df["Genel_%"] = 0.0
df.loc[df["Derinlik"] == 0, "Genel_%"] = df.loc[df["Derinlik"] == 0, "Yerel_%"]

for d in range(1, df["Derinlik"].max() + 1):
    for idx, row in df[df["Derinlik"] == d].iterrows():
        parent_genel = df.loc[df["WBS"] == row["Parent"], "Genel_%"].values
        if len(parent_genel) > 0:
            df.loc[idx, "Genel_%"] = (row["Yerel_%"] / 100.0) * parent_genel[0]
        else:
            df.loc[idx, "Genel_%"] = row["Yerel_%"] # Ebeveyni kayıpsa kök gibi davran

df["Bütçe (TL)"] = (df["Genel_%"] / 100.0) * st.session_state.bedel

df["Hakediş (TL)"] = 0.0
df.loc[df["Is_Leaf"], "Hakediş (TL)"] = df.loc[df["Is_Leaf"], "Bütçe (TL)"] * (df.loc[df["Is_Leaf"], "İlerleme_%"] / 100.0)

for d in range(df["Derinlik"].max() - 1, -1, -1):
    for idx, row in df[df["Derinlik"] == d].iterrows():
        child_mask = df["WBS"].str.startswith(row["WBS"] + ".") & df["Is_Leaf"]
        child_hakedis_sum = df.loc[child_mask, "Hakediş (TL)"].sum()
        if child_mask.sum() > 0:
            df.loc[idx, "Hakediş (TL)"] = child_hakedis_sum
            if df.loc[idx, "Bütçe (TL)"] > 0:
                calculated_progress = (child_hakedis_sum / df.loc[idx, "Bütçe (TL)"]) * 100.0
                df.loc[idx, "İlerleme_%"] = calculated_progress
                st.session_state.df.loc[idx, "İlerleme_%"] = calculated_progress

# --- 5. EXCEL İNDİRME FONKSİYONLARI ---
def excel_sablon_olustur():
    df_sablon = pd.DataFrame({"WBS": ["A", "A.1", "A.1.1", "A.1.2", "A.2"], "Ad": ["A Blok", "Üst Yapı", "Kaba İnşaat", "İnce İşler", "Alt Yapı"]})
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_sablon.to_excel(writer, index=False, sheet_name="WBS_Aktarim")
    return output.getvalue()

def excel_olustur(veri_df, proje_bedeli):
    wb = Workbook()
    ws = wb.active
    ws.title = "Pursantaj_İlerleme"
    ws["L1"] = "Sözleşme Bedeli:"
    ws["L1"].font = Font(bold=True)
    ws["M1"] = proje_bedeli
    ws["M1"].number_format = '#,##0.00 "₺"'
    
    headers = ["WBS", "Ağaç Görünümü", "Kalem Adı", "Yerel Pay (%)", "Genel Pay (%)", "İlerleme (%)", "Bütçe (TL)", "Hakediş (TL)", "Grup 1", "Grup 2", "Grup 3"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0F2044", end_color="0F2044", fill_type="solid")
        
    wbs_to_row = {wbs: idx + 3 for idx, wbs in enumerate(veri_df["WBS"])}
    
    for idx, row in veri_df.iterrows():
        r = idx + 3
        ws.cell(row=r, column=1, value=row["WBS"])
        ws.cell(row=r, column=2, value=row["Ağaç Görünümü"])
        ws.cell(row=r, column=3, value=row["Ad"])
        ws.cell(row=r, column=4, value=row["Yerel_%"] / 100.0).number_format = '0.00%'
        
        # Eğer öğenin parent'i Excel tablosunda yoksa veya kendisi ana kökse
        if row["Derinlik"] == 0 or row["Parent"] not in wbs_to_row: 
            ws.cell(row=r, column=5, value=f"=D{r}").number_format = '0.00%'
        else:
            p_row = wbs_to_row[row["Parent"]]
            ws.cell(row=r, column=5, value=f"=D{r}*E{p_row}").number_format = '0.00%'
            
        if row["Is_Leaf"]: 
            ws.cell(row=r, column=6, value=row["İlerleme_%"] / 100.0).number_format = '0.00%'
        else: 
            ws.cell(row=r, column=6, value=f"=IF(G{r}>0, H{r}/G{r}, 0)").number_format = '0.00%'
            
        ws.cell(row=r, column=7, value=f"=$M$1*E{r}").number_format = '#,##0.00 "₺"'
        
        if row["Is_Leaf"]: 
            ws.cell(row=r, column=8, value=f"=G{r}*F{r}").number_format = '#,##0.00 "₺"'
        else:
            children = veri_df[veri_df["Parent"] == row["WBS"]]["WBS"].tolist()
            valid_children = [c for c in children if c in wbs_to_row]
            if valid_children:
                child_rows = [str(wbs_to_row[c]) for c in valid_children]
                ws.cell(row=r, column=8, value="=" + "+".join([f"H{cr}" for cr in child_rows])).number_format = '#,##0.00 "₺"'
            else: 
                ws.cell(row=r, column=8, value=0).number_format = '#,##0.00 "₺"'
            
        ws.cell(row=r, column=9, value=row["Grup 1"])
        ws.cell(row=r, column=10, value=row["Grup 2"])
        ws.cell(row=r, column=11, value=row["Grup 3"])
        
    for col in ws.columns:
        max_length = 0
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 6. ANA EKRAN YERLEŞİMİ ---
st.title("🏗️ Profesyonel Pursantaj Yönetimi")

with st.expander("📂 Dosya İşlemleri (Excel Yükle / Rapor İndir)", expanded=False):
    t_c1, t_c2, t_c3 = st.columns([1, 1, 1.5])
    with t_c1:
        yeni_bedel = st.number_input("Sözleşme Bedeli (₺)", min_value=0.0, value=st.session_state.bedel, step=100000.0, format="%.2f")
        if yeni_bedel != st.session_state.bedel:
            st.session_state.bedel = yeni_bedel
            st.rerun()
        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
        st.download_button("📊 Raporu İndir (Formüllü Excel)", data=excel_olustur(df, st.session_state.bedel), file_name="pursantaj_raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
    with t_c2:
        st.download_button("📥 Excel Yükleme Şablonunu İndir", data=excel_sablon_olustur(), file_name="wbs_sablon.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.download_button("📥 Projeyi Yedekle (.json)", data=st.session_state.df.to_json(orient="records", force_ascii=False), file_name="pursantaj.json", mime="application/json", use_container_width=True)
        
    with t_c3:
        uploaded_file = st.file_uploader("📤 Excel'den Proje Yükle (WBS ve Ad)", type=["xlsx", "xls"])
        if "son_yuklenen_dosya_id" not in st.session_state: st.session_state.son_yuklenen_dosya_id = None
        if uploaded_file is not None:
            if st.session_state.son_yuklenen_dosya_id != uploaded_file.file_id:
                try:
                    imp_df = pd.read_excel(uploaded_file)
                    if "WBS" in imp_df.columns and "Ad" in imp_df.columns:
                        imp_df["WBS"] = imp_df["WBS"].astype(str).str.strip()
                        if "Yerel_%" not in imp_df.columns: imp_df["Yerel_%"] = 0.0
                        if "Kilit" not in imp_df.columns: imp_df["Kilit"] = False
                        if "İlerleme_%" not in imp_df.columns: imp_df["İlerleme_%"] = 0.0
                        if "Grup 1" not in imp_df.columns: imp_df["Grup 1"] = "-"
                        if "Grup 2" not in imp_df.columns: imp_df["Grup 2"] = "-"
                        if "Grup 3" not in imp_df.columns: imp_df["Grup 3"] = "-"
                        st.session_state.df = wbs_sirala(imp_df[["WBS", "Ad", "Yerel_%", "Kilit", "İlerleme_%", "Grup 1", "Grup 2", "Grup 3"]])
                        st.session_state.son_yuklenen_dosya_id = uploaded_file.file_id
                        st.session_state.secili_wbs_listesi = [st.session_state.df.iloc[0]["WBS"]]
                        st.rerun()
                    else: st.error("❌ Excel'de 'WBS' ve 'Ad' sütunları bulunamadı.")
                except Exception as e: st.error(f"❌ Dosya okuma hatası: {e}")
            else: st.success("✅ WBS Excel'den başarıyla yüklendi!")

st.divider()
tab1, tab2 = st.tabs(["📋 Planlama & Master Tablo", "📈 Pivot Rapor Analizi"])

# =========================================================================
# SEKME 1: PLANLAMA VE MASTER TABLO
# =========================================================================
with tab1:
    col_sol, col_sag = st.columns([3.5, 1.2])

    with col_sol:
        st.markdown("<div class='filter-container'>", unsafe_allow_html=True)
        f_c1, f_c2, f_c3, f_c4, f_c5 = st.columns([1.5, 1, 1, 1, 1.5])
        with f_c1: filtre_ad = st.text_input("Arama (Kalem Adı/WBS):")
        with f_c2: filtre_g1 = st.selectbox("Grup 1:", ["Tümü"] + sorted(list(df["Grup 1"].dropna().unique())))
        with f_c3: filtre_g2 = st.selectbox("Grup 2:", ["Tümü"] + sorted(list(df["Grup 2"].dropna().unique())))
        with f_c4: filtre_g3 = st.selectbox("Grup 3:", ["Tümü"] + sorted(list(df["Grup 3"].dropna().unique())))
        with f_c5:
            mevcut_seviyeler = sorted(df["Derinlik"].unique())
            secilen_seviyeler = st.multiselect("Gösterilecek Seviyeler:", options=mevcut_seviyeler, default=mevcut_seviyeler, format_func=lambda x: f"Seviye {x}")
        st.markdown("</div>", unsafe_allow_html=True)

        gosterilecek_df = df.copy()
        if filtre_ad: gosterilecek_df = gosterilecek_df[gosterilecek_df["Ad"].str.contains(filtre_ad, case=False) | gosterilecek_df["WBS"].astype(str).str.contains(filtre_ad)]
        if filtre_g1 != "Tümü": gosterilecek_df = gosterilecek_df[gosterilecek_df["Grup 1"] == filtre_g1]
        if filtre_g2 != "Tümü": gosterilecek_df = gosterilecek_df[gosterilecek_df["Grup 2"] == filtre_g2]
        if filtre_g3 != "Tümü": gosterilecek_df = gosterilecek_df[gosterilecek_df["Grup 3"] == filtre_g3]
        if secilen_seviyeler: gosterilecek_df = gosterilecek_df[gosterilecek_df["Derinlik"].isin(secilen_seviyeler)]

        edit_cols = ["WBS", "Ağaç Görünümü", "Ad", "Kilit", "Yerel_%", "İlerleme_%", "Bütçe (TL)", "Hakediş (TL)", "Grup 1", "Grup 2", "Grup 3"]
        
        edited_df = st.data_editor(
            gosterilecek_df[edit_cols],
            use_container_width=True, hide_index=True, num_rows="fixed", height=600,
            column_config={
                "WBS": st.column_config.TextColumn("✏️ WBS (Değiştir)", width="small"),
                "Ağaç Görünümü": st.column_config.TextColumn("Hiyerarşi (Görsel)", width="medium", disabled=True),
                "Ad": st.column_config.TextColumn("✏️ Kalem Adı"),
                "Kilit": st.column_config.CheckboxColumn("🔒 Kilit", width="small"),
                "Yerel_%": st.column_config.NumberColumn("Ağırlık (%)", format="%.2f %%", min_value=0.0, max_value=100.0, width="small"),
                "İlerleme_%": st.column_config.NumberColumn("✏️ İlerleme (%)", format="%.2f %%", min_value=0.0, max_value=100.0, width="small"),
                "Bütçe (TL)": st.column_config.NumberColumn("Bütçe", format="%,.0f ₺", disabled=True, width="small"),
                "Hakediş (TL)": st.column_config.NumberColumn("Hakediş", format="%,.0f ₺", disabled=True, width="small"),
                "Grup 1": st.column_config.TextColumn("Grup 1", width="small"),
                "Grup 2": st.column_config.TextColumn("Grup 2", width="small"),
                "Grup 3": st.column_config.TextColumn("Grup 3", width="small"),
            },
            key="master_editor"
        )
        
        if not edited_df.equals(gosterilecek_df[edit_cols]):
            wbs_degisti = False
            for i in range(len(gosterilecek_df)):
                eski_w = str(gosterilecek_df.iloc[i]["WBS"]).strip()
                yeni_w = str(edited_df.iloc[i]["WBS"]).strip()
                
                if eski_w != yeni_w:
                    if yeni_w in st.session_state.df["WBS"].values:
                        st.error(f"❌ Hata: '{yeni_w}' numaralı WBS kodu zaten kullanılıyor!")
                        st.stop()
                    elif not yeni_w:
                        st.error("❌ Hata: WBS kodu boş olamaz!")
                        st.stop()
                    else:
                        st.session_state.df = wbs_degistir(eski_w, yeni_w, st.session_state.df)
                        st.session_state.df = wbs_sirala(st.session_state.df)
                        if st.session_state.secili_wbs_listesi == [eski_w]:
                            st.session_state.secili_wbs_listesi = [yeni_w]
                        elif eski_w in st.session_state.secili_wbs_listesi:
                            st.session_state.secili_wbs_listesi = [yeni_w if x==eski_w else x for x in st.session_state.secili_wbs_listesi]
                        wbs_degisti = True
                        break 
                        
            if wbs_degisti: st.rerun()

            yeni_ana_df = st.session_state.df.copy()
            for idx, row in edited_df.iterrows():
                for col in ["Ad", "Yerel_%", "Kilit", "İlerleme_%", "Grup 1", "Grup 2", "Grup 3"]:
                    yeni_ana_df.loc[yeni_ana_df["WBS"] == row["WBS"], col] = row[col]
            st.session_state.df = akilli_dengele(st.session_state.df, yeni_ana_df)
            st.rerun()

    with col_sag:
        st.markdown("### 🛠️ İşlem Paneli")
        st.caption("Tablodaki kutucuklarla çoklu seçim yapabilirsiniz.")
        
        event = st.dataframe(
            gosterilecek_df[["WBS", "Ad"]],
            on_select="rerun",
            selection_mode="multi-row",
            hide_index=True,
            height=150,
            column_config={"WBS": st.column_config.TextColumn(width="small")}
        )
        
        if event.selection and event.selection.rows:
            st.session_state.secili_wbs_listesi = gosterilecek_df.iloc[event.selection.rows]["WBS"].tolist()
        else:
            if len(st.session_state.secili_wbs_listesi) == 0: 
                st.session_state.secili_wbs_listesi = [gosterilecek_df.iloc[0]["WBS"]]
            
        secili_wbsler = st.session_state.secili_wbs_listesi
        
        if len(secili_wbsler) == 1:
            ilk_secili_wbs = secili_wbsler[0]
            secili_row = df[df["WBS"] == ilk_secili_wbs].iloc[0]
            
            st.markdown(f"""
            <div class="selected-card">
                <div class="sc-title">SEÇİLİ ÖĞE</div>
                <div class="sc-name">{secili_row['Ad']}</div>
                <div class="sc-wbs">WBS: {secili_row['WBS']}</div>
                <div class="sc-metrics">
                    <div class="sc-metric-box">
                        <div class="sc-metric-lbl">BÜTÇE</div>
                        <div class="sc-metric-val">{secili_row['Bütçe (TL)']:,.0f} ₺</div>
                    </div>
                    <div class="sc-metric-box" style="text-align: right;">
                        <div class="sc-metric-lbl">GENEL PAY</div>
                        <div class="sc-metric-val" style="color: #60a5fa;">%{secili_row['Genel_%']:.2f}</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
        else:
            toplam_butce = df[df["WBS"].isin(secili_wbsler)]["Bütçe (TL)"].sum()
            toplam_pay = df[df["WBS"].isin(secili_wbsler)]["Genel_%"].sum()
            
            st.markdown(f"""
            <div class="selected-card" style="border-left-color: #f59e0b;">
                <div class="sc-title">ÇOKLU SEÇİM</div>
                <div class="sc-name">{len(secili_wbsler)} Öğe Seçildi</div>
                <div class="sc-metrics">
                    <div class="sc-metric-box">
                        <div class="sc-metric-lbl">TOPLAM BÜTÇE</div>
                        <div class="sc-metric-val" style="color: #fcd34d;">{toplam_butce:,.0f} ₺</div>
                    </div>
                    <div class="sc-metric-box" style="text-align: right;">
                        <div class="sc-metric-lbl">TOPLAM PAY</div>
                        <div class="sc-metric-val" style="color: #fcd34d;">%{toplam_pay:.2f}</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
        ilk_secili_wbs = secili_wbsler[0] 
        
        st.markdown("#### Hiyerarşi & Sıralama")
        s1, s2 = st.columns(2)
        with s1:
            if st.button("⬆️ Yukarı", use_container_width=True, disabled=len(secili_wbsler)>1): wbs_tasi(ilk_secili_wbs, "yukari"); st.rerun()
        with s2:
            if st.button("⬇️ Aşağı", use_container_width=True, disabled=len(secili_wbsler)>1): wbs_tasi(ilk_secili_wbs, "asagi"); st.rerun()
        i1, i2 = st.columns(2)
        with i1:
            if st.button("➡️ İçeri (Girinti)", use_container_width=True, disabled=len(secili_wbsler)>1): wbs_girinti_artir(ilk_secili_wbs); st.rerun()
        with i2:
            if st.button("⬅️ Dışarı", use_container_width=True, disabled=len(secili_wbsler)>1): wbs_girinti_azalt(ilk_secili_wbs); st.rerun()

        st.markdown("#### Ekle / Sil / Kopyala")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("➕ Alt Kırılım", use_container_width=True, disabled=len(secili_wbsler)>1):
                yeni_wbs = sonraki_wbs_uret(ilk_secili_wbs, st.session_state.df)
                st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([{"WBS": yeni_wbs, "Ad": "Yeni Kırılım", "Yerel_%": 0.0, "Kilit": False, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"}])], ignore_index=True)
                st.session_state.df = wbs_sirala(st.session_state.df)
                st.session_state.secili_wbs_listesi = [yeni_wbs]
                st.session_state.df = akilli_dengele(st.session_state.df, st.session_state.df); st.rerun()
        with c2:
            if st.button("↔️ Kardeş Ekle", use_container_width=True, disabled=len(secili_wbsler)>1):
                yeni_wbs = sonraki_wbs_uret(parent_bul(ilk_secili_wbs), st.session_state.df)
                st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([{"WBS": yeni_wbs, "Ad": "Yeni Kardeş", "Yerel_%": 0.0, "Kilit": False, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"}])], ignore_index=True)
                st.session_state.df = wbs_sirala(st.session_state.df)
                st.session_state.secili_wbs_listesi = [yeni_wbs]
                st.session_state.df = akilli_dengele(st.session_state.df, st.session_state.df); st.rerun()

        c_kopya1, c_kopya2 = st.columns([1, 1])
        with c_kopya1: kopya_adet = st.number_input("Adet", min_value=1, max_value=50, value=1)
        with c_kopya2:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("📋 Kopyala", use_container_width=True, disabled=len(secili_wbsler)>1): wbs_agaci_kopyala(ilk_secili_wbs, int(kopya_adet)); st.rerun()

        kokler = df[~df["WBS"].astype(str).str.contains(r'\.')]["WBS"].tolist()
        if not any(k in secili_wbsler for k in kokler):
            if st.button("🗑️ Seçili Kalemleri Sil", use_container_width=True, type="primary"):
                maskeler = [st.session_state.df["WBS"].astype(str).str.startswith(w + ".") | (st.session_state.df["WBS"] == w) for w in secili_wbsler]
                st.session_state.df = st.session_state.df[~pd.concat(maskeler, axis=1).any(axis=1)]
                st.session_state.secili_wbs_listesi = [kokler[0]] if kokler else ["1"]
                st.session_state.df = akilli_dengele(st.session_state.df, st.session_state.df); st.rerun()
                
        st.markdown("#### 📦 Çoklu Taşıma (Kes/Yapıştır)")
        k1, k2 = st.columns(2)
        with k1:
            if st.button("✂️ Kes", use_container_width=True):
                if any(k in secili_wbsler for k in kokler): st.error("Ana Blok / Kök kesilemez.")
                else: st.session_state.kesilen_wbs_listesi = secili_wbsler.copy(); st.rerun()
        with k2:
            yapistir_disabled = len(st.session_state.kesilen_wbs_listesi) == 0 or len(secili_wbsler) != 1
            if st.button("📋 İçine Yapıştır", use_container_width=True, disabled=yapistir_disabled):
                st.session_state.df = wbs_hedefe_tasi_coklu(st.session_state.kesilen_wbs_listesi, secili_wbsler[0])
                st.session_state.df = wbs_sirala(st.session_state.df)
                st.session_state.df = akilli_dengele(st.session_state.df, st.session_state.df)
                st.session_state.kesilen_wbs_listesi = []
                st.session_state.secili_wbs_listesi = [secili_wbsler[0]] 
                st.rerun()
                
        if st.session_state.kesilen_wbs_listesi: st.warning(f"📦 **{len(st.session_state.kesilen_wbs_listesi)}** öğe kesildi.")

# =========================================================================
# SEKME 2: PİVOT RAPOR ANALİZİ
# =========================================================================
with tab2:
    st.markdown("### 📊 Dinamik Pivot Rapor (Taşeron / Mahal / İş Tipi Analizi)")
    p_col1, p_col2 = st.columns([1, 3])
    with p_col1: grup_secim = st.selectbox("Raporlama Kriterini Seçin:", ["Grup 1", "Grup 2", "Grup 3"])
    with p_col2:
        pivot_data = df[df["Is_Leaf"]].groupby(grup_secim).agg(Kalem_Sayisi=("WBS", "count"), Toplam_Butce=("Bütçe (TL)", "sum"), Gerceklesen_Hakedis=("Hakediş (TL)", "sum")).reset_index()
        pivot_data["Gerçekleşme (%)"] = (pivot_data["Gerceklesen_Hakedis"] / pivot_data["Toplam_Butce"] * 100).fillna(0)
        st.dataframe(
            pivot_data, use_container_width=True, hide_index=True,
            column_config={
                grup_secim: st.column_config.TextColumn(grup_secim, width="medium"),
                "Kalem_Sayisi": st.column_config.NumberColumn("Kalem Adedi (Yaprak)"),
                "Toplam_Butce": st.column_config.NumberColumn("Bütçe (TL)", format="%,.0f ₺"),
                "Gerceklesen_Hakedis": st.column_config.NumberColumn("Hakediş (TL)", format="%,.0f ₺"),
                "Gerçekleşme (%)": st.column_config.ProgressColumn("Fiziksel İlerleme (%)", format="%.2f %%", min_value=0.0, max_value=100.0),
            }
        )
