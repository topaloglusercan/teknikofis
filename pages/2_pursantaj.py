import io
import os
import re
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# --- 1. SABİTLER VE TEMEL TEMİZLİK / TÜRKÇE FORMAT YARDIMCILARI ---
BASE_COLUMNS = ["WBS", "Ad", "Yerel_%", "Kilit", "İlerleme_%", "Grup 1", "Grup 2", "Grup 3"]
GRUP_COLUMNS = ["Grup 1", "Grup 2", "Grup 3"]


def temizle_metin(val, varsayilan=""):
    """None, NaN veya boşluklu değerleri güvenli temiz string'e çevirir."""
    if val is None:
        return varsayilan
    try:
        if pd.isna(val):
            return varsayilan
    except (TypeError, ValueError):
        pass
    metin = str(val).strip()
    return metin if metin else varsayilan


def temizle_grup_degeri(val):
    """Grup hücrelerindeki None/NaN/boş değerleri '-' yapar, baş/son boşlukları siler."""
    temiz = temizle_metin(val, varsayilan="-")
    if temiz.lower() in {"nan", "none", "null"}:
        return "-"
    return temiz


def guvenli_yuzde(val, varsayilan=0.0):
    """Yüzde değerlerini güvenli float'a çevirir ve [0.0, 100.0] aralığına sabitler."""
    if val is None:
        return float(varsayilan)
    try:
        if pd.isna(val):
            return float(varsayilan)
        f = float(val)
        if not np.isfinite(f):
            return float(varsayilan)
        return float(np.clip(f, 0.0, 100.0))
    except (TypeError, ValueError):
        return float(varsayilan)


def bedel_metni_cozumle(metin):
    """
    Türkçe ('1.000.000,50') ve uluslararası ('1000000.50' / '1,000,000.50')
    para metinlerini güvenli şekilde pozitif/sıfır float'a dönüştürür.
    Geçersiz girişlerde None döndürür.
    """
    ham = temizle_metin(metin, "")
    if not ham:
        return None
    sadece_rakam_ve_ayirici = re.sub(r"[^\d.,]", "", ham)
    if not sadece_rakam_ve_ayirici:
        return None

    virgul_sayisi = sadece_rakam_ve_ayirici.count(",")
    nokta_sayisi = sadece_rakam_ve_ayirici.count(".")

    if virgul_sayisi > 0 and nokta_sayisi > 0:
        if sadece_rakam_ve_ayirici.rfind(",") > sadece_rakam_ve_ayirici.rfind("."):
            temiz = sadece_rakam_ve_ayirici.replace(".", "").replace(",", ".")
        else:
            temiz = sadece_rakam_ve_ayirici.replace(",", "")
    elif virgul_sayisi == 1:
        sag = sadece_rakam_ve_ayirici.split(",")[1]
        if len(sag) in (1, 2):
            temiz = sadece_rakam_ve_ayirici.replace(",", ".")
        else:
            temiz = sadece_rakam_ve_ayirici.replace(",", "")
    elif virgul_sayisi > 1:
        temiz = sadece_rakam_ve_ayirici.replace(",", "")
    elif nokta_sayisi == 1:
        sag = sadece_rakam_ve_ayirici.split(".")[1]
        if len(sag) in (1, 2):
            temiz = sadece_rakam_ve_ayirici
        else:
            temiz = sadece_rakam_ve_ayirici.replace(".", "")
    elif nokta_sayisi > 1:
        temiz = sadece_rakam_ve_ayirici.replace(".", "")
    else:
        temiz = sadece_rakam_ve_ayirici

    try:
        deger = float(temiz)
        if not np.isfinite(deger) or deger < 0:
            return None
        return deger
    except ValueError:
        return None


def bedel_bicimlendir(deger):
    """Float değeri Türkçe sayı formatına ('1.000.000,00') çevirir."""
    try:
        val = float(deger)
        if not np.isfinite(val):
            val = 0.0
    except (TypeError, ValueError):
        val = 0.0
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def para_bicimlendir_tr(deger):
    """Float değeri Türkçe para formatına ('1.000.000,00 ₺') çevirir."""
    return f"{bedel_bicimlendir(deger)} ₺"


def yuzde_bicimlendir_tr(deger):
    """Float değeri Türkçe yüzde formatına ('%13,50') çevirir."""
    try:
        val = float(deger)
        if not np.isfinite(val):
            val = 0.0
    except (TypeError, ValueError):
        val = 0.0
    return f"%{val:.2f}".replace(".", ",")


# --- 2. VERİ MODELİ VE ALFANÜMERİK SIRALAMA ---
def wbs_sort_key(wbs_str):
    """Alfanümerik ve Türkçe karakter destekli kararlı WBS sıralama anahtarı."""
    parts = str(wbs_str).strip().split(".")
    key_parts = []
    for p in parts:
        p = p.strip()
        if p.isdigit():
            key_parts.append((0, int(p), ""))
        else:
            match = re.match(r"^([a-zA-ZçğıöşüÇĞIÖŞÜ]+)(\d+)$", p)
            if match:
                key_parts.append((1, match.group(1).upper(), int(match.group(2))))
            else:
                key_parts.append((2, p.upper(), 0))
    return tuple(key_parts)


def wbs_sirala(df_to_sort):
    """DataFrame'i mutasyona uğratmadan WBS koduna göre sıralar."""
    if df_to_sort is None or df_to_sort.empty:
        return pd.DataFrame(columns=BASE_COLUMNS) if df_to_sort is None else df_to_sort.copy().reset_index(drop=True)
    df_copy = df_to_sort.copy()
    df_copy["_sort_key"] = df_copy["WBS"].astype(str).apply(wbs_sort_key)
    df_copy = df_copy.sort_values(by="_sort_key", kind="mergesort").drop(columns=["_sort_key"])
    return df_copy.reset_index(drop=True)


def varsayilan_veri_olustur():
    data = [
        {"WBS": "DUV", "Ad": "DUVAR KAPLAMA İŞLERİ", "Yerel_%": 100.0, "Kilit": True, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"},
        {"WBS": "DUV.1", "Ad": "İZOLASYON İŞLERİ (ISI,SES VE SU)", "Yerel_%": 13.5, "Kilit": True, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"},
        {"WBS": "DUV.2", "Ad": "BRÜT YÜZEY TAMİR, KARA SIVA VE ALÇI SIVA İŞLERİ", "Yerel_%": 48.9, "Kilit": True, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"},
        {"WBS": "DUV.3", "Ad": "SERAMİK KAPLAMA İŞLERİ", "Yerel_%": 14.5, "Kilit": True, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"},
        {"WBS": "DUV.4", "Ad": "AHŞAP KAPLAMA İŞLERİ", "Yerel_%": 9.4, "Kilit": True, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"},
        {"WBS": "DUV.5", "Ad": "BOYA İŞLERİ", "Yerel_%": 13.7, "Kilit": False, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-"},
    ]
    return veri_cercevesini_normalize_et(pd.DataFrame(data))


def veri_cercevesini_normalize_et(df_raw):
    """Tüm zorunlu sütunları, tipleri ve grup temizliğini garanti altına alır."""
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=BASE_COLUMNS)

    df_norm = df_raw.copy()
    if "WBS" not in df_norm.columns or "Ad" not in df_norm.columns:
        raise ValueError("Veri tablosunda 'WBS' ve 'Ad' sütunları zorunludur.")

    df_norm["WBS"] = df_norm["WBS"].apply(lambda x: temizle_metin(x, ""))
    df_norm = df_norm[df_norm["WBS"] != ""].copy()

    df_norm["Ad"] = df_norm["Ad"].apply(lambda x: temizle_metin(x, "İsimsiz Kalem"))
    if "Yerel_%" not in df_norm.columns:
        df_norm["Yerel_%"] = 0.0
    df_norm["Yerel_%"] = df_norm["Yerel_%"].apply(guvenli_yuzde).round(4)

    if "Kilit" not in df_norm.columns:
        df_norm["Kilit"] = False
    df_norm["Kilit"] = df_norm["Kilit"].fillna(False).astype(bool)

    if "İlerleme_%" not in df_norm.columns:
        df_norm["İlerleme_%"] = 0.0
    df_norm["İlerleme_%"] = df_norm["İlerleme_%"].apply(guvenli_yuzde).round(2)

    for g_col in GRUP_COLUMNS:
        if g_col not in df_norm.columns:
            df_norm[g_col] = "-"
        df_norm[g_col] = df_norm[g_col].apply(temizle_grup_degeri)

    return wbs_sirala(df_norm[BASE_COLUMNS])


# --- 3. DİNAMİK EBEVEYN VE ALT AĞAÇ HARİTALARI (O(N)) ---
def _wbs_depth(wbs_str):
    return len(str(wbs_str).strip().split(".")) - 1


def ebeveyn_haritasi_olustur(df):
    """
    Sıralı WBS tablosunda her kalemin dinamik ebeveynini O(N) sürede hesaplar.
    Kural: Bir kalemin ebeveyni, kendisinden önce gelen ve derinliği kendisinden
    küçük olan en yakın kalemdir; yoksa 'KÖK'tür.
    """
    if df is None or df.empty:
        return {}
    df_sorted = wbs_sirala(df)
    parent_map = {}
    stack = []
    for wbs in df_sorted["WBS"].astype(str):
        depth = _wbs_depth(wbs)
        while stack and stack[-1][0] >= depth:
            stack.pop()
        parent_map[wbs] = stack[-1][1] if stack else "KÖK"
        stack.append((depth, wbs))
    return parent_map


def cocuk_haritasi_olustur(df, parent_map=None):
    """Sıralı tablo düzeninde her ebeveynin doğrudan çocuklarını listeler."""
    df_sorted = wbs_sirala(df)
    if parent_map is None:
        parent_map = ebeveyn_haritasi_olustur(df_sorted)
    children_map = {}
    for wbs in df_sorted["WBS"].astype(str):
        p = parent_map.get(wbs, "KÖK")
        children_map.setdefault(p, []).append(wbs)
    return children_map


def alt_agac_wbsleri(df, kok_wbs, parent_map=None):
    """Dinamik hiyerarşiye göre kok_wbs ve tüm alt çocuklarını sıralı liste olarak döndürür."""
    kok_str = str(kok_wbs)
    children_map = cocuk_haritasi_olustur(df, parent_map)
    sonuc = []
    ziyaret = set()

    def _gez(w):
        if w in ziyaret:
            return
        ziyaret.add(w)
        sonuc.append(w)
        for c in children_map.get(w, []):
            _gez(c)

    if kok_str in set(df["WBS"].astype(str).tolist()):
        _gez(kok_str)
    return sonuc


def get_dynamic_parent(df, target_wbs):
    target_wbs = str(target_wbs).strip()
    if _wbs_depth(target_wbs) == 0:
        return "KÖK"
    return ebeveyn_haritasi_olustur(df).get(target_wbs, "KÖK")


def gorsel_derinlik_haritasi_olustur(parent_map):
    """Her WBS'in ağaçtaki gerçek ebeveyn zinciri derinliğini (kök=0) hesaplar."""
    depth_map = {}
    for wbs in parent_map:
        d = 0
        curr = parent_map[wbs]
        ziyaret = {wbs}
        while curr != "KÖK" and curr not in ziyaret:
            ziyaret.add(curr)
            d += 1
            curr = parent_map.get(curr, "KÖK")
        depth_map[wbs] = d
    return depth_map


def etkin_grup_haritasi_olustur(df, grup_col, parent_map):
    """
    Her WBS satırı için etkin grubu belirler:
    - Satırın kendi grubu '-' değilse kendi grubunu kullanır.
    - Satırın kendi grubu '-' ise, ebeveyn zincirinde yukarı doğru ilk '-' olmayan
      grubun değerini miras alır.
    - Hiçbir üst kalemde grup yoksa '-' döndürür.
    """
    ham_harita = {
        str(w): temizle_grup_degeri(g)
        for w, g in zip(df["WBS"].astype(str), df[grup_col])
    }
    etkin_harita = {}

    for wbs in ham_harita:
        curr = wbs
        ziyaret = set()
        bulunan = "-"
        while curr != "KÖK" and curr not in ziyaret:
            ziyaret.add(curr)
            val = ham_harita.get(curr, "-")
            if val != "-":
                bulunan = val
                break
            curr = parent_map.get(curr, "KÖK")
        etkin_harita[wbs] = bulunan
    return etkin_harita


# --- 4. ALFANÜMERİK WBS İŞLEMLERİ (DİNAMİK ALT AĞAÇ DESTEKLİ) ---
def parent_bul(wbs):
    return str(wbs).rsplit(".", 1)[0] if "." in str(wbs) else "KÖK"


def sonraki_wbs_uret(parent_wbs, current_df):
    mevcut_set = set(current_df["WBS"].astype(str).tolist()) if not current_df.empty else set()
    if parent_wbs == "KÖK":
        sayilar = [int(w) for w in mevcut_set if "." not in w and w.isdigit()]
        aday_num = (max(sayilar) + 1) if sayilar else 1
        while str(aday_num) in mevcut_set:
            aday_num += 1
        return str(aday_num)
    else:
        prefix = str(parent_wbs) + "."
        parent_parts_len = len(str(parent_wbs).split("."))
        cocuklar = [
            w for w in mevcut_set
            if w.startswith(prefix) and len(w.split(".")) == parent_parts_len + 1
        ]
        max_num = 0
        for cocuk in cocuklar:
            son_kisim = cocuk.split(".")[-1]
            if son_kisim.isdigit():
                max_num = max(max_num, int(son_kisim))
            else:
                match = re.search(r"\d+$", son_kisim)
                if match:
                    max_num = max(max_num, int(match.group()))
        aday_num = max_num + 1
        while f"{prefix}{aday_num}" in mevcut_set:
            aday_num += 1
        return f"{prefix}{aday_num}"


def wbs_alt_agac_yeniden_numarala_haritasi(dataframe, eski_kok, yeni_kok):
    """
    Dinamik ağaç yapısını koruyarak eski_kok ve altındaki TÜM çocukları
    yeni_kok, yeni_kok.1, yeni_kok.1.1 şeklinde özyinelemeli olarak eşler.
    """
    eski_str = str(eski_kok)
    yeni_str = str(yeni_kok)
    parent_map = ebeveyn_haritasi_olustur(dataframe)
    children_map = cocuk_haritasi_olustur(dataframe, parent_map)

    yeni_harita = {}
    ziyaret = set()

    def _ata(eski_w, yeni_w):
        if eski_w in ziyaret:
            return
        ziyaret.add(eski_w)
        yeni_harita[eski_w] = yeni_w
        for idx_c, cocuk_w in enumerate(children_map.get(eski_w, []), start=1):
            _ata(cocuk_w, f"{yeni_w}.{idx_c}")

    if eski_str in set(dataframe["WBS"].astype(str).tolist()):
        _ata(eski_str, yeni_str)
    return yeni_harita


def wbs_degistir(eski_kok, yeni_kok, dataframe):
    """Dinamik alt ağaçtaki tüm kalemleri yeni_kok temelinde yeniden numaralandırır."""
    df_copy = dataframe.copy()
    yeni_harita = wbs_alt_agac_yeniden_numarala_haritasi(df_copy, eski_kok, yeni_kok)
    if not yeni_harita:
        return df_copy
    df_copy["WBS"] = df_copy["WBS"].astype(str).apply(lambda x: yeni_harita.get(x, x))
    return df_copy


def wbs_degistir_guvenli(eski_kok, yeni_kok, dataframe):
    """
    WBS yeniden adlandırma öncesi boşluk, tekillik ve alt-ağaç çakışma kontrolü yapar.
    Dönüş: (basarili: bool, yeni_df_veya_hata_mesaji, yeni_kod_haritasi: dict)
    """
    eski_str = temizle_metin(eski_kok, "")
    yeni_str = temizle_metin(yeni_kok, "")
    if not yeni_str:
        return False, "❌ Hata: WBS kodu boş olamaz!", {}
    if yeni_str == eski_str:
        return True, dataframe.copy(), {}

    yeni_harita = wbs_alt_agac_yeniden_numarala_haritasi(dataframe, eski_str, yeni_str)
    if not yeni_harita:
        return False, f"❌ Hata: '{eski_str}' numaralı WBS kodu bulunamadı!", {}

    wbs_series = dataframe["WBS"].astype(str)
    tasinan_set = set(yeni_harita.keys())
    diger_wbsler = set(wbs_series[~wbs_series.isin(tasinan_set)].tolist())
    yeni_degerler = set(yeni_harita.values())

    cakisanlar = diger_wbsler.intersection(yeni_degerler)
    if cakisanlar:
        ornek = sorted(cakisanlar)[0]
        return False, f"❌ Hata: '{ornek}' numaralı WBS kodu zaten kullanılıyor!", {}

    guncel = dataframe.copy()
    guncel["WBS"] = guncel["WBS"].astype(str).apply(lambda x: yeni_harita.get(x, x))
    return True, wbs_sirala(guncel), yeni_harita


def akilli_dengele(old_df, new_df):
    """
    Her dinamik ebeveyn grubunda Yerel_% toplamını 100.0000% olacak şekilde dengeler.
    """
    if new_df is None or new_df.empty:
        return pd.DataFrame(columns=BASE_COLUMNS) if new_df is None else new_df.copy()

    temp_df = new_df.copy().reset_index(drop=True)
    temp_df["Yerel_%"] = temp_df["Yerel_%"].apply(guvenli_yuzde)
    temp_df["Kilit"] = temp_df["Kilit"].fillna(False).astype(bool)

    parent_map = ebeveyn_haritasi_olustur(temp_df)
    temp_df["Dyn_Parent"] = temp_df["WBS"].astype(str).map(parent_map)

    old_yerel_map = {}
    if old_df is not None and not old_df.empty:
        old_yerel_map = {
            str(w): guvenli_yuzde(v)
            for w, v in zip(old_df["WBS"].astype(str), old_df["Yerel_%"])
        }

    for p in temp_df["Dyn_Parent"].unique():
        mask = temp_df["Dyn_Parent"] == p
        grup_idx = temp_df.index[mask].tolist()

        degisen_idx = []
        for idx in grup_idx:
            w = str(temp_df.at[idx, "WBS"])
            if w in old_yerel_map and not np.isclose(old_yerel_map[w], float(temp_df.at[idx, "Yerel_%"]), atol=0.001):
                degisen_idx.append(idx)

        if degisen_idx:
            temp_df.loc[degisen_idx, "Kilit"] = True

        locked_mask = mask & temp_df["Kilit"]
        unlocked_mask = mask & ~temp_df["Kilit"]
        locked_sum = float(temp_df.loc[locked_mask, "Yerel_%"].sum())

        if unlocked_mask.sum() == 0 and not np.isclose(locked_sum, 100.0, atol=0.01):
            acilacak_idx = [idx for idx in grup_idx if idx not in degisen_idx]
            if acilacak_idx:
                temp_df.loc[acilacak_idx, "Kilit"] = False
                locked_mask = mask & temp_df["Kilit"]
                unlocked_mask = mask & ~temp_df["Kilit"]
                locked_sum = float(temp_df.loc[locked_mask, "Yerel_%"].sum())

        unlocked_indices = temp_df.index[unlocked_mask].tolist()
        if unlocked_indices:
            remaining = max(0.0, round(100.0 - locked_sum, 4))
            unlocked_current_sum = float(temp_df.loc[unlocked_indices, "Yerel_%"].sum())

            if unlocked_current_sum > 0:
                yeni_degerler = [
                    round(float(temp_df.at[idx, "Yerel_%"]) * (remaining / unlocked_current_sum), 4)
                    for idx in unlocked_indices
                ]
            else:
                esit_pay = round(remaining / len(unlocked_indices), 4)
                yeni_degerler = [esit_pay for _ in unlocked_indices]

            if len(yeni_degerler) > 0:
                fark = round(remaining - sum(yeni_degerler), 4)
                yeni_degerler[-1] = round(max(0.0, yeni_degerler[-1] + fark), 4)

            for idx, val in zip(unlocked_indices, yeni_degerler):
                temp_df.at[idx, "Yerel_%"] = val

    temp_df["Yerel_%"] = temp_df["Yerel_%"].round(4)
    return temp_df.drop(columns=["Dyn_Parent"])


def wbs_hedefe_tasi_coklu_df(df, kaynak_wbs_listesi, hedef_parent_wbs):
    df_temp = df.copy()
    parent_map = ebeveyn_haritasi_olustur(df_temp)
    kokler = {w for w, p in parent_map.items() if p == "KÖK"}
    hedef_str = str(hedef_parent_wbs)

    adaylar = []
    for k_wbs in kaynak_wbs_listesi:
        k_str = str(k_wbs)
        if k_str in kokler or k_str == hedef_str or k_str not in parent_map:
            continue
        alt_set = set(alt_agac_wbsleri(df_temp, k_str, parent_map))
        if hedef_str in alt_set:
            continue
        adaylar.append(k_str)

    aday_set = set(adaylar)
    gecerli_kaynaklar = []
    for k_str in adaylar:
        ust_secili = False
        curr = parent_map.get(k_str, "KÖK")
        while curr != "KÖK":
            if curr in aday_set:
                ust_secili = True
                break
            curr = parent_map.get(curr, "KÖK")
        if not ust_secili:
            gecerli_kaynaklar.append(k_str)

    if not gecerli_kaynaklar:
        return df_temp

    tasinan_tum = set()
    for k in gecerli_kaynaklar:
        tasinan_tum.update(alt_agac_wbsleri(df_temp, k, parent_map))
    df_temp.loc[df_temp["WBS"].astype(str).isin(tasinan_tum), "Kilit"] = False

    for k_wbs in gecerli_kaynaklar:
        yeni_wbs = sonraki_wbs_uret(hedef_str, df_temp)
        df_temp = wbs_degistir(k_wbs, yeni_wbs, df_temp)

    df_temp = wbs_sirala(df_temp)
    return akilli_dengele(df_temp, df_temp)


def wbs_girinti_artir_df(df, kaynak_wbs):
    kaynak_str = str(kaynak_wbs)
    parent_map = ebeveyn_haritasi_olustur(df)
    parent = parent_map.get(kaynak_str, "KÖK")
    if parent == "KÖK":
        return df.copy(), kaynak_str

    kardesler = [w for w, p in parent_map.items() if p == parent]
    kardesler = wbs_sirala(pd.DataFrame({"WBS": kardesler}))["WBS"].tolist()
    if kaynak_str not in kardesler:
        return df.copy(), kaynak_str

    idx = kardesler.index(kaynak_str)
    if idx > 0:
        onceki_kardes = kardesler[idx - 1]
        yeni_wbs = sonraki_wbs_uret(onceki_kardes, df)
        yeni_df = wbs_degistir(kaynak_str, yeni_wbs, df)
        yeni_df = wbs_sirala(yeni_df)
        yeni_df = akilli_dengele(yeni_df, yeni_df)
        return yeni_df, yeni_wbs
    return df.copy(), kaynak_str


def wbs_girinti_azalt_df(df, kaynak_wbs):
    kaynak_str = str(kaynak_wbs)
    if "." not in kaynak_str:
        return df.copy(), kaynak_str
    parent_map = ebeveyn_haritasi_olustur(df)
    parent = parent_map.get(kaynak_str, "KÖK")
    if parent == "KÖK":
        return df.copy(), kaynak_str
    buyuk_baba = parent_map.get(parent, "KÖK")

    yeni_wbs = sonraki_wbs_uret(buyuk_baba, df)
    yeni_df = wbs_degistir(kaynak_str, yeni_wbs, df)
    yeni_df = wbs_sirala(yeni_df)
    yeni_df = akilli_dengele(yeni_df, yeni_df)
    return yeni_df, yeni_wbs


def wbs_tasi_df(df, kaynak_wbs, yon):
    kaynak_str = str(kaynak_wbs)
    parent_map = ebeveyn_haritasi_olustur(df)
    parent = parent_map.get(kaynak_str, "KÖK")
    kardesler = [w for w, p in parent_map.items() if p == parent]
    kardesler = wbs_sirala(pd.DataFrame({"WBS": kardesler}))["WBS"].tolist()

    if kaynak_str not in kardesler:
        return df.copy(), kaynak_str
    idx = kardesler.index(kaynak_str)
    hedef_wbs = None
    if yon == "yukari" and idx > 0:
        hedef_wbs = kardesler[idx - 1]
    elif yon == "asagi" and idx < len(kardesler) - 1:
        hedef_wbs = kardesler[idx + 1]

    if hedef_wbs:
        temp_wbs = f"__TEMP_SWAP_{kaynak_str}__"
        df_temp = wbs_degistir(kaynak_str, temp_wbs, df)
        df_temp = wbs_degistir(hedef_wbs, kaynak_str, df_temp)
        df_temp = wbs_degistir(temp_wbs, hedef_wbs, df_temp)
        return wbs_sirala(df_temp), hedef_wbs
    return df.copy(), kaynak_str


def wbs_agaci_kopyala_df(df, kaynak_wbs, kopya_sayisi):
    kaynak_str = str(kaynak_wbs)
    parent_map = ebeveyn_haritasi_olustur(df)
    alt_wbsler = alt_agac_wbsleri(df, kaynak_str, parent_map)
    if not alt_wbsler or kopya_sayisi < 1:
        return df.copy()

    kaynak_agac = df[df["WBS"].astype(str).isin(alt_wbsler)].copy()
    parent = parent_map.get(kaynak_str, "KÖK")
    df_temp = df.copy()

    for i in range(int(kopya_sayisi)):
        yeni_kok_wbs = sonraki_wbs_uret(parent, df_temp)
        yeni_harita = wbs_alt_agac_yeniden_numarala_haritasi(kaynak_agac, kaynak_str, yeni_kok_wbs)
        kopya_agac_i = kaynak_agac.copy()
        kopya_agac_i["WBS"] = kopya_agac_i["WBS"].astype(str).map(yeni_harita)
        kopya_agac_i.loc[kopya_agac_i["WBS"] == yeni_kok_wbs, "Ad"] += f" (Kopya {i + 1})"
        kopya_agac_i["İlerleme_%"] = 0.0
        kopya_agac_i["Kilit"] = False
        df_temp = pd.concat([df_temp, kopya_agac_i], ignore_index=True)

    df_temp = wbs_sirala(df_temp)
    return akilli_dengele(df_temp, df_temp)


def wbs_sil_df(df, silinecek_wbs_listesi):
    """
    Seçili WBS kalemlerini ve dinamik alt kırılımlarını siler.
    Tablonun tamamen boş kalmasını engeller.
    """
    if df.empty or not silinecek_wbs_listesi:
        return False, "Silinecek kalem seçilmedi."

    parent_map = ebeveyn_haritasi_olustur(df)
    silinecek_tum = set()
    for w in silinecek_wbs_listesi:
        silinecek_tum.update(alt_agac_wbsleri(df, str(w), parent_map))

    kalan = df[~df["WBS"].astype(str).isin(silinecek_tum)].copy().reset_index(drop=True)
    if kalan.empty:
        return False, "❌ Projedeki tüm kalemler silinemez; en az bir kalem kalmalıdır."

    kalan = wbs_sirala(kalan)
    kalan = akilli_dengele(kalan, kalan)
    return True, kalan


# --- 5. HESAPLAMA MOTORU VE PİVOT ---
def gorsel_isim_olustur(row):
    derinlik = int(row["Gorsel_Derinlik"])
    if row["Dyn_Parent"] == "KÖK":
        icon = "🏗️"
    elif derinlik == 1:
        icon = "🏢"
    elif derinlik == 2:
        icon = "🗂️"
    elif derinlik == 3:
        icon = "📂"
    elif derinlik == 4:
        icon = "📑"
    else:
        icon = "📄"
    prefix = "  " * derinlik + ("└─ " if derinlik > 0 else "")
    return f"{prefix}{icon} {row['Ad']}"


def hesapla_pursantaj(raw_df, proje_bedeli):
    """
    Ham WBS tablosundan tüm hiyerarşi, genel pay, bütçe, hakediş,
    Türkçe para görünüm sütunları, yaprak/ana-başlık ayrımı ve
    miraslı etkin grup sütunlarını hesaplar.
    """
    base_df = veri_cercevesini_normalize_et(raw_df)
    if base_df.empty:
        bos = base_df.copy()
        for col in [
            "Dyn_Parent", "Derinlik", "Gorsel_Derinlik", "Is_Leaf", "Kalem Tipi",
            "Ağaç Görünümü", "Genel_%", "Bütçe (TL)", "Hakediş (TL)",
            "Bütçe (TR)", "Hakediş (TR)",
            "Etkin_Grup 1", "Etkin_Grup 2", "Etkin_Grup 3",
        ]:
            bos[col] = pd.Series(dtype="object")
        return bos, base_df

    df = base_df.copy()
    parent_map = ebeveyn_haritasi_olustur(df)
    gorsel_derinlik_map = gorsel_derinlik_haritasi_olustur(parent_map)

    df["Dyn_Parent"] = df["WBS"].astype(str).map(parent_map)
    df["Derinlik"] = df["WBS"].astype(str).apply(_wbs_depth)
    df["Gorsel_Derinlik"] = df["WBS"].astype(str).map(gorsel_derinlik_map)

    ebeveyn_set = set(df["Dyn_Parent"].tolist()) - {"KÖK"}
    df["Is_Leaf"] = ~df["WBS"].astype(str).isin(ebeveyn_set)
    df["Kalem Tipi"] = np.where(df["Is_Leaf"], "✅ Aktivite", "🔒 Ana Başlık (Oto)")
    df["Ağaç Görünümü"] = df.apply(gorsel_isim_olustur, axis=1)

    for g_col in GRUP_COLUMNS:
        etkin_map = etkin_grup_haritasi_olustur(df, g_col, parent_map)
        df[f"Etkin_{g_col}"] = df["WBS"].astype(str).map(etkin_map)

    genel_map = {}
    max_g_derinlik = int(df["Gorsel_Derinlik"].max())
    for d in range(0, max_g_derinlik + 1):
        for idx, row in df[df["Gorsel_Derinlik"] == d].iterrows():
            w = str(row["WBS"])
            rp = row["Dyn_Parent"]
            if rp == "KÖK":
                genel_val = float(row["Yerel_%"])
            else:
                parent_genel = genel_map.get(rp, 100.0)
                genel_val = (float(row["Yerel_%"]) / 100.0) * parent_genel
            genel_map[w] = genel_val
            df.at[idx, "Genel_%"] = genel_val

    df["Bütçe (TL)"] = ((df["Genel_%"] / 100.0) * float(proje_bedeli)).round(2)

    df["Hakediş (TL)"] = 0.0
    leaf_mask = df["Is_Leaf"]
    df.loc[leaf_mask, "Hakediş (TL)"] = (
        df.loc[leaf_mask, "Bütçe (TL)"] * (df.loc[leaf_mask, "İlerleme_%"] / 100.0)
    ).round(2)

    for d in range(max_g_derinlik, -1, -1):
        for idx, row in df[df["Gorsel_Derinlik"] == d].iterrows():
            if not row["Is_Leaf"]:
                child_mask = df["Dyn_Parent"] == row["WBS"]
                child_hakedis_sum = float(df.loc[child_mask, "Hakediş (TL)"].sum())
                df.at[idx, "Hakediş (TL)"] = round(child_hakedis_sum, 2)
                butce = float(df.at[idx, "Bütçe (TL)"])
                if butce > 0:
                    calculated_progress = round(min(100.0, max(0.0, (child_hakedis_sum / butce) * 100.0)), 2)
                else:
                    calculated_progress = 0.0
                df.at[idx, "İlerleme_%"] = calculated_progress
                base_df.at[idx, "İlerleme_%"] = calculated_progress

    df["Bütçe (TR)"] = df["Bütçe (TL)"].apply(para_bicimlendir_tr)
    df["Hakediş (TR)"] = df["Hakediş (TL)"].apply(para_bicimlendir_tr)

    return df, base_df


def pivot_tablo_olustur(computed_df, grup_col):
    """
    Yaprak aktiviteler üzerinden etkin (miraslı) grup bazlı pivot tablo üretir.
    Dönüş: (pivot_df, ezilen_ebeveyn_gruplari: list[str])
    """
    if computed_df is None or computed_df.empty:
        bos = pd.DataFrame(columns=[
            grup_col, "Kalem_Sayisi", "Toplam_Butce", "Gerceklesen_Hakedis",
            "Bütçe (TR)", "Hakediş (TR)", "Gerçekleşme (%)"
        ])
        return bos, []

    etkin_col = f"Etkin_{grup_col}"
    if etkin_col not in computed_df.columns:
        parent_map = ebeveyn_haritasi_olustur(computed_df)
        etkin_map = etkin_grup_haritasi_olustur(computed_df, grup_col, parent_map)
        etkin_seri = computed_df["WBS"].astype(str).map(etkin_map)
    else:
        etkin_seri = computed_df[etkin_col]

    yaprak_df = computed_df[computed_df["Is_Leaf"]].copy()
    yaprak_df[grup_col] = etkin_seri[computed_df["Is_Leaf"]].apply(temizle_grup_degeri)

    if yaprak_df.empty:
        bos = pd.DataFrame(columns=[
            grup_col, "Kalem_Sayisi", "Toplam_Butce", "Gerceklesen_Hakedis",
            "Bütçe (TR)", "Hakediş (TR)", "Gerçekleşme (%)"
        ])
        return bos, []

    pivot_data = (
        yaprak_df.groupby(grup_col, dropna=False)
        .agg(
            Kalem_Sayisi=("WBS", "count"),
            Toplam_Butce=("Bütçe (TL)", "sum"),
            Gerceklesen_Hakedis=("Hakediş (TL)", "sum"),
        )
        .reset_index()
    )
    pivot_data["Toplam_Butce"] = pivot_data["Toplam_Butce"].round(2)
    pivot_data["Gerceklesen_Hakedis"] = pivot_data["Gerceklesen_Hakedis"].round(2)
    pivot_data["Bütçe (TR)"] = pivot_data["Toplam_Butce"].apply(para_bicimlendir_tr)
    pivot_data["Hakediş (TR)"] = pivot_data["Gerceklesen_Hakedis"].apply(para_bicimlendir_tr)
    pivot_data["Gerçekleşme (%)"] = np.where(
        pivot_data["Toplam_Butce"] > 0,
        np.clip((pivot_data["Gerceklesen_Hakedis"] / pivot_data["Toplam_Butce"]) * 100.0, 0.0, 100.0),
        0.0,
    ).round(2)

    ebeveyn_gruplari = set(
        computed_df.loc[~computed_df["Is_Leaf"], grup_col].apply(temizle_grup_degeri).tolist()
    ) - {"-"}
    pivottaki_gruplar = set(pivot_data[grup_col].tolist())
    ezilen_ebeveyn_gruplari = sorted(ebeveyn_gruplari - pivottaki_gruplar)

    return pivot_data, ezilen_ebeveyn_gruplari


def editor_degisikliklerini_uygula(ana_df, gosterilecek_df, edited_df):
    """
    Planlama & Master tablo değişikliklerini güvenli şekilde işler:
    - WBS değiştirildiğinde alt ağaç otomatik yeniden numaralandırılır.
    - İlerleme_% sütunu varsa bile ana başlıklara manuel ilerleme girişini engeller.
    """
    sonuc = {
        "df": ana_df.copy(),
        "degisim_var": False,
        "wbs_degisti": False,
        "hata": None,
        "engellenen_ebeveynler": [],
        "wbs_yeniden_esleme": {},
    }

    if gosterilecek_df.empty or edited_df.empty:
        return sonuc

    yeni_ana_df = ana_df.copy()

    for i in range(len(gosterilecek_df)):
        eski_w = temizle_metin(gosterilecek_df.iloc[i]["WBS"], "")
        yeni_w = temizle_metin(edited_df.iloc[i]["WBS"], "")
        if eski_w != yeni_w:
            basarili, res, yeni_harita = wbs_degistir_guvenli(eski_w, yeni_w, yeni_ana_df)
            if not basarili:
                sonuc["hata"] = res
                return sonuc
            sonuc["df"] = res
            sonuc["degisim_var"] = True
            sonuc["wbs_degisti"] = True
            sonuc["wbs_yeniden_esleme"] = yeni_harita
            return sonuc

    for i in range(len(gosterilecek_df)):
        eski_row = gosterilecek_df.iloc[i]
        yeni_row = edited_df.iloc[i]
        wbs = str(eski_row["WBS"])
        is_leaf = bool(eski_row["Is_Leaf"])
        hedef_mask = yeni_ana_df["WBS"].astype(str) == wbs

        eski_ad = temizle_metin(eski_row["Ad"], "İsimsiz Kalem")
        yeni_ad = temizle_metin(yeni_row["Ad"], "İsimsiz Kalem")
        if eski_ad != yeni_ad:
            yeni_ana_df.loc[hedef_mask, "Ad"] = yeni_ad
            sonuc["degisim_var"] = True

        eski_kilit = bool(eski_row["Kilit"])
        yeni_kilit = bool(yeni_row["Kilit"]) if pd.notna(yeni_row["Kilit"]) else False
        if eski_kilit != yeni_kilit:
            yeni_ana_df.loc[hedef_mask, "Kilit"] = yeni_kilit
            sonuc["degisim_var"] = True

        for g_col in GRUP_COLUMNS:
            if g_col in edited_df.columns:
                eski_g = temizle_grup_degeri(eski_row[g_col])
                yeni_g = temizle_grup_degeri(yeni_row[g_col])
                if eski_g != yeni_g:
                    yeni_ana_df.loc[hedef_mask, g_col] = yeni_g
                    sonuc["degisim_var"] = True

        if "Yerel_%" in edited_df.columns:
            eski_yerel = guvenli_yuzde(eski_row["Yerel_%"])
            yeni_yerel = guvenli_yuzde(yeni_row["Yerel_%"])
            if not np.isclose(eski_yerel, yeni_yerel, atol=0.005):
                yeni_ana_df.loc[hedef_mask, "Yerel_%"] = yeni_yerel
                sonuc["degisim_var"] = True

        if "İlerleme_%" in edited_df.columns:
            eski_il = guvenli_yuzde(eski_row["İlerleme_%"])
            yeni_il = guvenli_yuzde(yeni_row["İlerleme_%"])
            if not np.isclose(eski_il, yeni_il, atol=0.005):
                if is_leaf:
                    yeni_ana_df.loc[hedef_mask, "İlerleme_%"] = round(yeni_il, 2)
                    sonuc["degisim_var"] = True
                else:
                    sonuc["engellenen_ebeveynler"].append(wbs)

    if sonuc["degisim_var"]:
        dengeli_df = akilli_dengele(ana_df, yeni_ana_df)
        dengeli_df["Yerel_%"] = dengeli_df["Yerel_%"].round(4)
        sonuc["df"] = dengeli_df

    return sonuc


def yaprak_ilerleme_guncelle(ana_df, computed_df, wbs_ilerleme_eslesmesi):
    """
    Sadece çocuğu olmayan (Is_Leaf == True) aktivitelerin ilerleme yüzdelerini günceller.
    Çocuğu olan bir WBS verilirse reddeder ve engellenenler listesinde döndürür.
    """
    yeni_df = ana_df.copy()
    leaf_map = dict(zip(computed_df["WBS"].astype(str), computed_df["Is_Leaf"].astype(bool)))
    degisti = False
    engellenenler = []

    for wbs, yeni_val in wbs_ilerleme_eslesmesi.items():
        wbs_str = str(wbs)
        if not leaf_map.get(wbs_str, False):
            engellenenler.append(wbs_str)
            continue
        temiz_val = round(guvenli_yuzde(yeni_val), 2)
        mask = yeni_df["WBS"].astype(str) == wbs_str
        if mask.any():
            mevcut_val = guvenli_yuzde(yeni_df.loc[mask, "İlerleme_%"].iloc[0])
            if not np.isclose(mevcut_val, temiz_val, atol=0.005):
                yeni_df.loc[mask, "İlerleme_%"] = temiz_val
                degisti = True

    return yeni_df, degisti, engellenenler


def dosya_yukle_df(imp_df):
    """Excel veya JSON'dan okunan DataFrame'i doğrular, temizler ve dengeler."""
    if imp_df is None or imp_df.empty:
        raise ValueError("Yüklenen dosya boş.")
    if "WBS" not in imp_df.columns or "Ad" not in imp_df.columns:
        raise ValueError("Dosyada 'WBS' ve 'Ad' sütunları bulunamadı.")

    temiz = veri_cercevesini_normalize_et(imp_df)
    if temiz.empty:
        raise ValueError("Dosyada geçerli WBS satırı bulunamadı.")
    if temiz["WBS"].duplicated().any():
        tekrar = temiz.loc[temiz["WBS"].duplicated(), "WBS"].unique().tolist()
        raise ValueError(f"Dosyada mükerrer WBS kodları bulundu: {', '.join(tekrar[:5])}")

    return akilli_dengele(temiz, temiz)


# --- 6. EXCEL İNDİRME FONKSİYONLARI ---
def excel_sablon_olustur():
    df_sablon = pd.DataFrame({
        "WBS": ["DUV", "DUV.1", "DUV.1.1", "DUV.1.2", "DUV.2"],
        "Ad": ["Duvar İşleri", "Üst Yapı", "Kaba İnşaat", "İnce İşler", "Alt Yapı"],
    })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_sablon.to_excel(writer, index=False, sheet_name="WBS_Aktarim")
    return output.getvalue()


def excel_olustur(veri_df, proje_bedeli):
    veri_df = veri_df.reset_index(drop=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Pursantaj_İlerleme"
    ws["L1"] = "Sözleşme Bedeli:"
    ws["L1"].font = Font(bold=True)
    ws["M1"] = float(proje_bedeli)
    ws["M1"].number_format = '#,##0.00 "₺"'

    headers = [
        "WBS", "Ağaç Görünümü", "Kalem Adı", "Yerel Pay (%)", "Genel Pay (%)",
        "İlerleme (%)", "Bütçe (TL)", "Hakediş (TL)", "Grup 1", "Grup 2", "Grup 3"
    ]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0F2044", end_color="0F2044", fill_type="solid")

    wbs_to_row = {str(wbs): idx + 3 for idx, wbs in enumerate(veri_df["WBS"])}

    for idx, row in veri_df.iterrows():
        r = idx + 3
        ws.cell(row=r, column=1, value=str(row["WBS"]))
        ws.cell(row=r, column=2, value=str(row["Ağaç Görünümü"]))
        ws.cell(row=r, column=3, value=str(row["Ad"]))
        ws.cell(row=r, column=4, value=float(row["Yerel_%"]) / 100.0).number_format = "0.00%"

        if row["Dyn_Parent"] == "KÖK" or str(row["Dyn_Parent"]) not in wbs_to_row:
            ws.cell(row=r, column=5, value=f"=D{r}").number_format = "0.00%"
        else:
            p_row = wbs_to_row[str(row["Dyn_Parent"])]
            ws.cell(row=r, column=5, value=f"=D{r}*E{p_row}").number_format = "0.00%"

        if row["Is_Leaf"]:
            ws.cell(row=r, column=6, value=float(row["İlerleme_%"]) / 100.0).number_format = "0.00%"
        else:
            ws.cell(row=r, column=6, value=f"=IF(G{r}>0, H{r}/G{r}, 0)").number_format = "0.00%"

        ws.cell(row=r, column=7, value=f"=$M$1*E{r}").number_format = '#,##0.00 "₺"'

        if row["Is_Leaf"]:
            ws.cell(row=r, column=8, value=f"=G{r}*F{r}").number_format = '#,##0.00 "₺"'
        else:
            real_children = veri_df[veri_df["Dyn_Parent"] == row["WBS"]]["WBS"].astype(str).tolist()
            child_rows = [str(wbs_to_row[c]) for c in real_children if c in wbs_to_row]
            if child_rows:
                ws.cell(row=r, column=8, value="=" + "+".join([f"H{cr}" for cr in child_rows])).number_format = '#,##0.00 "₺"'
            else:
                ws.cell(row=r, column=8, value=0).number_format = '#,##0.00 "₺"'

        ws.cell(row=r, column=9, value=temizle_grup_degeri(row.get("Etkin_Grup 1", row["Grup 1"])))
        ws.cell(row=r, column=10, value=temizle_grup_degeri(row.get("Etkin_Grup 2", row["Grup 2"])))
        ws.cell(row=r, column=11, value=temizle_grup_degeri(row.get("Etkin_Grup 3", row["Grup 3"])))

    for col in ws.columns:
        max_length = 0
        for cell in col:
            val_len = len(str(cell.value)) if cell.value is not None else 0
            if val_len > max_length:
                max_length = val_len
        ws.column_dimensions[col[0].column_letter].width = min(max_length + 3, 60)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# --- 7. STREAMLIT ARAYÜZÜ ---
def render_app():
    st.markdown("""
    <style>
        .block-container {
            max-width: 95% !important;
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }
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

    if "bedel" not in st.session_state:
        st.session_state.bedel = 1_000_000.0
    if "bedel_str" not in st.session_state:
        st.session_state.bedel_str = bedel_bicimlendir(st.session_state.bedel)
    if "kesilen_wbs_listesi" not in st.session_state:
        st.session_state.kesilen_wbs_listesi = []
    if "df" not in st.session_state:
        st.session_state.df = varsayilan_veri_olustur()
    if "secili_wbs_listesi" not in st.session_state:
        st.session_state.secili_wbs_listesi = [st.session_state.df.iloc[0]["WBS"]] if not st.session_state.df.empty else []
    if "editor_version" not in st.session_state:
        st.session_state.editor_version = 0
    if "uyari_mesaji" not in st.session_state:
        st.session_state.uyari_mesaji = None

    df, synced_base = hesapla_pursantaj(st.session_state.df, st.session_state.bedel)
    st.session_state.df = synced_base

    # --- ANA BAŞLIK VE ÖZET METRİKLER ---
    st.title("🏗️ Profesyonel Pursantaj Yönetimi")

    toplam_butce = float(df.loc[df["Is_Leaf"], "Bütçe (TL)"].sum()) if not df.empty else 0.0
    toplam_hakedis = float(df.loc[df["Is_Leaf"], "Hakediş (TL)"].sum()) if not df.empty else 0.0
    genel_ilerleme = (toplam_hakedis / toplam_butce * 100.0) if toplam_butce > 0 else 0.0

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f"<div class='info-box'><div class='info-label'>Sözleşme Bedeli</div><div class='info-val'>{para_bicimlendir_tr(st.session_state.bedel)}</div></div>", unsafe_allow_html=True)
    with m2:
        st.markdown(f"<div class='info-box'><div class='info-label'>Toplam Hakediş</div><div class='info-val'>{para_bicimlendir_tr(toplam_hakedis)}</div></div>", unsafe_allow_html=True)
    with m3:
        st.markdown(f"<div class='info-box'><div class='info-label'>Genel Fiziksel İlerleme</div><div class='info-val'>{yuzde_bicimlendir_tr(genel_ilerleme)}</div></div>", unsafe_allow_html=True)

    if st.session_state.uyari_mesaji:
        st.warning(st.session_state.uyari_mesaji)
        st.session_state.uyari_mesaji = None

    with st.expander("📂 Dosya İşlemleri (Excel / JSON Yükle / Rapor İndir)", expanded=False):
        t_c1, t_c2, t_c3 = st.columns([1.2, 1, 1.5])
        with t_c1:
            yeni_bedel_str = st.text_input("Sözleşme Bedeli (₺)", value=st.session_state.bedel_str)
            yeni_bedel_float = bedel_metni_cozumle(yeni_bedel_str)
            if yeni_bedel_float is not None and not np.isclose(yeni_bedel_float, st.session_state.bedel, atol=0.005):
                st.session_state.bedel = yeni_bedel_float
                st.session_state.bedel_str = bedel_bicimlendir(yeni_bedel_float)
                st.rerun()

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
            st.download_button(
                "📊 Raporu İndir (Formüllü Excel)",
                data=excel_olustur(df, st.session_state.bedel),
                file_name="pursantaj_raporu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with t_c2:
            st.markdown("<div style='margin-top: 17px;'></div>", unsafe_allow_html=True)
            st.download_button(
                "📥 Excel Yükleme Şablonunu İndir",
                data=excel_sablon_olustur(),
                file_name="wbs_sablon.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.markdown("<div style='margin-top: 5px;'></div>", unsafe_allow_html=True)
            st.download_button(
                "📥 Projeyi Yedekle (.json)",
                data=st.session_state.df.to_json(orient="records", force_ascii=False),
                file_name="pursantaj.json",
                mime="application/json",
                use_container_width=True,
            )

        with t_c3:
            uploaded_file = st.file_uploader("📤 Proje Yükle (Excel veya JSON)", type=["xlsx", "xls", "json"])
            if "son_yuklenen_dosya_id" not in st.session_state:
                st.session_state.son_yuklenen_dosya_id = None

            if uploaded_file is not None:
                if st.session_state.son_yuklenen_dosya_id != uploaded_file.file_id:
                    try:
                        if uploaded_file.name.lower().endswith(".json"):
                            imp_df = pd.read_json(uploaded_file)
                        else:
                            imp_df = pd.read_excel(uploaded_file)

                        yuklenen = dosya_yukle_df(imp_df)
                        st.session_state.df = yuklenen
                        st.session_state.son_yuklenen_dosya_id = uploaded_file.file_id
                        st.session_state.secili_wbs_listesi = [yuklenen.iloc[0]["WBS"]]
                        st.session_state.editor_version += 1
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Dosya okuma hatası: {e}")
                else:
                    st.success("✅ Proje dosyası başarıyla yüklendi!")

    st.divider()
    tab1, tab2, tab3 = st.tabs([
        "📋 Planlama & Master Tablo",
        "✅ İlerleme Girişi (Alt Aktiviteler)",
        "📈 Pivot Rapor Analizi",
    ])

    # =========================================================================
    # SEKME 1: PLANLAMA VE MASTER TABLO (SADECE WBS, AĞIRLIK, BÜTÇE VE GRUPLAR)
    # =========================================================================
    with tab1:
        col_sol, col_sag = st.columns([3.5, 1.2])

        with col_sol:
            st.markdown("<div class='filter-container'>", unsafe_allow_html=True)
            f_c1, f_c2, f_c3, f_c4, f_c5 = st.columns([1.5, 1, 1, 1, 1.5])
            with f_c1:
                filtre_ad = st.text_input("Arama (Kalem Adı/WBS):")
            with f_c2:
                filtre_g1 = st.selectbox("Grup 1:", ["Tümü"] + sorted(df["Etkin_Grup 1"].unique().tolist()))
            with f_c3:
                filtre_g2 = st.selectbox("Grup 2:", ["Tümü"] + sorted(df["Etkin_Grup 2"].unique().tolist()))
            with f_c4:
                filtre_g3 = st.selectbox("Grup 3:", ["Tümü"] + sorted(df["Etkin_Grup 3"].unique().tolist()))
            with f_c5:
                mevcut_seviyeler = sorted(df["Gorsel_Derinlik"].unique().tolist())
                secilen_seviyeler = st.multiselect(
                    "Gösterilecek Seviyeler:",
                    options=mevcut_seviyeler,
                    default=mevcut_seviyeler,
                    format_func=lambda x: f"Seviye {x + 1}",
                )
            st.markdown("</div>", unsafe_allow_html=True)

            gosterilecek_df = df.copy()
            if filtre_ad:
                gosterilecek_df = gosterilecek_df[
                    gosterilecek_df["Ad"].astype(str).str.contains(filtre_ad, case=False, regex=False, na=False)
                    | gosterilecek_df["WBS"].astype(str).str.contains(filtre_ad, case=False, regex=False, na=False)
                ]
            if filtre_g1 != "Tümü":
                gosterilecek_df = gosterilecek_df[gosterilecek_df["Etkin_Grup 1"] == filtre_g1]
            if filtre_g2 != "Tümü":
                gosterilecek_df = gosterilecek_df[gosterilecek_df["Etkin_Grup 2"] == filtre_g2]
            if filtre_g3 != "Tümü":
                gosterilecek_df = gosterilecek_df[gosterilecek_df["Etkin_Grup 3"] == filtre_g3]
            if secilen_seviyeler:
                gosterilecek_df = gosterilecek_df[gosterilecek_df["Gorsel_Derinlik"].isin(secilen_seviyeler)]

            gosterilecek_df = gosterilecek_df.reset_index(drop=True)

            # Planlama tablosu: sadece WBS, hiyerarşi, ağırlık, bütçe ve gruplar (ilerleme ve hakediş yok)
            edit_cols = [
                "WBS", "Ağaç Görünümü", "Kalem Tipi", "Ad", "Kilit",
                "Yerel_%", "Bütçe (TR)", "Grup 1", "Grup 2", "Grup 3",
            ]

            edited_df = st.data_editor(
                gosterilecek_df[edit_cols],
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                height=600,
                column_config={
                    "WBS": st.column_config.TextColumn("✏️ WBS", width="small"),
                    "Ağaç Görünümü": st.column_config.TextColumn("Hiyerarşi (Görsel)", width="medium", disabled=True),
                    "Kalem Tipi": st.column_config.TextColumn("Durum", width="small", disabled=True),
                    "Ad": st.column_config.TextColumn("✏️ Kalem Adı"),
                    "Kilit": st.column_config.CheckboxColumn("🔒 Kilit", width="small"),
                    "Yerel_%": st.column_config.NumberColumn("Ağırlık (%)", format="%.2f %%", min_value=0.0, max_value=100.0, width="small"),
                    "Bütçe (TR)": st.column_config.TextColumn("Bütçe (₺)", disabled=True, width="small"),
                    "Grup 1": st.column_config.TextColumn("Grup 1", width="small"),
                    "Grup 2": st.column_config.TextColumn("Grup 2", width="small"),
                    "Grup 3": st.column_config.TextColumn("Grup 3", width="small"),
                },
                key=f"master_editor_{st.session_state.editor_version}",
            )

            st.caption(
                "ℹ️ Bir üst başlığın WBS kodunu değiştirdiğinizde (örn. `DUV` → `KAP`), alt kırılımları otomatik olarak `KAP.1`, `KAP.1.1` şeklinde güncellenir. "
                "İlerleme yüzdeleri ve hakediş takibi **✅ İlerleme Girişi (Alt Aktiviteler)** sekmesinden yapılır."
            )

            sonuc = editor_degisikliklerini_uygula(st.session_state.df, gosterilecek_df, edited_df)

            if sonuc["hata"]:
                st.error(sonuc["hata"])
                st.stop()

            if sonuc["degisim_var"]:
                st.session_state.df = sonuc["df"]
                if sonuc["wbs_degisti"] and sonuc["wbs_yeniden_esleme"]:
                    harita = sonuc["wbs_yeniden_esleme"]
                    st.session_state.secili_wbs_listesi = [
                        harita.get(x, x) for x in st.session_state.secili_wbs_listesi
                    ]
                    st.session_state.kesilen_wbs_listesi = [
                        harita.get(x, x) for x in st.session_state.kesilen_wbs_listesi
                    ]
                    st.session_state.editor_version += 1
                st.rerun()

        with col_sag:
            st.markdown("### 🛠️ İşlem Paneli")

            wbs_to_display = {w: f"[{w}] {a}" for w, a in zip(df["WBS"], df["Ad"])}
            display_to_wbs = {v: k for k, v in wbs_to_display.items()}

            gecerli_secimler = [w for w in st.session_state.secili_wbs_listesi if w in wbs_to_display]
            mevcut_display = [wbs_to_display[w] for w in gecerli_secimler]

            yeni_secim = st.multiselect(
                "👉 İşlem Yapılacak Kalem(ler):",
                options=list(wbs_to_display.values()),
                default=mevcut_display,
                help="Kalem arayabilir veya çoklu seçim yapabilirsiniz.",
            )

            if set(yeni_secim) != set(mevcut_display):
                st.session_state.secili_wbs_listesi = [display_to_wbs[v] for v in yeni_secim]
                st.rerun()

            secili_wbsler = gecerli_secimler
            is_single = len(secili_wbsler) == 1
            is_multi = len(secili_wbsler) > 0
            ilk_secili_wbs = secili_wbsler[0] if is_multi else None

            if not is_multi:
                st.info("👆 Lütfen işlem yapmak için yukarıdaki kutudan kalem seçiniz.")
            elif is_single:
                secili_row = df[df["WBS"] == ilk_secili_wbs].iloc[0]
                secili_is_leaf = bool(secili_row["Is_Leaf"])
                st.markdown(f"""
                <div class="selected-card">
                    <div class="sc-title">SEÇİLİ ÖĞE ({'ALT AKTİVİTE' if secili_is_leaf else 'ANA BAŞLIK'})</div>
                    <div class="sc-name">{secili_row['Ad']}</div>
                    <div class="sc-wbs">WBS: {secili_row['WBS']}</div>
                    <div class="sc-metrics">
                        <div class="sc-metric-box">
                            <div class="sc-metric-lbl">BÜTÇE</div>
                            <div class="sc-metric-val">{para_bicimlendir_tr(secili_row['Bütçe (TL)'])}</div>
                        </div>
                        <div class="sc-metric-box" style="text-align: right;">
                            <div class="sc-metric-lbl">GENEL PAY</div>
                            <div class="sc-metric-val" style="color: #60a5fa;">{yuzde_bicimlendir_tr(secili_row['Genel_%'])}</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            else:
                toplam_secili_butce = float(df[df["WBS"].isin(secili_wbsler)]["Bütçe (TL)"].sum())
                toplam_secili_pay = float(df[df["WBS"].isin(secili_wbsler)]["Genel_%"].sum())

                st.markdown(f"""
                <div class="selected-card" style="border-left-color: #f59e0b;">
                    <div class="sc-title">ÇOKLU SEÇİM</div>
                    <div class="sc-name">{len(secili_wbsler)} Öğe Seçildi</div>
                    <div class="sc-metrics">
                        <div class="sc-metric-box">
                            <div class="sc-metric-lbl">TOPLAM BÜTÇE</div>
                            <div class="sc-metric-val" style="color: #fcd34d;">{para_bicimlendir_tr(toplam_secili_butce)}</div>
                        </div>
                        <div class="sc-metric-box" style="text-align: right;">
                            <div class="sc-metric-lbl">TOPLAM PAY</div>
                            <div class="sc-metric-val" style="color: #fcd34d;">{yuzde_bicimlendir_tr(toplam_secili_pay)}</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("#### Hiyerarşi & Sıralama")
            s1, s2 = st.columns(2)
            with s1:
                if st.button("⬆️ Yukarı", use_container_width=True, disabled=not is_single):
                    yeni_df, yeni_sec = wbs_tasi_df(st.session_state.df, ilk_secili_wbs, "yukari")
                    st.session_state.df = yeni_df
                    st.session_state.secili_wbs_listesi = [yeni_sec]
                    st.session_state.editor_version += 1
                    st.rerun()
            with s2:
                if st.button("⬇️ Aşağı", use_container_width=True, disabled=not is_single):
                    yeni_df, yeni_sec = wbs_tasi_df(st.session_state.df, ilk_secili_wbs, "asagi")
                    st.session_state.df = yeni_df
                    st.session_state.secili_wbs_listesi = [yeni_sec]
                    st.session_state.editor_version += 1
                    st.rerun()
            i1, i2 = st.columns(2)
            with i1:
                if st.button("➡️ İçeri (Girinti)", use_container_width=True, disabled=not is_single):
                    yeni_df, yeni_sec = wbs_girinti_artir_df(st.session_state.df, ilk_secili_wbs)
                    st.session_state.df = yeni_df
                    st.session_state.secili_wbs_listesi = [yeni_sec]
                    st.session_state.editor_version += 1
                    st.rerun()
            with i2:
                if st.button("⬅️ Dışarı", use_container_width=True, disabled=not is_single):
                    yeni_df, yeni_sec = wbs_girinti_azalt_df(st.session_state.df, ilk_secili_wbs)
                    st.session_state.df = yeni_df
                    st.session_state.secili_wbs_listesi = [yeni_sec]
                    st.session_state.editor_version += 1
                    st.rerun()

            st.markdown("#### Ekle / Sil / Kopyala")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("➕ Alt Kırılım", use_container_width=True, disabled=not is_single):
                    yeni_wbs = sonraki_wbs_uret(ilk_secili_wbs, st.session_state.df)
                    yeni_satir = pd.DataFrame([{
                        "WBS": yeni_wbs, "Ad": "Yeni Kırılım", "Yerel_%": 0.0,
                        "Kilit": False, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-",
                    }])
                    birlesik = wbs_sirala(pd.concat([st.session_state.df, yeni_satir], ignore_index=True))
                    st.session_state.df = akilli_dengele(birlesik, birlesik)
                    st.session_state.secili_wbs_listesi = [yeni_wbs]
                    st.session_state.editor_version += 1
                    st.rerun()
            with c2:
                if st.button("↔️ Kardeş Ekle", use_container_width=True, disabled=not is_single):
                    yeni_wbs = sonraki_wbs_uret(get_dynamic_parent(st.session_state.df, ilk_secili_wbs), st.session_state.df)
                    yeni_satir = pd.DataFrame([{
                        "WBS": yeni_wbs, "Ad": "Yeni Kardeş", "Yerel_%": 0.0,
                        "Kilit": False, "İlerleme_%": 0.0, "Grup 1": "-", "Grup 2": "-", "Grup 3": "-",
                    }])
                    birlesik = wbs_sirala(pd.concat([st.session_state.df, yeni_satir], ignore_index=True))
                    st.session_state.df = akilli_dengele(birlesik, birlesik)
                    st.session_state.secili_wbs_listesi = [yeni_wbs]
                    st.session_state.editor_version += 1
                    st.rerun()

            c_kopya1, c_kopya2 = st.columns([1, 1])
            with c_kopya1:
                kopya_adet = st.number_input("Adet", min_value=1, max_value=50, value=1, disabled=not is_single)
            with c_kopya2:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                if st.button("📋 Kopyala", use_container_width=True, disabled=not is_single):
                    st.session_state.df = wbs_agaci_kopyala_df(st.session_state.df, ilk_secili_wbs, int(kopya_adet))
                    st.session_state.editor_version += 1
                    st.rerun()

            if st.button("🗑️ Seçili Kalemleri Sil", use_container_width=True, type="primary", disabled=not is_multi):
                basarili, res = wbs_sil_df(st.session_state.df, secili_wbsler)
                if not basarili:
                    st.error(res)
                else:
                    st.session_state.df = res
                    st.session_state.secili_wbs_listesi = [res.iloc[0]["WBS"]]
                    st.session_state.editor_version += 1
                    st.rerun()

            st.markdown("#### 📦 Çoklu Taşıma (Kes/Yapıştır)")
            kokler = df[df["Dyn_Parent"] == "KÖK"]["WBS"].astype(str).tolist()
            k1, k2 = st.columns(2)
            with k1:
                if st.button("✂️ Kes", use_container_width=True, disabled=not is_multi):
                    if any(k in secili_wbsler for k in kokler):
                        st.error("Ana Kök kesilemez.")
                    else:
                        st.session_state.kesilen_wbs_listesi = secili_wbsler.copy()
                        st.rerun()
            with k2:
                yapistir_disabled = len(st.session_state.kesilen_wbs_listesi) == 0 or not is_single
                if st.button("📋 İçine Yapıştır", use_container_width=True, disabled=yapistir_disabled):
                    st.session_state.df = wbs_hedefe_tasi_coklu_df(
                        st.session_state.df, st.session_state.kesilen_wbs_listesi, ilk_secili_wbs
                    )
                    st.session_state.kesilen_wbs_listesi = []
                    st.session_state.secili_wbs_listesi = [ilk_secili_wbs]
                    st.session_state.editor_version += 1
                    st.rerun()

            if st.session_state.kesilen_wbs_listesi:
                st.warning(f"📦 **{len(st.session_state.kesilen_wbs_listesi)}** öğe kesildi.")

    # =========================================================================
    # SEKME 2: SADECE ALT AKTİVİTELER İÇİN İLERLEME GİRİŞİ
    # =========================================================================
    with tab2:
        st.markdown("### ✅ Hızlı İlerleme Girişi (Sadece Çocuğu Olmayan Aktiviteler)")
        st.caption(
            "Bu bölüm yalnızca ilerleme yüzdelerinin girilmesi içindir. "
            "Alt kırılımı bulunan ana başlıklar aşağıda salt okunur özet olarak otomatik hesaplanır."
        )

        yaprak_giris_df = df[df["Is_Leaf"]][
            ["WBS", "Ağaç Görünümü", "Ad", "Genel_%", "Bütçe (TR)", "İlerleme_%", "Hakediş (TR)"]
        ].reset_index(drop=True)

        duzenlenen_yaprak = st.data_editor(
            yaprak_giris_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            height=420,
            column_config={
                "WBS": st.column_config.TextColumn("WBS", disabled=True, width="small"),
                "Ağaç Görünümü": st.column_config.TextColumn("Hiyerarşi", disabled=True, width="medium"),
                "Ad": st.column_config.TextColumn("Aktivite Adı", disabled=True),
                "Genel_%": st.column_config.NumberColumn("Genel Pay (%)", format="%.2f %%", disabled=True, width="small"),
                "Bütçe (TR)": st.column_config.TextColumn("Bütçe (₺)", disabled=True, width="small"),
                "İlerleme_%": st.column_config.NumberColumn("✏️ İlerleme (%)", format="%.2f %%", min_value=0.0, max_value=100.0, width="small"),
                "Hakediş (TR)": st.column_config.TextColumn("Hakediş (₺)", disabled=True, width="small"),
            },
            key=f"leaf_progress_editor_{st.session_state.editor_version}",
        )

        yaprak_eslesme = {
            str(duzenlenen_yaprak.iloc[i]["WBS"]): duzenlenen_yaprak.iloc[i]["İlerleme_%"]
            for i in range(len(duzenlenen_yaprak))
        }
        yeni_yaprak_df, yaprak_degisti, _ = yaprak_ilerleme_guncelle(st.session_state.df, df, yaprak_eslesme)
        if yaprak_degisti:
            st.session_state.df = yeni_yaprak_df
            st.session_state.editor_version += 1
            st.rerun()

        ana_basliklar_df = df[~df["Is_Leaf"]][
            ["WBS", "Ağaç Görünümü", "Genel_%", "Bütçe (TR)", "İlerleme_%", "Hakediş (TR)"]
        ].reset_index(drop=True)
        if not ana_basliklar_df.empty:
            st.markdown("#### 🔒 Ana Başlıklar (Salt Okunur — Alt Aktivitelerden Otomatik Hesaplanır)")
            st.dataframe(
                ana_basliklar_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "WBS": st.column_config.TextColumn("WBS", width="small"),
                    "Ağaç Görünümü": st.column_config.TextColumn("Ana Başlık", width="medium"),
                    "Genel_%": st.column_config.NumberColumn("Genel Pay (%)", format="%.2f %%", width="small"),
                    "Bütçe (TR)": st.column_config.TextColumn("Bütçe (₺)", width="small"),
                    "İlerleme_%": st.column_config.ProgressColumn("Otomatik İlerleme (%)", format="%.2f %%", min_value=0.0, max_value=100.0),
                    "Hakediş (TR)": st.column_config.TextColumn("Hakediş (₺)", width="small"),
                },
            )

    # =========================================================================
    # SEKME 3: PİVOT RAPOR ANALİZİ
    # =========================================================================
    with tab3:
        st.markdown("### 📊 Dinamik Pivot Rapor")
        p_col1, p_col2 = st.columns([1, 3])
        with p_col1:
            grup_secim = st.selectbox("Raporlama Kriterini Seçin:", GRUP_COLUMNS)
            st.caption(
                "ℹ️ Bir ana başlığa atanan grup, altındaki grupsuz (`-`) aktivitelere otomatik olarak aktarılır; "
                "böylece bütçe çift sayılmadan tüm gruplar pivota yansır."
            )
        with p_col2:
            pivot_data, ezilen_gruplar = pivot_tablo_olustur(df, grup_secim)
            pivot_gosterim = pivot_data[[
                grup_secim, "Kalem_Sayisi", "Bütçe (TR)", "Hakediş (TR)", "Gerçekleşme (%)"
            ]].copy()
            st.dataframe(
                pivot_gosterim,
                use_container_width=True,
                hide_index=True,
                column_config={
                    grup_secim: st.column_config.TextColumn(grup_secim, width="medium"),
                    "Kalem_Sayisi": st.column_config.NumberColumn("Kalem Adedi (Aktivite)", format="%d"),
                    "Bütçe (TR)": st.column_config.TextColumn("Bütçe (₺)", width="medium"),
                    "Hakediş (TR)": st.column_config.TextColumn("Hakediş (₺)", width="medium"),
                    "Gerçekleşme (%)": st.column_config.ProgressColumn(
                        "Fiziksel İlerleme (%)", format="%.2f %%", min_value=0.0, max_value=100.0
                    ),
                },
            )
            if ezilen_gruplar:
                st.info(
                    f"ℹ️ Not: Üst başlığa atanan **{', '.join(ezilen_gruplar)}** grubu, altındaki tüm aktivitelerin "
                    "kendi özel grubu bulunduğu için bütçe çift sayımını önlemek adına ayrı satır olarak eklenmedi."
                )


if not os.environ.get("PURSANTAJ_TEST_MODE"):
    render_app()
