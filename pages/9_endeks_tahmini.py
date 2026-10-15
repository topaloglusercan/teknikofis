"""
Fiyat Farkı Tahmin & Stokastik Endeks Simülatörü
Özellikler:
- DİNAMİK İŞÇİLİK (I o) SEÇİMİ (3 Basamak vs 4 Basamak):
  AltEndeks tablosundaki İşçilik (a -> I o) Temel Endeksi 3 basamaklı (< 500, ör. 100,421925)
  ise 2025=100 bazlı TÜFE serisiyle; 4 basamaklı (>= 500, ör. 3196,66) ise 2003=100 bazlı
  TÜFE serisiyle [Haziran 2026+ ayları =+(I14/H14)*H12 formülüyle türetilmiş olarak] ilerler.
- DİNAMİK YAKIT (Y o) SEÇİMİ (EPDK Motorin vs Yapım İşleri TÜİK 19):
  AltEndeks tablosundaki Yakıt (b3 -> Y o) Temel Endeksi EPDK ölçeğinde (< 500, ör. 44,717708
  veya 53,14131) ise doğrudan EPDK Aylık Ortalama Motorin verilerini (epdk_akaryakit_fiyatlari.xlsx)
  baz alarak oranlayıp ilerler; Yapım İşleri TÜİK 19 ölçeğinde (>= 500, ör. 7767,87) ise
  TÜİK 19 (Kok ve rafine petrol ürünleri) endeksiyle devam eder.
- ÇİFT KAYNAKLI WEB KAZIMA (Hakedis.org TÜİK + EPDK Motorin) ve Excel İçe Aktarma.
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import io
import os
import re
import datetime
import warnings
import requests
from bs4 import BeautifulSoup
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
    'a':  {'kolon': 'I o', 'resmi_kod': 'İn',    'kisa': 'İşçilik',         'ad': 'İşçilik (TÜFE Genel Endeksi - 3/4 Basamak Dinamik)'},
    'b1': {'kolon': 'Ç o', 'resmi_kod': 'Çn-23', 'kisa': 'Çimento/Mineral', 'ad': 'Metalik Olmayan Diğer Mineral Ürünler (23)'},
    'b2': {'kolon': 'D o', 'resmi_kod': 'Dn-24', 'kisa': 'Demir-Çelik',     'ad': 'Ana Metaller (24)'},
    'b3': {'kolon': 'Y o', 'resmi_kod': 'Ayn',   'kisa': 'Akaryakıt',       'ad': 'Akaryakıt (EPDK Motorin / TÜİK 19 Dinamik)'},
    'b4': {'kolon': 'K o', 'resmi_kod': 'Kn-16', 'kisa': 'Ağaç/Kereste',    'ad': 'Ağaç ve Mantar Ürünleri (16)'},
    'b5': {'kolon': 'G o', 'resmi_kod': 'Gn',    'kisa': 'Genel Yİ-ÜFE',    'ad': 'Genel Yurt İçi Üretici Fiyat Endeksi (Yİ-ÜFE)'},
    'c':  {'kolon': 'M o', 'resmi_kod': 'Mn-28', 'kisa': 'Makine/Ekipman',  'ad': 'Makine ve Ekipmanlar b.y.s. (28)'},
}

HEDEF_KOLONLAR = ["I o", "Ç o", "D o", "Y o", "K o", "G o", "M o"]
DEFAULT_HARITA = {'a': 'I o', 'b1': 'Ç o', 'b2': 'D o', 'b3': 'Y o', 'b4': 'K o', 'b5': 'G o', 'c': 'M o'}

# 2016 Ocak - 2026 Eylül Gömülü Çift Serili Master Kütük:
# İçerisinde hem I_2003 (4 basamaklı) hem I_2025 (3 basamaklı), hem Y_19 (Yapım İşleri) hem Y_EPDK (EPDK Motorin) yer alır.
BUILTIN_MASTER_KUTUK = json.loads('[{"AYLAR": "Ocak 2016", "I_2003": 274.44, "I_2025": 8.621434, "Ç o": 264.68, "D o": 285.53, "Y_19": 256.79, "Y_EPDK": 1.500647, "K o": 215.3, "G o": 250.67, "M o": 233.53}, {"AYLAR": "Şubat 2016", "I_2003": 274.38, "I_2025": 8.619549, "Ç o": 265.42, "D o": 282.84, "Y_19": 250.18, "Y_EPDK": 1.462019, "K o": 221.31, "G o": 250.16, "M o": 233.43}, {"AYLAR": "Mart 2016", "I_2003": 274.27, "I_2025": 8.616093, "Ç o": 267.1, "D o": 288.98, "Y_19": 273.47, "Y_EPDK": 1.598123, "K o": 221.39, "G o": 251.17, "M o": 233.35}, {"AYLAR": "Nisan 2016", "I_2003": 276.42, "I_2025": 8.683635, "Ç o": 267.54, "D o": 296.6, "Y_19": 280.11, "Y_EPDK": 1.636926, "K o": 217.94, "G o": 252.47, "M o": 233.55}, {"AYLAR": "Mayıs 2016", "I_2003": 278.02, "I_2025": 8.733898, "Ç o": 269.43, "D o": 314.19, "Y_19": 320.95, "Y_EPDK": 1.87559, "K o": 219.19, "G o": 256.21, "M o": 236.49}, {"AYLAR": "Haziran 2016", "I_2003": 279.33, "I_2025": 8.775052, "Ç o": 269.34, "D o": 311.05, "Y_19": 338.95, "Y_EPDK": 1.98078, "K o": 218.07, "G o": 257.27, "M o": 236.84}, {"AYLAR": "Temmuz 2016", "I_2003": 282.58, "I_2025": 8.877149, "Ç o": 269.7, "D o": 309.98, "Y_19": 329.85, "Y_EPDK": 1.927601, "K o": 217.82, "G o": 257.81, "M o": 239.27}, {"AYLAR": "Ağustos 2016", "I_2003": 281.76, "I_2025": 8.851389, "Ç o": 269.47, "D o": 309.45, "Y_19": 323.02, "Y_EPDK": 1.887687, "K o": 219.75, "G o": 258.01, "M o": 240.02}, {"AYLAR": "Eylül 2016", "I_2003": 282.27, "I_2025": 8.867411, "Ç o": 269.15, "D o": 306.43, "Y_19": 339.64, "Y_EPDK": 1.984812, "K o": 222.58, "G o": 258.77, "M o": 241.08}, {"AYLAR": "Ekim 2016", "I_2003": 286.33, "I_2025": 8.994954, "Ç o": 268.98, "D o": 313.19, "Y_19": 373.02, "Y_EPDK": 2.17988, "K o": 222.33, "G o": 260.94, "M o": 242.99}, {"AYLAR": "Kasım 2016", "I_2003": 287.81, "I_2025": 9.041448, "Ç o": 271.01, "D o": 344.91, "Y_19": 376.15, "Y_EPDK": 2.198171, "K o": 223.51, "G o": 266.16, "M o": 247.62}, {"AYLAR": "Aralık 2016", "I_2003": 292.54, "I_2025": 9.190039, "Ç o": 274.3, "D o": 374.67, "Y_19": 442.27, "Y_EPDK": 2.584567, "K o": 222.76, "G o": 274.09, "M o": 252.23}, {"AYLAR": "Ocak 2017", "I_2003": 299.74, "I_2025": 9.416224, "Ç o": 275.47, "D o": 403.2, "Y_19": 495.35, "Y_EPDK": 2.894759, "K o": 224.19, "G o": 284.99, "M o": 260.07}, {"AYLAR": "Şubat 2017", "I_2003": 302.17, "I_2025": 9.492562, "Ç o": 279.54, "D o": 406.89, "Y_19": 496.48, "Y_EPDK": 2.901363, "K o": 227.18, "G o": 288.59, "M o": 263.61}, {"AYLAR": "Mart 2017", "I_2003": 305.24, "I_2025": 9.589005, "Ç o": 283.92, "D o": 419.97, "Y_19": 474.66, "Y_EPDK": 2.77385, "K o": 227.33, "G o": 291.58, "M o": 265.73}, {"AYLAR": "Nisan 2017", "I_2003": 309.23, "I_2025": 9.714349, "Ç o": 286.64, "D o": 423.2, "Y_19": 476.41, "Y_EPDK": 2.784077, "K o": 230.97, "G o": 293.79, "M o": 267.01}, {"AYLAR": "Mayıs 2017", "I_2003": 310.61, "I_2025": 9.757701, "Ç o": 288.29, "D o": 419.01, "Y_19": 448.44, "Y_EPDK": 2.620624, "K o": 232.42, "G o": 295.31, "M o": 268.88}, {"AYLAR": "Haziran 2017", "I_2003": 309.78, "I_2025": 9.731627, "Ç o": 288.22, "D o": 417.34, "Y_19": 429.98, "Y_EPDK": 2.512746, "K o": 233.77, "G o": 295.52, "M o": 269.9}, {"AYLAR": "Temmuz 2017", "I_2003": 310.24, "I_2025": 9.746078, "Ç o": 290.73, "D o": 424.0, "Y_19": 439.97, "Y_EPDK": 2.571126, "K o": 238.03, "G o": 297.65, "M o": 275.66}, {"AYLAR": "Ağustos 2017", "I_2003": 311.85, "I_2025": 9.796656, "Ç o": 291.22, "D o": 441.09, "Y_19": 463.44, "Y_EPDK": 2.708282, "K o": 243.44, "G o": 300.18, "M o": 275.92}, {"AYLAR": "Eylül 2017", "I_2003": 313.88, "I_2025": 9.860427, "Ç o": 292.8, "D o": 454.67, "Y_19": 481.5, "Y_EPDK": 2.813822, "K o": 243.53, "G o": 300.9, "M o": 278.89}, {"AYLAR": "Ekim 2017", "I_2003": 320.4, "I_2025": 10.065251, "Ç o": 297.26, "D o": 470.77, "Y_19": 518.23, "Y_EPDK": 3.028467, "K o": 251.99, "G o": 306.04, "M o": 284.94}, {"AYLAR": "Kasım 2017", "I_2003": 325.18, "I_2025": 10.215413, "Ç o": 301.42, "D o": 492.76, "Y_19": 578.96, "Y_EPDK": 3.383365, "K o": 260.33, "G o": 312.21, "M o": 291.56}, {"AYLAR": "Aralık 2017", "I_2003": 327.41, "I_2025": 10.285467, "Ç o": 305.45, "D o": 508.97, "Y_19": 586.13, "Y_EPDK": 3.425266, "K o": 265.05, "G o": 316.48, "M o": 293.74}, {"AYLAR": "Ocak 2018", "I_2003": 330.75, "I_2025": 10.390392, "Ç o": 316.32, "D o": 515.63, "Y_19": 604.0, "Y_EPDK": 3.529696, "K o": 273.03, "G o": 319.6, "M o": 299.27}, {"AYLAR": "Şubat 2018", "I_2003": 333.17, "I_2025": 10.466416, "Ç o": 321.32, "D o": 523.27, "Y_19": 597.28, "Y_EPDK": 3.490425, "K o": 288.63, "G o": 328.17, "M o": 302.6}, {"AYLAR": "Mart 2018", "I_2003": 336.48, "I_2025": 10.570398, "Ç o": 325.53, "D o": 542.1, "Y_19": 610.59, "Y_EPDK": 3.568207, "K o": 294.14, "G o": 333.21, "M o": 306.54}, {"AYLAR": "Nisan 2018", "I_2003": 342.78, "I_2025": 10.76831, "Ç o": 333.86, "D o": 563.23, "Y_19": 673.8, "Y_EPDK": 3.937598, "K o": 301.33, "G o": 341.88, "M o": 312.35}, {"AYLAR": "Mayıs 2018", "I_2003": 348.34, "I_2025": 10.942976, "Ç o": 343.28, "D o": 589.48, "Y_19": 755.71, "Y_EPDK": 4.416269, "K o": 308.97, "G o": 354.85, "M o": 320.85}, {"AYLAR": "Haziran 2018", "I_2003": 357.44, "I_2025": 11.228849, "Ç o": 346.83, "D o": 618.99, "Y_19": 807.6, "Y_EPDK": 4.719507, "K o": 313.21, "G o": 365.6, "M o": 328.28}, {"AYLAR": "Temmuz 2018", "I_2003": 359.41, "I_2025": 11.290736, "Ç o": 349.88, "D o": 636.5, "Y_19": 815.14, "Y_EPDK": 4.76357, "K o": 321.04, "G o": 372.06, "M o": 338.42}, {"AYLAR": "Ağustos 2018", "I_2003": 367.66, "I_2025": 11.549907, "Ç o": 362.14, "D o": 713.3, "Y_19": 1019.82, "Y_EPDK": 5.959693, "K o": 329.52, "G o": 396.62, "M o": 359.95}, {"AYLAR": "Eylül 2018", "I_2003": 390.84, "I_2025": 12.278098, "Ç o": 381.31, "D o": 806.73, "Y_19": 1111.53, "Y_EPDK": 6.495634, "K o": 345.34, "G o": 439.78, "M o": 380.84}, {"AYLAR": "Ekim 2018", "I_2003": 401.27, "I_2025": 12.605753, "Ç o": 391.57, "D o": 773.58, "Y_19": 1090.44, "Y_EPDK": 6.372387, "K o": 370.67, "G o": 443.78, "M o": 383.95}, {"AYLAR": "Kasım 2018", "I_2003": 395.48, "I_2025": 12.423862, "Ç o": 394.22, "D o": 710.86, "Y_19": 921.55, "Y_EPDK": 5.385416, "K o": 369.63, "G o": 432.55, "M o": 371.32}, {"AYLAR": "Aralık 2018", "I_2003": 393.88, "I_2025": 12.373599, "Ç o": 392.45, "D o": 674.71, "Y_19": 800.44, "Y_EPDK": 4.677665, "K o": 359.05, "G o": 422.94, "M o": 370.77}, {"AYLAR": "Ocak 2019", "I_2003": 398.07, "I_2025": 12.505226, "Ç o": 400.1, "D o": 658.91, "Y_19": 776.81, "Y_EPDK": 4.539574, "K o": 362.93, "G o": 424.86, "M o": 381.81}, {"AYLAR": "Şubat 2019", "I_2003": 398.71, "I_2025": 12.525331, "Ç o": 403.93, "D o": 659.59, "Y_19": 814.37, "Y_EPDK": 4.75907, "K o": 364.51, "G o": 425.26, "M o": 381.69}, {"AYLAR": "Mart 2019", "I_2003": 402.81, "I_2025": 12.654131, "Ç o": 407.59, "D o": 681.05, "Y_19": 877.29, "Y_EPDK": 5.126766, "K o": 367.49, "G o": 431.98, "M o": 386.73}, {"AYLAR": "Nisan 2019", "I_2003": 409.63, "I_2025": 12.868379, "Ç o": 410.35, "D o": 713.26, "Y_19": 942.24, "Y_EPDK": 5.506325, "K o": 366.41, "G o": 444.85, "M o": 391.81}, {"AYLAR": "Mayıs 2019", "I_2003": 413.52, "I_2025": 12.990582, "Ç o": 418.94, "D o": 739.6, "Y_19": 1004.12, "Y_EPDK": 5.867943, "K o": 367.96, "G o": 456.74, "M o": 399.19}, {"AYLAR": "Haziran 2019", "I_2003": 413.63, "I_2025": 12.994038, "Ç o": 420.21, "D o": 719.65, "Y_19": 885.72, "Y_EPDK": 5.176029, "K o": 370.83, "G o": 457.16, "M o": 399.28}, {"AYLAR": "Temmuz 2019", "I_2003": 419.24, "I_2025": 13.170274, "Ç o": 420.0, "D o": 695.88, "Y_19": 886.34, "Y_EPDK": 5.179652, "K o": 372.49, "G o": 452.63, "M o": 395.18}, {"AYLAR": "Ağustos 2019", "I_2003": 422.84, "I_2025": 13.283367, "Ç o": 420.66, "D o": 682.4, "Y_19": 817.72, "Y_EPDK": 4.778646, "K o": 377.15, "G o": 449.96, "M o": 393.12}, {"AYLAR": "Eylül 2019", "I_2003": 427.04, "I_2025": 13.415308, "Ç o": 424.36, "D o": 673.37, "Y_19": 851.41, "Y_EPDK": 4.975526, "K o": 388.75, "G o": 450.55, "M o": 395.96}, {"AYLAR": "Ekim 2019", "I_2003": 435.59, "I_2025": 13.683903, "Ç o": 425.5, "D o": 658.51, "Y_19": 862.62, "Y_EPDK": 5.041036, "K o": 387.43, "G o": 451.31, "M o": 399.86}, {"AYLAR": "Kasım 2019", "I_2003": 437.25, "I_2025": 13.736052, "Ç o": 426.95, "D o": 646.65, "Y_19": 857.85, "Y_EPDK": 5.013161, "K o": 387.16, "G o": 450.97, "M o": 398.95}, {"AYLAR": "Aralık 2019", "I_2003": 440.5, "I_2025": 13.838149, "Ç o": 428.04, "D o": 656.09, "Y_19": 875.05, "Y_EPDK": 5.113676, "K o": 385.43, "G o": 454.08, "M o": 399.74}, {"AYLAR": "Ocak 2020", "I_2003": 446.45, "I_2025": 14.025066, "Ç o": 429.31, "D o": 673.29, "Y_19": 923.38, "Y_EPDK": 5.39611, "K o": 391.52, "G o": 462.42, "M o": 407.9}, {"AYLAR": "Şubat 2020", "I_2003": 448.02, "I_2025": 14.074387, "Ç o": 430.37, "D o": 688.78, "Y_19": 840.61, "Y_EPDK": 4.912413, "K o": 390.27, "G o": 464.64, "M o": 410.74}, {"AYLAR": "Mart 2020", "I_2003": 450.58, "I_2025": 14.154809, "Ç o": 430.08, "D o": 712.11, "Y_19": 651.81, "Y_EPDK": 3.809091, "K o": 385.02, "G o": 468.69, "M o": 421.12}, {"AYLAR": "Nisan 2020", "I_2003": 454.43, "I_2025": 14.275755, "Ç o": 430.03, "D o": 739.58, "Y_19": 466.31, "Y_EPDK": 2.725054, "K o": 386.68, "G o": 474.69, "M o": 431.26}, {"AYLAR": "Mayıs 2020", "I_2003": 460.62, "I_2025": 14.470212, "Ç o": 431.95, "D o": 749.43, "Y_19": 489.31, "Y_EPDK": 2.859463, "K o": 388.88, "G o": 482.02, "M o": 437.16}, {"AYLAR": "Haziran 2020", "I_2003": 465.84, "I_2025": 14.634196, "Ç o": 435.2, "D o": 741.78, "Y_19": 635.58, "Y_EPDK": 3.714245, "K o": 392.62, "G o": 485.37, "M o": 442.55}, {"AYLAR": "Temmuz 2020", "I_2003": 468.56, "I_2025": 14.719644, "Ç o": 439.42, "D o": 748.27, "Y_19": 710.02, "Y_EPDK": 4.149263, "K o": 391.5, "G o": 490.33, "M o": 447.88}, {"AYLAR": "Ağustos 2020", "I_2003": 472.61, "I_2025": 14.846873, "Ç o": 445.43, "D o": 801.59, "Y_19": 755.19, "Y_EPDK": 4.41323, "K o": 397.14, "G o": 501.85, "M o": 465.14}, {"AYLAR": "Eylül 2020", "I_2003": 477.21, "I_2025": 14.991381, "Ç o": 453.58, "D o": 847.59, "Y_19": 742.69, "Y_EPDK": 4.340182, "K o": 404.65, "G o": 515.13, "M o": 476.43}, {"AYLAR": "Ekim 2020", "I_2003": 487.38, "I_2025": 15.310868, "Ç o": 463.6, "D o": 896.36, "Y_19": 786.31, "Y_EPDK": 4.595091, "K o": 412.44, "G o": 533.44, "M o": 489.7}, {"AYLAR": "Kasım 2020", "I_2003": 498.58, "I_2025": 15.662711, "Ç o": 479.11, "D o": 962.99, "Y_19": 828.13, "Y_EPDK": 4.839481, "K o": 421.94, "G o": 555.18, "M o": 500.14}, {"AYLAR": "Aralık 2020", "I_2003": 504.81, "I_2025": 15.858425, "Ç o": 494.39, "D o": 1025.39, "Y_19": 916.84, "Y_EPDK": 5.35789, "K o": 427.23, "G o": 568.27, "M o": 505.58}, {"AYLAR": "Ocak 2021", "I_2003": 513.3, "I_2025": 16.125135, "Ç o": 518.33, "D o": 1056.08, "Y_19": 954.83, "Y_EPDK": 5.579898, "K o": 441.3, "G o": 583.38, "M o": 515.79}, {"AYLAR": "Şubat 2021", "I_2003": 517.96, "I_2025": 16.271527, "Ç o": 531.99, "D o": 1051.19, "Y_19": 1007.1, "Y_EPDK": 5.589593, "K o": 468.21, "G o": 590.52, "M o": 514.85}, {"AYLAR": "Mart 2021", "I_2003": 523.53, "I_2025": 16.446507, "Ç o": 548.42, "D o": 1146.37, "Y_19": 1148.4, "Y_EPDK": 5.532763, "K o": 479.13, "G o": 614.93, "M o": 526.25}, {"AYLAR": "Nisan 2021", "I_2003": 532.32, "I_2025": 16.722641, "Ç o": 564.82, "D o": 1269.94, "Y_19": 1210.23, "Y_EPDK": 5.449153, "K o": 499.54, "G o": 641.63, "M o": 539.87}, {"AYLAR": "Mayıs 2021", "I_2003": 537.05, "I_2025": 16.871233, "Ç o": 579.1, "D o": 1383.86, "Y_19": 1306.89, "Y_EPDK": 5.70272, "K o": 541.54, "G o": 666.79, "M o": 552.91}, {"AYLAR": "Haziran 2021", "I_2003": 547.48, "I_2025": 17.198887, "Ç o": 596.53, "D o": 1483.84, "Y_19": 1422.44, "Y_EPDK": 6.106068, "K o": 585.49, "G o": 693.54, "M o": 566.53}, {"AYLAR": "Temmuz 2021", "I_2003": 557.36, "I_2025": 17.509264, "Ç o": 611.85, "D o": 1560.13, "Y_19": 1497.97, "Y_EPDK": 6.116686, "K o": 624.04, "G o": 710.61, "M o": 568.87}, {"AYLAR": "Ağustos 2021", "I_2003": 563.6, "I_2025": 17.705291, "Ç o": 636.11, "D o": 1605.35, "Y_19": 1451.39, "Y_EPDK": 6.141331, "K o": 665.32, "G o": 730.28, "M o": 572.13}, {"AYLAR": "Eylül 2021", "I_2003": 570.66, "I_2025": 17.927079, "Ç o": 648.91, "D o": 1625.09, "Y_19": 1500.31, "Y_EPDK": 6.151144, "K o": 681.94, "G o": 741.58, "M o": 579.45}, {"AYLAR": "Ekim 2021", "I_2003": 584.32, "I_2025": 18.356203, "Ç o": 674.11, "D o": 1714.04, "Y_19": 1813.13, "Y_EPDK": 6.577627, "K o": 699.82, "G o": 780.45, "M o": 602.76}, {"AYLAR": "Kasım 2021", "I_2003": 604.84, "I_2025": 19.000831, "Ç o": 726.07, "D o": 1898.32, "Y_19": 2036.73, "Y_EPDK": 7.260619, "K o": 754.53, "G o": 858.43, "M o": 647.61}, {"AYLAR": "Aralık 2021", "I_2003": 686.95, "I_2025": 21.580287, "Ç o": 827.31, "D o": 2367.77, "Y_19": 2400.99, "Y_EPDK": 8.968381, "K o": 810.24, "G o": 1022.25, "M o": 737.28}, {"AYLAR": "Ocak 2022", "I_2003": 763.23, "I_2025": 23.976596, "Ç o": 998.57, "D o": 2495.39, "Y_19": 2680.66, "Y_EPDK": 11.624381, "K o": 877.87, "G o": 1129.03, "M o": 816.45}, {"AYLAR": "Şubat 2022", "I_2003": 799.93, "I_2025": 25.129513, "Ç o": 1116.24, "D o": 2568.25, "Y_19": 3010.0, "Y_EPDK": 13.021195, "K o": 960.44, "G o": 1210.6, "M o": 845.24}, {"AYLAR": "Mart 2022", "I_2003": 843.64, "I_2025": 26.502647, "Ç o": 1276.94, "D o": 2878.66, "Y_19": 4093.37, "Y_EPDK": 17.8095, "K o": 1018.77, "G o": 1321.9, "M o": 879.21}, {"AYLAR": "Nisan 2022", "I_2003": 904.79, "I_2025": 28.423653, "Ç o": 1463.59, "D o": 3069.68, "Y_19": 4230.39, "Y_EPDK": 18.189407, "K o": 1110.91, "G o": 1423.27, "M o": 902.82}, {"AYLAR": "Mayıs 2022", "I_2003": 931.76, "I_2025": 29.270905, "Ç o": 1587.9, "D o": 3183.54, "Y_19": 4522.98, "Y_EPDK": 18.950746, "K o": 1194.39, "G o": 1548.01, "M o": 945.82}, {"AYLAR": "Haziran 2022", "I_2003": 977.9, "I_2025": 30.720377, "Ç o": 1691.83, "D o": 3279.12, "Y_19": 5651.05, "Y_EPDK": 23.85522, "K o": 1274.91, "G o": 1652.75, "M o": 1006.95}, {"AYLAR": "Temmuz 2022", "I_2003": 1001.03, "I_2025": 31.446998, "Ç o": 1800.13, "D o": 3190.01, "Y_19": 5144.38, "Y_EPDK": 21.561898, "K o": 1318.77, "G o": 1738.21, "M o": 1045.22}, {"AYLAR": "Ağustos 2022", "I_2003": 1015.65, "I_2025": 31.90628, "Ç o": 1898.99, "D o": 3108.9, "Y_19": 4846.79, "Y_EPDK": 20.773246, "K o": 1373.36, "G o": 1780.05, "M o": 1080.3}, {"AYLAR": "Eylül 2022", "I_2003": 1046.89, "I_2025": 32.887673, "Ç o": 1992.85, "D o": 3132.95, "Y_19": 4709.49, "Y_EPDK": 20.596678, "K o": 1400.14, "G o": 1865.09, "M o": 1103.61}, {"AYLAR": "Ekim 2022", "I_2003": 1084.0, "I_2025": 34.05347, "Ç o": 2117.53, "D o": 3143.79, "Y_19": 5004.57, "Y_EPDK": 22.266034, "K o": 1430.51, "G o": 2011.13, "M o": 1148.3}, {"AYLAR": "Kasım 2022", "I_2003": 1115.26, "I_2025": 35.035492, "Ç o": 2202.64, "D o": 3121.86, "Y_19": 4903.7, "Y_EPDK": 21.355924, "K o": 1430.62, "G o": 2026.08, "M o": 1180.23}, {"AYLAR": "Aralık 2022", "I_2003": 1128.45, "I_2025": 35.449851, "Ç o": 2260.17, "D o": 3151.41, "Y_19": 4291.19, "Y_EPDK": 18.673424, "K o": 1430.96, "G o": 2021.19, "M o": 1204.94}, {"AYLAR": "Ocak 2023", "I_2003": 1203.48, "I_2025": 37.806892, "Ç o": 2447.42, "D o": 3218.34, "Y_19": 4448.04, "Y_EPDK": 19.164263, "K o": 1438.91, "G o": 2105.17, "M o": 1305.39}, {"AYLAR": "Şubat 2023", "I_2003": 1241.33, "I_2025": 38.995936, "Ç o": 2553.12, "D o": 3323.44, "Y_19": 4383.03, "Y_EPDK": 17.783441, "K o": 1471.93, "G o": 2138.04, "M o": 1333.49}, {"AYLAR": "Mart 2023", "I_2003": 1269.75, "I_2025": 39.88874, "Ç o": 2603.05, "D o": 3426.85, "Y_19": 4174.86, "Y_EPDK": 17.603373, "K o": 1527.59, "G o": 2147.44, "M o": 1357.76}, {"AYLAR": "Nisan 2023", "I_2003": 1300.04, "I_2025": 40.840289, "Ç o": 2636.27, "D o": 3506.86, "Y_19": 4186.93, "Y_EPDK": 17.340466, "K o": 1558.4, "G o": 2164.94, "M o": 1395.43}, {"AYLAR": "Mayıs 2023", "I_2003": 1300.6, "I_2025": 40.857881, "Ç o": 2653.73, "D o": 3515.08, "Y_19": 3827.0, "Y_EPDK": 16.035161, "K o": 1616.41, "G o": 2179.02, "M o": 1424.63}, {"AYLAR": "Haziran 2023", "I_2003": 1351.59, "I_2025": 42.459714, "Ç o": 2694.63, "D o": 3823.46, "Y_19": 4521.68, "Y_EPDK": 18.737144, "K o": 1696.66, "G o": 2320.72, "M o": 1526.03}, {"AYLAR": "Temmuz 2023", "I_2003": 1479.84, "I_2025": 46.488642, "Ç o": 2855.08, "D o": 4173.12, "Y_19": 5436.98, "Y_EPDK": 24.571817, "K o": 1767.03, "G o": 2511.75, "M o": 1693.95}, {"AYLAR": "Ağustos 2023", "I_2003": 1614.31, "I_2025": 50.712968, "Ç o": 3031.76, "D o": 4231.07, "Y_19": 6466.19, "Y_EPDK": 31.141767, "K o": 1864.37, "G o": 2659.6, "M o": 1762.8}, {"AYLAR": "Eylül 2023", "I_2003": 1691.04, "I_2025": 53.123414, "Ç o": 3166.62, "D o": 4226.72, "Y_19": 6854.79, "Y_EPDK": 32.651908, "K o": 1960.01, "G o": 2749.98, "M o": 1825.57}, {"AYLAR": "Ekim 2023", "I_2003": 1749.11, "I_2025": 54.947662, "Ç o": 3291.36, "D o": 4287.22, "Y_19": 6635.5, "Y_EPDK": 32.422008, "K o": 2041.36, "G o": 2803.29, "M o": 1869.2}, {"AYLAR": "Kasım 2023", "I_2003": 1806.5, "I_2025": 56.750548, "Ç o": 3355.34, "D o": 4376.36, "Y_19": 6401.68, "Y_EPDK": 31.399308, "K o": 2129.65, "G o": 2882.04, "M o": 1925.23}, {"AYLAR": "Aralık 2023", "I_2003": 1859.38, "I_2025": 58.411754, "Ç o": 3382.39, "D o": 4517.42, "Y_19": 6169.45, "Y_EPDK": 30.35225, "K o": 2208.32, "G o": 2915.02, "M o": 1968.58}, {"AYLAR": "Ocak 2024", "I_2003": 1984.02, "I_2025": 62.327275, "Ç o": 3490.15, "D o": 4760.87, "Y_19": 6343.72, "Y_EPDK": 32.878933, "K o": 2279.58, "G o": 3035.59, "M o": 2151.06}, {"AYLAR": "Şubat 2024", "I_2003": 2073.88, "I_2025": 65.150195, "Ç o": 3626.74, "D o": 4998.52, "Y_19": 6998.27, "Y_EPDK": 35.441908, "K o": 2394.26, "G o": 3149.03, "M o": 2212.02}, {"AYLAR": "Mart 2024", "I_2003": 2139.47, "I_2025": 67.210681, "Ç o": 3707.02, "D o": 5195.92, "Y_19": 6926.72, "Y_EPDK": 34.517708, "K o": 2417.47, "G o": 3252.79, "M o": 2274.78}, {"AYLAR": "Nisan 2024", "I_2003": 2207.5, "I_2025": 69.347819, "Ç o": 3928.56, "D o": 5265.54, "Y_19": 7047.27, "Y_EPDK": 34.878475, "K o": 2500.49, "G o": 3369.98, "M o": 2309.67}, {"AYLAR": "Mayıs 2024", "I_2003": 2281.85, "I_2025": 71.683497, "Ç o": 4086.13, "D o": 5199.87, "Y_19": 6683.76, "Y_EPDK": 33.690558, "K o": 2518.42, "G o": 3435.96, "M o": 2356.09}, {"AYLAR": "Haziran 2024", "I_2003": 2319.29, "I_2025": 72.859662, "Ç o": 4134.14, "D o": 5157.5, "Y_19": 6636.54, "Y_EPDK": 33.891525, "K o": 2576.92, "G o": 3483.25, "M o": 2393.21}, {"AYLAR": "Temmuz 2024", "I_2003": 2394.1, "I_2025": 75.209791, "Ç o": 4241.94, "D o": 5170.44, "Y_19": 7017.65, "Y_EPDK": 36.96005, "K o": 2609.4, "G o": 3550.88, "M o": 2436.29}, {"AYLAR": "Ağustos 2024", "I_2003": 2453.34, "I_2025": 77.070794, "Ç o": 4313.09, "D o": 5184.47, "Y_19": 6704.05, "Y_EPDK": 36.332092, "K o": 2702.43, "G o": 3610.51, "M o": 2525.85}, {"AYLAR": "Eylül 2024", "I_2003": 2526.16, "I_2025": 79.358408, "Ç o": 4317.68, "D o": 5279.37, "Y_19": 6334.84, "Y_EPDK": 35.296525, "K o": 2712.29, "G o": 3659.84, "M o": 2558.99}, {"AYLAR": "Ekim 2024", "I_2003": 2598.91, "I_2025": 81.643823, "Ç o": 4357.79, "D o": 5402.22, "Y_19": 6509.93, "Y_EPDK": 35.797808, "K o": 2761.78, "G o": 3707.1, "M o": 2572.66}, {"AYLAR": "Kasım 2024", "I_2003": 2657.23, "I_2025": 83.475925, "Ç o": 4382.41, "D o": 5420.45, "Y_19": 6482.77, "Y_EPDK": 36.000033, "K o": 2785.67, "G o": 3731.43, "M o": 2592.48}, {"AYLAR": "Aralık 2024", "I_2003": 2684.55, "I_2025": 84.334173, "Ç o": 4443.39, "D o": 5430.01, "Y_19": 6525.72, "Y_EPDK": 36.405508, "K o": 2816.9, "G o": 3746.52, "M o": 2626.84}, {"AYLAR": "Ocak 2025", "I_2003": 2819.65, "I_2025": 88.578291, "Ç o": 4537.23, "D o": 5406.99, "Y_19": 6911.75, "Y_EPDK": 38.578025, "K o": 2903.1, "G o": 3861.33, "M o": 2711.42}, {"AYLAR": "Şubat 2025", "I_2003": 2883.75, "I_2025": 90.59197, "Ç o": 4641.23, "D o": 5462.92, "Y_19": 7114.2, "Y_EPDK": 38.97835, "K o": 2936.95, "G o": 3943.01, "M o": 2808.87}, {"AYLAR": "Mart 2025", "I_2003": 2954.69, "I_2025": 92.820524, "Ç o": 4700.89, "D o": 5566.78, "Y_19": 6895.61, "Y_EPDK": 38.267208, "K o": 3052.11, "G o": 4017.3, "M o": 2880.53}, {"AYLAR": "Nisan 2025", "I_2003": 3043.23, "I_2025": 95.601977, "Ç o": 4743.75, "D o": 5750.9, "Y_19": 6760.36, "Y_EPDK": 37.7174, "K o": 3183.49, "G o": 4128.19, "M o": 2997.5}, {"AYLAR": "Mayıs 2025", "I_2003": 3089.74, "I_2025": 97.063071, "Ç o": 4808.92, "D o": 5882.18, "Y_19": 6752.22, "Y_EPDK": 37.880208, "K o": 3280.13, "G o": 4230.69, "M o": 3068.3}, {"AYLAR": "Haziran 2025", "I_2003": 3132.17, "I_2025": 98.395995, "Ç o": 4910.55, "D o": 5986.06, "Y_19": 7297.64, "Y_EPDK": 40.673817, "K o": 3375.09, "G o": 4334.94, "M o": 3133.65}, {"AYLAR": "Temmuz 2025", "I_2003": 3196.66, "I_2025": 100.421925, "Ç o": 4972.28, "D o": 6034.24, "Y_19": 7767.87, "Y_EPDK": 44.284425, "K o": 3481.27, "G o": 4409.73, "M o": 3218.0}, {"AYLAR": "Ağustos 2025", "I_2003": 3261.72, "I_2025": 102.465761, "Ç o": 5053.99, "D o": 6100.24, "Y_19": 7625.39, "Y_EPDK": 43.66425, "K o": 3530.4, "G o": 4518.89, "M o": 3259.16}, {"AYLAR": "Eylül 2025", "I_2003": 3367.22, "I_2025": 105.780006, "Ç o": 5110.09, "D o": 6187.73, "Y_19": 7854.35, "Y_EPDK": 45.1109, "K o": 3572.84, "G o": 4632.89, "M o": 3319.76}, {"AYLAR": "Ekim 2025", "I_2003": 3453.09, "I_2025": 108.477581, "Ç o": 5170.07, "D o": 6350.02, "Y_19": 7797.08, "Y_EPDK": 44.919783, "K o": 3577.46, "G o": 4708.2, "M o": 3368.81}, {"AYLAR": "Kasım 2025", "I_2003": 3482.96, "I_2025": 109.415936, "Ç o": 5216.99, "D o": 6453.42, "Y_19": 8398.17, "Y_EPDK": 47.560158, "K o": 3606.93, "G o": 4747.63, "M o": 3419.55}, {"AYLAR": "Aralık 2025", "I_2003": 3513.87, "I_2025": 110.386963, "Ç o": 5277.72, "D o": 6568.1, "Y_19": 7739.67, "Y_EPDK": 44.895592, "K o": 3667.48, "G o": 4783.04, "M o": 3475.04}, {"AYLAR": "Ocak 2026", "I_2003": 3683.83, "I_2025": 115.73, "Ç o": 5454.97, "D o": 6710.15, "Y_19": 7753.07, "Y_EPDK": 46.162817, "K o": 3786.33, "G o": 4910.53, "M o": 3624.85}, {"AYLAR": "Şubat 2026", "I_2003": 3793.05, "I_2025": 119.16, "Ç o": 5591.46, "D o": 6805.63, "Y_19": 8398.65, "Y_EPDK": 48.474967, "K o": 3895.97, "G o": 5029.76, "M o": 3708.22}, {"AYLAR": "Mart 2026", "I_2003": 3866.74, "I_2025": 121.47, "Ç o": 5683.71, "D o": 6960.9, "Y_19": 12354.52, "Y_EPDK": 56.330433, "K o": 3968.37, "G o": 5145.36, "M o": 3726.74}, {"AYLAR": "Nisan 2026", "I_2003": 4028.47, "I_2025": 126.55, "Ç o": 5841.56, "D o": 7209.79, "Y_19": 14725.17, "Y_EPDK": 61.81035, "K o": 4068.8, "G o": 5308.46, "M o": 3813.11}, {"AYLAR": "Mayıs 2026", "I_2003": 4097.55, "I_2025": 128.72, "Ç o": 5972.29, "D o": 7364.5, "Y_19": 14457.03, "Y_EPDK": 56.9463, "K o": 4233.67, "G o": 5454.58, "M o": 3857.2}, {"AYLAR": "Haziran 2026", "I_2003": 4137.98, "I_2025": 129.99, "Ç o": 6147.33, "D o": 7591.1, "Y_19": 12632.34, "Y_EPDK": 54.514233, "K o": 4325.13, "G o": 5552.52, "M o": 3907.73}, {"AYLAR": "Temmuz 2026", "I_2003": 4211.83, "I_2025": 132.31, "Ç o": 6227.36, "D o": 7700.25, "Y_19": 12912.92, "Y_EPDK": 59.506942, "K o": 4431.02, "G o": 5637.14, "M o": 3948.51}, {"AYLAR": "Ağustos 2026", "I_2003": 4289.5, "I_2025": 134.75, "Ç o": 6350.43, "D o": 7849.8, "Y_19": 15032.12, "Y_EPDK": 66.47305, "K o": 4458.92, "G o": 5781.74, "M o": 4036.91}, {"AYLAR": "Eylül 2026", "I_2003": 4368.45, "I_2025": 137.23, "Ç o": 6414.65, "D o": 8039.84, "Y_19": 17024.32, "Y_EPDK": 76.2859, "K o": 4523.21, "G o": 5901.26, "M o": 4089.47}]')

AYLAR_LIST = ['', 'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık']
MONTHS_MAP = {
    'oca': '01', 'ocak': '01', 'şub': '02', 'sub': '02', 'şubat': '02', 'subat': '02',
    'mar': '03', 'mart': '03', 'nis': '04', 'nisan': '04', 'may': '05', 'mayıs': '05', 'mayis': '05',
    'haz': '06', 'haziran': '06', 'tem': '07', 'temmuz': '07', 'ağu': '08', 'agu': '08', 'ağustos': '08', 'agustos': '08',
    'eyl': '09', 'eylül': '09', 'eylul': '09', 'eki': '10', 'ekim': '10', 'kas': '11', 'kasım': '11', 'kasim': '11',
    'ara': '12', 'aralık': '12', 'aralik': '12'
}

def parse_turkish_date(date_str):
    if pd.isna(date_str) or str(date_str).strip() == '':
        return pd.NaT
    s = str(date_str).strip().replace('.', ' ').replace('\n', ' ').replace('\r', ' ').lower()
    if s in ['none', 'nan', 'nat', '<na>']:
        return pd.NaT
    try:
        return pd.to_datetime(s).strftime('%Y-%m')
    except Exception:
        pass
    parts = [p for p in s.split() if p]
    if len(parts) >= 2:
        if parts[0] in MONTHS_MAP and parts[1].isdigit():
            return f"{parts[1] if len(parts[1])==4 else '20'+parts[1]}-{MONTHS_MAP[parts[0]]}"
        if parts[0].isdigit() and parts[1] in MONTHS_MAP:
            return f"{parts[0] if len(parts[0])==4 else '20'+parts[0]}-{MONTHS_MAP[parts[1]]}"
    return pd.NaT

def period_to_tr_full_str(period_val, short_year=False):
    yr = str(period_val.year)[-2:] if short_year else str(period_val.year)
    return f"{AYLAR_LIST[period_val.month]} {yr}"

def clean_decimal(val):
    if val is None or pd.isna(val):
        return Decimal('0.0')
    if isinstance(val, Decimal):
        return Decimal('0.0') if val.is_nan() else val
    if isinstance(val, (int, float, np.number)):
        if np.isnan(val):
            return Decimal('0.0')
        return Decimal(str(val))
    s = str(val).strip().replace('TL', '').replace('₺', '').replace('%', '').strip()
    if s.lower() in ['', 'none', 'nan', 'nat', '<na>']:
        return Decimal('0.0')
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    elif s.count('.') > 1:
        s = s.replace('.', '')
    try:
        d = Decimal(s)
        return Decimal('0.0') if d.is_nan() else d
    except Exception:
        return Decimal('0.0')

def tr_format(val, decimals=2):
    if pd.isna(val) or val == "":
        return ""
    try:
        fmt = f"{{:,.{decimals}f}}"
        return fmt.format(float(val)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)

def filter_empty_rows(df):
    if df is None or df.empty:
        return pd.DataFrame()
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
                except Exception:
                    pass
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
                    except Exception:
                        pass
            new_vals.append(s)
        df_clean[col] = new_vals
    return df_clean.astype(str).astype("string")

def ensure_text_df(df):
    if df is None or df.empty:
        return df
    return clean_df_for_ui(df)

def get_text_config(df):
    return {col: st.column_config.TextColumn(col, width="medium") for col in df.columns}

def ayristir_gercek_endeks(df_end):
    """Tahmin satırları eklenmiş bir endeks tablosundan yalnızca gerçekleşen satırları döndürür."""
    df_c = filter_empty_rows(df_end.copy())
    if df_c.empty:
        return ensure_text_df(df_end)
    end_col = 'AYLAR' if 'AYLAR' in df_c.columns else df_c.columns[0]
    num_cols = [c for c in df_c.columns if c != end_col]
    gecerli_satirlar = []
    for idx, row in df_c.iterrows():
        vals = [float(clean_decimal(row[c])) for c in num_cols]
        if any(v > 0 for v in vals):
            gecerli_satirlar.append(idx)
    if gecerli_satirlar:
        df_c = df_c.loc[gecerli_satirlar].reset_index(drop=True)

    try:
        if 'master_db' in st.session_state and not st.session_state.master_db.empty:
            db_df = st.session_state.master_db
            db_col = 'AYLAR' if 'AYLAR' in db_df.columns else db_df.columns[0]
            max_db_per = pd.to_datetime(db_df[db_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M').max()
            if pd.notna(max_db_per):
                c_pers = pd.to_datetime(df_c[end_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M')
                mask_real = c_pers <= max_db_per
                if mask_real.sum() >= 2:
                    df_c = df_c.loc[mask_real].reset_index(drop=True)
    except Exception:
        pass

    return ensure_text_df(df_c)

# ==========================================
# 2. WEB SCRAPING (TÜİK + EPDK) VE ÇİFT SERİLİ MASTER VERİTABANI
# ==========================================
MASTER_DB_FILE = "tuik_epdk_master_kutuk_v4.json"
LEGACY_TUIK_EXCEL = "endeksler_2016_2026.xlsx"
LEGACY_EPDK_EXCEL = "epdk_akaryakit_fiyatlari.xlsx"
URL_TUIK = "https://www.hakedis.org/endeksler/yapim-isleri-fiyat-farki-endeksleri-kasim-2013-sonrasi"
URL_EPDK = "https://www.hakedis.org/endeksler/epdk-akaryakit-bayi-satis-fiyatlari-aylik-ortalama"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
}
COLUMNS_RAW = ["ENDEKS KODU", "ENDEKS ADI", "OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN", "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]
COLUMNS_EPDK = ["AKARYAKIT TÜRÜ", "OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN", "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]

def clean_text(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()

def find_table_year(table_tag) -> int | None:
    prev_element = table_tag.find_previous(
        lambda tag: tag.name in ["h1", "h2", "h3", "h4", "h5", "h6", "div", "p", "strong", "b"]
        and re.search(r"\b(20\d{2})\s*Yılı", tag.get_text(), re.IGNORECASE)
    )
    if prev_element:
        match = re.search(r"\b(20\d{2})\b", prev_element.get_text())
        if match:
            return int(match.group(1))
    prev_text = table_tag.find_previous(string=re.compile(r"(20\d{2})\s*Yılı", re.IGNORECASE))
    if prev_text:
        match = re.search(r"\b(20\d{2})\b", prev_text)
        if match:
            return int(match.group(1))
    return None

def fetch_and_parse_indices() -> pd.DataFrame:
    """Hakedis.org Yapım İşleri TÜİK Endekslerini çeker."""
    start_year = 2016
    end_year = datetime.datetime.now().year + 5
    response = requests.get(URL_TUIK, headers=HEADERS, timeout=30)
    response.raise_for_status()
    response.encoding = response.apparent_encoding or "utf-8"
    soup = BeautifulSoup(response.text, "lxml")
    tables = soup.find_all("table")

    year_dfs = {}
    for table in tables:
        year = find_table_year(table)
        if not year or not (start_year <= year <= end_year) or year in year_dfs:
            continue
        rows = table.find_all("tr")
        data = []
        for tr in rows:
            cells = [clean_text(c.get_text()) for c in tr.find_all(["td", "th"])]
            if not cells or all(v == "" for v in cells):
                continue
            if any(h in " ".join(cells).upper() for h in ["ENDEKS KODU", "ENDEKS ADI", "OCAK"]):
                continue
            if len(cells) < len(COLUMNS_RAW):
                cells.extend([""] * (len(COLUMNS_RAW) - len(cells)))
            elif len(cells) > len(COLUMNS_RAW):
                cells = cells[:len(COLUMNS_RAW)]
            data.append(cells)
        if data:
            df = pd.DataFrame(data, columns=COLUMNS_RAW)
            df.insert(0, "YIL", year)
            year_dfs[year] = df

    sorted_years = sorted(year_dfs.keys())
    if not sorted_years:
        return pd.DataFrame()
    return pd.concat([year_dfs[y] for y in sorted_years], ignore_index=True)

def fetch_and_parse_epdk() -> pd.DataFrame:
    """Hakedis.org EPDK Aylık Ortalama Akaryakıt Fiyatlarını çeker."""
    start_year = 2016
    end_year = datetime.datetime.now().year + 5
    response = requests.get(URL_EPDK, headers=HEADERS, timeout=30)
    response.raise_for_status()
    response.encoding = response.apparent_encoding or "utf-8"
    soup = BeautifulSoup(response.text, "lxml")
    tables = soup.find_all("table")

    year_dfs = {}
    for table in tables:
        year = find_table_year(table)
        if not year or not (start_year <= year <= end_year) or year in year_dfs:
            continue
        rows = table.find_all("tr")
        data = []
        for tr in rows:
            cells = [clean_text(c.get_text()) for c in tr.find_all(["td", "th"])]
            if not cells or all(v == "" for v in cells):
                continue
            if any(h in " ".join(cells).upper() for h in ["AKARYAKIT TÜRÜ", "OCAK"]):
                continue
            if len(cells) < len(COLUMNS_EPDK):
                cells.extend([""] * (len(COLUMNS_EPDK) - len(cells)))
            elif len(cells) > len(COLUMNS_EPDK):
                cells = cells[:len(COLUMNS_EPDK)]
            data.append(cells)
        if data:
            df = pd.DataFrame(data, columns=COLUMNS_EPDK)
            df.insert(0, "YIL", year)
            year_dfs[year] = df

    sorted_years = sorted(year_dfs.keys())
    if not sorted_years:
        return pd.DataFrame()
    return pd.concat([year_dfs[y] for y in sorted_years], ignore_index=True)

def parse_float_safe(val):
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if s == "" or s.lower() in ['nan', 'none', 'nat']:
        return None
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        v = float(s)
        return v if v > 0 else None
    except Exception:
        return None

def build_master_db_from_dfs(tuik_raw_df, epdk_raw_df=None):
    """
    TÜİK ham tablosu ve EPDK ham tablosunu birleştirerek 129+ aylık Master Kütük üretir.
    Her satırda hem I_2003 (4 basamaklı) hem I_2025 (3 basamaklı),
    hem Y_19 (Yapım İşleri) hem Y_EPDK (EPDK Motorin) saklanır.
    """
    if tuik_raw_df is None or tuik_raw_df.empty:
        return pd.DataFrame(BUILTIN_MASTER_KUTUK)

    months = ['OCAK', 'ŞUBAT', 'MART', 'NİSAN', 'MAYIS', 'HAZİRAN', 'TEMMUZ', 'AĞUSTOS', 'EYLÜL', 'EKİM', 'KASIM', 'ARALIK']
    tr_months = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık']

    # Önce gömülü EPDK verilerini haritaya yükle, üstüne yeni epdk_raw_df varsa yaz
    epdk_map = {r["AYLAR"]: float(r["Y_EPDK"]) for r in BUILTIN_MASTER_KUTUK if r.get("Y_EPDK") is not None}
    if epdk_raw_df is not None and not epdk_raw_df.empty:
        turu_col = "AKARYAKIT TÜRÜ" if "AKARYAKIT TÜRÜ" in epdk_raw_df.columns else epdk_raw_df.columns[1]
        mot_rows = epdk_raw_df[epdk_raw_df[turu_col].astype(str).str.strip() == "Motorin"]
        if mot_rows.empty:
            mot_rows = epdk_raw_df[epdk_raw_df[turu_col].astype(str).str.contains("Motorin", case=False, na=False)]
        for _, r in mot_rows.iterrows():
            y = int(r["YIL"])
            for idx_m, m in enumerate(months):
                fv = parse_float_safe(r.get(m))
                if fv is not None:
                    # 10 Temmuz 2023 (7346 sayılı CB Kararı) öncesi KDV %18 (/1.18), Temmuz 2023 ve sonrası KDV %20 (/1.20)
                    kdv_oran = 1.18 if (y < 2023 or (y == 2023 and idx_m < 6)) else 1.20
                    epdk_map[f"{tr_months[idx_m]} {y}"] = round(fv / kdv_oran, 6)

    years = sorted(tuik_raw_df['YIL'].astype(int).unique())
    records = []
    prev_I03 = None
    prev_I25 = None
    prev_vals = {"Ç o": None, "D o": None, "Y_19": None, "K o": None, "G o": None, "M o": None}

    for year in years:
        df_year = tuik_raw_df[tuik_raw_df['YIL'].astype(int) == year]

        def get_row(code, substr=None):
            rows = df_year[df_year['ENDEKS KODU'].astype(str).str.strip() == code]
            if substr is not None:
                r_sub = rows[rows['ENDEKS ADI'].astype(str).str.contains(substr, na=False)]
                return r_sub.iloc[0] if not r_sub.empty else None
            return rows.iloc[0] if not rows.empty else None

        row_G = get_row('Yİ-ÜFE')
        row_I03 = get_row('I', '2003=100')
        row_I25 = get_row('I', '2025=100')
        row_C = get_row('23')
        row_D = get_row('24')
        row_Y = get_row('19')
        row_K = get_row('16')
        row_M = get_row('28')

        for i, m in enumerate(months):
            vG = parse_float_safe(row_G.get(m)) if row_G is not None else None
            vI03 = parse_float_safe(row_I03.get(m)) if row_I03 is not None else None
            vI25 = parse_float_safe(row_I25.get(m)) if row_I25 is not None else None
            vC = parse_float_safe(row_C.get(m)) if row_C is not None else None
            vD = parse_float_safe(row_D.get(m)) if row_D is not None else None
            vY19 = parse_float_safe(row_Y.get(m)) if row_Y is not None else None
            vK = parse_float_safe(row_K.get(m)) if row_K is not None else None
            vM = parse_float_safe(row_M.get(m)) if row_M is not None else None

            if vG is None and vC is None and vD is None and vI25 is None:
                continue

            g_ratio = (vG / prev_vals["G o"]) if (vG is not None and prev_vals["G o"] is not None and prev_vals["G o"] > 0) else 1.02

            # 4 Basamaklı İşçilik (I_2003): Mayıs 2026 sonrası =+(I14/H14)*H12 ile türetilir
            if vI03 is not None:
                final_I03 = vI03
            elif vI25 is not None and prev_I25 is not None and prev_I25 > 0 and prev_I03 is not None:
                final_I03 = (vI25 / prev_I25) * prev_I03
            elif prev_I03 is not None:
                final_I03 = prev_I03 * g_ratio
            else:
                final_I03 = 274.44

            # 3 Basamaklı İşçilik (I_2025): Varsa doğrudan al, yoksa I_2003 artış oranıyla türet
            if vI25 is not None:
                final_I25 = vI25
            elif final_I03 is not None and prev_I03 is not None and prev_I03 > 0 and prev_I25 is not None:
                final_I25 = (final_I03 / prev_I03) * prev_I25
            elif final_I03 is not None:
                final_I25 = final_I03 / 31.8322916667
            else:
                final_I25 = 100.0

            prev_I03 = final_I03
            prev_I25 = final_I25

            curr_map = {"Ç o": vC, "D o": vD, "Y_19": vY19, "K o": vK, "G o": vG, "M o": vM}
            for k_c in curr_map:
                if curr_map[k_c] is None and prev_vals[k_c] is not None:
                    curr_map[k_c] = prev_vals[k_c] * g_ratio
                if curr_map[k_c] is not None:
                    prev_vals[k_c] = curr_map[k_c]

            ay_label = f"{tr_months[i]} {year}"
            v_epdk = epdk_map.get(ay_label, None)

            records.append({
                "AYLAR": ay_label,
                "I_2003": round(final_I03, 2),
                "I_2025": round(final_I25, 6),
                "Ç o": round(curr_map["Ç o"], 2),
                "D o": round(curr_map["D o"], 2),
                "Y_19": round(curr_map["Y_19"], 2),
                "Y_EPDK": round(v_epdk, 6) if v_epdk is not None else None,
                "K o": round(curr_map["K o"], 2),
                "G o": round(curr_map["G o"], 2),
                "M o": round(curr_map["M o"], 2),
            })

    # Eksik EPDK aylarını Y_19 değişim oranıyla ileri/geri doldur
    valid_epdk_indices = [idx for idx, r in enumerate(records) if r["Y_EPDK"] is not None and r["Y_EPDK"] > 0]
    if valid_epdk_indices:
        first_idx = valid_epdk_indices[0]
        for idx in range(first_idx - 1, -1, -1):
            r_19 = records[idx]["Y_19"] / records[idx + 1]["Y_19"] if records[idx + 1]["Y_19"] > 0 else 0.98
            records[idx]["Y_EPDK"] = round(records[idx + 1]["Y_EPDK"] * r_19, 6)
        for idx in range(first_idx + 1, len(records)):
            if records[idx]["Y_EPDK"] is None or records[idx]["Y_EPDK"] <= 0:
                r_19 = records[idx]["Y_19"] / records[idx - 1]["Y_19"] if records[idx - 1]["Y_19"] > 0 else 1.02
                records[idx]["Y_EPDK"] = round(records[idx - 1]["Y_EPDK"] * r_19, 6)

    return pd.DataFrame(records)

def load_master_db():
    if os.path.exists(MASTER_DB_FILE):
        try:
            with open(MASTER_DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                df_m = pd.DataFrame(data)
                if len(df_m) >= 120 and "I_2025" in df_m.columns and "Y_EPDK" in df_m.columns:
                    return df_m
        except Exception:
            pass

    df_m = pd.DataFrame(BUILTIN_MASTER_KUTUK)
    try:
        df_m.to_json(MASTER_DB_FILE, orient="records", force_ascii=False, indent=2)
    except Exception:
        pass
    return df_m

def save_master_db(df_m):
    try:
        df_m.to_json(MASTER_DB_FILE, orient="records", force_ascii=False, indent=2)
    except Exception:
        pass

def update_master_from_web():
    try:
        tuik_raw = fetch_and_parse_indices()
        if tuik_raw.empty:
            return False, "Hakedis.org TÜİK tablosu çekilemedi.", pd.DataFrame()
        try:
            epdk_raw = fetch_and_parse_epdk()
        except Exception:
            epdk_raw = None
        df_m = build_master_db_from_dfs(tuik_raw, epdk_raw)
        if not df_m.empty:
            save_master_db(df_m)
            return True, f"TÜİK ve EPDK Kütüğü Başarıyla Güncellendi ({len(df_m)} Ay).", df_m
        return False, "Dönüştürme başarısız oldu.", pd.DataFrame()
    except Exception as e:
        return False, f"Güncelleme Hatası: {e}", pd.DataFrame()

# ==========================================
# 3. KULLANICININ TEMEL ENDEKSİNE GÖRE DİNAMİK KÜTÜK OLUŞTURUCU
# ==========================================
def tespit_et_ihale_baz_ayi(df_master, df_alt, varsayilan_ay="Temmuz 2025"):
    """
    AltEndeks tablosundaki Ç o, D o, K o, G o, M o temel endekslerinden
    ihalenin hangi aya ait olduğunu otomatik bulur (ör. Temmuz 2025).
    """
    if df_master is None or df_master.empty or df_alt is None or df_alt.empty:
        return varsayilan_ay
    temel_map = {str(r.get('Ağırlık', '')).strip().lower(): float(clean_decimal(r.get('Temel Endeks', 0))) for _, r in filter_empty_rows(df_alt).iterrows()}
    en_iyi_ay = varsayilan_ay
    en_yuksek_eslesme = 0
    for _, row in df_master.iterrows():
        eslesme = 0
        for k_kod, col_ad in [('b1', 'Ç o'), ('b2', 'D o'), ('b4', 'K o'), ('b5', 'G o'), ('c', 'M o')]:
            tv = temel_map.get(k_kod, 0.0)
            mv = float(row.get(col_ad, 0.0))
            if tv > 0 and mv > 0 and abs(tv - mv) / mv < 0.005:
                eslesme += 1
        if eslesme > en_yuksek_eslesme:
            en_yuksek_eslesme = eslesme
            en_iyi_ay = str(row["AYLAR"])
    return en_iyi_ay

def olustur_aktif_kütük(df_master, df_alt, secili_ihale_ayi="Temmuz 2025"):
    """
    Kullanıcının belirttiği kurala göre:
    1) İşçilik (a -> I o) Temel Endeksi:
       - 3 basamaklı (< 500, ör. 100,421925) ise -> I_2025 (2025=100) serisini kullanır.
       - 4 basamaklı (>= 500, ör. 3196,66) ise -> I_2003 (2003=100, Haziran 2026+ türetilmiş) serisini kullanır.
    2) Yakıt (b3 -> Y o) Temel Endeksi:
       - EPDK ölçeğinde (< 500, ör. 44,717708 veya 53,14131) ise -> Y_EPDK (epdk_akaryakit_fiyatlari.xlsx)
         serisini kullanır ve temel endeks KDV hariç (44,717708) girilmişse baz aydaki EPDK oranıyla
         tüm EPDK serisini birebir bağlantı kurarak ölçekler.
       - Yapım İşleri ölçeğinde (>= 500, ör. 7767,87) ise -> Y_19 (TÜİK 19 Kok ve rafine petrol) serisini kullanır.
    """
    if df_master is None or df_master.empty:
        return pd.DataFrame(), {}

    temel_map = {str(r.get('Ağırlık', '')).strip().lower(): float(clean_decimal(r.get('Temel Endeks', 0))) for _, r in filter_empty_rows(df_alt).iterrows()}
    temel_I = temel_map.get('a', 100.421925)
    temel_Y = temel_map.get('b3', 44.717708)

    baz_ay = tespit_et_ihale_baz_ayi(df_master, df_alt, varsayilan_ay=secili_ihale_ayi)
    baz_row_df = df_master[df_master["AYLAR"].astype(str).str.strip() == baz_ay]
    baz_row = baz_row_df.iloc[0] if not baz_row_df.empty else df_master.iloc[-1]

    # İşçilik modu: < 500 ise 3 basamaklı (2025=100), >= 500 ise 4 basamaklı (2003=100)
    iscilik_3_basamak = (0 < temel_I < 500)
    # Yakıt modu: < 500 ise EPDK Motorin, >= 500 ise Yapım İşleri TÜİK 19
    yakit_epdk_modu = (0 < temel_Y < 500)

    # Y_EPDK serisi doğrudan KDV hariç olarak (Haziran 2023 ve öncesi %18 -> /1.18, Temmuz 2023 sonrası %20 -> /1.20) hesaplanmıştır
    out_rows = []
    for _, r in df_master.iterrows():
        i_val = float(r["I_2025"]) if iscilik_3_basamak else float(r["I_2003"])
        y_val = float(r["Y_EPDK"]) if yakit_epdk_modu else float(r["Y_19"])

        i_str = tr_format(i_val, 6) if iscilik_3_basamak else tr_format(i_val, 2)
        y_str = tr_format(y_val, 6) if yakit_epdk_modu else tr_format(y_val, 2)

        out_rows.append({
            "AYLAR": str(r["AYLAR"]),
            "I o": i_str,
            "Ç o": tr_format(r["Ç o"], 2),
            "D o": tr_format(r["D o"], 2),
            "Y o": y_str,
            "K o": tr_format(r["K o"], 2),
            "G o": tr_format(r["G o"], 2),
            "M o": tr_format(r["M o"], 2),
        })

    bilgi = {
        "baz_ay": baz_ay,
        "iscilik_mod": "3 Basamaklı TÜFE (2025=100)" if iscilik_3_basamak else "4 Basamaklı TÜFE (2003=100, =+(I14/H14)*H12 Türetmeli)",
        "yakit_mod": "EPDK Motorin KDV Hariç Serisi (Temmuz 2023 Öncesi %18 KDV, Sonrası %20 KDV Düşülmüş)" if yakit_epdk_modu else "Yapım İşleri TÜİK 19 (Kok ve Rafine Petrol, 2003=100)",
        "iscilik_3_basamak": iscilik_3_basamak,
        "yakit_epdk_modu": yakit_epdk_modu,
    }
    return clean_df_for_ui(pd.DataFrame(out_rows)), bilgi

def kütükten_temel_endeks_getir(df_master, secilen_ay_str, df_alt_mevcut, iscilik_3_basamak=True, yakit_epdk_modu=True):
    """Seçilen ihale ayındaki değerleri (aktif 3/4 basamak ve EPDK/TÜİK19 moduna uygun olarak) AltEndeks'e yazar."""
    if df_master is None or df_master.empty or df_alt_mevcut is None or df_alt_mevcut.empty:
        return df_alt_mevcut
    row_match = df_master[df_master["AYLAR"].astype(str).str.strip() == str(secilen_ay_str).strip()]
    if row_match.empty:
        return df_alt_mevcut
    r_db = row_match.iloc[0]
    df_new_alt = df_alt_mevcut.copy()
    for idx, r in df_new_alt.iterrows():
        k_kod = str(r.get('Ağırlık', '')).strip().lower()
        sut = str(r.get('Endeks Sütunu', '')).strip() or DEFAULT_HARITA.get(k_kod, '')
        if sut == "I o":
            val = float(r_db["I_2025"]) if iscilik_3_basamak else float(r_db["I_2003"])
            df_new_alt.loc[idx, 'Temel Endeks'] = tr_format(val, 6 if iscilik_3_basamak else 2)
        elif sut == "Y o":
            val = float(r_db["Y_EPDK"]) if yakit_epdk_modu else float(r_db["Y_19"])
            df_new_alt.loc[idx, 'Temel Endeks'] = tr_format(val, 6 if yakit_epdk_modu else 2)
        elif sut in r_db:
            df_new_alt.loc[idx, 'Temel Endeks'] = tr_format(float(r_db[sut]), 2)
    return clean_df_for_ui(df_new_alt)

def hazirla_prog_ve_turetilen_ay(df_prog_raw, df_b_raw, tamamlama_aktif=True):
    df_p = filter_empty_rows(df_prog_raw.copy())
    df_b = filter_empty_rows(df_b_raw.copy())
    if df_p.empty or not tamamlama_aktif:
        return ensure_text_df(df_p), ensure_text_df(df_b), None
    ay_col = df_p.columns[0]
    prog_col = df_p.columns[1]
    iml_col = df_p.columns[2]
    bos_indisler = [i for i in range(len(df_p)) if str(df_p.iloc[i, 2]).strip() in ['', 'none', 'nan', '<na>']]
    if len(bos_indisler) > 0:
        for idx in bos_indisler:
            df_p.iloc[idx, 2] = df_p.iloc[idx, 1]
        ilk_bos = bos_indisler[0]
        baz_iml = clean_decimal(df_p.iloc[ilk_bos - 1, 2]) if ilk_bos > 0 else Decimal('0.0')
        kalan_tutar = clean_decimal(df_p.iloc[-1, 1]) - baz_iml
        return ensure_text_df(df_p), ensure_text_df(df_b), {
            'tip': 'mevcut_bos_aylar',
            'kalan_tutar': float(kalan_tutar),
            'aciklama': f"Tablodaki boş bırakılan **{len(bos_indisler)} gelecek ay** iş programına paralel olarak tamamlandı."
        }
    else:
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

def hesapla_gereken_ufuk(df_prog, df_db):
    try:
        df_p = filter_empty_rows(df_prog)
        df_e = filter_empty_rows(df_db)
        p_col = 'AYLAR' if 'AYLAR' in df_p.columns else df_p.columns[0]
        e_col = 'AYLAR' if 'AYLAR' in df_e.columns else df_e.columns[0]
        p_max = pd.to_datetime(df_p[p_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M').max()
        e_max = pd.to_datetime(df_e[e_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M').max()

        if pd.notna(p_max) and pd.notna(e_max) and p_max > e_max:
            diff = (p_max.year - e_max.year) * 12 + (p_max.month - e_max.month)
            return int(max(0, diff + 1))
    except Exception:
        pass
    return 0

# ==========================================
# 4. İLERİ DÜZEY ZAMAN SERİSİ VE BAZ ENDEKS TAHMİN MOTORU
# ==========================================
def tek_seri_tahmin_uret(y_arr, ek_ay, model_tipi, alpha=0.65, beta=0.35, phi=0.90, mc_quantile=80, n_sims=1000):
    """
    Gerçek geçmiş değerleri bozmadan (son 18 aylık güncel rejim penceresi üzerinde)
    Holt Çift Üstel, Sönümlü Trend (Damped), Monte Carlo GBM ve Otoregresif (AR-1 & EWMA)
    projeksiyonlarını üretir.
    """
    y_raw = [float(v) for v in y_arr if float(v) > 0]
    if len(y_raw) == 0:
        return [100.0] * max(ek_ay, 0)
    y_full = np.array(y_raw, dtype=float)
    if len(y_full) < 2 or ek_ay <= 0:
        return [float(y_full[-1])] * max(ek_ay, 0)

    y = y_full[-18:] if len(y_full) > 18 else y_full.copy()
    n = len(y)

    raw_pct = np.diff(y) / np.maximum(y[:-1], 1e-6)
    wins_pct = np.clip(raw_pct, -0.04, 0.065)
    init_step = float(np.median(np.diff(y[:min(4, n)]))) if n >= 3 else float(y[1] - y[0])

    preds = []
    if "Holt" in model_tipi:
        L = float(y[0])
        T = init_step
        for i in range(1, n):
            L_prev = L
            y_obs = float(np.clip(y[i], L_prev * 0.95, L_prev * 1.07))
            L = alpha * y_obs + (1.0 - alpha) * (L_prev + T)
            T = beta * (L - L_prev) + (1.0 - beta) * T
        L = float(y[-1])
        T = float(np.clip(T, y[-1] * 0.003, y[-1] * 0.042))
        for h in range(1, ek_ay + 1):
            preds.append(L + h * T)

    elif "Sönümlü" in model_tipi or "Damped" in model_tipi:
        L = float(y[0])
        T = init_step
        for i in range(1, n):
            L_prev = L
            y_obs = float(np.clip(y[i], L_prev * 0.95, L_prev * 1.07))
            L = alpha * y_obs + (1.0 - alpha) * (L_prev + phi * T)
            T = beta * (L - L_prev) + (1.0 - beta) * phi * T
        L = float(y[-1])
        T = float(np.clip(T, y[-1] * 0.003, y[-1] * 0.042))
        for h in range(1, ek_ay + 1):
            phi_sum = sum(phi ** j for j in range(1, h + 1))
            preds.append(L + phi_sum * T)

    elif "Monte Carlo" in model_tipi or "GBM" in model_tipi:
        np.random.seed(42)
        recent_pct = wins_pct[-12:] if len(wins_pct) >= 12 else wins_pct
        w_gbm = np.exp(np.linspace(-1.0, 0.0, len(recent_pct)))
        w_gbm /= w_gbm.sum()
        log_returns = np.log(1.0 + np.maximum(recent_pct, -0.05))
        mu = float(np.clip(np.sum(log_returns * w_gbm), 0.004, 0.036))
        sigma = float(np.clip(np.std(log_returns), 0.002, 0.018))
        if mc_quantile == 50:
            curr = float(y[-1])
            for h in range(1, ek_ay + 1):
                curr = curr * np.exp(mu)
                preds.append(curr)
        else:
            sim_paths = np.zeros((n_sims, ek_ay))
            for s in range(n_sims):
                curr = float(y[-1])
                for h in range(ek_ay):
                    z = np.random.normal(0, 1)
                    step_growth = np.exp((mu - 0.5 * (sigma ** 2)) + sigma * z)
                    step_growth = float(np.clip(step_growth, 0.98, 1.055))
                    curr = curr * step_growth
                    sim_paths[s, h] = curr
            preds = np.percentile(sim_paths, mc_quantile, axis=0).tolist()

    else:
        recent_pct = wins_pct[-12:] if len(wins_pct) >= 12 else wins_pct
        weights = np.exp(np.linspace(-1.5, 0.0, len(recent_pct)))
        weights /= weights.sum()
        ewma_r = float(np.clip(np.sum(recent_pct * weights), 0.004, 0.038))
        last_r = float(np.clip(recent_pct[-1], 0.002, 0.042))
        ar_rho = 0.35
        curr = float(y[-1])
        curr_r = last_r
        for h in range(1, ek_ay + 1):
            curr_r = ar_rho * curr_r + (1.0 - ar_rho) * ewma_r
            curr = curr * (1.0 + float(np.clip(curr_r, 0.002, 0.045)))
            preds.append(curr)

    safe_preds = []
    prev_val = float(y[-1])
    for p_val in preds:
        safe_val = float(np.clip(float(p_val), prev_val * 1.001, prev_val * 1.050))
        prev_val = safe_val
        safe_preds.append(safe_val)
    return safe_preds

def backtest_modelleri_hesapla(df_db_real, df_alt, alpha=0.65, beta=0.35, phi=0.90):
    df_clean = ayristir_gercek_endeks(df_db_real.copy())
    end_col = 'AYLAR' if 'AYLAR' in df_clean.columns else df_clean.columns[0]
    endeks_sutunlari = [c for c in df_clean.columns if c != end_col]

    katsayilar = {}
    for _, r in filter_empty_rows(df_alt).iterrows():
        k_kod = str(r['Ağırlık']).strip().lower()
        sut = str(r.get('Endeks Sütunu', KOD_BILGI.get(k_kod, {}).get('kolon', ''))).strip()
        katsayilar[sut] = float(clean_decimal(r['Katsayı']))

    modeller = [
        ("1. Holt's Çift Üstel", "Holt"),
        ("2. Sönümlü Trend (Damped)", "Sönümlü"),
        ("3. Monte Carlo GBM (P50 Medyan)", "Monte Carlo"),
        ("4. Otoregresif (AR-1 & EWMA)", "Otoregresif")
    ]
    satirlar = []
    karma_mape = {etiket: 0.0 for etiket, _ in modeller}
    hibrit_karma_mape = 0.0
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
        en_iyi_tahmin_serisi = []
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
                safe_g = np.maximum(gercekler_arr, 1e-6)
                mape = float(np.mean(np.abs((gercekler_arr - tahminler_arr) / safe_g)) * 100.0)
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
                en_iyi_tahmin_serisi = list(tahminler)

        hibrit_karma_mape += agirlik * en_iyi_mape
        detay_gecmis[col]["5. En İyi Karma (Şampiyon)"] = en_iyi_tahmin_serisi
        detay_gecmis[col]["Şampiyon Model"] = en_iyi_model_isim
        row_dict["En Düşük Hata (Önerilen)"] = f"🏆 {en_iyi_model_isim} (%{en_iyi_mape:.2f})"
        en_iyi_modeller_dict[col] = en_iyi_model_kod
        satirlar.append(row_dict)

    en_iyi_tekil_model = min(karma_mape, key=karma_mape.get)
    karma_row = {"Alt Endeks": "📌 AĞIRLIKLI KARMA Pn HATASI", "Sözleşme Ağırlığı": "%100"}
    for etiket, _ in modeller:
        karma_row[f"{etiket} MAPE"] = f"%{karma_mape[etiket]:.2f}"
    karma_row["En Düşük Hata (Önerilen)"] = f"🌟 5. En İyi Karma: %{hibrit_karma_mape:.2f} (En İyi Tekil: {en_iyi_tekil_model.split('.')[1].strip()} %{karma_mape[en_iyi_tekil_model]:.2f})"
    karma_mape["5. En İyi Karma (Hibrit)"] = hibrit_karma_mape
    satirlar.append(karma_row)

    return pd.DataFrame(satirlar), detay_gecmis, karma_mape, en_iyi_modeller_dict

def gelismis_endeks_tahmini(df_db_real, ek_ay, model_tipi, alpha=0.65, beta=0.35, phi=0.90, mc_quantile=80, n_sims=1000, df_alt=None, baz_endeks_kolon="G o", sabit_aylik_artis_yuzde=None):
    df_clean = ayristir_gercek_endeks(df_db_real.copy())
    if df_clean.empty or ek_ay <= 0:
        return ensure_text_df(df_clean)

    end_col = 'AYLAR' if 'AYLAR' in df_clean.columns else df_clean.columns[0]
    parsed_dates = df_clean[end_col].apply(parse_turkish_date)
    df_clean['_period'] = pd.to_datetime(parsed_dates, format='%Y-%m', errors='coerce').dt.to_period('M')
    df_clean = df_clean.dropna(subset=['_period']).drop_duplicates(subset=['_period']).sort_values('_period').reset_index(drop=True)

    if df_clean.empty:
        return ensure_text_df(df_db_real)

    son_period = df_clean['_period'].iloc[-1]
    endeks_sutunlari = [c for c in df_clean.columns if c not in [end_col, '_period']]

    gelecek_periods = [son_period + k for k in range(1, ek_ay + 1)]
    tahmin_sozluk = {end_col: [period_to_tr_full_str(p) for p in gelecek_periods]}

    if "Baz Endeks" in model_tipi:
        ref_col = baz_endeks_kolon if baz_endeks_kolon in endeks_sutunlari else "G o"
        y_ref = [float(clean_decimal(v)) for v in df_clean[ref_col] if float(clean_decimal(v)) > 0]
        son_ref = y_ref[-1] if len(y_ref) > 0 else 100.0
        if sabit_aylik_artis_yuzde is not None and sabit_aylik_artis_yuzde > 0:
            oran = 1.0 + (float(sabit_aylik_artis_yuzde) / 100.0)
            ref_preds = [son_ref * (oran ** h) for h in range(1, ek_ay + 1)]
        else:
            ref_preds = tek_seri_tahmin_uret(y_ref, ek_ay, "Sönümlü", alpha=alpha, beta=beta, phi=phi)

        aylik_carpanlar = []
        prev_r = son_ref
        for rp in ref_preds:
            aylik_carpanlar.append(rp / prev_r if prev_r > 0 else 1.02)
            prev_r = rp

        for col in endeks_sutunlari:
            y_c = [float(clean_decimal(v)) for v in df_clean[col] if float(clean_decimal(v)) > 0]
            curr_v = y_c[-1] if len(y_c) > 0 else 100.0
            dec_places = 6 if curr_v < 500 else 2
            col_preds = []
            for carp in aylik_carpanlar:
                curr_v = curr_v * carp
                col_preds.append(curr_v)
            tahmin_sozluk[col] = [tr_format(p, dec_places) for p in col_preds]
    else:
        best_models_dict = {}
        if "Karma" in model_tipi and df_alt is not None:
            _, _, _, best_models_dict = backtest_modelleri_hesapla(df_clean.drop(columns=['_period']), df_alt, alpha=alpha, beta=beta, phi=phi)

        for col in endeks_sutunlari:
            y = [float(clean_decimal(v)) for v in df_clean[col] if float(clean_decimal(v)) > 0]
            son_v = y[-1] if len(y) > 0 else 100.0
            dec_places = 6 if son_v < 500 else 2
            active_model = best_models_dict[col] if ("Karma" in model_tipi and col in best_models_dict) else model_tipi
            preds = tek_seri_tahmin_uret(y, ek_ay, active_model, alpha=alpha, beta=beta, phi=phi, mc_quantile=mc_quantile, n_sims=n_sims)
            tahmin_sozluk[col] = [tr_format(p, dec_places) for p in preds]

    df_tahmin = pd.DataFrame(tahmin_sozluk)
    df_birlesik = pd.concat([df_clean.drop(columns=['_period']), df_tahmin], ignore_index=True)
    return df_birlesik.astype(str).astype("string")

# ==========================================
# 5. RESMİ FİYAT FARKI HESAPLAMA MOTORU
# ==========================================
def hesapla(df_prog, df_endeks_auto, df_alt, df_b, b_sabit_override=None):
    df_prog = filter_empty_rows(df_prog.copy())
    df_endeks = filter_empty_rows(df_endeks_auto.copy())
    df_alt = filter_empty_rows(df_alt.copy())
    df_b = filter_empty_rows(df_b.copy())

    if df_prog.empty or df_endeks.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df_prog.columns = df_prog.columns.str.strip()

    end_col = 'AYLAR' if 'AYLAR' in df_endeks.columns else df_endeks.columns[0]
    df_endeks['AyKodu'] = pd.to_datetime(df_endeks[end_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M')
    df_endeks = df_endeks.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).sort_values('AyKodu').set_index('AyKodu')

    if not df_b.empty:
        b_ay_col = 'AYLAR' if 'AYLAR' in df_b.columns else df_b.columns[0]
        df_b['AyKodu'] = pd.to_datetime(df_b[b_ay_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M')
        df_b = df_b.dropna(subset=['AyKodu']).drop_duplicates(subset=['AyKodu']).set_index('AyKodu')

    p_ay_col = 'AYLAR' if 'AYLAR' in df_prog.columns else df_prog.columns[0]
    df_prog['AyKodu'] = pd.to_datetime(df_prog[p_ay_col].apply(parse_turkish_date), format='%Y-%m', errors='coerce').dt.to_period('M')
    df_prog = df_prog.dropna(subset=['AyKodu'])

    if df_endeks.empty or df_prog.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    son_endeks_ayi = df_endeks.index.max()
    katsayilar = {str(r['Ağırlık']).strip().lower(): clean_decimal(r['Katsayı']) for _, r in df_alt.iterrows()}
    temel_endeksler = {str(r['Ağırlık']).strip().lower(): clean_decimal(r['Temel Endeks']) for _, r in df_alt.iterrows()}

    if 'Endeks Sütunu' in df_alt.columns:
        endeks_haritasi = {
            str(r['Ağırlık']).strip().lower(): str(r['Endeks Sütunu']).strip()
            for _, r in df_alt.iterrows()
            if str(r.get('Endeks Sütunu', '')).strip() not in ['', 'nan', 'None']
        }
        if not endeks_haritasi:
            endeks_haritasi = DEFAULT_HARITA.copy()
    else:
        endeks_haritasi = DEFAULT_HARITA.copy()

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
            if guncel > Decimal('0.0'):
                onceki_imalat_kum = guncel
            continue

        if b_sabit_override is not None:
            b_kat = Decimal(str(b_sabit_override))
        else:
            b_val = df_b.loc[uyg_ayi, 'B'] if (not df_b.empty and uyg_ayi in df_b.index and 'B' in df_b.columns) else Decimal('0.90')
            b_kat = clean_decimal(b_val)
            if b_kat <= Decimal('0.0'):
                b_kat = Decimal('0.90')

        gercek_end_ayi = min(uyg_ayi, son_endeks_ayi)
        if gercek_end_ayi in df_endeks.index:
            endeks_uyg = df_endeks.loc[gercek_end_ayi]
        else:
            gecmis_aylar = df_endeks.index[df_endeks.index <= gercek_end_ayi]
            endeks_uyg = df_endeks.loc[gecmis_aylar.max()] if len(gecmis_aylar) > 0 else df_endeks.iloc[0]

        toplam_ff_aylik = Decimal('0.0')
        kalan = aylik

        for kova in kovalar:
            if kalan <= Decimal('0.0'):
                break
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
                    if e_temel > Decimal('0.0'):
                        pn += katsayi * (e_gecerli / e_temel)
                    elif katsayi > Decimal('0.0'):
                        pn += katsayi

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
            'Ay': str(uyg_ayi),
            'Aylık İmalat': float(aylik),
            'Kümülatif İmalat': float(guncel),
            'B Katsayısı': float(b_kat),
            'Aylık Fiyat Farkı': float(toplam_ff_aylik),
            'Kümülatif Fiyat Farkı': float(kumulatif_ff)
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
    else:
        df_pivot = pd.DataFrame()

    return df_sonuc, df_pivot, df_detay, df_aylik

# ==========================================
# 6. KAPSAMLI EXCEL VE WORD RAPORU MOTORU
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
    make_sheet(ws2, '2. ENDEKS TABLOSU (RESMİ TÜİK / EPDK VE TAHMİN)', 'Dinamik seçilen resmi endeksler ve tahmin edilen aylar', df_endeks, [18] + [14]*(len(df_endeks.columns)-1), True)
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
        make_sheet(ws6, '6. TAHMİN MODELLERİNİN FİYAT FARKI KIYASLAMASI', 'Toplam maliyet projeksiyonları', df_karsilastirma, [38, 22, 26, 20, 26], False)
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
        if df is None or df.empty:
            return "<p>Veri bulunamadı.</p>"
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

    html_content = (
        "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>"
        "<head><meta charset=\"utf-8\"><title>Teknik Ofis Fiyat Farkı ve Endeks Projeksiyon Raporu</title>"
        "<style>body { font-family: 'Calibri', 'Arial', sans-serif; color: #1e293b; line-height: 1.45; font-size: 10.5pt; } "
        "h1 { color: #1e3a8a; font-size: 16pt; border-bottom: 2px solid #1e3a8a; padding-bottom: 4px; } "
        "h2 { color: #1e40af; font-size: 12.5pt; margin-top: 18px; border-left: 4px solid #1e40af; padding-left: 8px; } "
        ".kpi-box { background-color: #f1f5f9; border: 1px solid #cbd5e1; padding: 10px; margin-bottom: 14px; }</style></head>"
        "<body>"
        "<h1>TEKNİK OFİS FİYAT FARKI TAHMİN VE ENDEKS SİMÜLASYON RAPORU</h1>"
        f"<p><b>Rapor Tarihi:</b> {tarih_str} &nbsp;|&nbsp; <b>Aktif Tahmin Algoritması:</b> {aktif_model}</p>"
        "<h2>1. YÖNETİCİ ÖZETİ VE BÜTÇE PROJEKSİYONU</h2>"
        "<div class=\"kpi-box\"><ul>"
        f"<li><b>Sözleşme Bedeli (BAC):</b> {tr_format(toplam_sozlesme)} TL</li>"
        f"<li><b>Hesaplanan Kümülatif İmalat Tutarı:</b> {tr_format(toplam_iml)} TL</li>"
        f"<li><b>Mevcut Duruma Kadar Fiili Fiyat Farkı (İdari Hakediş):</b> {tr_format(mevcut_ff)} TL</li>"
        f"<li><b>Kalan / Türetilen Ay Fiyat Farkı:</b> {tr_format(kalan_ff)} TL</li>"
        f"<li><b>Proje Sonu Toplam Fiyat Farkı (FF):</b> <b>{tr_format(toplam_ff)} TL</b> (İmalata Oranı: %{ff_oran:.2f})</li>"
        f"<li><b>Genel Hakediş Toplamı (İmalat + Fiyat Farkı):</b> <b>{tr_format(genel_toplam)} TL</b></li>"
        "</ul></div>"
        "<h2>2. TAHMİN MODELLERİNİN KARŞILAŞTIRMALI MALİYET TABLOSU</h2>"
        f"{df_to_html_table(df_karsilastirma)}"
        "<h2>3. MODEL DOĞRULUK TESTİ (WALK-FORWARD BACKTESTING)</h2>"
        f"{df_to_html_table(df_backtest)}"
        "<h2>4. HAKEDİŞ ÖDENEK DİLİMİ (KOVA) VE GECİKME MATRİSİ DETAYI</h2>"
        f"{df_to_html_table(df_det_short)}"
        "</body></html>"
    )
    return html_content.encode('utf-8')

# ==========================================
# 7. GLOBAL DEĞİŞKENLER & VARSAYILAN OKUL PROJESİ VERİLERİ
# ==========================================
df_detay = pd.DataFrame()
df_karsilastirma_out = pd.DataFrame()
df_backtest_out = pd.DataFrame()
toplam_sozlesme = 0.0
toplam_gerceklesen_iml = 0.0
toplam_ff_tutari = 0.0
mevcut_ff_tutari = 0.0
kalan_ff_tutari = 0.0

DEFAULT_PROG = [
    {"AYLAR": "Eylül 25",  "İŞ PROGRAMI KÜMÜLATİF": "2.165.869,79",   "İMALAT TUTARI KÜMÜLATİF": "0"},
    {"AYLAR": "Ekim 25",   "İŞ PROGRAMI KÜMÜLATİF": "6.470.938,48",   "İMALAT TUTARI KÜMÜLATİF": "0"},
    {"AYLAR": "Kasım 25",  "İŞ PROGRAMI KÜMÜLATİF": "17.293.937,21",  "İMALAT TUTARI KÜMÜLATİF": "0"},
    {"AYLAR": "Aralık 25", "İŞ PROGRAMI KÜMÜLATİF": "30.287.885,90",  "İMALAT TUTARI KÜMÜLATİF": "0"},
    {"AYLAR": "Ocak 26",   "İŞ PROGRAMI KÜMÜLATİF": "55.011.822,59",  "İMALAT TUTARI KÜMÜLATİF": "54.223.844,40"},
    {"AYLAR": "Şubat 26",  "İŞ PROGRAMI KÜMÜLATİF": "84.518.453,41",  "İMALAT TUTARI KÜMÜLATİF": "54.223.844,40"},
    {"AYLAR": "Mart 26",   "İŞ PROGRAMI KÜMÜLATİF": "116.193.494,10", "İMALAT TUTARI KÜMÜLATİF": "86.775.488,80"},
    {"AYLAR": "Nisan 26",  "İŞ PROGRAMI KÜMÜLATİF": "148.697.481,54", "İMALAT TUTARI KÜMÜLATİF": "128.039.357,60"},
    {"AYLAR": "Mayıs 26",  "İŞ PROGRAMI KÜMÜLATİF": "180.326.778,21", "İMALAT TUTARI KÜMÜLATİF": "162.021.367,20"}
]
# Varsayılan Okul Projesi (Temmuz 2025: I o = 3 basamaklı 100,421925 [2025=100]; Y o = EPDK 44,717708)
DEFAULT_ALT = [
    {"Ağırlık": "a",  "Katsayı": "0,15", "Temel Endeks": "100,421925",  "Endeks Sütunu": "I o"},
    {"Ağırlık": "b1", "Katsayı": "0,20", "Temel Endeks": "4972,280000", "Endeks Sütunu": "Ç o"},
    {"Ağırlık": "b2", "Katsayı": "0,20", "Temel Endeks": "6034,240000", "Endeks Sütunu": "D o"},
    {"Ağırlık": "b3", "Katsayı": "0,15", "Temel Endeks": "44,284425",   "Endeks Sütunu": "Y o"},
    {"Ağırlık": "b4", "Katsayı": "0,05", "Temel Endeks": "3481,270000", "Endeks Sütunu": "K o"},
    {"Ağırlık": "b5", "Katsayı": "0,10", "Temel Endeks": "4409,730000", "Endeks Sütunu": "G o"},
    {"Ağırlık": "c",  "Katsayı": "0,15", "Temel Endeks": "3218,000000", "Endeks Sütunu": "M o"}
]
DEFAULT_B = [
    {"AYLAR": "Eylül 25", "B": "0,9"},
    {"AYLAR": "Ekim 25",  "B": "0,9"},
    {"AYLAR": "Kasım 25", "B": "0,9"}
]

def load_from_excel(file):
    xls = pd.ExcelFile(file)
    dfs = {}
    if 'Tüm_Yıllar' in xls.sheet_names:
        raw_df = pd.read_excel(xls, sheet_name='Tüm_Yıllar')
        cols_up = [str(c).upper() for c in raw_df.columns]
        if 'AKARYAKIT TÜRÜ' in cols_up:
            dfs['epdk_raw_df'] = raw_df
        elif 'ENDEKS KODU' in cols_up:
            dfs['tuik_raw_df'] = raw_df
        return dfs

    sheet_map = {'IsProgrami': 'prog_df', 'Endeks': 'endeks_df', 'AltEndeks': 'alt_df', 'B': 'b_df'}
    for sheet in xls.sheet_names:
        if sheet in sheet_map:
            df_temp = pd.read_excel(xls, sheet_name=sheet, nrows=5)
            skip = 0
            cols = [str(c).upper() for c in df_temp.columns]
            if 'AYLAR' not in cols and 'AĞIRLIK' not in cols:
                for i, row in df_temp.iterrows():
                    if 'AYLAR' in [str(v).upper() for v in row.values] or 'AĞIRLIK' in [str(v).upper() for v in row.values]:
                        skip = i + 1
                        break
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=skip)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            if sheet == 'AltEndeks' and 'Endeks Sütunu' not in df.columns:
                df['Endeks Sütunu'] = df['Ağırlık'].astype(str).str.strip().str.lower().map(DEFAULT_HARITA).fillna('')
            dfs[sheet_map[sheet]] = clean_df_for_ui(df)
    return dfs

if 'ff_load_count' not in st.session_state:
    st.session_state.ff_load_count = 0
if 'ff_prog_df' not in st.session_state:
    st.session_state.ff_prog_df = clean_df_for_ui(pd.DataFrame(DEFAULT_PROG))
if 'ff_alt_df' not in st.session_state:
    st.session_state.ff_alt_df = clean_df_for_ui(pd.DataFrame(DEFAULT_ALT))
if 'ff_b_df' not in st.session_state:
    st.session_state.ff_b_df = clean_df_for_ui(pd.DataFrame(DEFAULT_B))
if 'ff_ek_ay_slider' not in st.session_state:
    st.session_state.ff_ek_ay_slider = 10
if 'master_db' not in st.session_state or st.session_state.master_db.empty:
    st.session_state.master_db = load_master_db()

# İlk açılışta aktif kütüğü oluştur
if 'ff_base_endeks_df' not in st.session_state or st.session_state.ff_base_endeks_df.empty:
    st.session_state.ff_base_endeks_df, st.session_state.aktif_kutuk_bilgi = olustur_aktif_kütük(
        st.session_state.master_db, st.session_state.ff_alt_df
    )

# ==========================================
# 8. YAN MENÜ VE VERİTABANI YÖNETİMİ
# ==========================================
with st.sidebar:
    st.header("📥 Proje veya Endeks Yükle")
    uploaded_file = st.file_uploader(
        "Proje (.xlsx/.json), TÜİK Excel veya EPDK Yakıt Excel",
        type=["xlsx", "xls", "json"]
    )
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        if st.session_state.get('ff_last_bytes') != file_bytes:
            try:
                if uploaded_file.name.endswith(('.xlsx', '.xls')):
                    dfs = load_from_excel(uploaded_file)
                    if 'epdk_raw_df' in dfs:
                        # Yalnızca EPDK yakıt tablosu yüklendiyse mevcut master_db içindeki Y_EPDK sütununu güncelle
                        epdk_raw = dfs['epdk_raw_df']
                        months_l = ['OCAK', 'ŞUBAT', 'MART', 'NİSAN', 'MAYIS', 'HAZİRAN', 'TEMMUZ', 'AĞUSTOS', 'EYLÜL', 'EKİM', 'KASIM', 'ARALIK']
                        tr_m_l = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık']
                        turu_c = "AKARYAKIT TÜRÜ" if "AKARYAKIT TÜRÜ" in epdk_raw.columns else epdk_raw.columns[1]
                        mot_r = epdk_raw[epdk_raw[turu_c].astype(str).str.strip() == "Motorin"]
                        if mot_r.empty:
                            mot_r = epdk_raw[epdk_raw[turu_c].astype(str).str.contains("Motorin", case=False, na=False)]
                        epdk_new_map = {}
                        for _, r_e in mot_r.iterrows():
                            y_e = int(r_e["YIL"])
                            for idx_m, m_e in enumerate(months_l):
                                fv = parse_float_safe(r_e.get(m_e))
                                if fv is not None:
                                    kdv_oran = 1.18 if (y_e < 2023 or (y_e == 2023 and idx_m < 6)) else 1.20
                                    epdk_new_map[f"{tr_m_l[idx_m]} {y_e}"] = round(fv / kdv_oran, 6)
                        df_m_curr = st.session_state.master_db.copy()
                        for idx_row, r_m in df_m_curr.iterrows():
                            ay_s = str(r_m["AYLAR"])
                            if ay_s in epdk_new_map:
                                df_m_curr.loc[idx_row, "Y_EPDK"] = epdk_new_map[ay_s]
                        st.session_state.master_db = df_m_curr
                        save_master_db(df_m_curr)
                    elif 'tuik_raw_df' in dfs:
                        df_m_new = build_master_db_from_dfs(dfs['tuik_raw_df'])
                        st.session_state.master_db = df_m_new
                        save_master_db(df_m_new)
                    else:
                        if 'prog_df' in dfs:
                            st.session_state.ff_prog_df = dfs['prog_df']
                        if 'alt_df' in dfs:
                            st.session_state.ff_alt_df = dfs['alt_df']
                        if 'b_df' in dfs:
                            st.session_state.ff_b_df = dfs['b_df']
                else:
                    data = json.load(uploaded_file)
                    if 'prog' in data:
                        st.session_state.ff_prog_df = clean_df_for_ui(pd.DataFrame(data['prog']))
                    if 'alt' in data:
                        df_a = pd.DataFrame(data['alt'])
                        if 'Endeks Sütunu' not in df_a.columns:
                            df_a['Endeks Sütunu'] = df_a['Ağırlık'].astype(str).str.strip().str.lower().map(DEFAULT_HARITA).fillna('')
                        st.session_state.ff_alt_df = clean_df_for_ui(df_a)
                    if 'b' in data:
                        st.session_state.ff_b_df = clean_df_for_ui(pd.DataFrame(data['b']))

                st.session_state.ff_base_endeks_df, st.session_state.aktif_kutuk_bilgi = olustur_aktif_kütük(
                    st.session_state.master_db, st.session_state.ff_alt_df
                )
                df_p_full, _, _ = hazirla_prog_ve_turetilen_ay(st.session_state.ff_prog_df, st.session_state.ff_b_df, tamamlama_aktif=True)
                st.session_state.ff_ek_ay_slider = max(6, hesapla_gereken_ufuk(df_p_full, st.session_state.ff_base_endeks_df))

                st.session_state.ff_load_count += 1
                st.session_state.ff_last_bytes = file_bytes
                st.session_state.ff_last_model_sig = None
                st.rerun()
            except Exception as e:
                st.error(f"Yükleme hatası: {e}")

    if st.button("🔄 Projeyi Sıfırla (Okul Projesi Örneği)", use_container_width=True):
        st.session_state.ff_prog_df = clean_df_for_ui(pd.DataFrame(DEFAULT_PROG))
        st.session_state.ff_alt_df = clean_df_for_ui(pd.DataFrame(DEFAULT_ALT))
        st.session_state.ff_b_df = clean_df_for_ui(pd.DataFrame(DEFAULT_B))
        st.session_state.ff_base_endeks_df, st.session_state.aktif_kutuk_bilgi = olustur_aktif_kütük(
            st.session_state.master_db, st.session_state.ff_alt_df
        )
        st.session_state.ff_ek_ay_slider = max(6, hesapla_gereken_ufuk(st.session_state.ff_prog_df, st.session_state.ff_base_endeks_df))
        st.session_state.ff_last_model_sig = None
        st.session_state.ff_load_count += 1
        st.rerun()

    st.markdown("---")
    st.header("🗄️ Merkezi TÜİK & EPDK Kütüğü")
    if not st.session_state.master_db.empty:
        son_ay = st.session_state.master_db.iloc[-1]['AYLAR']
        st.success(f"Kütük Aktif: {len(st.session_state.master_db)} Ay.\nSon Gerçek Veri: **{son_ay}**")

    if st.button("🔄 Hakedis.org'dan TÜİK + EPDK Güncelle", use_container_width=True):
        with st.spinner("Hakedis.org'dan TÜİK ve EPDK Motorin tabloları çekiliyor..."):
            basari, mesaj, yeni_m = update_master_from_web()
            if basari:
                st.session_state.master_db = yeni_m
                st.session_state.ff_base_endeks_df, st.session_state.aktif_kutuk_bilgi = olustur_aktif_kütük(
                    st.session_state.master_db, st.session_state.ff_alt_df
                )
                st.success(mesaj)
                st.session_state.ff_last_model_sig = None
                st.session_state.ff_load_count += 1
                st.rerun()
            else:
                st.info(mesaj)

    st.markdown("---")
    st.subheader("🔮 Gelecek Endeks Tahmini")
    model_secimi = st.selectbox(
        "İstatistiksel / Baz Endeks Modeli:",
        [
            "5. En İyi Karma (Alt Endekslere Göre Dinamik) 🌟",
            "6. Baz Endeks Oranlı Artış (Bir Endeksi Baz Al, Diğerlerini Artır) 📈",
            "1. Holt's Çift Üstel Düzleştirme (Level + Trend)",
            "2. Sönümlü Trend (Damped Trend - Dezenflasyon)",
            "3. Stokastik Monte Carlo / GBM (Olasılıksal)",
            "4. Otoregresif Momentum (AR-1 & EWMA)",
            "0. Tahmin Yok (Son Bilinen Endeks Sabit Kalsın)"
        ]
    )

    df_p_check, _, _ = hazirla_prog_ve_turetilen_ay(st.session_state.ff_prog_df, st.session_state.ff_b_df, tamamlama_aktif=True)
    hesaplanan_fark = hesapla_gereken_ufuk(df_p_check, st.session_state.ff_base_endeks_df)
    varsayilan_ufuk = max(6, hesaplanan_fark)
    if st.session_state.ff_ek_ay_slider < hesaplanan_fark:
        st.session_state.ff_ek_ay_slider = hesaplanan_fark
    ek_ay = st.slider("Tahmin Ufku (Ay)", min_value=0, max_value=60, value=varsayilan_ufuk, key="ff_ek_ay_slider")

    alpha_p, beta_p, phi_p, mc_q = 0.65, 0.35, 0.90, 80
    baz_endeks_sec_kolon = "G o"
    sabit_artis_oran = None

    if "Baz Endeks" in model_secimi:
        baz_endeks_sec_kod = st.selectbox(
            "Baz Alınacak Referans Endeks:",
            options=list(KOD_BILGI.keys()),
            index=5,
            format_func=lambda k: f"{KOD_BILGI[k]['kolon']} — {KOD_BILGI[k]['kisa']}"
        )
        baz_endeks_sec_kolon = KOD_BILGI[baz_endeks_sec_kod]['kolon']
        sabit_artis_aktif = st.checkbox("Sabit Aylık % Artış Oranı Kullan", value=False)
        if sabit_artis_aktif:
            sabit_artis_oran = st.slider("Aylık Sabit Artış (%)", 0.5, 6.0, 2.2, 0.1)
    if "Holt" in model_secimi or "Karma" in model_secimi:
        c_p1, c_p2 = st.columns(2)
        alpha_p = c_p1.slider("Seviye (α)", 0.1, 0.95, 0.65, 0.05)
        beta_p = c_p2.slider("Trend (β)", 0.05, 0.90, 0.35, 0.05)
    if "Sönümlü" in model_secimi or "Karma" in model_secimi or "Baz Endeks" in model_secimi:
        phi_p = st.slider("Sönümleme Katsayısı (ϕ)", 0.75, 0.98, 0.90, 0.01)
    if "Monte Carlo" in model_secimi or "Karma" in model_secimi:
        mc_q = st.select_slider("Güven Aralığı (Persentil)", options=[50, 70, 80, 90, 95], value=80)

# ==========================================
# 9. ANA EKRAN VE 5 GENİŞ SEKME
# ==========================================
st.title("📈 Fiyat Farkı Tahmin & Stokastik Endeks Simülatörü")
st.caption("İşçilik (3 Basamak 2025=100 / 4 Basamak 2003=100) & Yakıt (EPDK Motorin / Yapım İşleri TÜİK 19) Otomatik Algılama ve Bağlantı Sistemi")

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
        height=480,
        key=f"ff_prog_ed_{suffix}_{gelecek_imalat_modu[:8]}"
    )

    st.markdown("---")
    c_alt, c_b = st.columns([1.35, 1])
    with c_alt:
        st.subheader("2️⃣ Alt Endeks Ağırlıkları ve Temel Endeksler (İ_o)")
        st.caption(
            "💡 **Otomatik Algılama:** İşçilik (`a`) Temel Endeksi **3 basamaklı** (`<500`) ise `2025=100` serisi, **4 basamaklı** (`>=500`) ise `2003=100` serisi seçilir. "
            "Yakıt (`b3`) Temel Endeksi **EPDK ölçeğinde** (`<500`) ise **EPDK Motorin** serisi, **4 basamaklı** (`>=500`) ise **Yapım İşleri TÜİK 19** serisi kullanılır."
        )
        mevcut_aylar_list = st.session_state.master_db['AYLAR'].tolist() if not st.session_state.master_db.empty else ["Temmuz 2025"]
        varsayilan_ay_idx = mevcut_aylar_list.index("Temmuz 2025") if "Temmuz 2025" in mevcut_aylar_list else len(mevcut_aylar_list) - 1
        col_ih1, col_ih2, col_ih3 = st.columns([1.3, 1.1, 1.1])
        secilen_ihale_ayi = col_ih1.selectbox("📅 İhale Ayı:", options=mevcut_aylar_list, index=varsayilan_ay_idx)
        if col_ih2.button("📥 EPDK & 3B Getir", help="Seçili ayın 3 basamaklı İşçilik (2025=100) ve EPDK Motorin değerlerini getirir", use_container_width=True):
            st.session_state.ff_alt_df = kütükten_temel_endeks_getir(st.session_state.master_db, secilen_ihale_ayi, st.session_state.ff_alt_df, iscilik_3_basamak=True, yakit_epdk_modu=True)
            st.session_state.ff_load_count += 1
            st.session_state.ff_last_model_sig = None
            st.rerun()
        if col_ih3.button("📥 TÜİK 4B Getir", help="Seçili ayın 4 basamaklı İşçilik (2003=100) ve Yapım İşleri TÜİK 19 değerlerini getirir", use_container_width=True):
            st.session_state.ff_alt_df = kütükten_temel_endeks_getir(st.session_state.master_db, secilen_ihale_ayi, st.session_state.ff_alt_df, iscilik_3_basamak=False, yakit_epdk_modu=False)
            st.session_state.ff_load_count += 1
            st.session_state.ff_last_model_sig = None
            st.rerun()

        edited_alt = st.data_editor(
            st.session_state.ff_alt_df,
            column_config=get_text_config(st.session_state.ff_alt_df),
            num_rows="dynamic",
            use_container_width=True,
            height=290,
            key=f"ff_alt_ed_{suffix}"
        )

    with c_b:
        st.subheader("3️⃣ Aylık B Katsayısı Tablosu")
        edited_b = st.data_editor(
            df_b_working,
            column_config=get_text_config(df_b_working),
            num_rows="dynamic",
            use_container_width=True,
            height=355,
            key=f"ff_b_ed_{suffix}_{gelecek_imalat_modu[:8]}"
        )

# Kullanıcının düzenlediği AltEndeks (Temel Endeks) değerlerine göre (3B vs 4B İşçilik, EPDK vs TÜİK19 Yakıt) aktif kütüğü oluştur
alt_imza = "|".join([f"{r.get('Ağırlık','')}:{r.get('Katsayı','')}:{r.get('Temel Endeks','')}" for _, r in filter_empty_rows(edited_alt).iterrows()])
current_sig = f"{model_secimi}_{ek_ay}_{alpha_p}_{beta_p}_{phi_p}_{mc_q}_{baz_endeks_sec_kolon}_{sabit_artis_oran}_{len(st.session_state.master_db)}_{alt_imza}_{secilen_ihale_ayi}"

if (
    'ff_endeks_df' not in st.session_state
    or st.session_state.get('ff_last_model_sig') != current_sig
):
    st.session_state.ff_base_endeks_df, st.session_state.aktif_kutuk_bilgi = olustur_aktif_kütük(
        st.session_state.master_db, edited_alt, secili_ihale_ayi=secilen_ihale_ayi
    )
    if model_secimi.startswith("0.") or ek_ay == 0:
        st.session_state.ff_endeks_df = clean_df_for_ui(st.session_state.ff_base_endeks_df.copy())
    else:
        st.session_state.ff_endeks_df = gelismis_endeks_tahmini(
            st.session_state.ff_base_endeks_df,
            ek_ay,
            model_secimi,
            alpha=alpha_p,
            beta=beta_p,
            phi=phi_p,
            mc_quantile=mc_q,
            df_alt=edited_alt,
            baz_endeks_kolon=baz_endeks_sec_kolon,
            sabit_aylik_artis_yuzde=sabit_artis_oran
        )
    st.session_state.ff_last_model_sig = current_sig

# ---------------- SEKME 2: ENDEKS KÜTÜĞÜ VE TAHMİN EĞRİLERİ ----------------
with main_tab2:
    baz_df = st.session_state.ff_base_endeks_df.copy()
    k_bilgi = st.session_state.get('aktif_kutuk_bilgi', {})

    st.subheader(f"📈 Dinamik Endeks Kütüğü ve Gelecek Projeksiyonu — Aktif Model: {model_secimi}")
    st.info(
        f"🔎 **Otomatik Algılanan Seri Bağlantısı (İhale Baz Ayı: {k_bilgi.get('baz_ay', 'Temmuz 2025')}):**\n"
        f"- **İşçilik (`I o`):** {k_bilgi.get('iscilik_mod', '')}\n"
        f"- **Yakıt (`Y o`):** {k_bilgi.get('yakit_mod', '')}"
    )

    edited_endeks = st.data_editor(
        st.session_state.ff_endeks_df,
        column_config=get_text_config(st.session_state.ff_endeks_df),
        num_rows="dynamic",
        use_container_width=True,
        height=520,
        key=f"ff_end_ed_{suffix}_{ek_ay}_{alt_imza[:20]}"
    )

    st.markdown("---")
    st.subheader("🔬 Seçili Alt Endeks İçin Tahmin Modellerinin Karşılaştırmalı Projeksiyonu")
    secili_endeks_kodu = st.selectbox(
        "Grafikte İncelenecek Alt Endeks:",
        options=list(KOD_BILGI.keys()),
        format_func=lambda k: f"{KOD_BILGI[k]['kolon']} — {KOD_BILGI[k]['ad']} ({KOD_BILGI[k]['resmi_kod']})"
    )
    secili_kolon = KOD_BILGI[secili_endeks_kodu]['kolon']

    if not baz_df.empty and secili_kolon in baz_df.columns:
        ufuk_gosterim = max(ek_ay, 10)
        df_m1 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Holt", alpha=alpha_p, beta=beta_p, df_alt=edited_alt)
        df_m2 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Sönümlü", alpha=alpha_p, beta=beta_p, phi=phi_p, df_alt=edited_alt)
        df_m3 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Monte Carlo", mc_quantile=mc_q, df_alt=edited_alt)
        df_m4 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Otoregresif", df_alt=edited_alt)
        df_m6 = gelismis_endeks_tahmini(baz_df, ufuk_gosterim, "Baz Endeks", alpha=alpha_p, beta=beta_p, phi=phi_p, df_alt=edited_alt, baz_endeks_kolon=baz_endeks_sec_kolon, sabit_aylik_artis_yuzde=sabit_artis_oran)

        fig_end = go.Figure()
        n_hist = len(baz_df)
        ay_col = 'AYLAR' if 'AYLAR' in df_m1.columns else df_m1.columns[0]
        hist_start = max(0, n_hist - 36)

        fig_end.add_trace(go.Scatter(
            x=df_m1[ay_col][hist_start:n_hist],
            y=[float(clean_decimal(v)) for v in df_m1[secili_kolon][hist_start:n_hist]],
            mode='lines+markers', name='Gerçekleşen Resmi Veri', line=dict(color='#1f2937', width=4)
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
        fig_end.add_trace(go.Scatter(
            x=df_m6[ay_col][n_hist-1:],
            y=[float(clean_decimal(v)) for v in df_m6[secili_kolon][n_hist-1:]],
            mode='lines+markers', name=f"6. Baz Endeks Oranlı ({baz_endeks_sec_kolon})", line=dict(color='#f59e0b', width=2.5)
        ))
        fig_end.update_layout(
            title=f"{KOD_BILGI[secili_endeks_kodu]['ad']} — Geçmiş ve Gelecek Tahmin Eğrileri",
            xaxis_title="Aylar", yaxis_title="Endeks / Fiyat Değeri", template="simple_white", height=450, hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_end, use_container_width=True)

# ---------------- SEKME 3: FİYAT FARKI SONUÇLARI, MATRİS & MEVCUT DURUM AYRIŞTIRMASI ----------------
with main_tab3:
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
            fig_ff.add_trace(go.Scatter(
                x=df_aylik['Ay'], y=df_aylik['Kümülatif Fiyat Farkı'],
                name='Kümülatif Fiyat Farkı (Sağ Eksen)', yaxis='y2', mode='lines+markers',
                line=dict(color='#10b981', width=3)
            ))
            fig_ff.update_layout(
                title="Aylık İmalat, Fiyat Farkı ve Kümülatif Fiyat Farkı Gelişimi",
                yaxis=dict(title="Aylık Tutar (TL)"),
                yaxis2=dict(title="Kümülatif FF (TL)", overlaying='y', side='right'),
                barmode='group', template='simple_white', height=460, hovermode='x unified',
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_ff, use_container_width=True)

        with sub_t2:
            df_det_show = df_detay.copy()
            df_det_show['Kullanılan Tutar'] = df_det_show['Kullanılan Tutar'].apply(tr_format)
            df_det_show['Fiyat Farkı Tutarı'] = df_det_show['Fiyat Farkı Tutarı'].apply(tr_format)
            df_det_show['B Katsayısı'] = df_det_show['B Katsayısı'].apply(lambda x: tr_format(x, 2))
            df_det_show['Uygulanan Pn (15 Hane)'] = df_det_show['Uygulanan Pn (15 Hane)'].apply(
                lambda x: f"{x:.15f}".rstrip('0').rstrip('.').replace('.', ',')
            )
            st.dataframe(df_det_show, use_container_width=True, height=450)

        with sub_t3:
            if not df_pivot.empty:
                st.dataframe(
                    df_pivot.map(tr_format).style.set_properties(
                        subset=['HAKEDİŞ TUTARI (Toplam)'],
                        **{'font-weight': 'bold', 'background-color': '#e6f2ff'}
                    ),
                    use_container_width=True, height=450
                )

        with sub_t4:
            df_sonuc_gosterim = df_sonuc.copy()
            if 'AyKodu' in df_sonuc_gosterim.columns:
                df_sonuc_gosterim['AyKodu'] = df_sonuc_gosterim['AyKodu'].astype(str)
            for col in df_sonuc_gosterim.columns:
                if any(x in col.upper() for x in ['TUTAR', 'PROGRAM', 'FARKI']):
                    df_sonuc_gosterim[col] = df_sonuc_gosterim[col].apply(tr_format)
            st.dataframe(df_sonuc_gosterim, use_container_width=True, height=500)

# ---------------- SEKME 4: MODEL & SENARYO KARŞILAŞTIRMA ----------------
with main_tab4:
    st.subheader("⚖️ Tüm Tahmin Algoritmalarının ve Sabit Endeks Senaryosunun Kıyaslaması")
    st.caption("Mevcut durum + kalan ayların tamamlandığı senaryoda algoritmaların ürettiği Fiyat Farkı (FF) karşılaştırması:")

    df_prog_full, df_b_full, _ = hazirla_prog_ve_turetilen_ay(st.session_state.ff_prog_df, st.session_state.ff_b_df, tamamlama_aktif=True)

    if st.session_state.ff_base_endeks_df.empty:
        st.warning("⚠️ Karşılaştırma yapmak için geçmişe dönük en az 2 aylık endeks verisi gereklidir.")
    else:
        gereken_ufuk_comp = max(st.session_state.ff_ek_ay_slider, hesapla_gereken_ufuk(df_prog_full, st.session_state.ff_base_endeks_df))

        tum_modeller = [
            ("5. En İyi Karma (Dinamik Optimizasyon) 🌟", "Karma"),
            (f"6. Baz Endeks Oranlı Artış ({baz_endeks_sec_kolon} Bazlı) 📈", "Baz Endeks"),
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
                df_end_sim = gelismis_endeks_tahmini(
                    st.session_state.ff_base_endeks_df, gereken_ufuk_comp, m_kod,
                    alpha=alpha_p, beta=beta_p, phi=phi_p, mc_quantile=mc_q, df_alt=edited_alt,
                    baz_endeks_kolon=baz_endeks_sec_kolon, sabit_aylik_artis_yuzde=sabit_artis_oran
                )

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
        if not df_karsilastirma_out.empty:
            st.dataframe(
                df_karsilastirma_out.style.apply(
                    lambda x: ['background-color: #f0fdf4; font-weight: bold' if 'Karma' in str(x['Tahmin Modeli / Senaryo']) else '' for _ in x],
                    axis=1
                ),
                use_container_width=True, hide_index=True
            )
            fig_comp.update_layout(
                title="Algoritmalara Göre Kümülatif Fiyat Farkı Eğrileri",
                xaxis_title="Ay", yaxis_title="Kümülatif Fiyat Farkı (TL)",
                template="simple_white", height=450, hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_comp, use_container_width=True)

# ---------------- SEKME 5: AKADEMİK METODOLOJİ, TEORİ & BACKTESTING TABLOSU ----------------
with main_tab5:
    st.header("📚 Akademik Metodoloji, Backtesting (MAPE/RMSE) ve Matematiksel Çözüm Rehberi")
    st.markdown("---")
    st.subheader("🎯 1️⃣ Geçmiş Veriler Üzerinde Model Doğruluk Testi (Son 6 Ay Walk-Forward Backtesting)")
    st.info(
        "Aşağıdaki **Walk-Forward Backtesting** tablosu, her test ayında yalnızca o aya kadar olan geçmiş veriyi (`y[:t]`) kullanarak "
        "bir sonraki ayı (`y[t]`) tahmin eder ve gerçekleşen resmi verilerle (Nisan 2026 – Eylül 2026) karşılaştırır."
    )

    if not st.session_state.ff_base_endeks_df.empty and len(st.session_state.ff_base_endeks_df) > 6:
        df_backtest_out, detay_bt, karma_mape_dict, en_iyi_modeller_dict = backtest_modelleri_hesapla(
            st.session_state.ff_base_endeks_df, edited_alt, alpha=alpha_p, beta=beta_p, phi=phi_p
        )
        st.dataframe(df_backtest_out, use_container_width=True, hide_index=True)

        hibrit_err = karma_mape_dict.get("5. En İyi Karma (Hibrit)", 0.0)
        tekil_dict = {k: v for k, v in karma_mape_dict.items() if "Hibrit" not in k}
        en_iyi_tekil = min(tekil_dict, key=tekil_dict.get)

        c_top1, c_top2 = st.columns([1.1, 1.9])
        with c_top1:
            bt_secili_kod = st.selectbox(
                "🔬 Aylık Sapma Dökümü İncelenecek Alt Endeks:",
                options=list(KOD_BILGI.keys()),
                format_func=lambda k: f"{KOD_BILGI[k]['kolon']} — {KOD_BILGI[k]['ad']}"
            )
            bt_kolon = KOD_BILGI[bt_secili_kod]['kolon']
        with c_top2:
            sampiyon_kisa = detay_bt.get(bt_kolon, {}).get("Şampiyon Model", "")
            st.success(
                f"🌟 **5. Model (En İyi Karma / Hibrit) Ağırlıklı $P_n$ Hatası (MAPE): %{hibrit_err:.2f}**  |  "
                f"🏆 **En İyi Tekil Model:** {en_iyi_tekil} (**%{tekil_dict[en_iyi_tekil]:.2f}**)\n\n"
                f"Seçili **`{bt_kolon}`** serisi için en düşük hatayı veren şampiyon model: **🏆 {sampiyon_kisa}**"
            )

        if bt_kolon in detay_bt:
            bt_d = detay_bt[bt_kolon]
            fig_bt = go.Figure()
            fig_bt.add_trace(go.Scatter(
                x=bt_d["Aylar"], y=bt_d["Gerçekleşen"],
                mode='lines+markers', name='Gerçekleşen Resmi Veri', line=dict(color='#111827', width=4)
            ))
            fig_bt.add_trace(go.Scatter(
                x=bt_d["Aylar"], y=bt_d["1. Holt's Çift Üstel"],
                mode='lines+markers', name="1. Holt's Çift Üstel", line=dict(color='#2563eb', width=2, dash='dash')
            ))
            fig_bt.add_trace(go.Scatter(
                x=bt_d["Aylar"], y=bt_d["2. Sönümlü Trend (Damped)"],
                mode='lines+markers', name="2. Sönümlü Trend", line=dict(color='#10b981', width=2, dash='dot')
            ))
            fig_bt.add_trace(go.Scatter(
                x=bt_d["Aylar"], y=bt_d["3. Monte Carlo GBM (P50 Medyan)"],
                mode='lines+markers', name="3. Monte Carlo GBM (P50)", line=dict(color='#dc2626', width=2, dash='dashdot')
            ))
            fig_bt.add_trace(go.Scatter(
                x=bt_d["Aylar"], y=bt_d["4. Otoregresif (AR-1 & EWMA)"],
                mode='lines+markers', name="4. Otoregresif (AR-1)", line=dict(color='#9333ea', width=2)
            ))
            st.plotly_chart(
                fig_bt.update_layout(
                    title=f"{KOD_BILGI[bt_secili_kod]['ad']} — Son 6 Ay Walk-Forward Backtest Grafiği (Nisan 2026 – Eylül 2026)",
                    xaxis_title="Test Ayları", yaxis_title="Endeks / Fiyat Değeri",
                    template="simple_white", height=400, hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                ),
                use_container_width=True
            )

            aylik_sapma_rows = []
            for idx_m, ay_adi in enumerate(bt_d["Aylar"]):
                g_val = max(float(bt_d["Gerçekleşen"][idx_m]), 1e-6)
                dec_fmt = 4 if g_val < 100 else 2
                h_val = bt_d["1. Holt's Çift Üstel"][idx_m]
                s_val = bt_d["2. Sönümlü Trend (Damped)"][idx_m]
                mc_val = bt_d["3. Monte Carlo GBM (P50 Medyan)"][idx_m]
                ar_val = bt_d["4. Otoregresif (AR-1 & EWMA)"][idx_m]
                k_val = bt_d["5. En İyi Karma (Şampiyon)"][idx_m]
                aylik_sapma_rows.append({
                    "Test Ayı": ay_adi,
                    "Gerçekleşen": tr_format(g_val, dec_fmt),
                    "1. Holt": f"{tr_format(h_val, dec_fmt)} (Sapma: %{(abs(g_val - h_val) / g_val * 100):.2f})",
                    "2. Sönümlü Trend": f"{tr_format(s_val, dec_fmt)} (Sapma: %{(abs(g_val - s_val) / g_val * 100):.2f})",
                    "3. Monte Carlo GBM (P50)": f"{tr_format(mc_val, dec_fmt)} (Sapma: %{(abs(g_val - mc_val) / g_val * 100):.2f})",
                    "4. Otoregresif (AR-1)": f"{tr_format(ar_val, dec_fmt)} (Sapma: %{(abs(g_val - ar_val) / g_val * 100):.2f})",
                    f"🏆 5. En İyi Karma ({sampiyon_kisa})": f"{tr_format(k_val, dec_fmt)} (Sapma: %{(abs(g_val - k_val) / g_val * 100):.2f})",
                })
            st.caption(f"📋 **{bt_kolon} Alt Endeksi — Son 6 Ay Aylık Walk-Forward Tahmin ve Sapma Dökümü (Tam Genişlik):**")
            st.dataframe(pd.DataFrame(aylik_sapma_rows), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("📐 2️⃣ Dinamik İşçilik (3B/4B), KDV Arındırılmış EPDK Yakıt ve Fiyat Farkı Denklemleri")
    with st.expander("📖 Matematiksel Model ve Dinamik Geçiş Kurallarını Görüntüle", expanded=True):
        st.markdown("#### 1. İşçilik Endeksi (`I o`) 3 Basamaklı vs 4 Basamaklı Dinamik Seçim Kuralı")
        st.markdown(
            "- **3 Basamaklı Temel Endeks (`< 500`, ör. `100,421925`):** Doğrudan TÜİK **2025=100** bazlı TÜFE Genel Endeksi serisi (`100,421925` → `137,23`) kullanılır.\n"
            "- **4 Basamaklı Temel Endeks (`>= 500`, ör. `3.196,66`):** TÜİK **2003=100** bazlı TÜFE Genel Endeksi serisi kullanılır. Mayıs 2026 sonrasında yayından kaldırılan aylar (Haziran 2026+), Excel'deki `=+(I14/H14)*H12` zincirleme formülüyle türetilir:"
        )
        st.latex(r"I_{2003,\,t} = I_{2003,\,t-1} \times \frac{I_{2025,\,t}}{I_{2025,\,t-1}}")

        st.markdown("#### 2. Yakıt Endeksi (`Y o`) EPDK Motorin (KDV Hariç) vs Yapım İşleri TÜİK 19 Seçim Kuralı")
        st.markdown(
            "- **EPDK Ölçeğinde Temel Endeks (`< 500`, ör. `44,284425`):** `epdk_akaryakit_fiyatlari.xlsx` içerisindeki EPDK Aylık Ortalama Motorin Bayi Satış Fiyatları kullanılır. "
            "10 Temmuz 2023 tarihinde yürürlüğe giren **7346 sayılı Cumhurbaşkanı Kararı** uyarınca:\n"
            "  - **Ocak 2021 – Haziran 2023:** KDV oranı **%18** olduğundan fiyatlar `1,18`'e bölünür (ör. Ocak 2021: `6,58428 / 1,18 = 5,579898`).\n"
            "  - **Temmuz 2023 ve Sonrası:** KDV oranı **%20** olduğundan fiyatlar `1,20`'ye bölünür (ör. Temmuz 2025: `53,14131 / 1,20 = 44,284425`).\n"
            "  - **2021 Öncesi (2016–2020):** Ocak 2021 KDV hariç EPDK Motorin fiyatına (`5,579898`) TÜİK 19 nolu alt endeksin aylık değişim oranlarıyla geriye doğru orantılanır.\n"
            "- **Yapım İşleri Ölçeğinde Temel Endeks (`>= 500`, ör. `7.767,87`):** Doğrudan TÜİK 19 nolu **Kok ve Rafine Petrol Ürünleri (2003=100)** endeksi (`7.767,87` → `17.024,32`) kullanılır."
        )
        st.latex(r"Y_{\text{EPDK, KDV Hariç},\,t} = \begin{cases} \dfrac{\text{Motorin}_t}{1{,}18} & \text{Haziran 2023 ve öncesi (\%18 KDV)} \\[8pt] \dfrac{\text{Motorin}_t}{1{,}20} & \text{Temmuz 2023 ve sonrası (\%20 KDV)} \end{cases}")

        st.markdown("#### 3. 4734 Sayılı Kamu İhale Sözleşmeleri Fiyat Farkı ve Gecikme (Kova) Kuralı")
        st.latex(r"F_n = A_n \times B \times (P_n - 1), \qquad P_n = a\,\frac{I_n}{I_o} + b_1\,\frac{C_n}{C_o} + b_2\,\frac{D_n}{D_o} + b_3\,\frac{Y_n}{Y_o} + b_4\,\frac{K_n}{K_o} + b_5\,\frac{G_n}{G_o} + c\,\frac{M_n}{M_o}")
        st.markdown(
            "- **İş Programı Gecikmesi (Kova/Dilim Kuralı):** Yüklenicinin iş programının gerisinde kaldığı aylarda, geciken ödenek dilimi için uygulama ayı endeksi ile iş programı ayı endeksinden **küçük olanı** (`min(E_uyg, E_prog)`) esas alınır."
        )

# ==========================================
# 10. YAN MENÜ RAPOR VE ÇIKTI İNDİRME MERKEZİ
# ==========================================
with st.sidebar:
    st.markdown("---")
    st.header("📤 Rapor & Veri İndirme Merkezi")
    try:
        excel_bytes = generate_excel_download(
            edited_prog, edited_endeks, edited_alt, edited_b,
            df_detay=df_detay, df_karsilastirma=df_karsilastirma_out, df_backtest=df_backtest_out
        )
        st.download_button(
            label="📊 Kapsamlı Excel Raporu İndir (.xlsx)",
            data=excel_bytes,
            file_name="Fiyat_Farki_Simulasyon_Raporu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.caption(f"Excel hazırlanamadı: {e}")

    try:
        word_bytes = generate_word_report(
            model_secimi, toplam_sozlesme, toplam_gerceklesen_iml, toplam_ff_tutari,
            mevcut_ff_tutari, kalan_ff_tutari, df_karsilastirma_out, df_backtest_out, df_detay
        )
        st.download_button(
            label="📝 Yönetici Özeti Word Raporu (.doc)",
            data=word_bytes,
            file_name="Teknik_Ofis_Fiyat_Farki_Raporu.doc",
            mime="application/msword",
            use_container_width=True
        )
    except Exception as e:
        st.caption(f"Word raporu hazırlanamadı: {e}")
