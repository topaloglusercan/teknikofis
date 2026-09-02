import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter, FuncFormatter
from matplotlib.backends.backend_pdf import PdfPages
import re
import io

st.markdown("<h1 style='color: #2c3e50;'>📈 P6 S-Eğrisi (Parasal İlerleme)</h1>", unsafe_allow_html=True)
st.markdown("XER dosyasındaki aktivite ve kaynak verilerini tarayarak projeye ait aylık ve kümülatif S-Eğrisi (Nakit Akışı / İlerleme) grafiklerini oluşturun.")

def clean_dec_float(val):
    try:
        if pd.isna(val) or val is None: return 0.0
        s = str(val).strip()
        if not s: return 0.0
        if '.' in s and ',' in s:
            if s.rfind(',') > s.rfind('.'):
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '')
        elif ',' in s:
            s = s.replace(',', '.')
        return float(s)
    except:
        return 0.0

def format_tl(x, p=None):
    if x >= 1e9: return f"{x/1e9:.2f} Milyar"
    elif x >= 1e6: return f"{x/1e6:.1f} Milyon"
    elif x >= 1e3: return f"{x/1e3:.0f} Bin"
    else: return f"{int(x)}"

@st.cache_data
def parse_xer_for_scurve(file_bytes):
    text = file_bytes.decode('windows-1254', errors='ignore')
    lines = text.splitlines()
    tables = {}
    current_table = None
    columns = []
    
    for line in lines:
        if line.startswith('%T'):
            current_table = line.split('\t')[1].strip()
            tables[current_table] = []
        elif line.startswith('%F') and current_table:
            columns = [col.strip().lower() for col in line.split('\t')[1:]]
        elif line.startswith('%R') and current_table:
            values = line.split('\t')[1:]
            row_dict = {columns[i]: values[i].strip() if i < len(values) else "" for i in range(len(columns))}
            tables[current_table].append(row_dict)
            
    def clean_df(data_list):
        df = pd.DataFrame(data_list)
        if not df.empty:
            for col in df.columns: df[col] = df[col].astype(str).str.strip()
        return df

    return clean_df(tables.get('TASK', [])), clean_df(tables.get('TASKRSRC', [])), clean_df(tables.get('RSRC', []))

uploaded_xer = st.file_uploader("📂 XER Dosyasını Yükle (S-Eğrisi İçin)", type=['xer'])

if uploaded_xer:
    with st.spinner("P6 Veritabanı Taranıyor ve S-Eğrisi Hesaplanıyor..."):
        df_task, df_taskrsrc, df_rsrc = parse_xer_for_scurve(uploaded_xer.getvalue())
        
        if not df_task.empty and not df_taskrsrc.empty:
            t_cols = [c for c in ['task_id', 'target_start_date', 'target_end_date', 'early_start_date', 'early_end_date', 'act_start_date', 'act_end_date'] if c in df_task.columns]
            df_t = df_task[t_cols].copy()
            df_t = df_t.loc[:, ~df_t.columns.duplicated()].drop_duplicates(subset=['task_id'])
            df_t['task_id'] = df_t['task_id'].astype(str)
            
            df_tr = df_taskrsrc.copy()
            df_tr = df_tr.loc[:, ~df_tr.columns.duplicated()]
            df_tr['task_id'] = df_tr['task_id'].astype(str)
            
            cost_col = 'target_cost' if 'target_cost' in df_tr.columns else 'target_qty'
            if cost_col not in df_tr.columns:
                df_tr[cost_col] = 0.0
                
            df_merged = pd.merge(df_tr, df_t, on='task_id', how='inner')
            df_merged = df_merged.loc[:, ~df_merged.columns.duplicated()]
            
            if 'rsrc_id' in df_merged.columns:
                df_merged['rsrc_id'] = df_merged['rsrc_id'].astype(str)
                
                if not df_rsrc.empty and 'rsrc_id' in df_rsrc.columns:
                    df_r = df_rsrc.copy()
                    df_r = df_r.loc[:, ~df_r.columns.duplicated()].drop_duplicates(subset=['rsrc_id'])
                    df_r['rsrc_id'] = df_r['rsrc_id'].astype(str)
                    
                    name_col = next((c for c in ['rsrc_name', 'rsrc_short_name'] if c in df_r.columns), None)
                    
                    if name_col:
                        isim_sozlugu = df_r.set_index('rsrc_id')[name_col].to_dict()
                        df_merged['Kaynak_Adi'] = df_merged['rsrc_id'].map(isim_sozlugu).fillna('İsimsiz').astype(str)
                        df_merged['Kaynak_Gorunum'] = df_merged['rsrc_id'] + " - " + df_merged['Kaynak_Adi']
                    else:
                        df_merged['Kaynak_Gorunum'] = df_merged['rsrc_id']
                else:
                    df_merged['Kaynak_Gorunum'] = df_merged['rsrc_id']
            else:
                df_merged['Kaynak_Gorunum'] = 'Tüm Kaynaklar'

            df_merged['Baslangic'] = pd.NaT
            df_merged['Bitis'] = pd.NaT
            
            for col in ['target_start_date', 'early_start_date', 'act_start_date']:
                if col in df_merged.columns: df_merged['Baslangic'] = df_merged['Baslangic'].fillna(pd.to_datetime(df_merged[col], errors='coerce'))

            for col in ['target_end_date', 'early_end_date', 'act_end_date']:
                if col in df_merged.columns: df_merged['Bitis'] = df_merged['Bitis'].fillna(pd.to_datetime(df_merged[col], errors='coerce'))

            df_merged['Maliyet'] = df_merged[cost_col].apply(clean_dec_float)
            df_valid = df_merged.dropna(subset=['Baslangic', 'Bitis']).copy()
            df_valid = df_valid[df_valid['Maliyet'] > 0]
            
            if not df_valid.empty:
                kaynak_listesi = sorted([str(x) for x in df_valid['Kaynak_Gorunum'].unique()])
                secilen_kaynaklar = st.multiselect("📊 Analiz Edilecek Kaynakları Seçin (Örn: Excel'deki 'PARA' kaynağını arayıp seçin):", kaynak_listesi, default=[])
                
                if secilen_kaynaklar:
                    df_valid = df_valid[df_valid['Kaynak_Gorunum'].isin(secilen_kaynaklar)]
                
                df_valid['Toplam_Gun'] = (df_valid['Bitis'] - df_valid['Baslangic']).dt.days + 1
                df_valid['Toplam_Gun'] = df_valid['Toplam_Gun'].clip(lower=1).astype(int)
                df_valid['Gunluk_Maliyet'] = df_valid['Maliyet'] / df_valid['Toplam_Gun']
                
                df_spread = df_valid.loc[df_valid.index.repeat(df_valid['Toplam_Gun'])].copy()
                days_to_add = pd.to_timedelta(df_spread.groupby(level=0).cumcount(), unit='D')
                df_spread['Tarih'] = df_spread['Baslangic'] + days_to_add
                
                df_spread['Ay Sonu'] = df_spread['Tarih'] + pd.offsets.MonthEnd(0)
                df_monthly = df_spread.groupby('Ay Sonu')['Gunluk_Maliyet'].sum().reset_index()
                df_monthly.rename(columns={'Gunluk_Maliyet': 'Aylık Maliyet (TL)'}, inplace=True)
                
                total_budget = df_monthly['Aylık Maliyet (TL)'].sum()
                df_monthly['Kümülatif Maliyet (TL)'] = df_monthly['Aylık Maliyet (TL)'].cumsum()
                df_monthly['Aylık İlerleme (%)'] = (df_monthly['Aylık Maliyet (TL)'] / total_budget) * 100 if total_budget > 0 else 0
                df_monthly['Kümülatif İlerleme (%)'] = (df_monthly['Kümülatif Maliyet (TL)'] / total_budget) * 100 if total_budget > 0 else 0
                
                df_monthly['Ay Sonu Gösterim'] = df_monthly['Ay Sonu'].dt.strftime('%m-%Y')
                
                for col in ['Aylık Maliyet (TL)', 'Kümülatif Maliyet (TL)', 'Aylık İlerleme (%)', 'Kümülatif İlerleme (%)']:
                    df_monthly[col] = df_monthly[col].round(2)

                st.success(f"✅ Dağıtım Tamamlandı! Toplam Maliyet/Bütçe: **{total_budget:,.2f} TL**")
                
                # --- A3 GRAFİK ÇİZİMİ ---
                plt.style.use('ggplot')
                fig, (ax1, ax3) = plt.subplots(2, 1, figsize=(16.5, 11.7)) # A3 Landscape Boyutları
                fig.suptitle('Proje İlerleme ve Maliyet (S-Eğrisi) Raporu', fontsize=20, fontweight='bold', y=0.98)
                
                bars_pct = ax1.bar(df_monthly['Ay Sonu Gösterim'], df_monthly['Aylık İlerleme (%)'], color='#4C72B0', alpha=0.6, edgecolor='black', label='Aylık İlerleme (%)')
                ax1.set_ylabel('Aylık İlerleme (%)', color='#4C72B0', fontweight='bold', fontsize=12)
                ax1.tick_params(axis='y', labelcolor='#4C72B0', labelsize=10)
                ax1.set_xticklabels(df_monthly['Ay Sonu Gösterim'], rotation=45, ha='right', fontsize=10)
                ax1.yaxis.set_major_formatter(PercentFormatter())

                ax2 = ax1.twinx()
                line_pct = ax2.plot(df_monthly['Ay Sonu Gösterim'], df_monthly['Kümülatif İlerleme (%)'], color='#C44E52', marker='o', linewidth=3, markersize=8, label='Kümülatif İlerleme (%)')
                ax2.set_ylabel('Kümülatif İlerleme (%)', color='#C44E52', fontweight='bold', fontsize=12)
                ax2.tick_params(axis='y', labelcolor='#C44E52', labelsize=10)
                ax2.yaxis.set_major_formatter(PercentFormatter())

                for i, txt in enumerate(df_monthly['Kümülatif İlerleme (%)']):
                    ax2.annotate(f"{txt:.1f}%", (i, df_monthly['Kümülatif İlerleme (%)'][i]), textcoords="offset points", xytext=(5,10), ha='left', va='bottom', rotation=45, fontsize=10, fontweight='bold', color='#C44E52')

                for i, bar in enumerate(bars_pct):
                    val = df_monthly['Aylık İlerleme (%)'].iloc[i]
                    if val > 0.1:
                        ax1.annotate(f"{val:.1f}%", (bar.get_x() + bar.get_width() / 2, 0), textcoords="offset points", xytext=(0, 5), ha='center', va='bottom', rotation=90, fontsize=10, fontweight='bold', color='black')

                lines_1, labels_1 = ax1.get_legend_handles_labels()
                lines_2, labels_2 = ax2.get_legend_handles_labels()
                ax2.legend(lines_1 + lines_2, labels_1 + labels_2, loc='upper left', fontsize=10)

                bars_tl = ax3.bar(df_monthly['Ay Sonu Gösterim'], df_monthly['Aylık Maliyet (TL)'], color='#55A868', alpha=0.6, edgecolor='black', label='Aylık Maliyet')
                ax3.set_ylabel('Aylık Maliyet', color='#55A868', fontweight='bold', fontsize=12)
                ax3.tick_params(axis='y', labelcolor='#55A868', labelsize=10)
                ax3.set_xticklabels(df_monthly['Ay Sonu Gösterim'], rotation=45, ha='right', fontsize=10)
                ax3.yaxis.set_major_formatter(FuncFormatter(format_tl))

                ax4 = ax3.twinx()
                line_tl = ax4.plot(df_monthly['Ay Sonu Gösterim'], df_monthly['Kümülatif Maliyet (TL)'], color='#8C8C8C', marker='s', linewidth=3, markersize=8, label='Kümülatif Maliyet')
                ax4.set_ylabel('Kümülatif Maliyet', color='#8C8C8C', fontweight='bold', fontsize=12)
                ax4.tick_params(axis='y', labelcolor='#8C8C8C', labelsize=10)
                ax4.yaxis.set_major_formatter(FuncFormatter(format_tl))

                for i, txt in enumerate(df_monthly['Kümülatif Maliyet (TL)']):
                    ax4.annotate(f"{format_tl(txt)}", (i, df_monthly['Kümülatif Maliyet (TL)'][i]), textcoords="offset points", xytext=(5,10), ha='left', va='bottom', rotation=45, fontsize=10, fontweight='bold', color='#8C8C8C')

                for i, bar in enumerate(bars_tl):
                    val = df_monthly['Aylık Maliyet (TL)'].iloc[i]
                    if val > (total_budget * 0.005):
                        ax3.annotate(f"{format_tl(val)}", (bar.get_x() + bar.get_width() / 2, 0), textcoords="offset points", xytext=(0, 5), ha='center', va='bottom', rotation=90, fontsize=10, fontweight='bold', color='black')

                lines_3, labels_3 = ax3.get_legend_handles_labels()
                lines_4, labels_4 = ax4.get_legend_handles_labels()
                ax4.legend(lines_3 + lines_4, labels_3 + labels_4, loc='upper left', fontsize=10)

                plt.tight_layout(rect=[0, 0, 1, 0.96])
                st.pyplot(fig)
                
                # --- PDF OLUŞTURMA (Grafik ve Tablo) ---
                pdf_buffer = io.BytesIO()
                with PdfPages(pdf_buffer) as pdf:
                    # 1. Sayfa: A3 Grafik
                    pdf.savefig(fig, bbox_inches='tight')
                    
                    # 2. Sayfa: A3 Veri Tablosu
                    fig_table, ax_table = plt.subplots(figsize=(16.5, 11.7))
                    ax_table.axis('tight')
                    ax_table.axis('off')
                    ax_table.set_title("Proje İlerleme ve Maliyet Dağılım Tablosu", fontsize=20, fontweight='bold', pad=20)
                    
                    table_cols = ['Ay Sonu Gösterim', 'Aylık Maliyet (TL)', 'Kümülatif Maliyet (TL)', 'Aylık İlerleme (%)', 'Kümülatif İlerleme (%)']
                    table_data = df_monthly[table_cols].copy()
                    table_data['Aylık İlerleme (%)'] = table_data['Aylık İlerleme (%)'].astype(str) + ' %'
                    table_data['Kümülatif İlerleme (%)'] = table_data['Kümülatif İlerleme (%)'].astype(str) + ' %'
                    table_data['Aylık Maliyet (TL)'] = table_data['Aylık Maliyet (TL)'].apply(lambda x: f"{x:,.2f} ₺")
                    table_data['Kümülatif Maliyet (TL)'] = table_data['Kümülatif Maliyet (TL)'].apply(lambda x: f"{x:,.2f} ₺")
                    
                    table = ax_table.table(cellText=table_data.values, colLabels=table_data.columns, loc='center', cellLoc='center')
                    table.scale(1, 2.5) # Satır aralıklarını genişlet
                    table.auto_set_font_size(False)
                    table.set_fontsize(12)
                    
                    # Başlık hücrelerini şekillendirme
                    for (row, col), cell in table.get_celld().items():
                        if row == 0:
                            cell.set_text_props(weight='bold', color='white')
                            cell.set_facecolor('#2c3e50')
                        else:
                            cell.set_facecolor('#f8f9fa' if row % 2 == 0 else 'white')
                            
                    pdf.savefig(fig_table, bbox_inches='tight')
                    plt.close('all')
                
                st.markdown("### 📋 Dağıtım Tablosu (Rapor)")
                st.dataframe(df_monthly[['Ay Sonu Gösterim', 'Aylık Maliyet (TL)', 'Kümülatif Maliyet (TL)', 'Aylık İlerleme (%)', 'Kümülatif İlerleme (%)']], use_container_width=True)
                
                # --- İNDİRME BUTONLARI ---
                st.markdown("---")
                c1, c2 = st.columns(2)
                with c1:
                    pdf_bytes = pdf_buffer.getvalue()
                    st.download_button(label="📑 A3 PDF Raporu İndir", data=pdf_bytes, file_name='S_Egrisi_A3_Rapor.pdf', mime='application/pdf', use_container_width=True)
                with c2:
                    csv = df_monthly.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(label="💾 Excel/CSV Olarak İndir", data=csv, file_name='S_Egrisi_Raporu.csv', mime='text/csv', use_container_width=True)

            else:
                st.warning("⚠️ Seçilen dosyadaki aktivitelerde geçerli tarih veya bütçe/maliyet (Cost/Qty) değeri bulunamadı.")
        else:
            st.error("⚠️ XER dosyasında Kaynak Atamaları (TASKRSRC) veya Aktivite (TASK) tablosu okunamadı.")
