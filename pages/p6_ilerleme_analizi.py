import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import re
from datetime import datetime

st.set_page_config(page_title="P6 Matris Dashboard", layout="wide")

# --- SOFT UI & PASTEL CSS TASARIMI ---
st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    .soft-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 20px 10px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04);
        margin-bottom: 20px;
        text-align: center;
        border: 1px solid #f0f2f5;
        transition: transform 0.2s ease-in-out;
        height: 140px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .soft-card:hover { transform: translateY(-5px); box-shadow: 0 8px 25px rgba(0, 0, 0, 0.08); }
    .soft-card h3 { color: #8898aa; font-size: 12px; font-weight: 600; text-transform: uppercase; margin-bottom: 10px; }
    .soft-card h2 { font-size: 26px; font-weight: 700; margin: 0; }
    
    .val-blue { color: #5e72e4; }
    .val-green { color: #2dce89; }
    .val-red { color: #f5365c; }
    .val-orange { color: #fb6340; }
    .val-teal { color: #11cdef; }
    
    .kpi-note { font-size: 10px; color: #adb5bd; margin-top: 8px; display: block; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; font-size: 16px; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

def parse_xer(file_content):
    tables = {}
    current_table = None
    for line in file_content.splitlines():
        parts = line.split("\t")
        row_type = parts[0]
        if row_type == "%T":
            current_table = parts[1].strip()
            tables[current_table] = {"columns": [], "rows": []}
        elif row_type == "%F" and current_table:
            tables[current_table]["columns"] = parts[1:]
        elif row_type == "%R" and current_table:
            row_data = parts[1:]
            col_count = len(tables[current_table]["columns"])
            if len(row_data) < col_count:
                row_data.extend([''] * (col_count - len(row_data)))
            tables[current_table]["rows"].append(row_data[:col_count])
    return tables

def extract_kat(text):
    match = re.search(r'([1-9]BK|ZK|[1-9]NK)', str(text).upper())
    return match.group(1) if match else "Genel / Kapsam Dışı"

st.markdown("<h2 style='color: #32325d; margin-bottom: 20px;'>📊 Proje Performans ve İlerleme Paneli</h2>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("XER Dosyasını Yükle", type=["xer"])

if uploaded_file is not None:
    content = uploaded_file.getvalue().decode("windows-1254", errors="ignore")
    xer_data = parse_xer(content)
    
    if "TASK" in xer_data and "PROJECT" in xer_data:
        df_task = pd.DataFrame(xer_data["TASK"]["rows"], columns=xer_data["TASK"]["columns"])
        df_wbs = pd.DataFrame(xer_data["PROJWBS"]["rows"], columns=xer_data["PROJWBS"]["columns"])
        df_proj = pd.DataFrame(xer_data["PROJECT"]["rows"], columns=xer_data["PROJECT"]["columns"])
        
        # Data Date (Veri Tarihi) Okuma
        data_date_str = df_proj['last_recalc_date'].iloc[0] if 'last_recalc_date' in df_proj.columns and not df_proj.empty else None
        data_date = pd.to_datetime(data_date_str, errors='coerce')

        num_cols = ['phys_complete_pct', 'target_work_qty', 'act_work_qty', 'remain_work_qty', 
                    'target_equip_qty', 'act_equip_qty', 'target_drtn_hr_cnt', 'remain_drtn_hr_cnt']
        
        for col in num_cols:
            df_task[col] = pd.to_numeric(df_task.get(col, 0), errors='coerce').fillna(0)

        # Tarih Sütunları
        date_cols = ['target_start_date', 'target_end_date']
        for col in date_cols:
            if col in df_task.columns:
                df_task[col] = pd.to_datetime(df_task[col], errors='coerce')

        # EV Hesaplamaları
        df_task['EV_NonLabor'] = df_task['target_equip_qty'] * (df_task['phys_complete_pct'] / 100)
        df_task['EV_Labor'] = df_task['target_work_qty'] * (df_task['phys_complete_pct'] / 100)
        
        # PV Hesaplaması (Lineer Yaklaşım)
        def calculate_pv(row):
            start = row.get('target_start_date')
            end = row.get('target_end_date')
            budget_labor = row.get('target_work_qty', 0)
            budget_nonlabor = row.get('target_equip_qty', 0)
            total_budget = budget_labor + budget_nonlabor
            
            if pd.isna(start) or pd.isna(end) or pd.isna(data_date):
                return 0
            
            if data_date >= end:
                return total_budget
            elif data_date <= start:
                return 0
            else:
                total_duration = (end - start).total_seconds()
                elapsed_duration = (data_date - start).total_seconds()
                if total_duration > 0:
                    return total_budget * (elapsed_duration / total_duration)
                return 0

        df_task['PV_Linear'] = df_task.apply(calculate_pv, axis=1)

        df_merged = pd.merge(df_task, df_wbs[['wbs_id', 'wbs_name']], on='wbs_id', how='left')
        
        cat_columns = []
        if "TASKACTV" in xer_data and "ACTVCODE" in xer_data and "ACTVTYPE" in xer_data:
            df_taskactv = pd.DataFrame(xer_data["TASKACTV"]["rows"], columns=xer_data["TASKACTV"]["columns"])
            df_actvcode = pd.DataFrame(xer_data["ACTVCODE"]["rows"], columns=xer_data["ACTVCODE"]["columns"])
            df_actvtype = pd.DataFrame(xer_data["ACTVTYPE"]["rows"], columns=xer_data["ACTVTYPE"]["columns"])
            
            code_type_merged = pd.merge(df_actvcode, df_actvtype[['actv_code_type_id', 'actv_code_type']], on='actv_code_type_id')
            task_codes = pd.merge(df_taskactv, code_type_merged[['actv_code_id', 'actv_code_name', 'actv_code_type']], on='actv_code_id')
            
            pivot_codes = task_codes.pivot(index='task_id', columns='actv_code_type', values='actv_code_name').reset_index()
            pivot_codes = pivot_codes.fillna("Atanmamış") 
            df_merged = pd.merge(df_merged, pivot_codes, on='task_id', how='left')
            
            temp_cols = [col for col in pivot_codes.columns if col != 'task_id' and not str(col).startswith('Code-')]
            for col in temp_cols:
                if str(col).startswith('Global_'):
                    if str(col).replace('Global_', '') in temp_cols:
                        continue 
                cat_columns.append(col)

        upper_cats = [c.upper() for c in cat_columns]
        df_merged['Arama_Metni'] = df_merged['wbs_name'].fillna('') + " " + df_merged['task_name'].fillna('')
        
        if 'KAT' not in upper_cats:
            df_merged['KAT'] = df_merged['Arama_Metni'].apply(extract_kat)
            cat_columns.append('KAT')
            
        cat_columns = list(dict.fromkeys(cat_columns))
        
        unwanted = ['SÖZLEŞME']
        cat_columns = [c for c in cat_columns if c.upper() not in unwanted]
        
        preferred_order = ['ADA/PARSEL', 'BLOK', 'LVL1', 'LVL2', 'LVL3', 'KAT', 'DİSİPLİN']
        ordered_cat_columns = [col for col in preferred_order if col in cat_columns]
        ordered_cat_columns.extend([col for col in cat_columns if col not in ordered_cat_columns])
        cat_columns = ordered_cat_columns

        # --- DİNAMİK FİLTRELER ---
        st.markdown("<h4 style='color: #525f7f;'>Filtreler (Çoklu Seçim)</h4>", unsafe_allow_html=True)
        active_filters = {}
        num_filters = min(len(cat_columns), 8)
        
        cols1 = st.columns(4)
        for i in range(min(4, num_filters)):
            c_name = cat_columns[i]
            with cols1[i]:
                u_vals = sorted([str(x) for x in df_merged[c_name].dropna().unique()])
                active_filters[c_name] = st.multiselect(f"{c_name}", u_vals, placeholder="Tümü (Boş Bırakın)")
                
        if num_filters > 4:
            cols2 = st.columns(4)
            for i in range(4, num_filters):
                c_name = cat_columns[i]
                with cols2[i-4]:
                    u_vals = sorted([str(x) for x in df_merged[c_name].dropna().unique()])
                    active_filters[c_name] = st.multiselect(f"{c_name}", u_vals, placeholder="Tümü (Boş Bırakın)")
                
        df_filtered = df_merged.copy()
        for col_name, selected_vals in active_filters.items():
            if len(selected_vals) > 0:
                df_filtered = df_filtered[df_filtered[col_name].isin(selected_vals)]

        df_filtered['Total_Target'] = df_filtered['target_work_qty'] + df_filtered['target_equip_qty']
        df_filtered['Total_EV'] = df_filtered['EV_Labor'] + df_filtered['EV_NonLabor']
        df_filtered['Total_AC'] = df_filtered['act_work_qty'] + df_filtered['act_equip_qty']

        tot_target_all = df_filtered['Total_Target'].sum()
        tot_ev_all = df_filtered['Total_EV'].sum()
        tot_ac_all = df_filtered['Total_AC'].sum()
        tot_pv_all = df_filtered['PV_Linear'].sum()
        
        perf_pct = (tot_ev_all / tot_target_all * 100) if tot_target_all > 0 else 0
        
        tot_target_drtn = df_filtered['target_drtn_hr_cnt'].sum()
        tot_remain_drtn = df_filtered['remain_drtn_hr_cnt'].sum()
        dur_pct = ((tot_target_drtn - tot_remain_drtn) / tot_target_drtn * 100) if tot_target_drtn > 0 else 0
        
        tot_act_lab = df_filtered['act_work_qty'].sum()
        tot_at_comp_lab = tot_act_lab + df_filtered['remain_work_qty'].sum()
        lab_pct = (tot_act_lab / tot_at_comp_lab * 100) if tot_at_comp_lab > 0 else 0

        cpi = (tot_ev_all / tot_ac_all) if tot_ac_all > 0 else 0.0
        cpi_color = "val-green" if cpi >= 1.0 else ("val-orange" if cpi >= 0.85 else "val-red")
        
        spi = (tot_ev_all / tot_pv_all) if tot_pv_all > 0 else 0.0
        spi_color = "val-green" if spi >= 1.0 else ("val-orange" if spi >= 0.85 else "val-red")

        tab_yuzde, tab_para = st.tabs(["📊 Yüzdesel (Performans) Görünümü", "💰 Parasal / Miktar Görünümü"])
        
        show_cols = ['task_code', 'task_name', 'status_code', 'phys_complete_pct', 'Total_Target', 'Total_EV', 'Total_AC']
        if 'DİSİPLİN' in df_filtered.columns:
            show_cols.insert(2, 'DİSİPLİN')
            
        rename_dict = {'task_code': 'Aktivite Kodu', 'task_name': 'Aktivite Adı', 'status_code': 'Durum', 
                       'phys_complete_pct': 'Fiziksel (%)', 'Total_Target': 'Planlanan/Bütçe', 'Total_EV': 'Kazanılmış (EV)', 'Total_AC': 'Gerçekleşen (AC)'}

        # ==========================================
        # 1. SEKME: YÜZDESEL (PERCENTAGE)
        # ==========================================
        with tab_yuzde:
            st.write("")
            ky1, ky2, ky3, ky4, ky5 = st.columns(5)
            
            ky1.markdown(f"""<div class="soft-card"><h3>Performance % Complete</h3><h2 class="val-blue">%{perf_pct:.2f}</h2><span class="kpi-note">Bütçe Ağırlıklı İlerleme (EV / BAC)</span></div>""", unsafe_allow_html=True)
            ky2.markdown(f"""<div class="soft-card"><h3>Duration % Complete</h3><h2 class="val-teal">%{dur_pct:.2f}</h2><span class="kpi-note">Zaman Odaklı (Süre İlerlemesi)</span></div>""", unsafe_allow_html=True)
            ky3.markdown(f"""<div class="soft-card"><h3>Labor Units % Complete</h3><h2 class="val-teal">%{lab_pct:.2f}</h2><span class="kpi-note">Adam-Saat Tüketim Oranı</span></div>""", unsafe_allow_html=True)
            
            ky4.markdown(f"""<div class="soft-card"><h3>Cost Perf. Index (CPI)</h3><h2 class="{cpi_color}">{cpi:.2f}</h2><span class="kpi-note">Maliyet Verimliliği (EV / AC)</span></div>""", unsafe_allow_html=True)
            
            ky5.markdown(f"""
                <div class="soft-card" title="Data Date: {data_date_str}. Bu değer Lineer Yaklaşım ile hesaplanmış tahmini SPI değeridir.">
                    <h3>Schedule Perf. Index (SPI)</h3>
                    <h2 class="{spi_color}">{spi:.2f}</h2>
                    <span class="kpi-note">Zaman Performansı (Yaklaşık)</span>
                </div>
            """, unsafe_allow_html=True)

            if len(cat_columns) >= 2:
                my_col1, my_col2 = st.columns(2)
                with my_col1:
                    mat_y1 = st.selectbox("Satırlar (Yüzde)", cat_columns, index=0, key="y_yuzde")
                with my_col2:
                    def_x = next((i for i, col in enumerate(cat_columns) if col in ['DİSİPLİN', 'LVL3', 'KAT']), min(1, len(cat_columns)-1))
                    mat_x1 = st.selectbox("Sütunlar (Yüzde)", cat_columns, index=def_x, key="x_yuzde")

                if mat_y1 != mat_x1:
                    matrix_ev = df_filtered.groupby([mat_y1, mat_x1])['Total_EV'].sum().unstack(fill_value=0)
                    matrix_pv = df_filtered.groupby([mat_y1, mat_x1])['Total_Target'].sum().unstack(fill_value=0)
                    matrix_pct = (matrix_ev / matrix_pv * 100).fillna(0).replace([np.inf, -np.inf], 0).round(1)
                    
                    fig1 = px.imshow(matrix_pct, text_auto=True, aspect="auto", color_continuous_scale="PuBu", labels=dict(color="Performans (%)"))
                    fig1.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#525f7f'))
                    fig1.update_xaxes(side="top")
                    
                    sel1 = st.plotly_chart(fig1, use_container_width=True, on_select="rerun", selection_mode="points", key="fig_pct")
                    
                    st.markdown("<span class='kpi-note'>💡 Grafikte hücreye tıklayabilir veya tıklama çalışmıyorsa aşağıdaki kutulardan detay seçebilirsiniz:</span>", unsafe_allow_html=True)
                    
                    sel_y_val, sel_x_val = None, None
                    if sel1 and isinstance(sel1, dict) and "selection" in sel1 and sel1["selection"].get("points"):
                        pt = sel1["selection"]["points"][0]
                        sel_x_val, sel_y_val = pt.get("x"), pt.get("y")
                    
                    fc1, fc2, fc3 = st.columns([2,2,1])
                    m_y_val = fc1.selectbox("Detay Satır Seçimi", ['Seçiniz...'] + list(matrix_pct.index), index=0 if not sel_y_val else list(matrix_pct.index).index(sel_y_val)+1, key="man_y1")
                    m_x_val = fc2.selectbox("Detay Sütun Seçimi", ['Seçiniz...'] + list(matrix_pct.columns), index=0 if not sel_x_val else list(matrix_pct.columns).index(sel_x_val)+1, key="man_x1")
                    
                    if m_y_val != 'Seçiniz...' and m_x_val != 'Seçiniz...':
                        st.markdown(f"##### 🔍 Aktivite Detayları: {mat_y1} = **{m_y_val}** | {mat_x1} = **{m_x_val}**")
                        detail_df = df_filtered[(df_filtered[mat_y1] == m_y_val) & (df_filtered[mat_x1] == m_x_val)]
                        
                        csv = detail_df[show_cols].rename(columns=rename_dict).to_csv(index=False).encode('utf-8-sig')
                        fc3.download_button(label="📥 Veriyi Excel/CSV İndir", data=csv, file_name=f"Detay_{m_y_val}_{m_x_val}.csv", mime='text/csv')

                        st.dataframe(
                            detail_df[show_cols].rename(columns=rename_dict).style.format({'Fiziksel (%)': '{:.1f}%', 'Planlanan/Bütçe': '{:,.0f}', 'Kazanılmış Değer': '{:,.0f}', 'Gerçekleşen (AC)': '{:,.0f}'}),
                            use_container_width=True, hide_index=True
                        )

        # ==========================================
        # 2. SEKME: PARASAL (MONETARY / MİKTAR)
        # ==========================================
        with tab_para:
            st.write("")
            kp1, kp2, kp3, kp4 = st.columns(4)
            kp1.markdown(f"""<div class="soft-card"><h3>Toplam Planlanan (PV)</h3><h2 class="val-blue">{tot_target_all:,.0f}</h2><span class="kpi-note">Total Budgeted Units</span></div>""", unsafe_allow_html=True)
            kp2.markdown(f"""<div class="soft-card"><h3>Kazanılmış Değer (EV)</h3><h2 class="val-green">{tot_ev_all:,.0f}</h2><span class="kpi-note">Earned Value (Ağırlıklı İlerleme)</span></div>""", unsafe_allow_html=True)
            kp3.markdown(f"""<div class="soft-card"><h3>Gerçekleşen Miktar (AC)</h3><h2 class="val-red">{tot_ac_all:,.0f}</h2><span class="kpi-note">Actual Units</span></div>""", unsafe_allow_html=True)
            
            health_text = "🟢 İYİ (Bütçe İçi)" if cpi >= 1.0 else ("🟡 RİSKLİ" if cpi >= 0.85 else "🔴 KRİTİK (Bütçe Aşımı)")
            kp4.markdown(f"""<div class="soft-card"><h3>Proje Maliyet Sağlığı</h3><h2 class="{cpi_color}">{health_text}</h2><span class="kpi-note">CPI endeksine göre durum tespiti</span></div>""", unsafe_allow_html=True)

            if len(cat_columns) >= 2:
                mp_col1, mp_col2 = st.columns(2)
                with mp_col1:
                    mat_y2 = st.selectbox("Satırlar (Parasal)", cat_columns, index=0, key="y_para")
                with mp_col2:
                    def_x2 = next((i for i, col in enumerate(cat_columns) if col in ['DİSİPLİN', 'LVL3', 'KAT']), min(1, len(cat_columns)-1))
                    mat_x2 = st.selectbox("Sütunlar (Parasal)", cat_columns, index=def_x2, key="x_para")

                if mat_y2 != mat_x2:
                    matrix_ev = df_filtered.groupby([mat_y2, mat_x2])['Total_EV'].sum().unstack(fill_value=0).round(0)
                    
                    fig2 = px.imshow(matrix_ev, text_auto=True, aspect="auto", color_continuous_scale="Greens", labels=dict(color="Kazanılmış Değer (EV)"))
                    fig2.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#525f7f'))
                    fig2.update_xaxes(side="top")
                    
                    sel2 = st.plotly_chart(fig2, use_container_width=True, on_select="rerun", selection_mode="points", key="fig_ev")
                    
                    st.markdown("<span class='kpi-note'>💡 Grafikte hücreye tıklayabilir veya tıklama çalışmıyorsa aşağıdaki kutulardan detay seçebilirsiniz:</span>", unsafe_allow_html=True)
                    
                    sel_y_val2, sel_x_val2 = None, None
                    if sel2 and isinstance(sel2, dict) and "selection" in sel2 and sel2["selection"].get("points"):
                        pt = sel2["selection"]["points"][0]
                        sel_x_val2, sel_y_val2 = pt.get("x"), pt.get("y")
                    
                    fca, fcb, fcc = st.columns([2,2,1])
                    m_y_val2 = fca.selectbox("Detay Satır Seçimi", ['Seçiniz...'] + list(matrix_ev.index), index=0 if not sel_y_val2 else list(matrix_ev.index).index(sel_y_val2)+1, key="man_y2")
                    m_x_val2 = fcb.selectbox("Detay Sütun Seçimi", ['Seçiniz...'] + list(matrix_ev.columns), index=0 if not sel_x_val2 else list(matrix_ev.columns).index(sel_x_val2)+1, key="man_x2")

                    if m_y_val2 != 'Seçiniz...' and m_x_val2 != 'Seçiniz...':
                        st.markdown(f"##### 🔍 Aktivite Detayları: {mat_y2} = **{m_y_val2}** | {mat_x2} = **{m_x_val2}**")
                        detail_df2 = df_filtered[(df_filtered[mat_y2] == m_y_val2) & (df_filtered[mat_x2] == m_x_val2)]
                        
                        csv2 = detail_df2[show_cols].rename(columns=rename_dict).to_csv(index=False).encode('utf-8-sig')
                        fcc.download_button(label="📥 Veriyi Excel/CSV İndir", data=csv2, file_name=f"Detay_Parasal_{m_y_val2}_{m_x_val2}.csv", mime='text/csv')

                        st.dataframe(
                            detail_df2[show_cols].rename(columns=rename_dict).style.format({'Fiziksel (%)': '{:.1f}%', 'Planlanan/Bütçe': '{:,.0f}', 'Kazanılmış Değer': '{:,.0f}', 'Gerçekleşen (AC)': '{:,.0f}'}),
                            use_container_width=True, hide_index=True
                        )
