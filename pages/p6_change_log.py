import streamlit as st
import pandas as pd
import io
import plotly.express as px
import datetime

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(layout="wide", page_title="XER Teknik Ofis Paneli", page_icon="📊")
st.markdown("<h1 style='color: #2c3e50; font-weight: 800;'>📊 XER Teknik Ofis ve Kontrol Paneli</h1>", unsafe_allow_html=True)
st.markdown("Revizyon özetleri, saha ilerlemeleri ve Emlak Konut resmi Change Log raporu tek ekranda.")

# ==========================================
# 1. YARDIMCI FONKSİYONLAR (HIZLANDIRILMIŞ)
# ==========================================

@st.cache_data
def parse_xer(file_bytes):
    text = file_bytes.decode('windows-1254', errors='ignore')
    tables, current_table, columns = {}, None, []
    for line in text.splitlines():
        if line.startswith('%T'):
            current_table = line.split('\t')[1].strip()
            tables[current_table] = []
        elif line.startswith('%F') and current_table:
            columns = [col.strip().lower() for col in line.split('\t')[1:]]
        elif line.startswith('%R') and current_table:
            values = line.split('\t')[1:]
            tables[current_table].append({columns[i]: values[i].strip() if i < len(values) else "" for i in range(len(columns))})
    return {k: pd.DataFrame(v) for k, v in tables.items() if v}

def get_all_costs(trsrc_df, id_to_code_map):
    if trsrc_df.empty: return {}
    df = trsrc_df.copy()
    for col in ['target_cost', 'act_reg_cost']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        else:
            df[col] = 0
    grouped = df.groupby('task_id')[['target_cost', 'act_reg_cost']].sum()
    return {id_to_code_map.get(k, 'Unknown'): {'target': v['target_cost'], 'actual': v['act_reg_cost']} 
            for k, v in grouped.iterrows() if id_to_code_map.get(k)}

def map_relationships(pred_df, id_map):
    if pred_df.empty: return set()
    mapped = []
    for p_id, s_id, p_type, p_lag in zip(pred_df.get('pred_task_id', []), pred_df.get('task_id', []), pred_df.get('pred_type', []), pred_df.get('pred_lag_hr_cnt', [])):
        pred_code = id_map.get(p_id, 'Unknown')
        succ_code = id_map.get(s_id, 'Unknown')
        rel_type = str(p_type).replace('PR_', '') # FS, SS, FF, SF İlişki Tipleri
        lag = pd.to_numeric(str(p_lag).replace(',', '.'), errors='coerce') / 8
        lag_str = f" Lag: {lag}g" if pd.notna(lag) and lag != 0 else ""
        mapped.append(f"{pred_code} -> {succ_code} ({rel_type}{lag_str})")
    return set(mapped)

# ==========================================
# 2. ARAYÜZ VE DOSYA YÜKLEME
# ==========================================

st.markdown("### 📂 Dosya Yükleme Paneli")
c1, c2, c3 = st.columns(3)
with c1: baseline_file = st.file_uploader("1. Referans (Önceki) XER *", type=['xer'])
with c2: update_file = st.file_uploader("2. Güncel (Yeni) XER *", type=['xer'])
with c3:
    old_log_file = st.file_uploader("3. Eski Change Log (Geçmiş Notlar İçin)", type=['xlsx'])
    rapor_tarihi = st.date_input("Güncelleme Tarihi", datetime.date.today())

# ==========================================
# 3. ANA İŞLEM BLOĞU
# ==========================================

if baseline_file and update_file:
    with st.spinner("XER Verileri Analiz Ediliyor... (Hızlı Mod)"):
        db_base, db_upd = parse_xer(baseline_file.getvalue()), parse_xer(update_file.getvalue())
        task_b, task_u = db_base.get('TASK', pd.DataFrame()), db_upd.get('TASK', pd.DataFrame())
        trsrc_b, trsrc_u = db_base.get('TASKRSRC', pd.DataFrame()), db_upd.get('TASKRSRC', pd.DataFrame())
        wbs_b, wbs_u = db_base.get('PROJWBS', pd.DataFrame()), db_upd.get('PROJWBS', pd.DataFrame())
        
        if task_b.empty or task_u.empty:
            st.error("XER dosyalarından aktivite verisi okunamadı."); st.stop()

        id_code_b, id_code_u = dict(zip(task_b['task_id'], task_b['task_code'])), dict(zip(task_u['task_id'], task_u['task_code']))
        
        wbs_map = dict(zip(wbs_u['wbs_id'], wbs_u['wbs_name'])) if not wbs_u.empty else {}
        wbs_map.update(dict(zip(wbs_b['wbs_id'], wbs_b['wbs_name'])) if not wbs_b.empty else {})
        
        task_wbs_map_u = dict(zip(task_u['task_code'], task_u['wbs_id'])) if 'wbs_id' in task_u.columns else {}
        task_wbs_map_b = dict(zip(task_b['task_code'], task_b['wbs_id'])) if 'wbs_id' in task_b.columns else {}
        
        def get_wbs(t_id, source='u'): 
            w_id = task_wbs_map_u.get(t_id) if source == 'u' else task_wbs_map_b.get(t_id)
            return wbs_map.get(w_id, 'Bilinmiyor')
        
        float_u = dict(zip(task_u['task_code'], pd.to_numeric(task_u.get('total_float_hr_cnt', 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0) / 8))
        common_tasks = set(task_b['task_code']).intersection(set(task_u['task_code']))
        added_tasks = set(task_u['task_code']) - set(task_b['task_code'])
        deleted_tasks = set(task_b['task_code']) - set(task_u['task_code'])

        old_notes, old_cost_actuals = {}, {}
        if old_log_file:
            try:
                xls = pd.ExcelFile(old_log_file)
                if 'EKGYO_Change_Log' in xls.sheet_names:
                    old_df_rev = pd.read_excel(xls, sheet_name='EKGYO_Change_Log')
                    gerekce_sec = [c for c in old_df_rev.columns if 'Açıklaması' in c]
                    if gerekce_sec:
                        old_notes_df = old_df_rev.dropna(subset=[gerekce_sec[0]])
                        old_notes = dict(zip(old_notes_df['Activity ID'], old_notes_df[gerekce_sec[0]]))
                if 'Maliyet_ve_Hakedis' in xls.sheet_names:
                    old_df_cost = pd.read_excel(xls, sheet_name='Maliyet_ve_Hakedis')
                    if 'Aktivite ID' in old_df_cost.columns and 'Güncel Toplam Hakediş' in old_df_cost.columns:
                        old_cost_actuals = dict(zip(old_df_cost['Aktivite ID'], pd.to_numeric(old_df_cost['Güncel Toplam Hakediş'], errors='coerce').fillna(0)))
            except: pass

        cols = ['task_code', 'task_name', 'target_drtn_hr_cnt', 'remain_drtn_hr_cnt', 'phys_complete_pct', 'status_code', 'target_work_qty', 'target_equip_qty', 'act_equip_qty']
        df_join = task_b[task_b['task_code'].isin(common_tasks)][[c for c in cols if c in task_b.columns]].set_index('task_code').join(
                  task_u[task_u['task_code'].isin(common_tasks)][[c for c in cols if c in task_u.columns]].set_index('task_code'), lsuffix='_b', rsuffix='_u')

        # --- A. SAHA İLERLEMESİ ---
        rb, ru, pb, pu, sb, su = 'remain_drtn_hr_cnt_b', 'remain_drtn_hr_cnt_u', 'phys_complete_pct_b', 'phys_complete_pct_u', 'status_code_b', 'status_code_u'
        df_join[rb], df_join[ru] = pd.to_numeric(df_join.get(rb, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)/8, pd.to_numeric(df_join.get(ru, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)/8
        df_join[pb], df_join[pu] = pd.to_numeric(df_join.get(pb, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0), pd.to_numeric(df_join.get(pu, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        
        status_map = {'TK_NotStart': 'Başlamadı', 'TK_Active': 'Devam Ediyor', 'TK_Complete': 'Tamamlandı'}
        df_join[sb] = df_join.get(sb, '').replace(status_map)
        df_join[su] = df_join.get(su, '').replace(status_map)

        df_prog = df_join[(df_join[rb] != df_join[ru]) | (df_join[pb] != df_join[pu]) | (df_join.get(sb, '') != df_join.get(su, ''))].copy().reset_index()
        
        if not df_prog.empty:
            df_prog = df_prog.rename(columns={'task_code': 'Aktivite ID', 'task_name_b': 'Aktivite Adı', su: 'Durum', pb: 'Eski %', pu: 'Yeni %', rb: 'Eski Kalan', ru: 'Yeni Kalan'})
            df_prog['WBS'] = df_prog['Aktivite ID'].apply(lambda x: get_wbs(x, 'u'))
            df_prog['Bolluk'] = df_prog['Aktivite ID'].map(float_u).fillna(0)
            df_prog = df_prog[['WBS', 'Aktivite ID', 'Aktivite Adı', 'Durum', 'Eski %', 'Yeni %', 'Eski Kalan', 'Yeni Kalan', 'Bolluk']]
        else: 
            df_prog = pd.DataFrame(columns=['WBS', 'Aktivite ID', 'Aktivite Adı', 'Durum', 'Eski %', 'Yeni %', 'Eski Kalan', 'Yeni Kalan', 'Bolluk'])

        # İşçilik (Labor/AS) ve Non-Labor (Equip/Maliyet) Dönüşümleri (Virgülleri noktaya çevir)
        df_join['twq_b'] = pd.to_numeric(df_join.get('target_work_qty_b', 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_join['twq_u'] = pd.to_numeric(df_join.get('target_work_qty_u', 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_join['teq_b'] = pd.to_numeric(df_join.get('target_equip_qty_b', 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_join['teq_u'] = pd.to_numeric(df_join.get('target_equip_qty_u', 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_join['aeq_b'] = pd.to_numeric(df_join.get('act_equip_qty_b', 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_join['aeq_u'] = pd.to_numeric(df_join.get('act_equip_qty_u', 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

        # --- B. MALİYET VE HAKEDİŞ (DOĞRUDAN NON-LABOR/EQUIPMENT SÜTUNUNDAN) ---
        cost_data = []
        for tcode, row in df_join.iterrows():
            target_cost = row['teq_u']
            actual_cost = row['aeq_u']
            prev_actual = old_cost_actuals.get(tcode, row['aeq_b'])

            # Sadece geçmişte veya şimdi hakedişe girenler listelenir
            if prev_actual > 0 or actual_cost > 0:
                cost_data.append({
                    'Aktivite ID': tcode,
                    'Aktivite Adı': row.get('task_name_b', ''),
                    'Toplam Bütçe (Hedef)': target_cost,
                    'Önceki Ay Fiili Hakediş': prev_actual,
                    'Güncel Toplam Hakediş': actual_cost,
                    'Bu Ayki Artış (Fark)': actual_cost - prev_actual
                })
        df_cost = pd.DataFrame(cost_data)
        if not df_cost.empty:
            df_cost['WBS'] = df_cost['Aktivite ID'].apply(lambda x: get_wbs(x, 'u'))
            df_cost = df_cost[['WBS', 'Aktivite ID', 'Aktivite Adı', 'Toplam Bütçe (Hedef)', 'Önceki Ay Fiili Hakediş', 'Güncel Toplam Hakediş', 'Bu Ayki Artış (Fark)']]
        else:
            df_cost = pd.DataFrame(columns=['WBS', 'Aktivite ID', 'Aktivite Adı', 'Toplam Bütçe (Hedef)', 'Önceki Ay Fiili Hakediş', 'Güncel Toplam Hakediş', 'Bu Ayki Artış (Fark)'])

        # --- C. EMLAK KONUT FORMATINDA TEK TABLO (CHANGE LOG) ---
        tb, tu = ('target_drtn_hr_cnt_b' if 'target_drtn_hr_cnt_b' in df_join else 'target_drtn_b'), ('target_drtn_hr_cnt_u' if 'target_drtn_hr_cnt_u' in df_join else 'target_drtn_u')
        df_join[tb], df_join[tu] = pd.to_numeric(df_join.get(tb, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)/8, pd.to_numeric(df_join.get(tu, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)/8
        
        added_rels = map_relationships(db_upd.get('TASKPRED', pd.DataFrame()), id_code_u) - map_relationships(db_base.get('TASKPRED', pd.DataFrame()), id_code_b)
        deleted_rels = map_relationships(db_base.get('TASKPRED', pd.DataFrame()), id_code_b) - map_relationships(db_upd.get('TASKPRED', pd.DataFrame()), id_code_u)

        ekgyo_data = []
        row_no = 1
        tarih_str = rapor_tarihi.strftime("%d.%m.%Y")
        
        task_name_u_dict = dict(zip(task_u['task_code'], task_u['task_name']))
        task_drtn_u_dict = dict(zip(task_u['task_code'], pd.to_numeric(task_u.get('target_drtn_hr_cnt', 0), errors='coerce').fillna(0)/8))
        task_name_b_dict = dict(zip(task_b['task_code'], task_b['task_name']))
        task_drtn_b_dict = dict(zip(task_b['task_code'], pd.to_numeric(task_b.get('target_drtn_hr_cnt', 0), errors='coerce').fillna(0)/8))

        # 1. Süre Değişiklikleri
        df_diff_drtn = df_join[df_join[tb] != df_join[tu]].copy().reset_index()
        for _, row in df_diff_drtn.iterrows():
            act_id = row['task_code']
            ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(act_id, 'u'), 'Activity ID': act_id, 'Activity Name': row.get('task_name_u', ''), 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Süre', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(act_id, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': '', 'İlk Durum (Süre)': row[tb], 'Son Durum (Süre)': row[tu]})
            row_no += 1

        # 2. Adam-Saat (Labor Units) Değişiklikleri (> 0.5 Hassasiyet Toleransı ile)
        df_diff_as = df_join[abs(df_join['twq_b'] - df_join['twq_u']) > 0.5].copy().reset_index()
        for _, row in df_diff_as.iterrows():
            act_id = row['task_code']
            aciklama = f"Eski Adam-Saat: {row['twq_b']:g} -> Yeni Adam-Saat: {row['twq_u']:g}. {old_notes.get(act_id, '')}"
            ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(act_id, 'u'), 'Activity ID': act_id, 'Activity Name': row.get('task_name_u', ''), 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Adam-Saat (Miktar)', 'Değişiklik Nedeni ve Açıklaması': aciklama, 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': '', 'İlk Durum (Süre)': '', 'Son Durum (Süre)': ''})
            row_no += 1

        # 3. Parasal Kaynak (Maliyet/Bütçe) Değişiklikleri (> 1.0 TL Hassasiyet Toleransı ile)
        df_diff_cost = df_join[abs(df_join['teq_b'] - df_join['teq_u']) > 1.0].copy().reset_index()
        for _, row in df_diff_cost.iterrows():
            act_id = row['task_code']
            aciklama = f"Eski Bütçe: ₺{row['teq_b']:,.2f} -> Yeni Bütçe: ₺{row['teq_u']:,.2f}. {old_notes.get(act_id, '')}"
            ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(act_id, 'u'), 'Activity ID': act_id, 'Activity Name': row.get('task_name_u', ''), 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Parasal Kaynak (Bütçe)', 'Değişiklik Nedeni ve Açıklaması': aciklama, 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': '', 'İlk Durum (Süre)': '', 'Son Durum (Süre)': ''})
            row_no += 1

        # 4. İlişki Eklendi / Lag (FS, SS vb.) Değişti
        for rel in added_rels:
            try:
                pred, rest = rel.split(' -> ')
                succ, rel_type = rest.split(' (')[0], rest.split(' (')[1].replace(')', '')
                ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(pred, 'u'), 'Activity ID': pred, 'Activity Name': task_name_u_dict.get(pred, ''), 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Bağlantı (Yeni/Revize)', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(pred, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': f"{succ} ({rel_type})", 'İlk Durum (Süre)': '', 'Son Durum (Süre)': ''})
                row_no += 1
            except: pass

        # 5. İlişki Silindi
        for rel in deleted_rels:
            try:
                pred, rest = rel.split(' -> ')
                succ, rel_type = rest.split(' (')[0], rest.split(' (')[1].replace(')', '')
                ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(pred, 'b'), 'Activity ID': pred, 'Activity Name': task_name_b_dict.get(pred, ''), 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Bağlantı (Silindi/Değişti)', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(pred, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': f"{succ} ({rel_type})", 'Ardıl Son Durum': '', 'İlk Durum (Süre)': '', 'Son Durum (Süre)': ''})
                row_no += 1
            except: pass

        # 6. Aktivite Eklendi
        for act in added_tasks:
            try:
                ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(act, 'u'), 'Activity ID': act, 'Activity Name': task_name_u_dict.get(act, ''), 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Yeni Aktivite', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(act, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': '', 'İlk Durum (Süre)': '', 'Son Durum (Süre)': task_drtn_u_dict.get(act, '')})
                row_no += 1
            except: pass

        # 7. Aktivite Silindi
        for act in deleted_tasks:
            try:
                ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(act, 'b'), 'Activity ID': act, 'Activity Name': task_name_b_dict.get(act, ''), 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Silinen Aktivite', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(act, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': '', 'İlk Durum (Süre)': task_drtn_b_dict.get(act, ''), 'Son Durum (Süre)': ''})
                row_no += 1
            except: pass

        df_ekgyo = pd.DataFrame(ekgyo_data)
        if df_ekgyo.empty:
            df_ekgyo = pd.DataFrame(columns=['No.', 'Proje Adı', 'WBS', 'Activity ID', 'Activity Name', 'Değişiklik Tarihi', 'Değişiklik Yapılan Yer', 'Değişiklik Nedeni ve Açıklaması', 'Öncül İlk Durum', 'Öncül Son Durum', 'Ardıl İlk Durum', 'Ardıl Son Durum', 'İlk Durum (Süre)', 'Son Durum (Süre)'])

        # ==========================================
        # 4. GÖRSELLEŞTİRME VE FİLTRELER
        # ==========================================
        st.markdown("---")
        all_wbs = sorted(list(set(df_prog['WBS'].tolist() + df_ekgyo['WBS'].tolist() + df_cost['WBS'].tolist())))
        selected_wbs = st.multiselect("🏗️ WBS Filtresi (Tüm sekmeleri günceller)", options=all_wbs, default=all_wbs) if all_wbs else []
        
        if selected_wbs:
            df_ekgyo = df_ekgyo[df_ekgyo['WBS'].isin(selected_wbs)] if not df_ekgyo.empty else df_ekgyo
            df_prog = df_prog[df_prog['WBS'].isin(selected_wbs)] if not df_prog.empty else df_prog
            df_cost = df_cost[df_cost['WBS'].isin(selected_wbs)] if not df_cost.empty else df_cost

        df_ekgyo_export = df_ekgyo.drop(columns=['WBS']) if 'WBS' in df_ekgyo.columns else df_ekgyo

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("⏱️ Toplam Revizyon Hareketi", len(df_ekgyo))
        k2.metric("📈 İlerleme Girilen İş", len(df_prog))
        k3.metric("💰 Hakediş Kalemi", len(df_cost[df_cost['Güncel Toplam Hakediş'] > 0]) if not df_cost.empty else 0)
        k4.metric("🚨 Geciken İş (Bolluk ≤ 0)", sum(1 for act_id, val in float_u.items() if val <= 0 and get_wbs(act_id, 'u') in (selected_wbs or all_wbs)), delta_color="inverse")

        # ==========================================
        # 5. TABLOLAR VE DASHBOARD
        # ==========================================
        st.markdown("### 📋 Analiz Detayları")
        t1, t2, t3, t4 = st.tabs(["📊 Revizyon Özeti", "📈 Saha İlerleme", "💰 Maliyet ve Hakediş", "📑 EKGYO Resmi Change Log"])
        
        with t1:
            r1, r2 = st.columns(2)
            with r1:
                if not df_ekgyo.empty:
                    rev_counts = df_ekgyo['Değişiklik Yapılan Yer'].value_counts().reset_index()
                    rev_counts.columns = ['Değişiklik Türü', 'Adet']
                    fig_rev = px.pie(rev_counts, values='Adet', names='Değişiklik Türü', title='Değişiklik Dağılımı', hole=0.4)
                    st.plotly_chart(fig_rev, use_container_width=True)
            with r2:
                if not df_diff_drtn.empty:
                    df_diff_drtn['Süre Farkı'] = pd.to_numeric(df_diff_drtn[tu] - df_diff_drtn[tb], errors='coerce')
                    top_diff = df_diff_drtn.sort_values(by='Süre Farkı', key=abs, ascending=False).head(5)
                    fig_diff = px.bar(top_diff, x='task_code', y='Süre Farkı', title='Süresi En Çok Sapan 5 Aktivite (Gün)')
                    st.plotly_chart(fig_diff, use_container_width=True)
            d1, d2, d3 = st.columns(3)
            d1.metric("Toplam Planlanan Bütçe", f"₺ {df_cost['Toplam Bütçe (Hedef)'].sum():,.2f}" if not df_cost.empty else "₺ 0.00")
            d2.metric("Bugüne Kadarki Toplam Hakediş", f"₺ {df_cost['Güncel Toplam Hakediş'].sum():,.2f}" if not df_cost.empty else "₺ 0.00")
            d3.metric("Bu Dönem Yapılan Net Hakediş", f"₺ {df_cost['Bu Ayki Artış (Fark)'].sum():,.2f}" if not df_cost.empty else "₺ 0.00")

        with t2: st.dataframe(df_prog, use_container_width=True, hide_index=True)
        with t3: st.dataframe(df_cost, use_container_width=True, hide_index=True)
        with t4: st.dataframe(df_ekgyo_export, use_container_width=True, hide_index=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_prog.to_excel(writer, sheet_name='Saha_Ilerlemesi', index=False)
            df_cost.to_excel(writer, sheet_name='Maliyet_ve_Hakedis', index=False)
            df_ekgyo_export.to_excel(writer, sheet_name='EKGYO_Change_Log', index=False)
            
        st.download_button("💾 Excel Raporunu İndir", data=output.getvalue(), file_name=f"EKGYO_P6_Raporu_{rapor_tarihi.strftime('%d.%m.%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
