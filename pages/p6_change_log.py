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
# 1. YARDIMCI FONKSİYONLAR
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
    for _, row in pred_df.iterrows():
        pred_code = id_map.get(row.get('pred_task_id'), 'Unknown')
        succ_code = id_map.get(row.get('task_id'), 'Unknown')
        rel_type = row.get('pred_type', '').replace('PR_', '')
        mapped.append(f"{pred_code} -> {succ_code} ({rel_type})")
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
    with st.spinner("XER Verileri Analiz Ediliyor..."):
        db_base, db_upd = parse_xer(baseline_file.getvalue()), parse_xer(update_file.getvalue())
        task_b, task_u = db_base.get('TASK', pd.DataFrame()), db_upd.get('TASK', pd.DataFrame())
        trsrc_b, trsrc_u = db_base.get('TASKRSRC', pd.DataFrame()), db_upd.get('TASKRSRC', pd.DataFrame())
        wbs_b, wbs_u = db_base.get('PROJWBS', pd.DataFrame()), db_upd.get('PROJWBS', pd.DataFrame())
        
        if task_b.empty or task_u.empty:
            st.error("XER dosyalarından aktivite verisi okunamadı."); st.stop()

        id_code_b, id_code_u = dict(zip(task_b['task_id'], task_b['task_code'])), dict(zip(task_u['task_id'], task_u['task_code']))
        
        wbs_map = dict(zip(wbs_u['wbs_id'], wbs_u['wbs_name'])) if not wbs_u.empty else {}
        wbs_map.update(dict(zip(wbs_b['wbs_id'], wbs_b['wbs_name'])) if not wbs_b.empty else {})
        def get_wbs(t_id, df): return wbs_map.get(df.loc[df['task_code'] == t_id, 'wbs_id'].values[0] if not df.loc[df['task_code'] == t_id].empty else None, 'Bilinmiyor')
        
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

        cols = ['task_code', 'task_name', 'target_drtn_hr_cnt', 'target_drtn', 'remain_drtn_hr_cnt', 'phys_complete_pct', 'act_start_date', 'status_code']
        df_join = task_b[task_b['task_code'].isin(common_tasks)][[c for c in cols if c in task_b.columns]].set_index('task_code').join(
                  task_u[task_u['task_code'].isin(common_tasks)][[c for c in cols if c in task_u.columns]].set_index('task_code'), lsuffix='_b', rsuffix='_u')

        # --- A. SAHA İLERLEMESİ ---
        rb, ru, pb, pu, sb, su = 'remain_drtn_hr_cnt_b', 'remain_drtn_hr_cnt_u', 'phys_complete_pct_b', 'phys_complete_pct_u', 'status_code_b', 'status_code_u'
        df_join[rb], df_join[ru] = pd.to_numeric(df_join.get(rb, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)/8, pd.to_numeric(df_join.get(ru, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)/8
        df_join[pb], df_join[pu] = pd.to_numeric(df_join.get(pb, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0), pd.to_numeric(df_join.get(pu, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_join[su] = df_join.get(su, '').replace({'TK_NotStart': 'Başlamadı', 'TK_Active': 'Devam Ediyor', 'TK_Complete': 'Tamamlandı'})

        df_prog = df_join[(df_join[rb] != df_join[ru]) | (df_join[pb] != df_join[pu]) | (df_join.get(sb, '') != df_join.get(su, ''))].copy().reset_index()
        if not df_prog.empty:
            df_prog = df_prog.rename(columns={'task_code': 'Aktivite ID', 'task_name_b': 'Aktivite Adı', su: 'Durum', pb: 'Eski %', pu: 'Yeni %', rb: 'Eski Kalan', ru: 'Yeni Kalan'})
            df_prog['WBS'] = df_prog['Aktivite ID'].apply(lambda x: get_wbs(x, task_u))
            df_prog['Bolluk'] = df_prog['Aktivite ID'].map(float_u).fillna(0)
            df_prog = df_prog[['WBS', 'Aktivite ID', 'Aktivite Adı', 'Durum', 'Eski %', 'Yeni %', 'Eski Kalan', 'Yeni Kalan', 'Bolluk']]
        else: 
            df_prog = pd.DataFrame(columns=['WBS', 'Aktivite ID', 'Aktivite Adı', 'Durum', 'Eski %', 'Yeni %', 'Eski Kalan', 'Yeni Kalan', 'Bolluk'])

        # --- B. MALİYET VE HAKEDİŞ ---
        cost_b = get_all_costs(trsrc_b, id_code_b)
        cost_u = get_all_costs(trsrc_u, id_code_u)
        cost_data = []
        for tcode in common_tasks:
            cb = cost_b.get(tcode, {'target': 0, 'actual': 0})
            cu = cost_u.get(tcode, {'target': 0, 'actual': 0})
            prev_actual = old_cost_actuals.get(tcode, cb['actual'])
            if cu['target'] > 0 or prev_actual > 0 or cu['actual'] > 0:
                cost_data.append({
                    'Aktivite ID': tcode, 'Aktivite Adı': df_join.loc[tcode, 'task_name_b'] if tcode in df_join.index else '',
                    'Toplam Bütçe (Hedef)': cu['target'], 'Önceki Ay Fiili Hakediş': prev_actual,
                    'Güncel Toplam Hakediş': cu['actual'], 'Bu Ayki Artış (Fark)': cu['actual'] - prev_actual
                })
        df_cost = pd.DataFrame(cost_data)
        if not df_cost.empty:
            df_cost['WBS'] = df_cost['Aktivite ID'].apply(lambda x: get_wbs(x, task_u))
            df_cost = df_cost[['WBS', 'Aktivite ID', 'Aktivite Adı', 'Toplam Bütçe (Hedef)', 'Önceki Ay Fiili Hakediş', 'Güncel Toplam Hakediş', 'Bu Ayki Artış (Fark)']]
        else:
            df_cost = pd.DataFrame(columns=['WBS', 'Aktivite ID', 'Aktivite Adı', 'Toplam Bütçe (Hedef)', 'Önceki Ay Fiili Hakediş', 'Güncel Toplam Hakediş', 'Bu Ayki Artış (Fark)'])

        # --- C. EMLAK KONUT FORMATINDA TEK TABLO (EKGYO_Change_Log) ---
        tb, tu = ('target_drtn_hr_cnt_b' if 'target_drtn_hr_cnt_b' in df_join else 'target_drtn_b'), ('target_drtn_hr_cnt_u' if 'target_drtn_hr_cnt_u' in df_join else 'target_drtn_u')
        df_join[tb], df_join[tu] = pd.to_numeric(df_join.get(tb, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)/8, pd.to_numeric(df_join.get(tu, 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0)/8
        
        df_diff = df_join[df_join[tb] != df_join[tu]].copy().reset_index()
        if not df_diff.empty:
            df_diff = df_diff.rename(columns={'task_code': 'Aktivite ID', tb: 'Önceki Hedef', tu: 'Güncel Hedef'})

        added_rels = map_relationships(db_upd.get('TASKPRED', pd.DataFrame()), id_code_u) - map_relationships(db_base.get('TASKPRED', pd.DataFrame()), id_code_b)
        deleted_rels = map_relationships(db_base.get('TASKPRED', pd.DataFrame()), id_code_b) - map_relationships(db_upd.get('TASKPRED', pd.DataFrame()), id_code_u)

        ekgyo_data = []
        row_no = 1
        tarih_str = rapor_tarihi.strftime("%d.%m.%Y")
        
        for _, row in df_diff.iterrows():
            act_id = row['Aktivite ID']
            ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(act_id, task_u), 'Activity ID': act_id, 'Activity Name': row.get('task_name_u', ''), 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Süre', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(act_id, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': '', 'İlk Durum (Süre)': row['Önceki Hedef'], 'Son Durum (Süre)': row['Güncel Hedef']})
            row_no += 1

        for rel in added_rels:
            try:
                pred, rest = rel.split(' -> ')
                succ, rel_type = rest.split(' (')[0], rest.split(' (')[1].replace(')', '')
                ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(pred, task_u), 'Activity ID': pred, 'Activity Name': df_join.loc[pred, 'task_name_u'] if pred in df_join.index else '', 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Bağlantı (Yeni)', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(pred, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': f"{succ} {rel_type}", 'İlk Durum (Süre)': '', 'Son Durum (Süre)': ''})
                row_no += 1
            except: pass

        for rel in deleted_rels:
            try:
                pred, rest = rel.split(' -> ')
                succ, rel_type = rest.split(' (')[0], rest.split(' (')[1].replace(')', '')
                ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(pred, task_b), 'Activity ID': pred, 'Activity Name': df_join.loc[pred, 'task_name_b'] if pred in df_join.index else '', 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Bağlantı (Silindi)', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(pred, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': f"{succ} {rel_type}", 'Ardıl Son Durum': '', 'İlk Durum (Süre)': '', 'Son Durum (Süre)': ''})
                row_no += 1
            except: pass

        for act in added_tasks:
            try:
                name = task_u.loc[task_u['task_code'] == act, 'task_name'].values[0]
                drtn = pd.to_numeric(task_u.loc[task_u['task_code'] == act, 'target_drtn_hr_cnt'].values[0])/8
                ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(act, task_u), 'Activity ID': act, 'Activity Name': name, 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Yeni Aktivite', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(act, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': '', 'İlk Durum (Süre)': '', 'Son Durum (Süre)': drtn})
                row_no += 1
            except: pass

        for act in deleted_tasks:
            try:
                name = task_b.loc[task_b['task_code'] == act, 'task_name'].values[0]
                drtn = pd.to_numeric(task_b.loc[task_b['task_code'] == act, 'target_drtn_hr_cnt'].values[0])/8
                ekgyo_data.append({'No.': row_no, 'Proje Adı': 'Proje', 'WBS': get_wbs(act, task_b), 'Activity ID': act, 'Activity Name': name, 'Değişiklik Tarihi': tarih_str, 'Değişiklik Yapılan Yer': 'Silinen Aktivite', 'Değişiklik Nedeni ve Açıklaması': old_notes.get(act, ''), 'Öncül İlk Durum': '', 'Öncül Son Durum': '', 'Ardıl İlk Durum': '', 'Ardıl Son Durum': '', 'İlk Durum (Süre)': drtn, 'Son Durum (Süre)': ''})
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
        selected_wbs = st.multiselect("🏗️ WBS Filtresi (Uygulanan filtre Yönetici Özeti dahil tüm sekmeleri günceller)", options=all_wbs, default=all_wbs) if all_wbs else []
        
        if selected_wbs:
            df_ekgyo = df_ekgyo[df_ekgyo['WBS'].isin(selected_wbs)] if not df_ekgyo.empty else df_ekgyo
            df_prog = df_prog[df_prog['WBS'].isin(selected_wbs)] if not df_prog.empty else df_prog
            df_cost = df_cost[df_cost['WBS'].isin(selected_wbs)] if not df_cost.empty else df_cost

        df_ekgyo_export = df_ekgyo.drop(columns=['WBS']) if 'WBS' in df_ekgyo.columns else df_ekgyo

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("⏱️ Toplam Revizyon Hareketi", len(df_ekgyo))
        k2.metric("📈 İlerleme Girilen İş", len(df_prog))
        k3.metric("💰 Hakediş Kalemi", len(df_cost[df_cost['Güncel Toplam Hakediş'] > 0]) if not df_cost.empty else 0)
        k4.metric("🚨 Geciken İş (Bolluk ≤ 0)", sum(1 for act_id, val in float_u.items() if val <= 0 and get_wbs(act_id, task_u) in (selected_wbs or all_wbs)), delta_color="inverse")

        # ==========================================
        # 5. TABLOLAR VE DASHBOARD
        # ==========================================
        st.markdown("### 📋 Analiz Detayları")
        t1, t2, t3, t4 = st.tabs(["📊 Revizyon Özeti (Dashboard)", "📈 Saha İlerleme (Süre/Yüzde)", "💰 Maliyet ve Hakediş", "📑 EKGYO Resmi Change Log"])
        
        with t1:
            st.markdown("#### 🔄 Değişiklik Dağılımları")
            r1, r2 = st.columns(2)
            with r1:
                if not df_ekgyo.empty:
                    rev_counts = df_ekgyo['Değişiklik Yapılan Yer'].value_counts().reset_index()
                    rev_counts.columns = ['Değişiklik Türü', 'Adet']
                    fig_rev = px.pie(rev_counts, values='Adet', names='Değişiklik Türü', title='Hangi Tür Müdahaleler Yapıldı?', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    st.plotly_chart(fig_rev, use_container_width=True)
                else:
                    st.info("Bu dönemde herhangi bir yapısal revizyon tespit edilmemiştir.")
            
            with r2:
                if not df_diff.empty:
                    df_diff['Süre Farkı'] = pd.to_numeric(df_diff['Güncel Hedef'] - df_diff['Önceki Hedef'], errors='coerce')
                    top_diff = df_diff.sort_values(by='Süre Farkı', key=abs, ascending=False).head(5)
                    fig_diff = px.bar(top_diff, x='Aktivite ID', y='Süre Farkı', title='Süresi En Çok Sapan 5 Aktivite (Gün)', color='Süre Farkı', color_continuous_scale='RdBu')
                    st.plotly_chart(fig_diff, use_container_width=True)
                else:
                    st.info("Bu dönemde hedef sürelerde (Target Duration) bir sapma olmamıştır.")
            
            st.markdown("---")
            st.markdown("#### 🏆 Finansal Durum (Filtrelenmiş)")
            d1, d2, d3 = st.columns(3)
            d1.metric("Toplam Planlanan Bütçe", f"₺ {df_cost['Toplam Bütçe (Hedef)'].sum():,.2f}" if not df_cost.empty else "₺ 0.00")
            d2.metric("Bugüne Kadarki Toplam Hakediş", f"₺ {df_cost['Güncel Toplam Hakediş'].sum():,.2f}" if not df_cost.empty else "₺ 0.00")
            d3.metric("Bu Dönem Yapılan Net Hakediş", f"₺ {df_cost['Bu Ayki Artış (Fark)'].sum():,.2f}" if not df_cost.empty else "₺ 0.00")

        with t2: 
            st.info("💡 Fiziksel ilerlemeler ve kalan süre değişimleri burada takip edilir.")
            st.dataframe(df_prog, use_container_width=True, hide_index=True)
        with t3: 
            st.info("💡 Bütçe ve fiili hakediş (gerçekleşen maliyet) takibi.")
            st.dataframe(df_cost, use_container_width=True, hide_index=True)
        with t4: 
            st.success("✅ Emlak Konut şablonuna göre düzenlenmiş yapısal değişiklik raporu. Çıktı alıp doğrudan kullanabilirsiniz.")
            st.dataframe(df_ekgyo_export, use_container_width=True, hide_index=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_prog.to_excel(writer, sheet_name='Saha_Ilerlemesi', index=False)
            df_cost.to_excel(writer, sheet_name='Maliyet_ve_Hakedis', index=False)
            df_ekgyo_export.to_excel(writer, sheet_name='EKGYO_Change_Log', index=False)
            
        st.download_button("💾 Excel Raporunu İndir", data=output.getvalue(), file_name=f"EKGYO_P6_Raporu_{rapor_tarihi.strftime('%d.%m.%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)