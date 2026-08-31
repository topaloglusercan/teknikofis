import streamlit as st
import pandas as pd
import networkx as nx

st.markdown("<h1 style='color: #2c3e50;'>🔗 P6 Gereksiz Bağ Analizi (Redundant Logic)</h1>", unsafe_allow_html=True)
st.markdown("XER dosyasındaki dolaylı (geçişli) ilişkileri tarayarak, projenin mantığını bozmadan silinebilecek gereksiz fazlalık bağları tespit edin.")

@st.cache_data
def parse_xer_relationships(file_bytes):
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

    return clean_df(tables.get('TASK', [])), clean_df(tables.get('TASKPRED', []))

uploaded_xer = st.file_uploader("📂 XER Dosyasını Yükle (Gereksiz Bağ Taraması İçin)", type=['xer'])

if uploaded_xer:
    with st.spinner("Şebeke Mantığı (Network Logic) Taranıyor..."):
        df_task, df_rels = parse_xer_relationships(uploaded_xer.getvalue())
        
        if not df_task.empty and not df_rels.empty:
            # 1. Aktivite İsimlerini Sözlüğe Al
            df_t = df_task[['task_id', 'task_code', 'task_name']].drop_duplicates(subset=['task_id']).copy()
            act_dict = df_t.set_index('task_id')['task_name'].to_dict()
            code_dict = df_t.set_index('task_id')['task_code'].to_dict()

            # 2. NetworkX Yönlü Graf (DiGraph) Oluştur
            G = nx.DiGraph()
            
            # TASKPRED tablosundaki ilişkileri grafa kenar (edge) olarak ekle
            for _, row in df_rels.iterrows():
                pred = row.get('pred_task_id')
                succ = row.get('task_id')
                rel_type = row.get('pred_type', 'FS')
                if pred and succ:
                    G.add_edge(pred, succ, type=rel_type)

            gereksiz_baglar = []

            # 3. Dolaylı Yol (Transitive Reduction) Kontrolü
            # Her bir doğrudan bağı (A -> B) koparıp, "Acaba A'dan B'ye gidecek başka bir yol (A -> C -> B) var mı?" diye bakıyoruz.
            edges = list(G.edges(data=True))
            for u, v, data in edges:
                G.remove_edge(u, v)
                
                # Eğer aradaki bağı kopardığımız halde hala u'dan v'ye bir yol varsa, bu kopardığımız bağ gereksizdir.
                if nx.has_path(G, u, v):
                    gereksiz_baglar.append({
                        'Predecessor ID': code_dict.get(u, u),
                        'Öncül Aktivite': act_dict.get(u, 'Bilinmeyen'),
                        'Successor ID': code_dict.get(v, v),
                        'Ardıl Aktivite': act_dict.get(v, 'Bilinmeyen'),
                        'İlişki Tipi': data.get('type', 'FS'),
                        'Delete This Row': 'd'
                    })
                
                # İnceleme bitti, bağı yerine geri koy (diğer bağları test etmek için)
                G.add_edge(u, v, **data)

            # 4. Sonuçları Raporlama
            if gereksiz_baglar:
                df_gereksiz = pd.DataFrame(gereksiz_baglar)
                st.warning(f"⚠️ Şebekede {len(gereksiz_baglar)} adet dolaylı/gereksiz bağ (Redundant Link) tespit edildi.")
                
                st.markdown("### 🗑️ Silinmesi Önerilen Bağlar")
                st.dataframe(df_gereksiz[['Predecessor ID', 'Öncül Aktivite', 'Successor ID', 'Ardıl Aktivite', 'İlişki Tipi']], use_container_width=True)
                
                st.info("💡 **Nasıl Silinir?** İndirdiğiniz CSV dosyasını P6'ya import (Update Existing) ettiğinizde 'Delete This Row = d' komutu sayesinde bu gereksiz bağlar P6'dan otomatik silinecektir.")
                
                # P6 Import Formatında CSV
                csv = df_gereksiz.to_csv(index=False).encode('utf-8-sig')
                st.download_button(label="💾 P6 Silme Şablonunu İndir (CSV)", data=csv, file_name='Gereksiz_Baglar_Silme.csv', mime='text/csv')
            else:
                st.success("✅ Harika! Şebekede hiçbir gereksiz (dolaylı) bağ bulunamadı. Aktiviteler kusursuz bağlanmış.")
        else:
            st.error("⚠️ XER dosyasında İlişkiler (TASKPRED) tablosu okunamadı.")