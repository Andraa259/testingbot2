import streamlit as st
import docx
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Word to Excel Score Extractor", layout="wide")

st.title("📊 Word to Excel: Score Extractor")
st.write("Unggah file Word Anda untuk mengambil skor penilaian secara otomatis.")

# 1. File Uploader
uploaded_file = st.file_uploader("Pilih file Word (.docx)", type="docx")

def extract_scores(file):
    doc = docx.Document(file)
    data_skor = []
    
    for table in doc.tables:
        for row in table.rows:
            # Mengambil teks tiap sel dalam baris
            cells = [cell.text.strip() for cell in row.cells]
            
            # Berdasarkan dokumen sumber, kolom skor ada di indeks 3, 4, dan 5 
            if len(cells) >= 6:
                aitem = cells[2]
                k = cells[3] # Kejelasan
                r = cells[4] # Relevansi
                s = cells[5] # Kesesuaian
                
                # Validasi: Hanya ambil jika sel berisi angka 1-4 [cite: 10, 16]
                if k.isdigit() or r.isdigit() or s.isdigit():
                    data_skor.append({
                        "Pernyataan/Aitem": aitem,
                        "Skor Kejelasan": k,
                        "Skor Relevansi": r,
                        "Skor Kesesuaian": s
                    })
    return pd.DataFrame(data_skor)

if uploaded_file is not None:
    # 2. Proses Data
    df = extract_scores(uploaded_file)
    
    if not df.empty:
        st.success(f"Berhasil mengekstrak {len(df)} baris skor!")
        
        # Tampilkan Preview secara Horizontal
        st.subheader("Preview Data")
        st.dataframe(df, use_container_width=True)
        
        # 3. Download Button (Excel)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Skor_Validasi')
        
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Download File Excel",
            data=processed_data,
            file_name="Rekap_Skor_Validasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Tidak ditemukan data skor angka di dalam tabel dokumen tersebut.")
