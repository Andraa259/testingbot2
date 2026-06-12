import streamlit as st
import pandas as pd
import io

# Konfigurasi halaman
st.set_page_config(page_title="Pengacak Kelas - Wonderful Class", page_icon="🔀", layout="centered")

st.title("🔀 Sorting & Shuffling Wonderful Class")
st.write("Unggah file hasil Google Form (.csv atau .xlsx) untuk mengacak baris dan mengurutkan berdasarkan Kelas B -> Kelas A.")

# File uploader (menerima CSV dan Excel)
uploaded_file = st.file_uploader("Pilih file Google Form", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        # Membaca file berdasarkan formatnya
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success("File berhasil diunggah!")
        
        # Menampilkan data asli (opsional)
        with st.expander("👁️ Lihat Data Asli"):
            st.dataframe(df)
            
        # Validasi apakah kolom 'Wonderful class' ada
        target_col = 'Wonderful class'
        if target_col not in df.columns:
            st.error(f"Kolom '{target_col}' tidak ditemukan di dalam file. Pastikan nama kolom sesuai.")
        else:
            # 1. Acak seluruh data terlebih dahulu (agar urutan orang per kelas menjadi acak)
            # random_state bisa dihapus jika ingin hasil acak yang selalu berubah setiap klik
            df_shuffled = df.sample(frac=1).reset_index(drop=True)
            
            # 2. Pisahkan data menjadi kelas B dan kelas A
            df_b = df_shuffled[df_shuffled[target_col].astype(str).str.upper() == 'B']
            df_a = df_shuffled[df_shuffled[target_col].astype(str).str.upper() == 'A']
            df_others = df_shuffled[~df_shuffled[target_col].astype(str).str.upper().isin(['A', 'B'])]
            
            # 3. Gabungkan kembali dengan urutan B dulu baru A, lalu data lain jika ada
            df_final = pd.concat([df_b, df_a, df_others], ignore_index=True)
            
            st.subheader("📊 Hasil Pemrosesan Data")
            st.dataframe(df_final)
            
            # --- Fitur Download Hasil ---
            # Mengonversi DataFrame ke Excel di dalam memori
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Hasil_Acak')
            
            st.download_button(
                label="📥 Download Hasil (Excel)",
                data=buffer.getvalue(),
                file_name="Hasil_Sorting_Wonderful_Class.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
