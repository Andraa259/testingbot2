import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Konfigurasi halaman
st.set_page_config(page_title="Pengacak Kelas - Wonderful Class", page_icon="🔀", layout="centered")

st.title("🔀 Sorting & Shuffling Wonderful Class")
st.write("Unggah file hasil Google Form (.csv atau .xlsx) untuk mengacak baris, mengurutkan Kelas B -> A, serta mengekspor dengan format khusus.")

# File uploader (menerima CSV dan Excel)
uploaded_file = st.file_uploader("Pilih file Google Form", type=["csv", "xlsx"])

def generate_styled_excel(df_data, include_link=True):
    """
    Fungsi untuk membuat file Excel di memori dengan kustomisasi layout dan style:
    - Kosongkan Kolom A (Data dimulai dari Kolom B)
    - 6 Baris awal template judul (Baris 3 & 4 digabung/merge & bold)
    - Header kolom di baris 7 (Warna Hijau Muda & bold)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Hasil_Acak"
    
    # 1. Setup Judul di baris 3 dan 4
    ws['B3'] = "REKAPITULASI URUTAN SHADOWING PRACTICING"
    ws['B4'] = "LAST MEETING WONDERFUL CLASS 2026"
    
    # Menentukan rentang kolom yang akan di-merge untuk judul berdasarkan jumlah kolom data
    max_col_letter = 'F' if include_link else 'E' # Karena data bergeser ke kanan, kolom berakhir di E atau F
    
    ws.merge_cells(f'B3:{max_col_letter}3')
    ws.merge_cells(f'B4:{max_col_letter}4')
    
    # Style Judul
    title_font = Font(name='Arial', size=12, bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    
    ws['B3'].font = title_font
    ws['B3'].alignment = center_align
    ws['B4'].font = title_font
    ws['B4'].alignment = center_align
    
    # Set tinggi baris agar lebih rapi
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 20
    
    # 2. Setup Header Kolom di Baris 7
    headers = ["No.", "Nama Lengkap", "Class", "Status Keangggotaan"]
    if include_link:
        headers.append("Link Drive")
        
    header_font = Font(name='Arial', size=11, bold=True)
    header_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Hijau muda
    
    for col_num, header_title in enumerate(headers, start=2): # Start=2 artinya dimulai dari Kolom B
        cell = ws.cell(row=7, column=col_num)
        cell.value = header_title
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
    
    # 3. Masukkan Data Orang ke Excel mulai dari Baris 8, Kolom B
    for row_num, row_data in enumerate(dataframe_to_rows(df_data, index=False, header=False), start=8):
        # Kolom B: Nomor urut (No. 1 - sekian)
        no_cell = ws.cell(row=row_num, column=2, value=row_num - 7)
        no_cell.alignment = center_align
        
        # Kolom C dst: Data Nama, Class, Status, Link
        for col_idx, value in enumerate(row_data, start=3):
            # Jika tidak include_link, baris data link terluar otomatis terpotong dari seleksi DataFrame
            ws.cell(row=row_num, column=col_idx, value=value)
            
    # Auto-fit lebar kolom agar teks tidak terpotong (opsional tapi sangat membantu)
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        if col_letter != 'A': # Abaikan kolom A yang kosong
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    # Atur kolom A tetap ramping karena sengaja dikosongkan
    ws.column_dimensions['A'].width = 3

    # Simpan ke dalam format buffer bytes
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success("File berhasil diunggah!")
        
        with st.expander("👁️ Lihat Data Asli"):
            st.dataframe(df)
            
        target_col = 'Wonderful class'
        if target_col not in df.columns:
            st.error(f"Kolom '{target_col}' tidak ditemukan di dalam file. Pastikan nama kolom sesuai.")
        else:
            # 1. Acak seluruh baris data mentah
            df_shuffled = df.sample(frac=1).reset_index(drop=True)
            
            # 2. Pisahkan dan Urutkan: Kelas B duluan, lalu Kelas A
            df_b = df_shuffled[df_shuffled[target_col].astype(str).str.upper() == 'B']
            df_a = df_shuffled[df_shuffled[target_col].astype(str).str.upper() == 'A']
            df_others = df_shuffled[~df_shuffled[target_col].astype(str).str.upper().isin(['A', 'B'])]
            
            df_sorted = pd.concat([df_b, df_a, df_others], ignore_index=True)
            
            # 3. Mapping struktur kolom baru dan membersihkan data lama
            # Sesuai urutan target: Nama Lengkap; Class; Status Keangggotaan; Link Drive
            # Menggunakan penamaan kolom dinamis berdasarkan posisi/nama asli file gform
            nama_col = 'Nama Lengkap'
            status_col = 'Status Peserta'
            link_col = df.columns[5] # Kolom ke-6 adalah instruksi/Link Drive Video Practice
            
            # Buat DataFrame dasar yang susunan kolomnya sudah rapi (tanpa Timestamp)
            df_clean_full = df_sorted[[nama_col, target_col, status_col, link_col]].copy()
            df_clean_no_link = df_sorted[[nama_col, target_col, status_col]].copy()
            
            st.subheader("📊 Preview Hasil Struktur Baru")
            st.write("Data di bawah ini adalah representasi mentahnya. Saat di-download, data otomatis bergeser mulai dari **Kolom B** dan dilengkapi template judul atas.")
            st.dataframe(df_clean_full)
            
            # --- Proses Generate Excel Menggunakan OpenPyXL ---
            excel_full = generate_styled_excel(df_clean_full, include_link=True)
            excel_no_link = generate_styled_excel(df_clean_no_link, include_link=False)
            
            # --- Area Download Tombol ---
            st.markdown("---")
            st.subheader("📥 Download Pilihan File Output")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.info("📂 **Output 1: Versi Lengkap**\n\nMemiliki struktur penuh termasuk kolom Link Drive.")
                st.download_button(
                    label="📥 Download Versi Lengkap (Excel)",
                    data=excel_full,
                    file_name="Rekap_Shadowing_Wonderful_Class_Lengkap.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_full"
                )
                
            with col2:
                st.warning("📂 **Output 2: Tanpa Link Drive**\n\nKolom F atau bagian Link Drive dihapus dari tabel.")
                st.download_button(
                    label="📥 Download Tanpa Link Drive (Excel)",
                    data=excel_no_link,
                    file_name="Rekap_Shadowing_Wonderful_Class_Tanpa_Link.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_no_link"
                )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
