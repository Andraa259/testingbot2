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
    Fungsi untuk membuat file Excel di memori dengan kustomisasi layout baru:
    - Baris 1: Kosong
    - Baris 2 & 3: Judul (Merged & Bold)
    - Baris 4: Kosong
    - Baris 5: Header Tabel (Warna Hijau Lebih Gelap & Bold)
    - Baris 6 dst: Data (Nomor dimulai dari 1, Teks Nama Capitalize Each Word)
    - Kolom A kosong total. Semua data bergeser mulai Kolom B.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Hasil_Acak"
    
    # 1. Setup Judul di baris 2 dan 3
    ws['B2'] = "REKAPITULASI URUTAN SHADOWING PRACTICING"
    ws['B3'] = "LAST MEETING WONDERFUL CLASS 2026"
    
    # Menentukan rentang kolom untuk merge judul
    max_col_letter = 'F' if include_link else 'E' 
    
    ws.merge_cells(f'B2:{max_col_letter}2')
    ws.merge_cells(f'B3:{max_col_letter}3')
    
    # Style Judul
    title_font = Font(name='Arial', size=12, bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    
    ws['B2'].font = title_font
    ws['B2'].alignment = center_align
    ws['B3'].font = title_font
    ws['B3'].alignment = center_align
    
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20
    
    # 2. Setup Header Kolom di Baris 5
    headers = ["No.", "Nama Lengkap", "Class", "Status Keangggotaan"]
    if include_link:
        headers.append("Link Drive")
        
    header_font = Font(name='Arial', size=11, bold=True)
    # Menggunakan warna hijau yang lebih gelap/solid (Medium Sea Green / Mint Darker)
    header_fill = PatternFill(start_color="93D1A3", end_color="93D1A3", fill_type="solid") 
    
    for col_num, header_title in enumerate(headers, start=2): # Start dari Kolom B
        cell = ws.cell(row=5, column=col_num)
        cell.value = header_title
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
    
    ws.row_dimensions[5].height = 22
    
    # 3. Masukkan Data ke Excel mulai dari Baris 6
    for row_idx, row_data in enumerate(dataframe_to_rows(df_data, index=False, header=False), start=6):
        # Kolom B: Nomor Urut (Dipastikan mulai tepat dari 1)
        current_no = row_idx - 5
        no_cell = ws.cell(row=row_idx, column=2, value=current_no)
        no_cell.alignment = center_align
        
        # Kolom C dst: Isi data
        for col_offset, value in enumerate(row_data):
            target_col = col_offset + 3 # Mulai dari Kolom C (indeks 3)
            
            # Evaluasi khusus untuk Kolom C (Nama Lengkap) agar otomatis Capitalize Each Word (.title())
            if target_col == 3 and isinstance(value, str):
                value = value.title()
                
            ws.cell(row=row_idx, column=target_col, value=value)
            
    # 4. Auto-fit Lebar Kolom Secara Akurat (Termasuk Kolom No.)
    for col in ws.columns:
        col_letter = col[0].column_letter
        if col_letter == 'A':
            ws.column_dimensions['A'].width = 3  # Kolom A dibiarkan tetap ramping
            continue
            
        max_len = 0
        for cell in col:
            # Lewati baris 2 dan 3 saat menghitung panjang kolom agar efek 'merged cells' tidak merusak kalkulasi lebar
            if cell.row in [2, 3]:
                continue
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
                
        # Berikan padding agar ruang teks aman dan tidak memicu "###" atau terpotong
        ws.column_dimensions[col_letter].width = max(max_len + 4, 10)

    # Simpan ke format buffer bytes
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
            
            # 3. Mapping struktur kolom baru
            nama_col = 'Nama Lengkap'
            status_col = 'Status Peserta'
            link_col = df.columns[5] # Kolom instruksi/Link Drive Video
            
            df_clean_full = df_sorted[[nama_col, target_col, status_col, link_col]].copy()
            df_clean_no_link = df_sorted[[nama_col, target_col, status_col]].copy()
            
            # Format preview DataFrame di Streamlit agar Nama tampil Capitalize Each Word
            df_clean_full[nama_col] = df_clean_full[nama_col].astype(str).str.title()
            df_clean_no_link[nama_col] = df_clean_no_link[nama_col].astype(str).str.title()
            
            st.subheader("📊 Preview Hasil Struktur Baru")
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
