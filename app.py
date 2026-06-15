import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Konfigurasi halaman
st.set_page_config(page_title="Rekap Kelulusan - Wonderful Class", page_icon="🎓", layout="centered")

st.title("🎓 Rekapitulasi Kelulusan Wonderful Class")
st.write("Unggah file rekap presensi (.csv atau .xlsx) untuk memproses status kelulusan peserta secara otomatis.")

# File uploader
uploaded_file = st.file_uploader("Pilih file Presensi Wonderful Class", type=["csv", "xlsx"])

def generate_graduation_excel(df_processed):
    """
    Fungsi untuk memformat output Excel sesuai kustomisasi:
    - Kolom A dikosongkan.
    - Baris 2 & 3 berisi judul (Bold, Merged, Center).
    - Baris 5 adalah header tabel (Hijau Tua, Font Bold Putih, Center).
    - Semua sel data diberikan border hitam tipis.
    - Alignment: Nama Lengkap (Left), Kolom lainnya (Center).
    - Auto-fit konten untuk semua kolom.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap_Kelulusan"
    
    # 1. Setup Judul di baris 2 dan 3
    title_text = "REKAPITULASI KELULUSAN WONDERFUL CLASS 2026"
    ws['B2'] = title_text
    ws['B3'] = title_text
    
    # Merge dari kolom B sampai E (karena ada 4 kolom data)
    ws.merge_cells('B2:E2')
    ws.merge_cells('B3:E3')
    
    title_font = Font(name='Arial', size=12, bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    
    ws['B2'].font = title_font
    ws['B2'].alignment = center_align
    ws['B3'].font = title_font
    ws['B3'].alignment = center_align
    
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20
    
    # Definisi Border Tipis untuk Tabel
    thin_side = Side(border_style="thin", color="000000")
    table_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # 2. Setup Header Kolom di Baris 5
    headers = ["No.", "Nama Lengkap", "NBI", "Syarat Kelulusan"]
    
    # Menggunakan warna Hijau Tua (Hex: 2E7D32) dengan font putih agar kontras
    header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    
    for col_num, header_title in enumerate(headers, start=2): # Start=2 artinya Kolom B
        cell = ws.cell(row=5, column=col_num)
        cell.value = header_title
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = table_border
        
    ws.row_dimensions[5].height = 24
    
    # Alignment isi data
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    
    # 3. Memasukkan Data Peserta (Mulai Baris 6)
    for idx, row in df_processed.iterrows():
        current_row = idx + 6
        
        # Kolom B: No. (Center)
        no_cell = ws.cell(row=current_row, column=2, value=idx + 1)
        no_cell.alignment = align_center
        no_cell.border = table_border
        
        # Kolom C: Nama Lengkap (Left & Capitalize Each Word)
        nama_cell = ws.cell(row=current_row, column=3, value=str(row['Nama']).title())
        nama_cell.alignment = align_left
        nama_cell.border = table_border
        
        # Kolom D: NBI (Center)
        nbi_cell = ws.cell(row=current_row, column=4, value=str(row['NIM / NPM']))
        nbi_cell.alignment = align_center
        nbi_cell.border = table_border
        
        # Kolom E: Syarat Kelulusan (Center - Hanya Angka 1-4)
        status_cell = ws.cell(row=current_row, column=5, value=int(row['Syarat Kelulusan']))
        status_cell.alignment = align_center
        status_cell.border = table_border
        
        ws.row_dimensions[current_row].height = 20
        
    # 4. Auto-fit Lebar Kolom Secara Akurat
    for col in ws.columns:
        col_letter = col[0].column_letter
        if col_letter == 'A':
            ws.column_dimensions['A'].width = 3 # Kolom A tetap ramping kosong
            continue
            
        max_len = 0
        for cell in col:
            if cell.row in [2, 3]: # Lewati baris judul agar tidak merusak kalkulasi
                continue
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
                
        ws.column_dimensions[col_letter].width = max(max_len + 4, 10)
        
    # Simpan ke format buffer bytes
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


if uploaded_file is not None:
    try:
        # Membaca data mentah
        if uploaded_file.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_file, header=None)
        else:
            raw_df = pd.read_excel(uploaded_file, header=None)
            
        # Mencari posisi baris header tabel presensi secara dinamis
        header_row_idx = None
        for idx, row in raw_df.iterrows():
            row_str_list = [str(x).strip() for x in row.values if pd.notna(x)]
            if 'Nama' in row_str_list and ('NIM / NPM' in row_str_list or 'Kelas' in row_str_list):
                header_row_idx = idx
                break
                
        if header_row_idx is None:
            st.error("Struktur file tidak dikenali. Pastikan kolom 'Nama' dan 'NIM / NPM' tersedia di file presensi.")
        else:
            st.success("File presensi berhasil dimuat!")
            
            # Memotong dataframe mulai dari baris header yang ditemukan
            df = raw_df.iloc[header_row_idx:].copy()
            df.columns = [str(c).strip() for c in df.iloc[0]]
            df = df[1:].reset_index(drop=True)
            
            # Membersihkan data dari baris kosong
            df = df.dropna(subset=['Nama', 'NIM / NPM'], how='all')
            
            # List nama kolom pertemuan 1-11
            p_cols = [f"Pertemuan {i}" for i in range(1, 12)]
            
            # 5. Logika Baru Pemetaan Angka Kelulusan
            def hitung_kelulusan(row):
                # Hitung jumlah checkmark di pertemuan 1 s/d 11
                total_hadir_reguler = 0
                for col in p_cols:
                    if col in df.columns and str(row[col]).strip() == '✅':
                        total_hadir_reguler += 1
                
                # Cek kehadiran di Pertemuan Final
                ikut_final = 'Pertemuan Final' in df.columns and str(row['Pertemuan Final']).strip() == '✅'
                
                # Aturan Percabangan Baru Berbasis Angka Murni:
                if total_hadir_reguler >= 6 and ikut_final:
                    return 1  # 1 untuk Keduanya
                elif total_hadir_reguler >= 6 and not ikut_final:
                    return 2  # 2 untuk pertemuan 6 kali atau lebih
                elif total_hadir_reguler < 6 and ikut_final:
                    return 3  # 3 untuk pertemuan final
                else:
                    return 4  # 4 untuk tidak lulus
            
            # Terapkan logika angka ke baris data
            df['Syarat Kelulusan'] = df.apply(hitung_kelulusan, axis=1)
            
            # --- FITUR SORTING OTOMATIS ---
            # Mengurutkan baris data berdasarkan angka Syarat Kelulusan (1 ke 4)
            df = df.sort_values(by='Syarat Kelulusan', ascending=True).reset_index(drop=True)
            
            # Pilih kolom utama untuk ditampilkan di preview UI Streamlit
            df_preview = df[['Nama', 'NIM / NPM', 'Syarat Kelulusan']].copy()
            df_preview['Nama'] = df_preview['Nama'].astype(str).str.title()
            
            st.subheader("📊 Preview Hasil Logika Kelulusan (Sudah Disortir 1-4)")
            st.dataframe(df_preview)
            
            # --- Proses Pembuatan Excel Hasil ---
            excel_data = generate_graduation_excel(df)
            
            st.markdown("---")
            st.subheader("📥 Download Hasil Akhir")
            st.download_button(
                label="📥 Download Rekap Kelulusan (Excel)",
                data=excel_data,
                file_name="Rekap_Kelulusan_Wonderful_Class_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data presensi: {e}")
