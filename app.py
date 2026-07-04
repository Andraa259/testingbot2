import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import matplotlib.pyplot as plt
import io

# 1. Konfigurasi Dasar Halaman Web
st.set_page_config(page_title="Psikometri Auto-Analyzer & Plotter", layout="wide")

st.title("📊 Automated Psychometric Plotter & Analyzer (v4.0)")
st.write("Sistem otomatisasi pengisian Kolom O dan pembuatan Grafik Kuadrant Area di Sheet 3 secara presisi.")

# 2. Komponen Input (Gateway)
uploaded_file = st.file_uploader("Upload File Excel 'hasil analisis quiz kls A.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Load workbook asli untuk mempertahankan seluruh data, rumus, dan format
        file_bytes = uploaded_file.read()
        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
        
        if 'RANGKUMAN' not in wb.sheetnames:
            st.error("Error: Sheet bernama 'RANGKUMAN' tidak ditemukan!")
        else:
            ws_rangkuman = wb['RANGKUMAN']
            
            # Membaca data sekunder lewat pandas untuk kebutuhan kalkulasi & koordinat grafik
            df_calc = pd.read_excel(io.BytesIO(file_bytes), sheet_name='RANGKUMAN', header=0)
            
            # Helper untuk merapikan nama opsi (A, B, & C)
            def format_opsi_nama(daftar_opsi):
                if not daftar_opsi:
                    return ""
                if len(daftar_opsi) == 1:
                    return daftar_opsi[0]
                if len(daftar_opsi) == 2:
                    return f"{daftar_opsi[0]} & {daftar_opsi[1]}"
                return ", ".join(daftar_opsi[:-1]) + f", & {daftar_opsi[-1]}"

            # --- PROSES 1: OTOMASI ANALISIS DISTRAKTOR (KOLOM O) ---
            with st.spinner("Sistem sedang memproses Kolom O berbasis warna kuning KJ..."):
                for idx in range(1, len(df_calc)):
                    row = df_calc.iloc[idx]
                    excel_row_num = idx + 2
                    
                    try:
                        mean_val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                        if mean_val >= 1.0:
                            ws_rangkuman.cell(row=excel_row_num, column=15, value="Semua Distraktor tidak efektif")
                            continue
                            
                        opsi_pct = {
                            'A': float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0,
                            'B': float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0,
                            'C': float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0,
                            'D': float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0
                        }
                        
                        kolom_mapping = {'A': 7, 'B': 9, 'C': 11, 'D': 13}
                        kunci_jawaban = None
                        for opsi, col_idx in kolom_mapping.items():
                            cell_obj = ws_rangkuman.cell(row=excel_row_num, column=col_idx)
                            cell_warna = str(cell_obj.fill.start_color.rgb).upper().strip()
                            if cell_obj.fill.fill_type is not None and cell_warna not in ['00000000', '000000', 'FFFFFFFF', 'NONE', 'INDEXED']:
                                kunci_jawaban = opsi
                                break
                        
                        if not kunci_jawaban:
                            kunci_jawaban = min(opsi_pct, key=lambda k: abs(opsi_pct[k] - mean_val))
                    except Exception:
                        continue

                    if max(opsi_pct.values()) == 0.0:
                        ws_rangkuman.cell(row=excel_row_num, column=15, value="Semua Distraktor tidak efektif")
                        continue

                    tidak_efektif, cukup_efektif, sangat_efektif, overpowered = [], [], [], []
                    for opsi, pct in opsi_pct.items():
                        if opsi == kunci_jawaban: continue
                        if pct > opsi_pct[kunci_jawaban]: overpowered.append(opsi)
                        elif mean_val >= 0.90 and pct > 0: cukup_efektif.append(opsi)
                        elif pct >= 0.10: sangat_efektif.append(opsi)
                        elif pct >= 0.05: cukup_efektif.append(opsi)
                        else: tidak_efektif.append(opsi)

                    kalimat_final = []
                    if overpowered: kalimat_final.append(f"Disatraktor {format_opsi_nama(overpowered)} sangat efektif bahkan cenderung dipilih dibanding kunci jawaban")
                    if sangat_efektif: kalimat_final.append(f"Distraktor {format_opsi_nama(sangat_efektif)} sangat efektif")
                    if cukup_efektif: kalimat_final.append(f"Distraktor {format_opsi_nama(cukup_efektif)} cukup efektif")
                    if tidak_efektif: kalimat_final.append(f"Distraktor {format_opsi_nama(tidak_efektif)} tidak efektif")

                    text_kesimpulan = "; ".join(kalimat_final) if overpowered else ", ".join(kalimat_final)
                    ws_rangkuman.cell(row=excel_row_num, column=15, value=text_kesimpulan)

            # --- PROSES 2: AUTOMATED GRAPHIC PLOTTER KHUSUS SHEET 3 ---
            with st.spinner("Sistem sedang menggambar dan menyuntikkan Grafik Area ke Sheet 3..."):
                # 1. Bersihkan data untuk plotting grafik (Hapus baris kosong/sub-header)
                df_clean = df_calc.dropna(subset=['No Item', 'Mean', 'Corrected Item-Total Correlation']).copy()
                df_clean = df_clean[df_clean['No Item'].str.contains('VAR', na=False)]
                
                # 2. Konfigurasi Kanvas Grafik Baru
                fig, ax = plt.subplots(figsize=(12, 9))
                
                # 3. Plot Seluruh 85 Koordinat Aitem (Bentuk Bintang Oranye Ukuran Besar)
                ax.scatter(df_clean['Corrected Item-Total Correlation'], df_clean['Mean'], 
                           color='#FF8C00', marker='*', s=150, edgecolor='black', linewidth=0.5, label='Butir Soal')
                
                # 4. Membuat Garis Batas Kuadrant Sesuai Rumus Teori Dosen
                ax.axhline(y=0.5, color='red', linestyle='--', alpha=0.6, linewidth=1.5, label='Batas Batas Kesulitan (Mean = 0.50)')
                ax.axvline(x=0.3, color='blue', linestyle='--', alpha=0.6, linewidth=1.5, label='Batas Daya Beda (Korelasi = 0.30)')
                ax.axvline(x=0.2, color='purple', linestyle=':', alpha=0.5, linewidth=1.2, label='Batas Batas Minimum Gugur (Korelasi = 0.20)')
                
                # 5. Melabeli Setiap Titik Soal Secara Otomatis (VAR00001 - VAR00085)
                for _, row in df_clean.iterrows():
                    ax.annotate(str(row['No Item']), 
                                (float(row['Corrected Item-Total Correlation']), float(row['Mean'])),
                                textcoords="offset points", 
                                xytext=(0, 6), 
                                ha='center', 
                                fontsize=7, 
                                fontweight='semibold')
                
                # Dekorasi Estetika Grafik Kartesius
                ax.set_title("PETA KEDUDUKAN KUALITAS AITEM QUIZ (AUTOMATED CLUSTERING)", fontsize=14, fontweight='bold', pad=15)
                ax.set_xlabel("Daya Beda (Corrected Item-Total Correlation)", fontsize=11, fontweight='bold', labelpad=10)
                ax.set_ylabel("Tingkat Kesulitan (Mean)", fontsize=11, fontweight='bold', labelpad=10)
                ax.set_xlim(-0.3, 0.8) # Mengakomodasi item bernilai korelasi minus (seperti VAR010)
                ax.set_ylim(-0.05, 1.05)
                ax.grid(True, linestyle=':', alpha=0.5)
                ax.legend(loc='lower right', fontsize=9)
                
                # Teks Penanda Area Kuadrant pada Bidang Grafik
                ax.text(0.5, 0.8, 'AREA F\n(Diterima Prima)', fontsize=12, color='green', alpha=0.7, fontweight='bold', ha='center')
                ax.text(0.1, 0.8, 'AREA E\n(Soal Mudah / Revisi)', fontsize=12, color='darkorange', alpha=0.7, fontweight='bold', ha='center')
                ax.text(0.5, 0.2, 'AREA D\n(Soal Sukar / Cukup)', fontsize=12, color='blue', alpha=0.7, fontweight='bold', ha='center')
                ax.text(-0.15, 0.5, 'WILAYAH\nGUGUR', fontsize=12, color='red', alpha=0.7, fontweight='bold', ha='center')
                
                plt.tight_layout()
                
                # 6. Simpan Grafik ke Objek Memori (Bytes Stream)
                img_buf = io.BytesIO()
                plt.savefig(img_buf, format='png', dpi=180)
                img_buf.seek(0)
                plt.close(fig)
                
                # 7. Ambil Target Sheet 3 (Gunakan Sheet yang ada atau buat baru jika terhapus)
                if 'GRAFIK POSISI AITEM' in wb.sheetnames:
                    ws_sheet3 = wb['GRAFIK POSISI AITEM']
                    # Bersihkan objek gambar lama jika ada agar tidak menumpuk double
                    ws_sheet3._images.clear() 
                else:
                    ws_sheet3 = wb.create_sheet(title='GRAFIK POSISI AITEM')
                
                # 8. Suntikkan Gambar Grafik Baru Tepat Mulai Cell B3 di Sheet 3
                xl_img = OpenpyxlImage(img_buf)
                ws_sheet3.add_image(xl_img, 'B3')

            st.success("Sukses Mutlak! Kolom O Terisi dan Grafik Kuadrant Sheet 3 Selesai Dirender Otomatis.")

            # 9. Export Gateway
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Excel Hasil Pemrosesan Final (Kolom O + Grafik Sheet 3 Aman)",
                data=processed_data,
                file_name="hasil_analisis_quiz_kls_A_FINAL_ALL_SHEETS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error Deteksi Sistem: {e}. Pastikan format file Excel sesuai.")
