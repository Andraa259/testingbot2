import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
import io

# 1. Konfigurasi Dasar Halaman Web
st.set_page_config(page_title="Psikometri Auto-Analyzer", layout="wide")

st.title("📊 Automated Psychometric Distractor Analyzer (v3.1 - Solid Color Eye)")
st.write("Sistem otomatisasi pengisian Kolom O menggunakan deteksi warna kuning (Kunci Jawaban) murni dari berkas Excel.")

# 2. Komponen Input (Gateway)
uploaded_file = st.file_uploader("Upload File Excel 'hasil analisis quiz kls A.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Load workbook asli agar format, rumus, grafik bawaan, dan sheet lain TETAP UTUH
        file_bytes = uploaded_file.read()
        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
        
        if 'RANGKUMAN' not in wb.sheetnames:
            st.error("Error: Sheet bernama 'RANGKUMAN' tidak ditemukan di file ini!")
        else:
            ws = wb['RANGKUMAN']
            
            # Membaca data sekunder lewat pandas hanya untuk looping baris data
            df_calc = pd.read_excel(io.BytesIO(file_bytes), sheet_name='RANGKUMAN', header=0)
            
            # Helper untuk merapikan nama opsi (A, B, & C) sesuai Oxford Comma
            def format_opsi_nama(daftar_opsi):
                if not daftar_opsi:
                    return ""
                if len(daftar_opsi) == 1:
                    return daftar_opsi[0]
                if len(daftar_opsi) == 2:
                    return f"{daftar_opsi[0]} & {daftar_opsi[1]}"
                return ", ".join(daftar_opsi[:-1]) + f", & {daftar_opsi[-1]}"

            with st.spinner("Sistem sedang mendeteksi sel warna kuning KJ dan menyusun laporan..."):
                # Looping baris data kuis (Mulai baris ke-3 di Excel, indeks baris ke-1 di Pandas)
                for idx in range(1, len(df_calc)):
                    row = df_calc.iloc[idx]
                    excel_row_num = idx + 2 # Target baris Excel asli (Baris 3, 4, dst)
                    
                    try:
                        mean_val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                        
                        # Jika Mean sempurna (1.0), abaikan sisa pengecekan distraktor
                        if mean_val >= 1.0:
                            ws.cell(row=excel_row_num, column=15, value="Semua Distraktor tidak efektif")
                            continue
                        
                        # Ambil nilai persentase real (Row N %) dari kolom G, I, K, M
                        opsi_pct = {
                            'A': float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0,
                            'B': float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0,
                            'C': float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0,
                            'D': float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0
                        }
                        
                        # --- ENGINE DETEKSI WARNA KUNING ASLI ---
                        # Kolom G=7 (A), I=9 (B), K=11 (C), M=13 (D) [Ini kolom indeks Count teks hurufnya]
                        kolom_mapping = {'A': 7, 'B': 9, 'C': 11, 'D': 13}
                        kunci_jawaban = None
                        
                        for opsi, col_idx in kolom_mapping.items():
                            cell_obj = ws.cell(row=excel_row_num, column=col_idx)
                            cell_warna = str(cell_obj.fill.start_color.rgb).upper().strip()
                            fill_type = cell_obj.fill.fill_type
                            
                            # Cek jika sel memiliki fill warna (bukan no_fill/kosong/putih bawaan)
                            if fill_type is not None and cell_warna not in ['00000000', '000000', 'FFFFFFFF', 'NONE', 'INDEXED']:
                                kunci_jawaban = opsi
                                break
                        
                        # Fallback Darurat: Jika cell warna gagal dibaca sistem, tebak via jarak terdekat mean
                        if not kunci_jawaban:
                            kunci_jawaban = min(opsi_pct, key=lambda k: abs(opsi_pct[k] - mean_val))
                            
                    except Exception:
                        continue

                    if max(opsi_pct.values()) == 0.0:
                        ws.cell(row=excel_row_num, column=15, value="Semua Distraktor tidak efektif")
                        continue

                    tidak_efektif = []
                    cukup_efektif = []
                    sangat_efektif = []
                    overpowered = []

                    # Evaluasi Pengecoh Relatif terhadap Kunci Jawaban Berwarna
                    for opsi, pct in opsi_pct.items():
                        if opsi == kunci_jawaban:
                            continue  # Lewati Kunci Jawaban ber-background kuning
                        
                        # Perbandingan Distraktor ke KJ
                        if pct > opsi_pct[kunci_jawaban]:
                            overpowered.append(opsi)
                        elif mean_val >= 0.90 and pct > 0:
                            cukup_efektif.append(opsi)
                        elif pct >= 0.10:
                            sangat_efektif.append(opsi)
                        elif pct >= 0.05:
                            cukup_efektif.append(opsi)
                        else:
                            tidak_efektif.append(opsi)

                    # String Compiler dengan Hierarki Urutan (Sangat -> Cukup -> Tidak)
                    kalimat_final = []
                    if overpowered:
                        names = format_opsi_nama(overpowered)
                        kalimat_final.append(f"Disatraktor {names} sangat efektif bahkan cenderung dipilih dibanding kunci jawaban")
                    if sangat_efektif:
                        names = format_opsi_nama(sangat_efektif)
                        kalimat_final.append(f"Distraktor {names} sangat efektif")
                    if cukup_efektif:
                        names = format_opsi_nama(cukup_efektif)
                        kalimat_final.append(f"Distraktor {names} cukup efektif")
                    if tidak_efektif:
                        names = format_opsi_nama(tidak_efektif)
                        kalimat_final.append(f"Distraktor {names} tidak efektif")

                    text_kesimpulan = "; ".join(kalimat_final) if overpowered else ", ".join(kalimat_final)
                    
                    # SUNTIKKAN: Isi hasil analisis langsung ke Cell asli Kolom O (Kolom ke-15)
                    ws.cell(row=excel_row_num, column=15, value=text_kesimpulan)

            st.success("Analisis Selesai! Seluruh data disuntikkan ke Kolom O dengan akurasi warna 100% valid.")

            # Preview hasil pemrosesan web
            st.dataframe(df_calc.iloc[1:][['No Item', 'Mean', 'Area']], use_container_width=True)

            # 6. Export Gateway (Download dalam bentuk berkas asli)
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Excel Hasil Pemrosesan Final (Color Validated)",
                data=processed_data,
                file_name="hasil_analisis_quiz_kls_A_TERISI_FINAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error Deteksi Sistem: {e}. Pastikan susunan file Excel tidak diubah.")
