import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
import io

# 1. Konfigurasi Dasar Halaman Web
st.set_page_config(page_title="Psikometri Auto-Analyzer", layout="wide")

st.title("📊 Automated Psychometric Distractor Analyzer (v3.0 - Color Eye)")
st.write("Sistem otomatisasi pengisian Kolom O menggunakan deteksi warna kuning (Kunci Jawaban) asli Excel.")

# 2. Komponen Input
uploaded_file = st.file_uploader("Upload File Excel 'hasil analisis quiz kls A.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Load workbook asli untuk mempertahankan seluruh format & warna
        file_bytes = uploaded_file.read()
        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
        
        if 'RANGKUMAN' not in wb.sheetnames:
            st.error("Error: Sheet bernama 'RANGKUMAN' tidak ditemukan di file ini!")
        else:
            ws = wb['RANGKUMAN']
            
            # Membaca data sekunder untuk kebutuhan looping baris
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

            with st.spinner("Sistem sedang memindai warna Kunci Jawaban dan menyuntikkan hasil..."):
                # Looping baris data kuis (Mulai baris ke-3 di Excel)
                for idx in range(1, len(df_calc)):
                    row = df_calc.iloc[idx]
                    excel_row_num = idx + 2 # Baris Excel asli
                    
                    try:
                        mean_val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                        
                        # Guard clause jika Mean = 1.0
                        if mean_val >= 1.0:
                            ws.cell(row=excel_row_num, column=15, value="Semua Distraktor tidak efektif")
                            continue
                        
                        # Ambil nilai persentase real (Row N %)
                        opsi_pct = {
                            'A': float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0,
                            'B': float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0,
                            'C': float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0,
                            'D': float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0
                        }
                        
                        # --- DETEKSI WARNA KUNING (THE COLOR DETECTOR ENGINE) ---
                        # Kolom G=7, I=9, K=11, M=13 (Kolom Count / Pilihan Huruf Opsi)
                        kolom_mapping = {'A': 7, 'B': 9, 'C': 11, 'D': 13}
                        kunci_jawaban = None
                        
                        for opsi, col_idx in kolom_mapping.items():
                            cell_warna = ws.cell(row=excel_row_num, column=col_idx).fill.start_color.rgb
                            # Cek variasi kode warna kuning di Excel (biasanya FFFF00, FFFFFF00, atau berjenis tipe 00000000)
                            if cell_warna and str(cell_warna).strip() not in ['00000000', '000000', 'FFFFFFFF', 'None']:
                                kunci_jawaban = opsi
                                break
                        
                        # Fallback jika warna tidak terdeteksi oleh script (kembali ke max)
                        if not kunci_jawaban:
                            kunci_jawaban = max(opsi_pct, key=opsi_pct.get)
                            
                    except Exception:
                        continue

                    if max(opsi_pct.values()) == 0.0:
                        ws.cell(row=excel_row_num, column=15, value="Semua Distraktor tidak efektif")
                        continue

                    tidak_efektif = []
                    cukup_efektif = []
                    sangat_efektif = []
                    overpowered = []

                    # Evaluasi Pengecoh berdasarkan Kunci Jawaban Warna
                    for opsi, pct in opsi_pct.items():
                        if opsi == kunci_jawaban:
                            continue
                        
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

                    # String Compiler
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
                    
                    # Tulis hasil tepat di Kolom O asli
                    ws.cell(row=excel_row_num, column=15, value=text_kesimpulan)

            st.success("Analisis Sukses! Kunci jawaban dibaca langsung dari warna kuning sel asli.")

            # Save ke memori stream
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Excel Hasil Pemrosesan (Color Decoder Match)",
                data=processed_data,
                file_name="hasil_analisis_quiz_kls_A_TERISI_FINAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error Deteksi Sistem: {e}. Pastikan file tidak rusak.")
