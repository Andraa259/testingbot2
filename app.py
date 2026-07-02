import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
import io

# 1. Konfigurasi Dasar Halaman Web
st.set_page_config(page_title="Psikometri Auto-Analyzer", layout="wide")

st.title("📊 Automated Psychometric Distractor Analyzer (v2.3 - Format Keeper)")
st.write("Sistem otomatisasi pengisian Kolom O tanpa merusak sheet lain, warna, atau format asli Excel.")

# 2. Komponen Input
uploaded_file = st.file_uploader("Upload File Excel 'hasil analisis quiz kls A.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    try:
        # --- SOLUSI: Load workbook asli agar format & semua sheet terjaga ---
        file_bytes = uploaded_file.read()
        wb = load_workbook(io.BytesIO(file_bytes), data_only=False) # data_only=False menjaga rumus asli
        
        if 'RANGKUMAN' not in wb.sheetnames:
            st.error("Error: Sheet bernama 'RANGKUMAN' tidak ditemukan di file ini!")
        else:
            ws = wb['RANGKUMAN']
            
            # Membaca data untuk komputasi (Pandas hanya untuk nyari logic saja)
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

            with st.spinner("Sistem sedang menyuntikkan hasil analisis ke Kolom O asli..."):
                # Looping baris data kuis (Mulai baris ke-3 di Excel, indeks baris ke-1 di Pandas)
                for idx in range(1, len(df_calc)):
                    row = df_calc.iloc[idx]
                    excel_row_num = idx + 2 # Menyesuaikan posisi baris asli di Excel (baris 3, 4, dst)
                    
                    try:
                        mean_val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                        
                        # Guard clause jika Mean = 1.0
                        if mean_val >= 1.0:
                            ws.cell(row=excel_row_num, column=15, value="Semua Distraktor tidak efektif")
                            continue
                            
                        opsi_pct = {
                            'A': float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0,
                            'B': float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0,
                            'C': float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0,
                            'D': float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0
                        }
                    except Exception:
                        continue

                    # Deteksi Kunci Jawaban (Selisih terdekat dengan Mean)
                    kunci_jawaban = min(opsi_pct, key=lambda k: abs(opsi_pct[k] - mean_val))
                    
                    if max(opsi_pct.values()) == 0.0:
                        ws.cell(row=excel_row_num, column=15, value="Semua Distraktor tidak efektif")
                        continue

                    tidak_efektif = []
                    cukup_efektif = []
                    sangat_efektif = []
                    overpowered = []

                    # Evaluasi Pengecoh
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
                    
                    # --- WRITING INJECTION: Isi langsung ke cell asli Kolom O (Column 15) ---
                    ws.cell(row=excel_row_num, column=15, value=text_kesimpulan)

            st.success("Analisis Sukses! Data disuntikkan langsung ke dalam file asli tanpa merusak format.")

            # 6. Export Gateway (Menyimpan workbook asli langsung ke memory stream)
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Excel Hasil Pemrosesan (Format & Sheet Aman)",
                data=processed_data,
                file_name="hasil_analisis_quiz_kls_A_TERISI_FINAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error Deteksi Sistem: {e}. Pastikan file tidak rusak.")
