import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re

# 1. Konfigurasi Halaman Web
st.set_page_config(page_title="Psikometri Cell Plotter v8.0", layout="wide")

st.title("📊 Automated Psychometric Cell Plotter (v8.0 - Fresh System)")
st.write("Sistem Baru: Membaca data aman via Pandas, mengisi data Kolom O, dan memplot nomor aitem (VAR00010 -> 10) langsung ke koordinat sel Sheet 3 tanpa menggunakan berkas rusak.")

# 2. Gateway Input Berkas Excel dari User
uploaded_file = st.file_uploader("Upload File Excel Analisis Kuis", type=["xlsx"])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        
        # --- LANGKAH 1: BYPASS ENGINE (PANDAS MEMBACA MENTAH DATA EXCEL) ---
        with st.spinner("Membaca data kuis secara aman via Pandas..."):
            # Pandas tidak peduli dengan XML stylesheet yang rusak, jadi dijamin 100% lolos eror
            df_calc = pd.read_excel(io.BytesIO(file_bytes), sheet_name='RANGKUMAN', header=0)
            
        # Fungsi ekstraksi label: VAR00010 -> 10
        def ekstrak_angka_item(teks_item):
            match = re.search(r'\d+', str(teks_item))
            return int(match.group()) if match else str(teks_item)

        # Helper untuk format nama opsi teks kuis
        def format_opsi_nama(daftar_opsi):
            if not daftar_opsi: return ""
            if len(daftar_opsi) == 1: return daftar_opsi[0]
            if len(daftar_opsi) == 2: return f"{daftar_opsi[0]} & {daftar_opsi[1]}"
            return ", ".join(daftar_opsi[:-1]) + f", & {daftar_opsi[-1]}"

        # --- LANGKAH 2: HITUNG LOGIKA DISTRAKTOR PADA KOLOM O ---
        hasil_kolom_o = [""] # Baris sub-header dikosongkan
        
        for idx in range(1, len(df_calc)):
            row = df_calc.iloc[idx]
            try:
                mean_val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                if mean_val >= 1.0:
                    hasil_kolom_o.append("Semua Distraktor tidak efektif")
                    continue
                    
                opsi_pct = {
                    'A': float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0,
                    'B': float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0,
                    'C': float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0,
                    'D': float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0
                }
                # Lock KJ menggunakan rumus jarak selisih absolut terdekat dengan Mean
                kunci_jawaban = min(opsi_pct, key=lambda k: abs(opsi_pct[k] - mean_val))
            except Exception:
                hasil_kolom_o.append("")
                continue

            if max(opsi_pct.values()) == 0.0:
                hasil_kolom_o.append("Semua Distraktor tidak efektif")
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
            hasil_kolom_o.append(text_kesimpulan)

        df_calc['Analisis Distraktor'] = hasil_kolom_o

        # --- LANGKAH 3: MEMBUAT BERKAS WORKBOOK EXCEL BARU YANG STERIL ---
        wb_new = openpyxl.Workbook()
        
        # Lembar Kerja 1: RANGKUMAN
        ws1 = wb_new.active
        ws1.title = "RANGKUMAN"
        
        # Pindahkan data bersih dari Pandas ke sheet baru
        for r_idx, row in enumerate(dataframe_to_rows(df_calc, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws1.cell(row=r_idx, column=c_idx, value=value)

        # Lembar Kerja 2: GRAFIK POSISI AITEM (Membuat Grid Baru)
        ws2 = wb_new.create_sheet(title="GRAFIK POSISI AITEM")
        
        # Layout Judul Header di Sheet 3
        ws2["B2"] = "PETA SEBARAN MATRIKS KUALITAS AITEM"
        ws2["B2"].font = Font(size=14, bold=True, name="Arial")
        ws2["B3"] = "Koordinat ditentukan berdasarkan: Sumbu Horizontal (X) = Mean | Sumbu Vertikal (Y) = CITC"
        ws2["B3"].font = Font(size=10, italic=True, name="Arial")

        # Ambil hanya item soal valid yang mengandung teks 'VAR'
        df_clean = df_calc.dropna(subset=['No Item', 'Mean', 'Corrected Item-Total Correlation']).copy()
        df_clean = df_clean[df_clean['No Item'].str.contains('VAR', na=False)]

        # --- PROCESSOR PEMETAAN KOORDINAT SEL (ANTI-MENIMPA) ---
        with st.spinner("Memetakan nomor butir soal ke dalam grid sel Sheet 3..."):
            for _, row in df_clean.iterrows():
                label_singkat = ekstrak_angka_item(row['No Item'])
                mean_val = float(row['Mean'])
                citc_val = float(row['Corrected Item-Total Correlation'])
                
                # RUMUS KONVERSI MATEMATIKA KE SEL GRID EXCEL:
                # Nilai Mean (0.0 sampai 1.0) dikonversi proporsional ke Kolom E sampai O (Kolom 5 s/d 15)
                target_col = int(5 + (mean_val * 10))
                
                # Nilai CITC (0.8 turun ke -0.2) dikonversi proporsional ke Baris 6 s/d 26
                target_row = int(6 + ((0.8 - citc_val) * 20))
                
                # Kunci pengaman koordinat agar tidak keluar dari perimeter grid chart
                target_col = max(5, min(15, target_col))
                target_row = max(6, min(26, target_row))
                
                # Baca isi sel tujuan saat ini
                current_val = ws2.cell(row=target_row, column=target_col).value
                
                # SYARAT UTAMA: Jika koordinat sel sudah ada nilainya, gabungkan dengan koma agar tidak menimpa!
                if current_val:
                    ws2.cell(row=target_row, column=target_col, value=f"{current_val}, {label_singkat}")
                else:
                    ws2.cell(row=target_row, column=target_col, value=label_singkat)
                
                # Hias sel titik koordinat agar rapi dan kontras
                ws2.cell(row=target_row, column=target_col).font = Font(bold=True, color="000000", name="Arial")
                ws2.cell(row=target_row, column=target_col).alignment = Alignment(horizontal="center")
                ws2.cell(row=target_row, column=target_col).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

        # --- LANGKAH 4: PROSES EKSPOR DAN GENERATE FILE ---
        output = io.BytesIO()
        wb_new.save(output)
        processed_data = output.getvalue()
        
        st.success("Sukses Total! Masalah arsitektur file Excel bawaan SPSS berhasil di-bypass.")
        
        st.download_button(
            label="📥 Download Excel Hasil Pemrosesan Baru (100% Bersih)",
            data=processed_data,
            file_name="hasil_analisis_quiz_KLS_A_FIXED_TOTAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Eror Kritis Sistem: {e}. Pastikan file input sesuai.")
