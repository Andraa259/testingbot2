import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import io
import re

# 1. Konfigurasi Halaman Web
st.set_page_config(page_title="Psikometri Cell Plotter v7.0", layout="wide")

st.title("📊 Automated Psychometric Cell Plotter (v7.0 - Anti-Crash System)")
st.write("Sistem baru: Mengisi data Kolom O dan memplot nomor aitem langsung ke koordinat grid sel Sheet 3 tanpa merusak layout.")

# 2. Gateway Input Berkas
uploaded_file = st.file_uploader("Upload File Excel Analisis Kuis", type=["xlsx"])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        
        # --- LANGKAH 1: BACA DATA VALID DENGAN PANDAS ENGINE ---
        with st.spinner("Membaca arsitektur data kuis via Pandas..."):
            df_calc = pd.read_excel(io.BytesIO(file_bytes), sheet_name='RANGKUMAN', header=0)
            
        # Ekstraksi angka saja dari nama item (Contoh: VAR00010 -> 10)
        def ekstrak_angka_item(teks_item):
            match = re.search(r'\d+', str(teks_item))
            return int(match.group()) if match else str(teks_item)

        # Helper untuk merapikan nama opsi (A, B, & C)
        def format_opsi_nama(daftar_opsi):
            if not daftar_opsi: return ""
            if len(daftar_opsi) == 1: return daftar_opsi[0]
            if len(daftar_opsi) == 2: return f"{daftar_opsi[0]} & {daftar_opsi[1]}"
            return ", ".join(daftar_opsi[:-1]) + f", & {daftar_opsi[-1]}"

        # --- LANGKAH 2: CORE ENGINE KOMPUTASI DISTRAKTOR (KOLOM O) ---
        hasil_kolom_o = [""] # Baris sub-header kosong
        
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
                # Kunci KJ berdasarkan selisih terdekat Mean (Solusi Kasus VAR047)
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

        # --- LANGKAH 3: MEMBUAT WORKBOOK BARU YANG 100% BERSIH DARI EROR XML ---
        wb_new = openpyxl.Workbook()
        
        # Sheet 1: RANGKUMAN
        ws1 = wb_new.active
        ws1.title = "RANGKUMAN"
        
        # Tulis ulang data hasil komputasi dari Pandas
        for r_idx, row in enumerate(dataframe_to_rows(df_calc, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws1.cell(row=r_idx, column=c_idx, value=value)

        # Sheet 2: GRAFIK POSISI AITEM (Membangun Grid Sel tanpa Menimpa Struktur)
        ws2 = wb_new.create_sheet(title="GRAFIK POSISI AITEM")
        
        # Membuat judul area grafik di sel atas
        ws2["B2"] = "PETA SEBARAN MATRIKS KUALITAS AITEM (CELL COORDINATE PLOTTER)"
        ws2["B2"].font = Font(size=14, bold=True)
        
        # Keterangan Sumbu Grafik pada grid sel
        ws2["B4"] = "Sumbu Vertikal (Y) = Daya Beda (CITC) | Sumbu Horizontal (X) = Kesulitan (Mean)"
        ws2["B4"].font = Font(italic=True)

        # Filter data aitem valid untuk plotting koordinat
        df_clean = df_calc.dropna(subset=['No Item', 'Mean', 'Corrected Item-Total Correlation']).copy()
        df_clean = df_clean[df_clean['No Item'].str.contains('VAR', na=False)]

        # --- ENGINE CELL COORDINATE MAPPING ---
        # Kita petakan data statistik ke koordinat baris & kolom grid Excel:
        # Sumbu Horizontal (Mean: 0.0 sampai 1.0) dipetakan ke Kolom E sampai O (Kolom 5 - 15)
        # Sumbu Vertikal (CITC: -0.2 sampai 0.8) dipetakan ke Baris 6 sampai 26
        
        with st.spinner("Memetakan nomor aitem ke koordinat sel Sheet 3..."):
            for _, row in df_clean.iterrows():
                label_singkat = ekstrak_angka_item(row['No Item'])
                mean_val = float(row['Mean'])
                citc_val = float(row['Corrected Item-Total Correlation'])
                
                # Formula Konversi Matematika ke Titik Koordinat Grid Sel Excel
                # Kolom X: Rentang Mean 0.0 - 1.0 dipetakan proporsional ke indeks kolom 5 hingga 15
                target_col = int(5 + (mean_val * 10))
                
                # Baris Y: Rentang CITC 0.8 down to -0.2 dipetakan ke indeks baris 6 hingga 26
                target_row = int(6 + ((0.8 - citc_val) * 20))
                
                # Amankan batas koordinat agar tidak keluar dari grid grafik
                target_col = max(5, min(15, target_col))
                target_row = max(6, min(26, target_row))
                
                # Ambil nilai sel tujuan koordinat saat ini
                current_val = ws2.cell(row=target_row, column=target_col).value
                
                # Jika koordinat sel tersebut sudah diisi aitem lain, gabungkan agar tidak menimpa!
                if current_val:
                    ws2.cell(row=target_row, column=target_col, value=f"{current_val}, {label_singkat}")
                else:
                    ws2.cell(row=target_row, column=target_col, value=label_singkat)
                
                # Beri gaya visual bintang/penanda teks pada sel koordinat tersebut
                ws2.cell(row=target_row, column=target_col).font = Font(bold=True, color="FF0000")
                ws2.cell(row=target_row, column=target_col).alignment = Alignment(horizontal="center")
                ws2.cell(row=target_row, column=target_col).fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")

        # --- PROSES DOWNLOAD ---
        output = io.BytesIO()
        wb_new.save(output)
        processed_data = output.getvalue()
        
        st.success("Sukses Mutlak! Berkas baru yang bersih berhasil dibuat dengan pemetaan koordinat sel presisi.")
        
        st.download_button(
            label="📥 Download Excel Hasil Pemrosesan Final (Bebas Eror XML SPSS)",
            data=processed_data,
            file_name="hasil_analisis_quiz_KLS_A_ALL_OK.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Eror Kritis Sistem: {e}. Pastikan susunan tabel kuis valid.")
