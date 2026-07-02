import streamlit as st
import pandas as pd
import numpy as np
import io

# 1. Konfigurasi Dasar Halaman Web
st.set_page_config(page_title="Psikometri Auto-Analyzer", layout="wide")

st.title("📊 Automated Psychometric Distractor Analyzer (v2.1)")
st.write("Sistem otomatisasi pengisian kolom analisis distraktor berbasis aturan evaluasi dosen.")

# 2. Komponen Input (Gateway)
uploaded_file = st.file_uploader("Upload File Excel 'hasil analisis quiz kls A.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    try:
        # PENGATURAN EVALUASI: Membaca SHEET "RANGKUMAN" secara eksplisit
        df_raw = pd.read_excel(uploaded_file, sheet_name='RANGKUMAN', header=0)
        
        # Slicing data kuis asli (Mulai baris ke-3 di Excel)
        df_data = df_raw.iloc[1:].copy().reset_index(drop=True)

        # Fungsi Helper untuk merapikan gabungan nama opsi (A, B, & C)
        def format_opsi_nama(daftar_opsi):
            if not daftar_opsi:
                return ""
            if len(daftar_opsi) == 1:
                return daftar_opsi[0]
            if len(daftar_opsi) == 2:
                return f"{daftar_opsi[0]} & {daftar_opsi[1]}"
            # Jika lebih dari 2 opsi, gunakan koma sebelum ampersand (Oxford Comma style)
            return ", ".join(daftar_opsi[:-1]) + f", & {daftar_opsi[-1]}"

        # 3. Core Engine Logika Analisis Distraktor (Kolom G sampai O)
        def proses_distraktor_per_baris(row):
            try:
                mean_val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                
                # EVALUASI 2: Jika Mean Sempurna (1.0), abaikan sisa pengecekan distraktor
                if mean_val >= 1.0:
                    return "Semua Distraktor tidak efektif"
                
                opsi_pct = {
                    'A': float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0,
                    'B': float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0,
                    'C': float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0,
                    'D': float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0
                }
            except Exception:
                return ""

            # Menentukan Kunci Jawaban (Mencari nilai tertinggi sebagai jangkar baseline)
            kunci_jawaban = max(opsi_pct, key=opsi_pct.get)
            
            if max(opsi_pct.values()) == 0.0:
                return "Semua Distraktor tidak efektif"

            tidak_efektif = []
            cukup_efektif = []
            sangat_efektif = []
            overpowered = []

            # Evaluasi Pengecoh
            for opsi, pct in opsi_pct.items():
                if opsi == kunci_jawaban:
                    continue  # Skip Kunci Jawaban berpola kuning
                
                # Rule Anomali Kasus VAR057
                if pct > opsi_pct[kunci_jawaban]:
                    overpowered.append(opsi)
                # Rule Toleransi Kasus VAR001 (Kuis Terlalu Mudah)
                elif mean_val >= 0.90 and pct > 0:
                    cukup_efektif.append(opsi)
                # Threshold Normal
                elif pct >= 0.10:
                    sangat_efektif.append(opsi)
                elif pct >= 0.05:
                    cukup_efektif.append(opsi)
                else:
                    tidak_efektif.append(opsi)

            # 4. String Compiler: Aturan Urutan & Struktur Kalimat Evaluasi 2
            kalimat_final = []
            
            if overpowered:
                names = format_opsi_nama(overpowered)
                # Typo pada laporan asli "Disatraktor" dipertahankan agar match 100% dengan dosen
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

            # Menggabungkan komponen kalimat menggunakan titik koma (;) khusus jika ada anomali overpowered
            if overpowered:
                return "; ".join(kalimat_final)
            return ", ".join(kalimat_final)

        # 5. Batch Processing running system
        with st.spinner("Sistem sedang memproses data kuis Kelas A..."):
            hasil_analisis = df_data.apply(proses_distraktor_per_baris, axis=1)
            
            # Memasukkan kembali hasil kerja program ke struktur file asli di Kolom O (Indeks 14)
            df_raw.iloc[1:, 14] = hasil_analisis.values

        st.success("Analisis Sukses! Aturan Evaluasi 2 berhasil diterapkan ke dalam sistem.")

        # Preview Data Terupdate
        st.dataframe(df_raw.iloc[1:][['No Item', 'Mean', 'Area', 'Analisis Distraktor']], use_container_width=True)

        # 6. Export Gateway (Download hasil dalam bentuk .XLSX asli)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_raw.to_excel(writer, sheet_name='RANGKUMAN', index=False)
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Download Excel Hasil Pemrosesan Terbaru (v2.1)",
            data=processed_data,
            file_name="hasil_analisis_quiz_kls_A_TERISI_V2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error Deteksi Sistem: {e}. Pastikan file tidak sedang dibuka di komputer lain.")
