import streamlit as st
import pandas as pd
import numpy as np
import io

# 1. Konfigurasi Dasar Halaman Web
st.set_page_config(page_title="Psikometri Auto-Analyzer", layout="wide")

st.title("📊 Automated Psychometric Distractor Analyzer (v2.0)")
st.write("Sistem otomatisasi pengisian kolom analisis distraktor berbasis aturan evaluasi dosen.")

# 2. Komponen Input (Gateway)
uploaded_file = st.file_uploader("Upload File Excel 'hasil analisis quiz kls A.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    try:
        # PENGATURAN EVALUASI: Membaca SHEET "RANGKUMAN" secara eksplisit
        # Header=0 agar baris pertama (No Item, Mean, dsb) dibaca sebagai nama kolom utama
        df_raw = pd.read_excel(uploaded_file, sheet_name='RANGKUMAN', header=0)
        
        # Slicing data: Mengambil baris data kuis asli (Mulai dari indeks ke-1 di Pandas, yang merupakan baris 3 di Excel)
        # Menghapus baris sub-header ['Count', 'Row N %'] agar tidak ikut terhitung
        df_data = df_raw.iloc[1:].copy().reset_index(drop=True)

        # 3. Core Engine Logika Analisis Distraktor (Kolom G sampai O)
        def proses_distraktor_per_baris(row):
            # Mapping data kolom G sampai N berdasarkan letak indeks kolom asli kuis
            # G=6, H=7 (A) | I=8, J=9 (B) | K=10, L=11 (C) | M=12, N=13 (D)
            try:
                opsi_pct = {
                    'A': float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0,
                    'B': float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0,
                    'C': float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0,
                    'D': float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0
                }
                mean_val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
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
                    overpowered.append(f"Disatraktor {opsi} sangat efektif bahkan cenderung dipilih dibanding kunci jawaban")
                # Rule Toleransi Kasus VAR001 (Kuis Terlalu Mudah)
                elif mean_val >= 0.90 and pct > 0:
                    cukup_efektif.append(f"Distraktor {opsi} cukup efektif")
                # Threshold Normal
                elif pct >= 0.10:
                    sangat_efektif.append(f"Distraktor {opsi} sangat efektif")
                elif pct >= 0.05:
                    cukup_efektif.append(f"Distraktor {opsi} cukup efektif")
                else:
                    tidak_efektif.append(f"Distraktor {opsi} tidak efektif")

            # 4. String Compiler: Aturan Urutan (Sangat -> Cukup -> Tidak Efektif)
            kalimat_final = []
            
            if overpowered:
                kalimat_final.extend(overpowered)
                
            if sangat_efektif:
                # Contoh: "Distraktor A & B sangat efektif"
                names = " & ".join([x.split()[1] for x in sangat_efektif])
                kalimat_final.append(f"Distraktor {names} sangat efektif")
                
            if cukup_efektif:
                names = " & ".join([x.split()[1] for x in cukup_efektif])
                kalimat_final.append(f"Distraktor {names} cukup efektif")
                
            if tidak_efektif:
                names = " & ".join([x.split()[1] for x in tidak_efektif])
                kalimat_final.append(f"Distraktor {names} tidak efektif")

            return ", ".join(kalimat_final)

        # 5. Batch Processing running system
        with st.spinner("Sistem sedang memproses data kuis Kelas A..."):
            hasil_analisis = df_data.apply(proses_distraktor_per_baris, axis=1)
            
            # Memasukkan kembali hasil kerja program ke struktur file asli di Kolom O (Indeks 14)
            df_raw.iloc[1:, 14] = hasil_analisis.values

        st.success("Analisis Sukses! Kolom O telah terisi otomatis dengan urutan hierarki yang benar.")

        # Preview Data Terupdate
        st.dataframe(df_raw.iloc[1:][['No Item', 'Mean', 'Area', 'Analisis Distraktor']], use_container_width=True)

        # 6. Export Gateway (Download hasil dalam bentuk .XLSX asli)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_raw.to_excel(writer, sheet_name='RANGKUMAN', index=False)
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Download Excel Hasil Pemrosesan (Kolom O Terisi)",
            data=processed_data,
            file_name="hasil_analisis_quiz_kls_A_TERISI.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error Deteksi Sistem: {e}. Pastikan file tidak sedang dibuka di komputer lain.")
