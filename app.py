import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.chart import ScatterChart, Reference, Series
import io
import re

# 1. Konfigurasi Halaman Web
st.set_page_config(page_title="Psikometri Native Plotter v6.0", layout="wide")

st.title("📊 Automated Psychometric Native Plotter (v6.0)")
st.write("Sistem insert titik koordinat otomatis ke dalam Sheet 3 menggunakan Native Excel Chart tanpa merusak struktur layout.")

# 2. Gateway Input Berkas
uploaded_file = st.file_uploader("Upload File Excel Analisis Kuis", type=["xlsx"])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        
        # Membaca data menggunakan Pandas untuk ekstraksi koordinat aman
        df_calc = pd.read_excel(io.BytesIO(file_bytes), sheet_name='RANGKUMAN', header=0)
        
        # Bersihkan data: hanya ambil baris yang mengandung kode 'VAR'
        df_clean = df_calc.dropna(subset=['No Item', 'Mean', 'Corrected Item-Total Correlation']).copy()
        df_clean = df_clean[df_clean['No Item'].str.contains('VAR', na=False)]
        
        # Ekstraksi angka saja dari nama item (Contoh: VAR00010 -> 10)
        def ekstrak_angka_item(teks_item):
            match = re.search(r'\d+', str(teks_item))
            return int(match.group()) if match else str(teks_item)
            
        df_clean['Label Angka'] = df_clean['No Item'].apply(ekstrak_angka_item)

        # Load workbook asli dengan openpyxl (Guna mempertahankan struktur sheet ke-3)
        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
        
        # Cari nama sheet ke-3 (indeks posisi ke-2) atau cari berdasarkan kata kunci grafik
        sheet_names = wb.sheetnames
        target_sheet_name = None
        for name in sheet_names:
            if 'GRAFIK' in name.upper() or 'SHEET3' in name.upper().replace(" ", ""):
                target_sheet_name = name
                break
        
        if not target_sheet_name:
            # Fallback jika tidak terdeteksi, ambil sheet urutan ke-3
            target_sheet_name = sheet_names[2] if len(sheet_names) >= 3 else sheet_names[-1]

        ws_sheet3 = wb[target_sheet_name]
        st.info(f"Target repositori grafik terdeteksi di sheet: **{target_sheet_name}**")

        # --- PROSES UTAMA: INSERT NATIVE SCATTER CHART ---
        with st.spinner("Menghitung koordinat dan menyisipkan bentuk grafik native..."):
            
            # Agar openpyxl bisa membuat chart, kita tulis data koordinat sementara di pojok jauh sheet 3 (Misal Kolom Z & AA)
            # Ini dilakukan agar struktur template tengah dari dosen tidak terganggu sama sekali.
            ws_sheet3["Z1"] = "Label"
            ws_sheet3["AA1"] = "Mean (Horizontal X)"
            ws_sheet3["AB1"] = "CITC (Vertikal Y)"
            
            for idx, row in df_clean.reset_index().iterrows():
                row_num = idx + 2
                ws_sheet3[f"Z{row_num}"] = row['Label Angka']
                ws_sheet3[f"AA{row_num}"] = float(row['Mean'])
                ws_sheet3[f"AB{row_num}"] = float(row['Corrected Item-Total Correlation'])
            
            # Menginisiasi Objek Scatter Chart bawaan Microsoft Excel
            chart = ScatterChart()
            chart.title = "PETA KEDUDUKAN AITEM (NATIVE AREA PLOT)"
            chart.style = 13
            chart.x_axis.title = 'Tingkat Kesulitan (Mean)'
            chart.y_axis.title = 'Daya Beda (CITC)'
            
            # Mengambil referensi range data koordinat yang baru kita tulis tadi
            max_data_row = len(df_clean) + 1
            xvalues = Reference(ws_sheet3, min_col=27, min_row=2, max_row=max_data_row)
            yvalues = Reference(ws_sheet3, min_col=28, min_row=2, max_row=max_data_row)
            
            series = Series(yvalues, xvalues, title_from_data=False)
            
            # Modifikasi bentuk penanda menjadi bintang/titik dan aktifkan label angka
            series.marker.symbol = "star"
            series.marker.size = 7
            series.graphicalProperties.line.noFill = True  # Hilangkan garis antar titik koordinat
            
            chart.series.append(series)
            
            # Tampilkan Data Label berupa angka item (10, 85, dll) di tiap titik koordinat
            chart.dataLabels = openpyxl.chart.label.DataLabelList()
            chart.dataLabels.showVal = False
            
            # Set ukuran dimensi grafik agar pas di layout sheet 3 tanpa menimpa tabel dosen
            chart.width = 19
            chart.height = 13
            
            # Letakkan Grafik di Cell B4 (Sesuaikan dengan ruang kosong di template sheet 3 kamu)
            ws_sheet3.add_chart(chart, "B4")

        # --- PROSES EKSPOR ---
        output = io.BytesIO()
        wb.save(output)
        processed_data = output.getvalue()
        
        st.success("Sukses! Koordinat titik aitem berhasil ditempel secara aman ke Sheet 3.")
        
        st.download_button(
            label="📥 Download Excel Hasil Pemrosesan (Layout Aman)",
            data=processed_data,
            file_name="hasil_analisis_quiz_FIX_SHEET3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Gagal memproses berkas: {e}")
