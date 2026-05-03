import docx
import pandas as pd

def extract_only_scores(file_path, output_name):
    doc = docx.Document(file_path)
    all_scores = []
    
    for table in doc.tables:
        for row in table.rows:
            # Mengambil teks dari sel 0-5 (No, Dimensi, Aitem, Kejelasan, Relevansi, Kesesuaian)
            cells = [cell.text.strip() for cell in row.cells]
            
            if len(cells) >= 6:
                aitem_text = cells[2]
                s_kejelasan = cells[3]
                s_relevansi = cells[4]
                s_kesesuaian = cells[5]
                
                # Filter: Pastikan kita hanya mengambil baris yang memiliki nilai angka
                # Ini menghindari pengambilan header tabel atau baris kosong
                if any(s.isdigit() for s in [s_kejelasan, s_relevansi, s_kesesuaian]):
                    all_scores.append({
                        "Aitem Skala": aitem_text,
                        "Skor Kejelasan": s_kejelasan,
                        "Skor Relevansi": s_relevansi,
                        "Skor Kesesuaian": s_kesesuaian
                    })

    # Membuat DataFrame
    df = pd.DataFrame(all_scores)
    
    # Simpan ke Excel
    df.to_excel(output_name, index=False)
    print(f"Selesai! {len(df)} baris skor berhasil diekstrak ke {output_name}")

# Jalankan dengan nama file Anda
file_input = "Form_Validasi_Expert_Judgement_Forgiveness_Qori_atul_Tri_Setya_A.docx"
file_output = "Rekap_Skor_Horizontal.xlsx"
extract_only_scores(file_input, file_output)
