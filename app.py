import docx
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_scores_from_word():
    # 1. Menampilkan jendela pilih file
    root = tk.Tk()
    root.withdraw() # Sembunyikan jendela utama tkinter
    
    file_path = filedialog.askopenfilename(
        title="Pilih File Word Validasi",
        filetypes=[("Word files", "*.docx")]
    )
    
    if not file_path:
        print("Pemilihan file dibatalkan.")
        return

    try:
        doc = docx.Document(file_path)
        data_skor = []
        
        # 2. Proses ekstraksi tabel
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                # Sesuai struktur dokumen:
                # Kolom 3: Kejelasan | Kolom 4: Relevansi | Kolom 5: Kesesuaian
                if len(cells) >= 6:
                    k = cells[3]
                    r = cells[4]
                    s = cells[5]
                    
                    # Validasi: Hanya ambil jika berisi angka (Skor 1-4) [cite: 10]
                    if k.isdigit() or r.isdigit() or s.isdigit():
                        data_skor.append({
                            "Aitem": cells[2], # Teks pernyataan 
                            "Kejelasan": k,
                            "Relevansi": r,
                            "Kesesuaian": s
                        })
        
        if data_skor:
            # 3. Simpan ke Excel secara horizontal
            df = pd.DataFrame(data_skor)
            output_file = file_path.replace(".docx", "_HASIL_SKOR.xlsx")
            df.to_excel(output_file, index=False)
            
            messagebox.showinfo("Berhasil", f"Data berhasil diekstrak ke:\n{output_file}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ditemukan data skor angka di dalam tabel.")

    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")

if __name__ == "__main__":
    extract_scores_from_word()
