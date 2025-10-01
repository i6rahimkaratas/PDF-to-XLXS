import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
import pandas as pd
import re

def pdf_to_excel():
    pdf_path = pdf_entry.get()
    excel_path = excel_entry.get()
    
    if not pdf_path or not excel_path:
        messagebox.showerror("Hata", "Lütfen hem PDF hem de Excel dosya yolunu seçin!")
        return
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            all_text = ""
            
            for page in pdf_reader.pages:
                all_text += page.extract_text() + "\n"
        
        lines = all_text.strip().split('\n')
        data = []
        
        for line in lines:
            if line.strip():
                cells = re.split(r'\s{2,}|\t', line.strip())
                data.append(cells)
        
        max_cols = max(len(row) for row in data) if data else 0
        for row in data:
            while len(row) < max_cols:
                row.append("")
        
        df = pd.DataFrame(data)
        df.to_excel(excel_path, index=False, header=False)
        
        messagebox.showinfo("Başarılı", "PDF başarıyla Excel'e dönüştürüldü!")
        
    except Exception as e:
        messagebox.showerror("Hata", f"Dönüştürme sırasında hata oluştu:\n{str(e)}")

def select_pdf():
    filename = filedialog.askopenfilename(
        title="PDF Dosyası Seç",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    if filename:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, filename)

def select_excel():
    filename = filedialog.asksaveasfilename(
        title="Excel Dosyasını Kaydet",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if filename:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, filename)

root = tk.Tk()
root.title("PDF'den Excel'e Dönüştürücü")
root.geometry("600x250")
root.resizable(False, False)

main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(fill=tk.BOTH, expand=True)

pdf_label = tk.Label(main_frame, text="PDF Dosyası:", font=("Arial", 10))
pdf_label.grid(row=0, column=0, sticky="w", pady=5)

pdf_entry = tk.Entry(main_frame, width=50, font=("Arial", 10))
pdf_entry.grid(row=0, column=1, padx=10, pady=5)

pdf_button = tk.Button(main_frame, text="Gözat", command=select_pdf, width=10, font=("Arial", 10))
pdf_button.grid(row=0, column=2, pady=5)

excel_label = tk.Label(main_frame, text="Excel Dosyası:", font=("Arial", 10))
excel_label.grid(row=1, column=0, sticky="w", pady=5)

excel_entry = tk.Entry(main_frame, width=50, font=("Arial", 10))
excel_entry.grid(row=1, column=1, padx=10, pady=5)

excel_button = tk.Button(main_frame, text="Gözat", command=select_excel, width=10, font=("Arial", 10))
excel_button.grid(row=1, column=2, pady=5)

convert_button = tk.Button(main_frame, text="Dönüştür", command=pdf_to_excel, 
                          width=20, height=2, font=("Arial", 12, "bold"), 
                          bg="#4CAF50", fg="white", cursor="hand2")
convert_button.grid(row=2, column=0, columnspan=3, pady=30)

root.mainloop()
