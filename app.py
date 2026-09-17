import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import ollama

import config
import database

database.init_db()

def extract_text(pdf_path, max_pages=3):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:max_pages]:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Xato matn o'qishda: {e}")
    return text.strip()

def analyze_with_ai(text):
    prompt = f"""
Quyidagi rasmiy hujjat matnini o'rganib chiqib, uchta narsani aniqlang:
1. Qaysi viloyat yoki hududga tegishli? (Variantlar: {', '.join(config.REGIONS)})
2. Qaysi tashkilot yoki organga tegishli? (Variantlar: Kadastr_Agentligi, Davlat_Kadastrlari_Palatasi, Bosh_Prokuratura, IIV, DXX, Sudlar, Boshqa_Organlar)
3. Hujjat turi nima? (Variantlar: Buyruqlar, Topshiriqlar, Xatlar, Taqdimnomalar, Arizalar, Boshqa)

Javobni FAQAT quyidagi formatda bering, ortiqcha so'z yozmang:
HUDUD: <aniqlangan_hudud>
TASHKILOT: <aniqlangan_tashkilot>
TUR: <aniqlangan_tur>

Hujjat matni:
{text[:2500]}
"""
    try:
        response = ollama.chat(
            model="llama3",
            messages=[{"role": "user", "content": prompt}]
        )
        ans = response["message"]["content"].strip().split("\n")
        
        region = "Toshkent_shahri"
        org = "Kadastr_Agentligi"
        doc_type = "Xatlar"
        
        for line in ans:
            if "HUDUD:" in line:
                val = line.replace("HUDUD:", "").strip()
                for r in config.REGIONS:
                    if r.lower() in val.lower():
                        region = r
                        break
            elif "TASHKILOT:" in line:
                val = line.replace("TASHKILOT:", "").strip()
                for o in config.ORGANIZATIONS + config.LAW_ENFORCEMENT:
                    if o.lower() in val.lower():
                        org = o
                        break
            elif "TUR:" in line:
                val = line.replace("TUR:" , "").strip()
                for d in config.DOC_TYPES:
                    if d.lower() in val.lower():
                        doc_type = d
                        break
        return region, org, doc_type
    except Exception as e:
        print(f"AI tahlilida xatolik: {e}")
        return "Toshkent_shahri", "Kadastr_Agentligi", "Boshqa"

def process_single_file():
    file_path = filedialog.askopenfilename(
        title="PDF hujjatni tanlang",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not file_path:
        return

    status_var.set("Hujjat o'qilmoqda...")
    root.update()

    text = extract_text(file_path)
    if not text:
        messagebox.showerror("Xato", "PDF matnini o'qib bo'lmadi!")
        status_var.set("Tayyor")
        return

    status_var.set("AI tahlil qilmoqda (Oflayn)...")
    root.update()

    region, org, doc_type = analyze_with_ai(text)

    if org in config.LAW_ENFORCEMENT:
        target_dir = os.path.join(config.BASE_DIR, "Huquqni_muhofaza_organlari", org, doc_type)
    else:
        target_dir = os.path.join(config.BASE_DIR, region, org, doc_type)

    os.makedirs(target_dir, exist_ok=True)
    file_name = os.path.basename(file_path)
    dest_path = os.path.join(target_dir, file_name)

    shutil.copy2(file_path, dest_path)

    saved_time = database.save_document_record(file_name, dest_path, region, org, doc_type, text)

    status_var.set(f"Joylandi: {org} -> {doc_type} ({saved_time})")
    messagebox.showinfo("Muvaffaqiyatli", f"Fayl saqlandi:\n{dest_path}\nVaqti: {saved_time}")

def run_search():
    query = search_entry.get().strip()
    if not query:
        messagebox.showwarning("Diqqat", "Qidiruv so'zini kiriting!")
        return

    for item in result_tree.get_children():
        result_tree.delete(item)

    results = database.search_documents(query)
    for row in results:
        result_tree.insert("", "end", values=(row[0], row[2], row[3], row[4], row[5], row[1]))

def open_selected_file(event):
    selected = result_tree.selection()
    if not selected:
        return
    item = result_tree.item(selected[0])
    file_path = item["values"][5]
    if os.path.exists(file_path):
        os.startfile(file_path)
    else:
        messagebox.showerror("Xato", "Fayl topilmadi!")

root = tk.Tk()
root.title("Kadastr Hujjatlar Arxivi va AI Saralovchi")
root.geometry("850x550")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

tab_upload = ttk.Frame(notebook)
notebook.add(tab_upload, text="Hujjat Yuklash va Saralash")

lbl_main = tk.Label(tab_upload, text="PDF Hujjatlarni Avtomatik Arxivga Saralash", font=("Arial", 14, "bold"))
lbl_main.pack(pady=30)

btn_upload = tk.Button(tab_upload, text="PDF Hujjatni Tanlash", command=process_single_file, font=("Arial", 12), bg="#007acc", fg="white", padx=15, pady=8)
btn_upload.pack(pady=15)

status_var = tk.StringVar(value="Tizim tayyor")
lbl_status = tk.Label(tab_upload, textvariable=status_var, font=("Arial", 11), fg="green")
lbl_status.pack(pady=20)

tab_search = ttk.Frame(notebook)
notebook.add(tab_search, text="Hujjatlar Bazasidan Qidirish")

search_frame = tk.Frame(tab_search)
search_frame.pack(fill="x", padx=10, pady=10)

search_entry = tk.Entry(search_frame, font=("Arial", 12))
search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

search_btn = tk.Button(search_frame, text="Qidirish", command=run_search, font=("Arial", 11), bg="#28a745", fg="white")
search_btn.pack(side="right")

cols = ("Fayl nomi", "Hudud", "Tashkilot", "Turi", "Saqlangan vaqti", "Manzil")
result_tree = ttk.Treeview(tab_search, columns=cols, show="headings")
for col in cols:
    result_tree.heading(col, text=col)
    result_tree.column(col, width=120)

result_tree.column("Fayl nomi", width=180)
result_tree.column("Manzil", width=0, stretch=False)
result_tree.pack(fill="both", expand=True, padx=10, pady=10)
result_tree.bind("<Double-1>", open_selected_file)

lbl_hint = tk.Label(tab_search, text="Faylni ochish uchun ustiga sichqonchani ikki marta bosing", font=("Arial", 9), fg="gray")
lbl_hint.pack(pady=5)

if __name__ == "__main__":
    root.mainloop()
