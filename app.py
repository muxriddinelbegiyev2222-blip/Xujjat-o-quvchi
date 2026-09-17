import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import ollama

import config
import database

database.init_db()

# --- Matnni sug'urish va AI tahlili ---
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
                val = line.replace("TUR:", "").strip()
                for d in config.DOC_TYPES:
                    if d.lower() in val.lower():
                        doc_type = d
                        break
        return region, org, doc_type
    except Exception as e:
        print(f"AI tahlilida xatolik: {e}")
        return "Toshkent_shahri", "Kadastr_Agentligi", "Boshqa"

# --- PDF Viewer Oynasi ---
class PDFViewerWindow(tk.Toplevel):
    def __init__(self, parent, pdf_path):
        super().__init__(parent)
        self.title(f"Ko'rish: {os.path.basename(pdf_path)}")
        self.geometry("850x900")
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.current_page = 0
        self.zoom = 1.2

        # Boshqaruv tugmalari
        ctrl_frame = tk.Frame(self, bg="#2c3e50", pady=5)
        ctrl_frame.pack(fill="x")

        self.btn_prev = tk.Button(ctrl_frame, text="◀ Oldingi", command=self.prev_page, bg="#34495e", fg="white")
        self.btn_prev.pack(side="left", padx=10)

        self.lbl_page = tk.Label(ctrl_frame, text="", bg="#2c3e50", fg="white", font=("Arial", 11, "bold"))
        self.lbl_page.pack(side="left", padx=10)

        self.btn_next = tk.Button(ctrl_frame, text="Keyingi ▶", command=self.next_page, bg="#34495e", fg="white")
        self.btn_next.pack(side="left", padx=10)

        btn_open_ext = tk.Button(ctrl_frame, text="Tashqi dasturda ochish", command=lambda: os.startfile(self.pdf_path), bg="#27ae60", fg="white")
        btn_open_ext.pack(side="right", padx=10)

        # Rasm ko'rsatish qismi
        self.canvas_frame = tk.Frame(self)
        self.canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.canvas_frame, bg="#7f8c8d")
        self.scrollbar_y = tk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar_x = tk.Scrollbar(self.canvas_frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.scrollbar_x.set, yscrollcommand=self.scrollbar_y.set)

        self.scrollbar_y.pack(side="right", fill="y")
        self.scrollbar_x.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.render_page()

    def render_page(self):
        page = self.doc.load_page(self.current_page)
        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.photo = ImageTk.PhotoImage(img)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.photo)
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self.lbl_page.config(text=f"Sahifa: {self.current_page + 1} / {len(self.doc)}")

    def next_page(self):
        if self.current_page < len(self.doc) - 1:
            self.current_page += 1
            self.render_page()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()

# --- Asosiy Ilova ---
class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Davlat Kadastri va Organlar Hujjatlari Axborot-Boshqaruv Tizimi")
        self.geometry("1100x750")
        self.configure(bg="#f5f6fa")

        # Uslub
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", font=("Arial", 10), rowheight=26)
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"), background="#dcdde1")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        # Vkladkalar
        self.tab_dashboard = ttk.Frame(self.notebook)
        self.tab_upload = ttk.Frame(self.notebook)
        self.tab_explorer = ttk.Frame(self.notebook)
        self.tab_search = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_dashboard, text=" 📊 Monitoring & Dashboard ")
        self.notebook.add(self.tab_upload, text=" 📥 Hujjat Qabul Qilish ")
        self.notebook.add(self.tab_explorer, text=" 📂 Arxiv Explorer ")
        self.notebook.add(self.tab_search, text=" 🔍 Qidiruv Tizimi ")
        self.notebook.add(self.tab_settings, text=" ⚙️ Admin & Sozlamalar ")

        self.init_dashboard_tab()
        self.init_upload_tab()
        self.init_explorer_tab()
        self.init_search_tab()
        self.init_settings_tab()

        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def on_tab_change(self, event):
        selected = self.notebook.select()
        if selected == str(self.tab_dashboard):
            self.update_dashboard()
        elif selected == str(self.tab_explorer):
            self.load_tree_directories()

    # ================= 1. DASHBOARD TAB =================
    def init_dashboard_tab(self):
        header = tk.Frame(self.tab_dashboard, bg="#2f3640", height=60)
        header.pack(fill="x")
        tk.Label(header, text="KADASTR VA HUQUQ-TARTIBOT HUJJATLARI MONITORINGI", fg="white", bg="#2f3640", font=("Arial", 14, "bold")).pack(pady=15)

        # Yuqori ko'rsatkichlar kartalari
        cards_frame = tk.Frame(self.tab_dashboard, bg="#f5f6fa", pady=15)
        cards_frame.pack(fill="x", padx=20)

        self.card_total = self.create_card(cards_frame, "JAMI HUJJATLAR", "0 ta", "#487eb0", 0)
        self.card_regions = self.create_card(cards_frame, "HUDUDLAR KESIMIDA", f"{len(config.REGIONS)} ta hudud", "#4cd137", 1)
        self.card_orgs = self.create_card(cards_frame, "IDORALAR", f"{len(config.ORGANIZATIONS + config.LAW_ENFORCEMENT)} ta organ", "#e1b12c", 2)

        # Statistika jadvallari
        tables_frame = tk.Frame(self.tab_dashboard, bg="#f5f6fa")
        tables_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Hududlar statistikasi jadvali
        reg_box = tk.LabelFrame(tables_frame, text="Hududlar bo'yicha hujjatlar soni", font=("Arial", 11, "bold"), padx=10, pady=10)
        reg_box.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self.tree_reg_stat = ttk.Treeview(reg_box, columns=("Hudud", "Soni"), show="headings", height=12)
        self.tree_reg_stat.heading("Hudud", text="Hudud / Viloyat")
        self.tree_reg_stat.heading("Soni", text="Hujjatlar soni")
        self.tree_reg_stat.pack(fill="both", expand=True)

        # Toifalar statistikasi jadvali
        type_box = tk.LabelFrame(tables_frame, text="Hujjat turlari taqsimoti", font=("Arial", 11, "bold"), padx=10, pady=10)
        type_box.pack(side="right", fill="both", expand=True)

        self.tree_type_stat = ttk.Treeview(type_box, columns=("Tur", "Soni"), show="headings", height=12)
        self.tree_type_stat.heading("Tur", text="Hujjat Toifasi")
        self.tree_type_stat.heading("Soni", text="Soni")
        self.tree_type_stat.pack(fill="both", expand=True)

    def create_card(self, parent, title, value, color, col):
        frame = tk.Frame(parent, bg=color, padx=20, pady=15, relief="raised", bd=1)
        frame.grid(row=0, column=col, padx=15, sticky="nsew")
        parent.grid_columnconfigure(col, weight=1)

        tk.Label(frame, text=title, font=("Arial", 10, "bold"), fg="white", bg=color).pack()
        lbl_val = tk.Label(frame, text=value, font=("Arial", 16, "bold"), fg="white", bg=color, pady=5)
        lbl_val.pack()
        return lbl_val

    def update_dashboard(self):
        stats = database.get_statistics()
        self.card_total.config(text=f"{stats['total']} ta")

        for item in self.tree_reg_stat.get_children():
            self.tree_reg_stat.delete(item)
        for r, cnt in stats["regions"]:
            self.tree_reg_stat.insert("", "end", values=(r.replace("_", " "), f"{cnt} ta"))

        for item in self.tree_type_stat.get_children():
            self.tree_type_stat.delete(item)
        for t, cnt in stats["types"]:
            self.tree_type_stat.insert("", "end", values=(t, f"{cnt} ta"))

    # ================= 2. HUJJAT YUKLASH TAB =================
    def init_upload_tab(self):
        frame = tk.Frame(self.tab_upload, bg="#f5f6fa", padx=30, pady=30)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Hujjatlarni Saralash va Avtomatik Arxivlash", font=("Arial", 16, "bold"), bg="#f5f6fa").pack(pady=10)
        tk.Label(frame, text="PDF fayllarni yakka, guruh yoki butun boshli papka bo'yicha tizimga kiritishingiz mumkin.", font=("Arial", 10), fg="#718093", bg="#f5f6fa").pack()

        btn_box = tk.Frame(frame, bg="#f5f6fa", pady=25)
        btn_box.pack()

        tk.Button(btn_box, text="📄 Bir nechta PDF tanlash", command=self.select_files, font=("Arial", 11, "bold"), bg="#0984e3", fg="white", padx=20, pady=12, relief="flat").pack(side="left", padx=10)
        tk.Button(btn_box, text="📁 Papkani tanlash (Barcha PDF)", command=self.select_folder, font=("Arial", 11, "bold"), bg="#00b894", fg="white", padx=20, pady=12, relief="flat").pack(side="left", padx=10)

        self.progress_bar = ttk.Progressbar(frame, orient="horizontal", length=650, mode="determinate")
        self.progress_bar.pack(pady=20)

        self.status_var = tk.StringVar(value="Tizim fayllarni qabul qilishga tayyor.")
        tk.Label(frame, textvariable=self.status_var, font=("Arial", 11), bg="#f5f6fa", fg="#2f3640").pack()

    def process_file_list(self, files):
        if not files:
            return
        total = len(files)
        success = 0
        self.progress_bar["maximum"] = total
        self.progress_bar["value"] = 0

        for idx, file_path in enumerate(files, 1):
            fname = os.path.basename(file_path)
            self.status_var.set(f"[{idx}/{total}] Tahlil qilinmoqda: {fname}")
            self.update()

            text = extract_text(file_path)
            if not text:
                continue

            region, org, doc_type = analyze_with_ai(text)

            if org in config.LAW_ENFORCEMENT:
                target_dir = os.path.join(config.BASE_DIR, "Huquqni_muhofaza_organlari", org, doc_type)
            else:
                target_dir = os.path.join(config.BASE_DIR, region, org, doc_type)

            os.makedirs(target_dir, exist_ok=True)
            dest_path = os.path.join(target_dir, fname)

            shutil.copy2(file_path, dest_path)
            database.save_document_record(fname, dest_path, region, org, doc_type, text)
            success += 1

            self.progress_bar["value"] = idx
            self.update()

        self.status_var.set(f"Saralash yakunlandi. Jami {success}/{total} ta hujjat muvaffaqiyatli arxivlandi.")
        messagebox.showinfo("Bajarildi", f"{success} ta hujjat tegishli papkalariga joylandi!")

    def select_files(self):
        f = filedialog.askopenfilenames(title="PDF tanlang", filetypes=[("PDF files", "*.pdf")])
        if f:
            self.process_file_list(list(f))

    def select_folder(self):
        folder = filedialog.askdirectory(title="Papkani tanlang")
        if not folder:
            return
        files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        if not files:
            messagebox.showwarning("Bo'sh", "Bu papkada PDF fayl yo'q!")
            return
        self.process_file_list(files)

    # ================= 3. ARXIV EXPLORER TAB (ICHKI KO'RUVCHI) =================
    def init_explorer_tab(self):
        paned = tk.PanedWindow(self.tab_explorer, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # Chap panel: Papkalar daraxti
        tree_frame = tk.Frame(paned)
        paned.add(tree_frame, minsize=300)

        tk.Label(tree_frame, text="Papkalar Iyerarxiyasi", font=("Arial", 11, "bold")).pack(anchor="w", pady=5)
        self.folder_tree = ttk.Treeview(tree_frame)
        self.folder_tree.pack(fill="both", expand=True)
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_folder_selected)

        # O'ng panel: Papka ichidagi fayllar ro'yxati
        files_frame = tk.Frame(paned)
        paned.add(files_frame, minsize=600)

        tk.Label(files_frame, text="Papka ichidagi hujjatlar (Ko'rish uchun ikki marta bosing)", font=("Arial", 11, "bold")).pack(anchor="w", pady=5)
        
        self.files_list = ttk.Treeview(files_frame, columns=("Nomi", "Hajmi", "Yo'li"), show="headings")
        self.files_list.heading("Nomi", text="Fayl Nomi")
        self.files_list.heading("Hajmi", text="Hajmi (KB)")
        self.files_list.heading("Yo'li", text="To'liq Yo'li")
        self.files_list.column("Yo'li", width=0, stretch=False)
        self.files_list.pack(fill="both", expand=True)
        self.files_list.bind("<Double-1>", self.open_pdf_viewer)

    def load_tree_directories(self):
        for item in self.folder_tree.get_children():
            self.folder_tree.delete(item)

        if not os.path.exists(config.BASE_DIR):
            os.makedirs(config.BASE_DIR, exist_ok=True)

        root_node = self.folder_tree.insert("", "end", text=config.BASE_DIR, values=(config.BASE_DIR,))
        self.populate_tree(root_node, config.BASE_DIR)

    def populate_tree(self, parent_node, path):
        try:
            for item in sorted(os.listdir(path)):
                item_path = os.path.join(path, item)
                if os.path.isdir(item_path):
                    node = self.folder_tree.insert(parent_node, "end", text=item, values=(item_path,))
                    self.populate_tree(node, item_path)
        except Exception:
            pass

    def on_folder_selected(self, event):
        selected = self.folder_tree.selection()
        if not selected:
            return
        folder_path = self.folder_tree.item(selected[0])["values"][0]

        for item in self.files_list.get_children():
            self.files_list.delete(item)

        if os.path.exists(folder_path):
            for f in os.listdir(folder_path):
                f_path = os.path.join(folder_path, f)
                if os.path.isfile(f_path) and f.lower().endswith(".pdf"):
                    size_kb = f"{os.path.getsize(f_path) // 1024} KB"
                    self.files_list.insert("", "end", values=(f, size_kb, f_path))

    def open_pdf_viewer(self, event):
        selected = self.files_list.selection()
        if not selected:
            return
        f_path = self.files_list.item(selected[0])["values"][2]
        PDFViewerWindow(self, f_path)

    # ================= 4. QIDIRUV TAB =================
    def init_search_tab(self):
        top_frame = tk.Frame(self.tab_search, bg="#f5f6fa", pady=15, padx=20)
        top_frame.pack(fill="x")

        tk.Label(top_frame, text="Kalit so'z yoki hujjat nomini yozing:", font=("Arial", 11, "bold"), bg="#f5f6fa").pack(anchor="w")
        
        entry_frame = tk.Frame(top_frame, bg="#f5f6fa", pady=5)
        entry_frame.pack(fill="x")

        self.search_entry = tk.Entry(entry_frame, font=("Arial", 12))
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        tk.Button(entry_frame, text="🔍 Bazasidan Qidirish", command=self.search_db, font=("Arial", 11, "bold"), bg="#27ae60", fg="white", padx=15).pack(side="right")

        # Natijalar jadvali
        cols = ("Fayl nomi", "Hudud", "Tashkilot", "Turi", "Saqlangan vaqti", "Manzil")
        self.search_tree = ttk.Treeview(self.tab_search, columns=cols, show="headings")
        for col in cols:
            self.search_tree.heading(col, text=col)
            self.search_tree.column(col, width=130)

        self.search_tree.column("Fayl nomi", width=220)
        self.search_tree.column("Manzil", width=0, stretch=False)
        self.search_tree.pack(fill="both", expand=True, padx=20, pady=10)
        self.search_tree.bind("<Double-1>", self.open_search_pdf)

        tk.Label(self.tab_search, text="Hujjatni o'rnatilgan viewer'da ochish uchun ustiga ikki marta bosing", font=("Arial", 9), fg="gray").pack(pady=5)

    def search_db(self):
        query = self.search_entry.get().strip()
        if not query:
            return
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)

        results = database.search_documents(query)
        for row in results:
            self.search_tree.insert("", "end", values=(row[0], row[2], row[3], row[4], row[5], row[1]))

    def open_search_pdf(self, event):
        selected = self.search_tree.selection()
        if not selected:
            return
        f_path = self.search_tree.item(selected[0])["values"][5]
        if os.path.exists(f_path):
            PDFViewerWindow(self, f_path)
        else:
            messagebox.showerror("Xato", "Ushbu hujjat joylashgan joyida topilmadi!")

    # ================= 5. ADMIN VA SOZLAMALAR TAB =================
    def init_settings_tab(self):
        frame = tk.Frame(self.tab_settings, bg="#f5f6fa", padx=40, pady=30)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Tizim Sozlamalari va Boshqaruv", font=("Arial", 14, "bold"), bg="#f5f6fa").pack(anchor="w", pady=10)

        # Ma'lumotlar papkasi yo'li
        path_box = tk.LabelFrame(frame, text="Arxiv Bosh Papkasi (Katalog)", padx=15, pady=15, bg="#f5f6fa", font=("Arial", 10, "bold"))
        path_box.pack(fill="x", pady=10)
        tk.Label(path_box, text=f"Hozirgi saqlash joyi: {os.path.abspath(config.BASE_DIR)}", font=("Arial", 10), bg="#f5f6fa").pack(anchor="w")

        # Ma'lumotlar bazasi operatsiyalari
        db_box = tk.LabelFrame(frame, text="Ma'lumotlar Bazasi Xizmati", padx=15, pady=15, bg="#f5f6fa", font=("Arial", 10, "bold"))
        db_box.pack(fill="x", pady=10)

        tk.Button(db_box, text="Bazani qayta indekslash (Update)", command=self.update_dashboard, bg="#2980b9", fg="white", font=("Arial", 10)).pack(side="left", padx=5)
        tk.Button(db_box, text="Bazani butunlay tozalash (Reset DB)", command=self.clear_database_prompt, bg="#c0392b", fg="white", font=("Arial", 10)).pack(side="left", padx=15)

    def clear_database_prompt(self):
        if messagebox.askyesno("Tasdiqlash", "Haqiqatan ham barcha indekslangan yozuvlarni bazadan o'chirmoqchimisiz?"):
            database.clear_all_data()
            self.update_dashboard()
            messagebox.showinfo("Tozalandi", "Ma'lumotlar bazasi tozalandi.")

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
