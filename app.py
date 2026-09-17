import os
import re
import shutil
import hashlib
import threading
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import fitz  # PyMuPDF
from PIL import Image, ImageTk
from docx import Document
import ollama

import config
import database

database.init_db()

# --- Yordamchi Funksiyalar ---
def calculate_md5(file_path):
    hasher = hashlib.md5()
    with open(file_path, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()

def extract_pdf_data(pdf_path):
    full_text = ""
    pages_text = []
    try:
        # 1-bosqich: Oddiy matn sifatida o'qish (pdfplumber)
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                pages_text.append(txt)
                full_text += txt + "\n"

        # 2-bosqich: Agar matn topilmasa (skaner bo'lsa), PyMuPDF orqali chuqur qatlamni o'qish
        if not full_text.strip():
            full_text = ""
            pages_text = []
            doc = fitz.open(pdf_path)
            for page in doc:
                txt = page.get_text("text") or ""
                pages_text.append(txt)
                full_text += txt + "\n"

        # 3-bosqich: Agar mutlaqo matn bo'lmasa (toza rasm/skaner), fayl nomidagi ma'lumotlarni tahlilga berish
        if not full_text.strip():
            file_name = os.path.basename(pdf_path)
            full_text = f"Hujjat fayl nomi: {file_name}. Ushbu hujjat skaner qilingan rasm bo'lib, uning mazmuni va toifasini fayl nomidagi so'zlar va belgilarga qarab aniqlang."
            pages_text = [f"[Skanerlangan rasm shaklidagi hujjat: {file_name}]"]

    except Exception as e:
        print(f"Xato matn o'qishda: {e}")
        file_name = os.path.basename(pdf_path)
        full_text = f"Hujjat fayl nomi: {file_name}"
        pages_text = [f"Fayl nomi: {file_name}"]

    return full_text.strip(), pages_text

def export_pdf_to_word(pages_text, target_docx_path):
    try:
        doc = Document()
        for idx, page_txt in enumerate(pages_text, 1):
            doc.add_heading(f"Sahifa {idx}", level=2)
            doc.add_paragraph(page_txt.strip() if page_txt.strip() else "[Matn aniqlanmadi yoki skaner rasm]")
            if idx < len(pages_text):
                doc.add_page_break()
        doc.save(target_docx_path)
    except Exception as e:
        print(f"Word yaratishda xato: {e}")

def analyze_with_ai(text):
    prompt = f"""
Quyidagi rasmiy hujjat matnini yoki nomini tahlil qiling va qat'iy quyidagi kalitlar bo'yicha javob bering:
1. HUDUD: ({', '.join(config.REGIONS)}) orasidan eng mosi. Agar aniq bo'lmasa: Toshkent_shahri
2. TASHKILOT: (Kadastr_Agentligi, Davlat_Kadastrlari_Palatasi, Bosh_Prokuratura, IIV, DXX, Sudlar, Boshqa_Organlar)
3. TUR: (Buyruqlar, Topshiriqlar, Xatlar, Taqdimnomalar, Arizalar, Boshqa)
4. SANA: Hujjat qabul qilingan sana (Format: YYYY-MM-DD. Agar topilmasa: {datetime.now().strftime('%Y-%m-%d')})
5. RAQAM: Hujjatning qayd raqami (masalan: 12-son, 04/18-22. Topilmasa: Nomsiz)

Javobni FAQAT quyidagi formatda qaytaring:
HUDUD: <natija>
TASHKILOT: <natija>
TUR: <natija>
SANA: <YYYY-MM-DD>
RAQAM: <natija>

Hujjat matni:
{text[:3000]}
"""
    try:
        response = ollama.chat(
            model="llama3.2:1b",
            messages=[{"role": "user", "content": prompt}]
        )
        lines = response["message"]["content"].strip().split("\n")
        data = {
            "HUDUD": "Toshkent_shahri",
            "TASHKILOT": "Kadastr_Agentligi",
            "TUR": "Xatlar",
            "SANA": datetime.now().strftime("%Y-%m-%d"),
            "RAQAM": "Noma'lum"
        }
        for line in lines:
            for k in data.keys():
                if line.startswith(f"{k}:"):
                    val = line.replace(f"{k}:", "").strip()
                    if val:
                        data[k] = val

        sana_match = re.search(r"(\d{4})-(\d{2})-(\d{2})", data["SANA"])
        if sana_match:
            year, month, day = sana_match.groups()
        else:
            now = datetime.now()
            year, month, day = str(now.year), f"{now.month:02d}", f"{now.day:02d}"

        return data["HUDUD"], data["TASHKILOT"], data["TUR"], year, month, day, data["RAQAM"]
    except Exception as e:
        print(f"AI Xatolik: {e}")
        now = datetime.now()
        return "Toshkent_shahri", "Kadastr_Agentligi", "Boshqa", str(now.year), f"{now.month:02d}", f"{now.day:02d}", "Noma'lum"

# --- PDF Viewer & AI Chat Oynasi ---
class ModernPDFViewer(tk.Toplevel):
    def __init__(self, parent, pdf_path):
        super().__init__(parent)
        self.title(f"Hujjat Ko'rish va AI Chat: {os.path.basename(pdf_path)}")
        self.geometry("1150x850")
        self.configure(bg="#0f172a")
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.current_page = 0
        self.zoom = 1.15

        paned = tk.PanedWindow(self, orient="horizontal", bg="#1e293b", sashwidth=4)
        paned.pack(fill="both", expand=True)

        left_frame = tk.Frame(paned, bg="#0f172a")
        paned.add(left_frame, minsize=700)

        nav = tk.Frame(left_frame, bg="#1e293b", pady=8)
        nav.pack(fill="x")
        tk.Button(nav, text="◀ Oldingi", command=self.prev_p, bg="#334155", fg="white", relief="flat").pack(side="left", padx=10)
        self.lbl_p = tk.Label(nav, text="", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 10, "bold"))
        self.lbl_p.pack(side="left", padx=10)
        tk.Button(nav, text="Keyingi ▶", command=self.next_p, bg="#334155", fg="white", relief="flat").pack(side="left", padx=10)
        tk.Button(nav, text="Word (.docx) yuklash", command=self.manual_word_export, bg="#0284c7", fg="white", relief="flat").pack(side="right", padx=10)

        self.canvas = tk.Canvas(left_frame, bg="#334155", highlightthickness=0)
        sc = tk.Scrollbar(left_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sc.set)
        sc.pack(side="right", fill="y")
        self.canvas.pack(fill="both", expand=True)

        right_frame = tk.Frame(paned, bg="#0f172a", padx=10, pady=10)
        paned.add(right_frame, minsize=400)

        tk.Label(right_frame, text="💬 Hujjat bo'yicha AI Chat", font=("Segoe UI", 12, "bold"), fg="#38bdf8", bg="#0f172a").pack(anchor="w", pady=(0, 5))
        self.chat_history = tk.Text(right_frame, bg="#1e293b", fg="#f8fafc", font=("Consolas", 10), wrap="word", relief="flat", padx=10, pady=10)
        self.chat_history.pack(fill="both", expand=True)
        self.chat_history.insert("end", "AI: Ushbu hujjat bo'yicha savollaringizni yozishingiz mumkin.\n")
        self.chat_history.config(state="disabled")

        chat_input_frame = tk.Frame(right_frame, bg="#0f172a", pady=5)
        chat_input_frame.pack(fill="x")
        self.chat_entry = tk.Entry(chat_input_frame, bg="#1e293b", fg="white", font=("Segoe UI", 11), relief="flat", insertbackground="white")
        self.chat_entry.pack(side="left", fill="x", expand=True, padx=(0, 5), ipady=4)
        self.chat_entry.bind("<Return>", lambda e: self.send_chat())
        tk.Button(chat_input_frame, text="Yuborish", command=self.send_chat, bg="#38bdf8", fg="#0f172a", font=("Segoe UI", 10, "bold"), relief="flat").pack(side="right")

        self.render()

    def render(self):
        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.photo = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.photo)
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self.lbl_p.config(text=f"{self.current_page + 1} / {len(self.doc)}")

    def next_p(self):
        if self.current_page < len(self.doc) - 1:
            self.current_page += 1
            self.render()

    def prev_p(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render()

    def manual_word_export(self):
        dest = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if dest:
            _, pages = extract_pdf_data(self.pdf_path)
            export_pdf_to_word(pages, dest)
            messagebox.showinfo("Bajarildi", "Word fayli saqlandi!")

    def send_chat(self):
        query = self.chat_entry.get().strip()
        if not query:
            return
        self.chat_entry.delete(0, "end")
        self.append_chat(f"Siz: {query}\n")
        threading.Thread(target=self.query_ai_bg, args=(query,), daemon=True).start()

    def query_ai_bg(self, query):
        full_txt, _ = extract_pdf_data(self.pdf_path)
        prompt = f"Hujjat mazmuni:\n{full_txt[:4000]}\n\nFoydalanuvchi savoli: {query}\nJavobni o'zbek tilida lo'nda va aniq bering."
        try:
            res = ollama.chat(model="llama3.2:1b", messages=[{"role": "user", "content": prompt}])
            reply = res["message"]["content"].strip()
        except Exception as e:
            reply = f"Xatolik: {e}"
        self.append_chat(f"AI: {reply}\n\n")

    def append_chat(self, msg):
        self.chat_history.config(state="normal")
        self.chat_history.insert("end", msg)
        self.chat_history.see("end")
        self.chat_history.config(state="disabled")

# --- Asosiy Zamonaviy Dastur ---
class MasterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Smart Kadastr & Organlar Arxiv Tizimi (AI Integrated)")
        self.geometry("1240x820")
        self.configure(bg="#0f172a")

        self.setup_styles()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.tab_dash = ttk.Frame(self.notebook)
        self.tab_upload = ttk.Frame(self.notebook)
        self.tab_explorer = ttk.Frame(self.notebook)
        self.tab_search = ttk.Frame(self.notebook)
        self.tab_admin = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_dash, text=" 📊 Zamonaviy Dashboard ")
        self.notebook.add(self.tab_upload, text=" 📥 Hujjatlarni Saralash & Word ")
        self.notebook.add(self.tab_explorer, text=" 📂 Dynamic Vaqtli Arxiv ")
        self.notebook.add(self.tab_search, text=" 🔍 Qidiruv & Eksport ")
        self.notebook.add(self.tab_admin, text=" ⚙️ Tizim Boshqaruvi ")

        self.init_dash()
        self.init_upload()
        self.init_explorer()
        self.init_search()
        self.init_admin()

        self.notebook.bind("<<NotebookTabChanged>>", self.tab_switch)

    def setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TNotebook", background="#0f172a", borderwidth=0)
        s.configure("TNotebook.Tab", background="#1e293b", foreground="#94a3b8", padding=[15, 8], font=("Segoe UI", 10, "bold"))
        s.map("TNotebook.Tab", background=[("selected", "#0284c7")], foreground=[("selected", "#ffffff")])
        s.configure("Treeview", background="#1e293b", foreground="#f8fafc", fieldbackground="#1e293b", rowheight=28, font=("Segoe UI", 10))
        s.configure("Treeview.Heading", background="#334155", foreground="#38bdf8", font=("Segoe UI", 10, "bold"))

    def tab_switch(self, event):
        tab = self.notebook.select()
        if tab == str(self.tab_dash):
            self.refresh_dash()
        elif tab == str(self.tab_explorer):
            self.refresh_explorer()

    # --- 1. Dashboard ---
    def init_dash(self):
        top = tk.Frame(self.tab_dash, bg="#0f172a", pady=20, padx=25)
        top.pack(fill="x")
        tk.Label(top, text="KADASTR VA HUQUQIY HUJJATLAR MONITORINGI", font=("Segoe UI", 16, "bold"), fg="#f8fafc", bg="#0f172a").pack(anchor="w")
        tk.Label(top, text="Hujjatlarning yillar, oylar, hududlar va toifalar kesimidagi real vaqt tahlili", font=("Segoe UI", 10), fg="#64748b", bg="#0f172a").pack(anchor="w")

        cards = tk.Frame(self.tab_dash, bg="#0f172a", padx=20)
        cards.pack(fill="x")
        self.c_total = self.build_card(cards, "JAMI ARXIVLANIGAN", "0 ta", "#0284c7", 0)
        self.c_year = self.build_card(cards, "JORIY YIL BO'YICHA", "0 ta", "#0d9488", 1)
        self.c_types = self.build_card(cards, "HUJJAT TURLARI", f"{len(config.DOC_TYPES)} toifa", "#7c3aed", 2)

        mid = tk.Frame(self.tab_dash, bg="#0f172a", padx=20, pady=15)
        mid.pack(fill="both", expand=True)

        box_reg = tk.LabelFrame(mid, text="Hududiy Taqsimot", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 11, "bold"), padx=10, pady=10)
        box_reg.pack(side="left", fill="both", expand=True, padx=(0, 10))
        self.tree_reg = ttk.Treeview(box_reg, columns=("H", "S"), show="headings")
        self.tree_reg.heading("H", text="Viloyat / Hudud")
        self.tree_reg.heading("S", text="Soni")
        self.tree_reg.pack(fill="both", expand=True)

        box_tm = tk.LabelFrame(mid, text="Yillar va Toifalar", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 11, "bold"), padx=10, pady=10)
        box_tm.pack(side="right", fill="both", expand=True)
        self.tree_types = ttk.Treeview(box_tm, columns=("T", "S"), show="headings")
        self.tree_types.heading("T", text="Hujjat Toifasi")
        self.tree_types.heading("S", text="Soni")
        self.tree_types.pack(fill="both", expand=True)

    def build_card(self, p, title, val, col, i):
        f = tk.Frame(p, bg=col, padx=20, pady=15)
        f.grid(row=0, column=i, padx=10, sticky="nsew")
        p.grid_columnconfigure(i, weight=1)
        tk.Label(f, text=title, font=("Segoe UI", 10, "bold"), fg="#e2e8f0", bg=col).pack()
        lbl = tk.Label(f, text=val, font=("Segoe UI", 18, "bold"), fg="white", bg=col, pady=5)
        lbl.pack()
        return lbl

    def refresh_dash(self):
        st = database.get_statistics()
        self.c_total.config(text=f"{st['total']} ta")
        cy = str(datetime.now().year)
        y_cnt = next((cnt for y, cnt in st["years"] if str(y) == cy), 0)
        self.c_year.config(text=f"{y_cnt} ta ({cy})")

        for item in self.tree_reg.get_children():
            self.tree_reg.delete(item)
        for r, c in st["regions"]:
            self.tree_reg.insert("", "end", values=(r.replace("_", " "), f"{c} ta"))

        for item in self.tree_types.get_children():
            self.tree_types.delete(item)
        for t, c in st["types"]:
            self.tree_types.insert("", "end", values=(t, f"{c} ta"))

    # --- 2. Ommaviy Yuklash & Word Generatsiya ---
    def init_upload(self):
        f = tk.Frame(self.tab_upload, bg="#0f172a", padx=40, pady=30)
        f.pack(fill="both", expand=True)

        tk.Label(f, text="Ommaviy Qabul, Vaqtli Papkalash va Word Avtomatlashtirish", font=("Segoe UI", 16, "bold"), fg="#f8fafc", bg="#0f172a").pack(anchor="w")
        tk.Label(f, text="Hujjatlar o'qiladi, sanasi/oy/yili aniqlanadi, Word formatiga o'giriladi va vaqt zanjiri bo'yicha papkalanadi.", font=("Segoe UI", 10), fg="#64748b", bg="#0f172a").pack(anchor="w", pady=(0, 20))

        btn_box = tk.Frame(f, bg="#0f172a")
        btn_box.pack(anchor="w", pady=10)

        tk.Button(btn_box, text="📄 PDF Fayllarni Tanlash", command=self.btn_select_files, bg="#0284c7", fg="white", font=("Segoe UI", 11, "bold"), padx=18, pady=10, relief="flat").pack(side="left", padx=(0, 15))
        tk.Button(btn_box, text="📁 Butun Papkani Tanlash", command=self.btn_select_folder, bg="#0d9488", fg="white", font=("Segoe UI", 11, "bold"), padx=18, pady=10, relief="flat").pack(side="left")

        self.pbar = ttk.Progressbar(f, orient="horizontal", length=800, mode="determinate")
        self.pbar.pack(anchor="w", pady=25)

        self.txt_status = tk.StringVar(value="Tizim tayyor holatda.")
        tk.Label(f, textvariable=self.txt_status, font=("Segoe UI", 11), fg="#38bdf8", bg="#0f172a").pack(anchor="w")

        tk.Label(f, text="Amallar jurnali (Log):", font=("Segoe UI", 10, "bold"), fg="#94a3b8", bg="#0f172a").pack(anchor="w", pady=(20, 5))
        self.log_box = tk.Text(f, bg="#1e293b", fg="#e2e8f0", height=12, font=("Consolas", 9), relief="flat")
        self.log_box.pack(fill="x")

    def log_msg(self, m):
        self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {m}\n")
        self.log_box.see("end")

    def btn_select_files(self):
        f = filedialog.askopenfilenames(title="PDF tanlang", filetypes=[("PDF", "*.pdf")])
        if f:
            threading.Thread(target=self.process_batch, args=(list(f),), daemon=True).start()

    def btn_select_folder(self):
        folder = filedialog.askdirectory(title="Papkani tanlang")
        if folder:
            flist = [os.path.join(folder, x) for x in os.listdir(folder) if x.lower().endswith(".pdf")]
            if flist:
                threading.Thread(target=self.process_batch, args=(flist,), daemon=True).start()
            else:
                messagebox.showwarning("Bo'sh", "Ushbu papkada PDF fayl yo'q!")

    def process_batch(self, files):
        tot = len(files)
        self.pbar["maximum"] = tot
        self.pbar["value"] = 0
        succ = 0

        for i, path in enumerate(files, 1):
            fname = os.path.basename(path)
            self.txt_status.set(f"[{i}/{tot}] Tahlil va konvertatsiya: {fname}")
            self.log_msg(f"Fayl boshlandi: {fname}")

            # 1. Hesh & Dublikat tekshiruvi
            f_hash = calculate_md5(path)
            dup = database.is_duplicate(f_hash)
            if dup:
                self.log_msg(f"OGOHLANTIRISH: {fname} avval yuklangan ({dup[1]}). O'tkazib yuborildi.")
                self.pbar["value"] = i
                continue

            # 2. Matn olish & AI tahlil
            full_txt, pages = extract_pdf_data(path)
            reg, org, d_type, y, m, d, doc_num = analyze_with_ai(full_txt)

            # 3. Yil / Oy / Hudud / Tashkilot / Tur bo'yicha papkalash
            target_dir = os.path.join(config.BASE_DIR, str(y), f"{str(m)}-oy", reg, org, d_type)
            os.makedirs(target_dir, exist_ok=True)

            # PDF ni ko'chirish
            dest_pdf = os.path.join(target_dir, fname)
            shutil.copy2(path, dest_pdf)

            # Avtomatik Word (.docx) nusxa yaratish
            word_name = os.path.splitext(fname)[0] + "_nusxa.docx"
            dest_word = os.path.join(target_dir, word_name)
            export_pdf_to_word(pages, dest_word)

            # Bazaga yozish
            database.save_document_record(fname, dest_pdf, y, m, d, reg, org, d_type, doc_num, f_hash, full_txt)
            self.log_msg(f"Muvaffaqiyatli: {y}/{m}-oy/{reg}/{d_type} papkasiga joylandi va Word yaratildi.")
            succ += 1
            self.pbar["value"] = i

        self.txt_status.set(f"Jarayon yakunlandi: {succ}/{tot} ta hujjat muvaffaqiyatli arxivlandi.")
        messagebox.showinfo("Tayyor", f"{succ} ta hujjat arxivlandi va Word nusxalari yaratildi!")

    # --- 3. Arxiv Explorer ---
    def init_explorer(self):
        paned = tk.PanedWindow(self.tab_explorer, orient="horizontal", bg="#0f172a", sashwidth=4)
        paned.pack(fill="both", expand=True, padx=15, pady=15)

        left = tk.Frame(paned, bg="#1e293b")
        paned.add(left, minsize=350)
        tk.Label(left, text="Vaqtli va Hududiy Papkalar Zanjiri", font=("Segoe UI", 11, "bold"), fg="#38bdf8", bg="#1e293b", pady=8).pack(anchor="w", padx=10)

        self.exp_tree = ttk.Treeview(left)
        self.exp_tree.pack(fill="both", expand=True, padx=5, pady=5)
        self.exp_tree.bind("<<TreeviewSelect>>", self.on_exp_select)

        right = tk.Frame(paned, bg="#1e293b")
        paned.add(right, minsize=650)
        tk.Label(right, text="Papka tarkibi (PDF ustiga 2 marta bossangiz viewer ochiladi)", font=("Segoe UI", 11, "bold"), fg="#38bdf8", bg="#1e293b", pady=8).pack(anchor="w", padx=10)

        self.exp_files = ttk.Treeview(right, columns=("Nom", "Format", "Hajm", "Path"), show="headings")
        self.exp_files.heading("Nom", text="Hujjat Nomi")
        self.exp_files.heading("Format", text="Format")
        self.exp_files.heading("Hajm", text="Hajmi (KB)")
        self.exp_files.heading("Path", text="Path")
        self.exp_files.column("Path", width=0, stretch=False)
        self.exp_files.pack(fill="both", expand=True, padx=5, pady=5)
        self.exp_files.bind("<Double-1>", self.open_exp_file)

    def refresh_explorer(self):
        for itm in self.exp_tree.get_children():
            self.exp_tree.delete(itm)
        if not os.path.exists(config.BASE_DIR):
            os.makedirs(config.BASE_DIR, exist_ok=True)
        root_node = self.exp_tree.insert("", "end", text="ASOSIY ARXIV", values=(config.BASE_DIR,))
        self.fill_exp_tree(root_node, config.BASE_DIR)

    def fill_exp_tree(self, parent, path):
        try:
            for item in sorted(os.listdir(path)):
                p = os.path.join(path, item)
                if os.path.isdir(p):
                    n = self.exp_tree.insert(parent, "end", text=item, values=(p,))
                    self.fill_exp_tree(n, p)
        except Exception:
            pass

    def on_exp_select(self, e):
        sel = self.exp_tree.selection()
        if not sel:
            return
        p = self.exp_tree.item(sel[0])["values"][0]
        for itm in self.exp_files.get_children():
            self.exp_files.delete(itm)

        if os.path.exists(p) and os.path.isdir(p):
            for f in sorted(os.listdir(p)):
                fp = os.path.join(p, f)
                if os.path.isfile(fp):
                    fmt = "Word (.docx)" if f.endswith(".docx") else ("PDF" if f.endswith(".pdf") else "Boshqa")
                    sz = f"{os.path.getsize(fp)//1024} KB"
                    self.exp_files.insert("", "end", values=(f, fmt, sz, fp))

    def open_exp_file(self, e):
        sel = self.exp_files.selection()
        if not sel:
            return
        fp = self.exp_files.item(sel[0])["values"][3]
        if fp.lower().endswith(".pdf"):
            ModernPDFViewer(self, fp)
        elif fp.lower().endswith(".docx"):
            os.startfile(fp)

    # --- 4. Qidiruv ---
    def init_search(self):
        top = tk.Frame(self.tab_search, bg="#0f172a", padx=25, pady=20)
        top.pack(fill="x")
        tk.Label(top, text="Aqlli Indeksli Qidiruv Tizimi", font=("Segoe UI", 16, "bold"), fg="#f8fafc", bg="#0f172a").pack(anchor="w")

        sf = tk.Frame(top, bg="#0f172a", pady=10)
        sf.pack(fill="x")
        self.search_in = tk.Entry(sf, bg="#1e293b", fg="white", font=("Segoe UI", 12), relief="flat", insertbackground="white")
        self.search_in.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=6)
        tk.Button(sf, text="Qidirish", command=self.do_search, bg="#0284c7", fg="white", font=("Segoe UI", 11, "bold"), padx=20, relief="flat").pack(side="right")

        cols = ("Nom", "Hudud", "Organ", "Tur", "Sana", "Raqam", "Path")
        self.srch_tree = ttk.Treeview(self.tab_search, columns=cols, show="headings")
        for c in cols:
            self.srch_tree.heading(c, text=c)
            self.srch_tree.column(c, width=120)
        self.srch_tree.column("Nom", width=220)
        self.srch_tree.column("Path", width=0, stretch=False)
        self.srch_tree.pack(fill="both", expand=True, padx=25, pady=10)
        self.srch_tree.bind("<Double-1>", self.open_srch_pdf)

    def do_search(self):
        q = self.search_in.get().strip()
        if not q:
            return
        for itm in self.srch_tree.get_children():
            self.srch_tree.delete(itm)
        res = database.search_documents(q)
        for r in res:
            self.srch_tree.insert("", "end", values=r)

    def open_srch_pdf(self, e):
        sel = self.srch_tree.selection()
        if not sel:
            return
        fp = self.srch_tree.item(sel[0])["values"][6]
        if os.path.exists(fp):
            ModernPDFViewer(self, fp)

    # --- 5. Admin ---
    def init_admin(self):
        f = tk.Frame(self.tab_admin, bg="#0f172a", padx=40, pady=30)
        f.pack(fill="both", expand=True)
        tk.Label(f, text="Tizim Sozlamalari va Baza Xizmati", font=("Segoe UI", 16, "bold"), fg="#f8fafc", bg="#0f172a").pack(anchor="w")

        b1 = tk.LabelFrame(f, text="Arxivlash Qoidasi", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 11, "bold"), padx=15, pady=15)
        b1.pack(fill="x", pady=15)
        tk.Label(b1, text="Tuzilma: /Arxiv/<Yil>/<Oy>/<Hudud>/<Idora>/<Hujjat_Turi>/[PDF va DOCX]", fg="#e2e8f0", bg="#1e293b", font=("Segoe UI", 10)).pack(anchor="w")

        b2 = tk.LabelFrame(f, text="Bazani Tozalash", bg="#1e293b", fg="#ef4444", font=("Segoe UI", 11, "bold"), padx=15, pady=15)
        b2.pack(fill="x", pady=10)
        tk.Button(b2, text="Barcha indekslarni tozalash (Reset DB)", command=self.reset_db, bg="#dc2626", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=10, pady=5).pack(side="left")

    def reset_db(self):
        if messagebox.askyesno("Tasdiq", "Rostdan ham barcha ma'lumotlar bazasini o'chirmoqchimisiz?"):
            database.clear_all_data()
            messagebox.showinfo("Bajarildi", "Baza tozalandi.")

if __name__ == "__main__":
    app = MasterApp()
    app.mainloop()
