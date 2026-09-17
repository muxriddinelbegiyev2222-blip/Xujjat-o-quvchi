import os
import re
import shutil
import hashlib
import threading
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import fitz  # PyMuPDF
from PIL import Image, ImageTk
from docx import Document
import ollama

import config
import database

database.init_db()

# --- Yordamchi Funksiyalar ---
def calculate_md5(file_path):
    try:
        hasher = hashlib.md5()
        with open(file_path, "rb") as f:
            while chunk := f.read(65536):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception:
        return ""

def extract_pdf_content(pdf_path):
    """Faqat PDF ichidagi haqiqiy matnni ajratib olish (fayl nomiga bog'lanmagan)"""
    full_text = ""
    pages_text = []
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            txt = page.get_text("text") or ""
            pages_text.append(txt)
            full_text += f"\n--- SAHIFA {page_num + 1} ---\n" + txt
        doc.close()
    except Exception as e:
        print(f"PDF o'qishda xato: {e}")

    return full_text.strip(), pages_text

def export_single_pdf_to_word(pdf_path, target_docx_path):
    try:
        _, pages_text = extract_pdf_content(pdf_path)
        doc = Document()
        for idx, page_txt in enumerate(pages_text, 1):
            doc.add_heading(f"Sahifa {idx}", level=2)
            clean_txt = page_txt.strip() if page_txt.strip() else "[Ushbu sahifada kompyuter matni yo'q (skaner rasm)]"
            doc.add_paragraph(clean_txt)
            if idx < len(pages_text):
                doc.add_page_break()
        doc.save(target_docx_path)
        return True
    except Exception as e:
        print(f"Word yaratishda xato: {e}")
        return False

def extract_dates_from_text(text):
    """Matn ichidan rasmiy sanalarni qidirish (regex)"""
    # 1. 2026-09-17 yoki 2026.09.17 yoki 2026/09/17
    m1 = re.search(r"\b(202[0-9])[\.\-\/](0[1-9]|1[0-2])[\.\-\/](0[1-9]|[12][0-9]|3[01])\b", text)
    if m1:
        return m1.group(1), m1.group(2), m1.group(3)

    # 2. 17.09.2026 yoki 17-09-2026
    m2 = re.search(r"\b(0[1-9]|[12][0-9]|3[01])[\.\-\/](0[1-9]|1[0-2])[\.\-\/](202[0-9])\b", text)
    if m2:
        return m2.group(3), m2.group(2), m2.group(1)

    # 3. "17" sentyabr 2026 yil
    months_uz = {
        "yanvar": "01", "fevral": "02", "mart": "03", "aprel": "04", "may": "05", "iyun": "06",
        "iyul": "07", "avgust": "08", "sentyabr": "09", "oktyabr": "10", "noyabr": "11", "dekabr": "12"
    }
    for m_name, m_num in months_uz.items():
        m3 = re.search(rf"(\d{{1,2}})\s*[-–—\s]*{m_name}\s*[-–—\s]*(202[0-9])", text, re.IGNORECASE)
        if m3:
            day_str = f"{int(m3.group(1)):02d}"
            year_str = m3.group(2)
            return year_str, m_num, day_str

    now = datetime.now()
    return str(now.year), f"{now.month:02d}", f"{now.day:02d}"

def analyze_document_by_content(text):
    """PDF ICHIDAGI MATNNI O'QIB TAHLIL QILISH (Fayl nomi hisobga olinmaydi)"""
    if not text or len(text.strip()) < 15:
        now = datetime.now()
        return "Toshkent_shahri", "Kadastr_Agentligi", "Boshqa", str(now.year), f"{now.month:02d}", f"{now.day:02d}", "Noma'lum"

    prompt = f"""
Siz rasmiy idoraviy hujjatlar bo'yicha ekspert-arxivchisisiz.
Quyidagi hujjat matnini boshidan oxirigacha sinchiklab o'rganing. Hujjatning sarlavhasi, maqsadi, matnidagi mazmun-mohiyatiga qarab to'g'ri toifalarga ajrating.

Qoidalar:
1. HUDUD: Hujjat qaysi viloyat yoki hududga tegishli yoki qayerga yuborilgan?
Variantlar: {', '.join(config.REGIONS)}. Agar butun respublika yoki agentlik markazi bo'lsa: Toshkent_shahri.
2. TASHKILOT: Qaysi idoradan chiqqan yoki qaysi organga tegishli?
Variantlar: Kadastr_Agentligi, Davlat_Kadastrlari_Palatasi, Bosh_Prokuratura, IIV, DXX, Sudlar, Boshqa_Organlar.
3. TUR: Hujjatning rasmiy toifasi/nomi nima?
Sarlavha va matniga qarang:
- Agar buyruq bo'lsa -> Buyruqlar
- Agar taqdimnoma bo'lsa -> Taqdimnomalar
- Agar bildirishnoma yoki xabarnoma bo'lsa -> Bildirishnomalar
- Agar ma'lumotnoma (spravka) bo'lsa -> Malumotnomalar
- Agar topshiriq yoki chora-tadbirlar rejasi bo'lsa -> Topshiriqlar
- Agar rasmiy xat yoki jo'natma bo'lsa -> Xatlar
- Agar ariza yoki murojaat bo'lsa -> Arizalar
- Agar qaror bo'lsa -> Qarorlar
- Agar bayonnoma bo'lsa -> Bayonnomalar
- Variantlar: {', '.join(config.DOC_TYPES)}
4. SANA: Hujjat matnida ko'rsatilgan sanani toping (Format: YYYY-MM-DD).
5. RAQAM: Hujjatning qayd raqami (masalan: 12-son, 02-14/56).

Javobni FAQAT quyidagi formatda bering:
HUDUD: <hudud>
TASHKILOT: <tashkilot>
TUR: <tur>
SANA: <YYYY-MM-DD>
RAQAM: <raqam>

HUJJAT MATNI:
{text[:3500]}
"""
    try:
        client = ollama.Client(timeout=12)
        response = client.chat(
            model="llama3.2:1b",
            messages=[{"role": "user", "content": prompt}]
        )
        ans = response["message"]["content"].strip().split("\n")

        region = "Toshkent_shahri"
        org = "Kadastr_Agentligi"
        doc_type = "Xatlar"
        doc_num = "Noma'lum"
        parsed_date = None

        for line in ans:
            l = line.strip()
            if l.startswith("HUDUD:"):
                val = l.replace("HUDUD:", "").strip()
                for r in config.REGIONS:
                    if r.lower() in val.lower():
                        region = r
                        break
            elif l.startswith("TASHKILOT:"):
                val = l.replace("TASHKILOT:", "").strip()
                for o in config.ORGANIZATIONS + config.LAW_ENFORCEMENT:
                    if o.lower() in val.lower():
                        org = o
                        break
            elif l.startswith("TUR:"):
                val = l.replace("TUR:", "").strip()
                for t in config.DOC_TYPES:
                    if t.lower() in val.lower():
                        doc_type = t
                        break
            elif l.startswith("SANA:"):
                val = l.replace("SANA:", "").strip()
                sm = re.search(r"(\d{4})-(\d{2})-(\d{2})", val)
                if sm:
                    parsed_date = sm.groups()
            elif l.startswith("RAQAM:"):
                doc_num = l.replace("RAQAM:", "").strip() or "Noma'lum"

        if parsed_date:
            year, month, day = parsed_date
        else:
            year, month, day = extract_dates_from_text(text)

        return region, org, doc_type, year, month, day, doc_num

    except Exception as e:
        print(f"AI tahlilida xato: {e}")
        # Matn ichidan qo'lda qidirish zaxirasi
        t_low = text.lower()
        doc_type = "Xatlar"
        if "buyruq" in t_low:
            doc_type = "Buyruqlar"
        elif "taqdimnoma" in t_low:
            doc_type = "Taqdimnomalar"
        elif "bildirishnoma" in t_low or "bildirgi" in t_low:
            doc_type = "Bildirishnomalar"
        elif "ma'lumotnoma" in t_low or "malumotnoma" in t_low or "spravka" in t_low:
            doc_type = "Malumotnomalar"
        elif "topshiriq" in t_low:
            doc_type = "Topshiriqlar"
        elif "ariza" in t_low:
            doc_type = "Arizalar"
        elif "qaror" in t_low:
            doc_type = "Qarorlar"

        year, month, day = extract_dates_from_text(text)
        return "Toshkent_shahri", "Kadastr_Agentligi", doc_type, year, month, day, "Noma'lum"

# --- PDF Viewer Oynasi ---
class ModernPDFViewer(tk.Toplevel):
    def __init__(self, parent, pdf_path):
        super().__init__(parent)
        self.title(f"Hujjat: {os.path.basename(pdf_path)}")
        self.geometry("950x850")
        self.configure(bg="#0f172a")
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.current_page = 0

        nav = tk.Frame(self, bg="#1e293b", pady=8)
        nav.pack(fill="x")
        
        tk.Button(nav, text="◀ Oldingi", command=self.prev_p, bg="#334155", fg="white", relief="flat").pack(side="left", padx=10)
        self.lbl_p = tk.Label(nav, text="", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 10, "bold"))
        self.lbl_p.pack(side="left", padx=10)
        tk.Button(nav, text="Keyingi ▶", command=self.next_p, bg="#334155", fg="white", relief="flat").pack(side="left", padx=10)

        tk.Button(nav, text="📝 Word (.docx) qilib olish", command=self.export_word, bg="#0284c7", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10).pack(side="right", padx=10)
        tk.Button(nav, text="Tashqi ochish", command=lambda: os.startfile(self.pdf_path), bg="#334155", fg="white", relief="flat").pack(side="right", padx=5)

        self.canvas = tk.Canvas(self, bg="#334155", highlightthickness=0)
        sc = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sc.set)
        sc.pack(side="right", fill="y")
        self.canvas.pack(fill="both", expand=True)

        self.render()

    def render(self):
        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.15, 1.15))
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

    def export_word(self):
        default_name = os.path.splitext(os.path.basename(self.pdf_path))[0] + ".docx"
        dest = filedialog.asksaveasfilename(
            initialfile=default_name,
            defaultextension=".docx",
            filetypes=[("Word Hujjati", "*.docx")]
        )
        if dest:
            if export_single_pdf_to_word(self.pdf_path, dest):
                messagebox.showinfo("Muvaffaqiyatli", f"Hujjat Word (.docx) formatiga o'tkazildi:\n{dest}")
            else:
                messagebox.showerror("Xato", "Word formatiga o'tkazishda xatolik yuz berdi!")

# --- Asosiy Dastur ---
class MasterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Smart Kadastr & Organlar Arxiv Tizimi")
        self.geometry("1180x760")
        self.configure(bg="#0f172a")

        self.setup_styles()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.tab_dash = ttk.Frame(self.notebook)
        self.tab_upload = ttk.Frame(self.notebook)
        self.tab_explorer = ttk.Frame(self.notebook)
        self.tab_search = ttk.Frame(self.notebook)
        self.tab_admin = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_dash, text=" 📊 Dashboard ")
        self.notebook.add(self.tab_upload, text=" 📥 Hujjatlarni Saralash ")
        self.notebook.add(self.tab_explorer, text=" 📂 Arxiv Explorer ")
        self.notebook.add(self.tab_search, text=" 🔍 Qidiruv ")
        self.notebook.add(self.tab_admin, text=" ⚙️ Sozlamalar ")

        self.init_dash()
        self.init_upload()
        self.init_explorer()
        self.init_search()
        self.init_admin()

        self.refresh_dash()
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
        
        header_box = tk.Frame(top, bg="#0f172a")
        header_box.pack(fill="x")
        tk.Label(header_box, text="HUJJATLAR MONITORINGI VA STATISTIKA", font=("Segoe UI", 16, "bold"), fg="#f8fafc", bg="#0f172a").pack(side="left")
        tk.Button(header_box, text="🔄 Yangilash", command=self.refresh_dash, bg="#0284c7", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=4).pack(side="right")

        cards = tk.Frame(self.tab_dash, bg="#0f172a", padx=20)
        cards.pack(fill="x")
        self.c_total = self.build_card(cards, "JAMI ARXIVLANGAN", "0 ta", "#0284c7", 0)
        self.c_year = self.build_card(cards, "JORIY YIL", "0 ta", "#0d9488", 1)
        self.c_types = self.build_card(cards, "TOIFALAR", f"{len(config.DOC_TYPES)} toifa", "#7c3aed", 2)

        mid = tk.Frame(self.tab_dash, bg="#0f172a", padx=20, pady=15)
        mid.pack(fill="both", expand=True)

        box_reg = tk.LabelFrame(mid, text="Hududlar Taqsimoti", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 11, "bold"), padx=10, pady=10)
        box_reg.pack(side="left", fill="both", expand=True, padx=(0, 10))
        self.tree_reg = ttk.Treeview(box_reg, columns=("H", "S"), show="headings")
        self.tree_reg.heading("H", text="Viloyat / Hudud")
        self.tree_reg.heading("S", text="Soni")
        self.tree_reg.pack(fill="both", expand=True)

        box_tm = tk.LabelFrame(mid, text="Toifalar Taqsimoti", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 11, "bold"), padx=10, pady=10)
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
        try:
            st = database.get_statistics()
            self.c_total.config(text=f"{st['total']} ta")

            cy = str(datetime.now().year)
            y_cnt = 0
            for y, cnt in st.get("years", []):
                if str(y) == cy:
                    y_cnt += cnt
            self.c_year.config(text=f"{y_cnt} ta ({cy})")

            for item in self.tree_reg.get_children():
                self.tree_reg.delete(item)
            for r, c in st.get("regions", []):
                self.tree_reg.insert("", "end", values=(str(r).replace("_", " "), f"{c} ta"))

            for item in self.tree_types.get_children():
                self.tree_types.delete(item)
            for t, c in st.get("types", []):
                self.tree_types.insert("", "end", values=(str(t), f"{c} ta"))
        except Exception as e:
            print(f"Dashboard xato: {e}")

    # --- 2. Hujjat Yuklash & Mazmun Bo'yicha Saralash ---
    def init_upload(self):
        f = tk.Frame(self.tab_upload, bg="#0f172a", padx=40, pady=30)
        f.pack(fill="both", expand=True)

        tk.Label(f, text="PDF Hujjatlarni Mazmuni va Matni Bo'yicha Saralash", font=("Segoe UI", 16, "bold"), fg="#f8fafc", bg="#0f172a").pack(anchor="w")
        tk.Label(f, text="Dastur fayl nomiga qaramaydi: PDF ichidagi sarlavha, matn, sana va organni tahlil qilib saralaydi.", font=("Segoe UI", 10), fg="#64748b", bg="#0f172a").pack(anchor="w", pady=(0, 15))

        btn_box = tk.Frame(f, bg="#0f172a")
        btn_box.pack(anchor="w", pady=10)

        tk.Button(btn_box, text="📄 Bir nechta PDF Tanlash", command=self.btn_select_files, bg="#0284c7", fg="white", font=("Segoe UI", 11, "bold"), padx=18, pady=10, relief="flat").pack(side="left", padx=(0, 15))
        tk.Button(btn_box, text="📁 Butun Papkani Tanlash", command=self.btn_select_folder, bg="#0d9488", fg="white", font=("Segoe UI", 11, "bold"), padx=18, pady=10, relief="flat").pack(side="left")

        self.pbar = ttk.Progressbar(f, orient="horizontal", length=800, mode="determinate")
        self.pbar.pack(anchor="w", pady=20)

        self.txt_status = tk.StringVar(value="Tizim tayyor.")
        tk.Label(f, textvariable=self.txt_status, font=("Segoe UI", 11), fg="#38bdf8", bg="#0f172a").pack(anchor="w")

        tk.Label(f, text="Amallar jurnali (Tahlil natijalari):", font=("Segoe UI", 10, "bold"), fg="#94a3b8", bg="#0f172a").pack(anchor="w", pady=(15, 5))
        self.log_box = tk.Text(f, bg="#1e293b", fg="#e2e8f0", height=12, font=("Consolas", 9), relief="flat")
        self.log_box.pack(fill="x")

    def log_msg(self, m):
        self.after(0, self._log_msg_ui, m)

    def _log_msg_ui(self, m):
        self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {m}\n")
        self.log_box.see("end")

    def update_status_ui(self, text, val=None):
        self.after(0, self._update_status, text, val)

    def _update_status(self, text, val):
        self.txt_status.set(text)
        if val is not None:
            self.pbar["value"] = val

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
                messagebox.showwarning("Bo'sh", "Ushbu papkada PDF fayllar topilmadi!")

    def process_batch(self, files):
        tot = len(files)
        self.after(0, lambda: self.pbar.configure(maximum=tot, value=0))
        succ = 0

        for i, path in enumerate(files, 1):
            fname = os.path.basename(path)
            self.update_status_ui(f"[{i}/{tot}] PDF ichi o'qilmoqda: {fname}", i)
            self.log_msg(f"O'qilmoqda: {fname}")

            try:
                f_hash = calculate_md5(path)
                dup = database.is_duplicate(f_hash)
                if dup:
                    self.log_msg(f"MAVJUD: {fname} allaqachon arxivda bor. O'tkazildi.")
                    continue

                # 1. Faqat PDF ichidagi matnni olish
                full_txt, _ = extract_pdf_content(path)
                
                if not full_txt or len(full_txt.strip()) < 10:
                    self.log_msg(f"DIQQAT: {fname} skaner qilingan rasm bo'lib chiqdi (matn yo'q).")

                # 2. Mazmun bo'yicha tahlil
                reg, org, d_type, y, m, d, doc_num = analyze_document_by_content(full_txt)

                # 3. Yil / Oy / Hudud / Idora / Toifa papkasiga faqat PDF ni nusxalash
                target_dir = os.path.join(config.BASE_DIR, str(y), f"{str(m)}-oy", reg, org, d_type)
                os.makedirs(target_dir, exist_ok=True)

                dest_pdf = os.path.join(target_dir, fname)
                shutil.copy2(path, dest_pdf)

                database.save_document_record(fname, dest_pdf, y, m, d, reg, org, d_type, doc_num, f_hash, full_txt)
                self.log_msg(f"Aniqlangan toifa: [{d_type}] | Sana: {y}-{m}-{d} | Hudud: {reg}")
                succ += 1

            except Exception as err:
                self.log_msg(f"Xato ({fname}): {err}")

        self.update_status_ui(f"Tugallandi: {succ}/{tot} ta PDF saralandi.", tot)
        self.after(0, self.refresh_dash)
        self.after(0, lambda: messagebox.showinfo("Bajarildi", f"{succ} ta PDF hujjat mazmuni bo'yicha saralandi!"))

    # --- 3. Explorer ---
    def init_explorer(self):
        paned = tk.PanedWindow(self.tab_explorer, orient="horizontal", bg="#0f172a", sashwidth=4)
        paned.pack(fill="both", expand=True, padx=15, pady=15)

        left = tk.Frame(paned, bg="#1e293b")
        paned.add(left, minsize=350)
        tk.Label(left, text="Vaqtli va Hududiy Papkalar", font=("Segoe UI", 11, "bold"), fg="#38bdf8", bg="#1e293b", pady=8).pack(anchor="w", padx=10)

        self.exp_tree = ttk.Treeview(left)
        self.exp_tree.pack(fill="both", expand=True, padx=5, pady=5)
        self.exp_tree.bind("<<TreeviewSelect>>", self.on_exp_select)

        right = tk.Frame(paned, bg="#1e293b")
        paned.add(right, minsize=650)
        tk.Label(right, text="PDF Hujjatlar (Ko'rish uchun ustiga 2 marta bosing)", font=("Segoe UI", 11, "bold"), fg="#38bdf8", bg="#1e293b", pady=8).pack(anchor="w", padx=10)

        self.exp_files = ttk.Treeview(right, columns=("Nom", "Hajm", "Path"), show="headings")
        self.exp_files.heading("Nom", text="Hujjat Nomi")
        self.exp_files.heading("Hajm", text="Hajmi")
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
                if os.path.isfile(fp) and f.lower().endswith(".pdf"):
                    sz = f"{os.path.getsize(fp)//1024} KB"
                    self.exp_files.insert("", "end", values=(f, sz, fp))

    def open_exp_file(self, e):
        sel = self.exp_files.selection()
        if not sel:
            return
        fp = self.exp_files.item(sel[0])["values"][2]
        if fp.lower().endswith(".pdf"):
            ModernPDFViewer(self, fp)

    # --- 4. Qidiruv ---
    def init_search(self):
        top = tk.Frame(self.tab_search, bg="#0f172a", padx=25, pady=20)
        top.pack(fill="x")
        tk.Label(top, text="Qidiruv Tizimi", font=("Segoe UI", 16, "bold"), fg="#f8fafc", bg="#0f172a").pack(anchor="w")

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

    # --- 5. Sozlamalar ---
    def init_admin(self):
        f = tk.Frame(self.tab_admin, bg="#0f172a", padx=40, pady=30)
        f.pack(fill="both", expand=True)
        tk.Label(f, text="Tizim Sozlamalari", font=("Segoe UI", 16, "bold"), fg="#f8fafc", bg="#0f172a").pack(anchor="w")

        b2 = tk.LabelFrame(f, text="Bazani Tozalash", bg="#1e293b", fg="#ef4444", font=("Segoe UI", 11, "bold"), padx=15, pady=15)
        b2.pack(fill="x", pady=15)
        tk.Button(b2, text="Barcha ma'lumotlarni tozalash (Reset DB)", command=self.reset_db, bg="#dc2626", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=10, pady=5).pack(side="left")

    def reset_db(self):
        if messagebox.askyesno("Tasdiq", "Rostdan ham barcha ma'lumotlar bazasini o'chirmoqchimisiz?"):
            database.clear_all_data()
            self.refresh_dash()
            messagebox.showinfo("Bajarildi", "Baza tozalandi.")

if __name__ == "__main__":
    app = MasterApp()
    app.mainloop()
