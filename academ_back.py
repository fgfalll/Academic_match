import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog, filedialog
import threading
import requests
import time
import random
import string
import uuid
from datetime import datetime
from collections import Counter
from scholarly import scholarly
from fake_useragent import UserAgent
import webbrowser
import re
import json
import os

try:
    from ai_advisor import launch_ai_advisor

    AI_ADVISOR_AVAILABLE = True
except ImportError:
    AI_ADVISOR_AVAILABLE = False

from crypto_utils import decrypt_with_embedded_pin_hash


# --- НАЛАШТУВАННЯ SCHOLARLY ---
def setup_scholarly():
    """Налаштовує бібліотеку для імітації реального браузера."""
    try:
        ua = UserAgent()
        current_ua = ua.random
        scholarly._current_ua = current_ua
        headers = {
            "User-Agent": current_ua,
            "Accept-Language": "uk-UA,uk;q=0.9,en-US;q=0.8,en;q=0.7",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
            "Referer": "https://scholar.google.com/",
        }
        if hasattr(scholarly, "nav"):
            scholarly.nav.session.headers.update(headers)
    except Exception as e:
        print(f"Помилка налаштування: {e}")


setup_scholarly()


# --- ДОПОМІЖНІ ФУНКЦІЇ ---


def decode_openalex_abstract(inverted_index):
    """Декодує анотацію з формату inverted index OpenAlex."""
    if not inverted_index:
        return ""
    try:
        word_index = []
        for word, locations in inverted_index.items():
            for loc in locations:
                word_index.append((loc, word))
        word_index.sort(key=lambda x: x[0])
        return " ".join(word for index, word in word_index)
    except:
        return ""


def get_author_info_openalex(orcid):
    """Отримує офіційне ім'я автора через OpenAlex Author API."""
    headers = {"User-Agent": "AcademicMatch/1.0 (mailto:mon-phd-check@example.com)"}
    url = f"https://api.openalex.org/authors/https://orcid.org/{orcid}"
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.json().get("display_name", "Невідомо")
    except:
        pass
    return "Невідомо"


def heuristic_score(
    title,
    concepts,
    author_keywords,
    manual_keywords,
    target_keywords,
    banned_keywords=[],
    abstract="",
):
    """Алгоритм оцінки релевантності за назвою, ключовими словами та анотацією з урахуванням чорного списку."""
    score = 0

    def norm(txt):
        return (
            str(txt).lower().replace("'", "'").replace("`", "'").replace("'", "'")
            if txt
            else ""
        )

    t_l = norm(title)
    ab_l = norm(abstract)
    c_l = [norm(c) for c in concepts]
    ak_l = [norm(k) for k in author_keywords]
    mk_l = norm(manual_keywords)
    matched = []

    banned_set = set([norm(b).strip() for b in banned_keywords if b.strip()])

    for kw in target_keywords:
        kw = norm(kw).strip(string.punctuation + " ")
        if not kw or kw in banned_set:
            continue
        pat = rf"(?u)(?<!\w){re.escape(kw)}(?!\w)"

        found_in_title = False
        if re.search(pat, t_l):
            score += 5
            matched.append(f"'{kw}' (Назва:+5)")
            found_in_title = True

        combined_kw = ak_l + [mk_l] if mk_l else ak_l
        found_kw = False
        for kw_src in combined_kw:
            if re.search(pat, kw_src):
                score += 4
                matched.append(f"'{kw}' (Ключове слово:+4)")
                found_kw = True
                break
        if found_kw:
            continue

        found_c = False
        for c in c_l:
            if re.search(pat, c):
                score += 3
                matched.append(f"'{kw}' (Напрям:+3)")
                found_c = True
                break
        if found_c:
            continue

        if not found_in_title and ab_l and re.search(pat, ab_l):
            score += 2
            matched.append(f"'{kw}' (Анотація:+2)")

    return score, list(set(matched))


class MonCouncilProApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Аналітика разової вченої ради (КМУ №44)")
        self.root.geometry("1200x900")
        self.all_candidates = {}
        self.all_papers = {}
        self.cutoff_year = 2022
        self.target_keywords = []
        self.current_cand_filter = None
        self.years_back_var = tk.IntVar(value=4)
        self.current_author_keywords = []
        self.current_banned_keywords = []
        self.global_banned_keywords = []
        self.ai_advisor_instance = None
        self._current_file_path = None
        self.create_widgets()
        self.update_keyword_preview()

    def create_widgets(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(
            label="Зберегти сесію", command=self.save_session, accelerator="Ctrl+S"
        )
        file_menu.add_command(
            label="Зберегти сесію як...",
            command=self.save_session_as,
            accelerator="Ctrl+Shift+S",
        )
        file_menu.add_command(
            label="Завантажити сесію", command=self.load_session, accelerator="Ctrl+O"
        )
        file_menu.add_separator()
        file_menu.add_command(label="Вихід", command=self.root.quit)
        self.root.bind("<Control-s>", lambda e: self.save_session())
        self.root.bind("<Control-o>", lambda e: self.load_session())
        self.root.bind("<Control-s>", lambda e: self.save_session())
        self.root.bind("<Control-o>", lambda e: self.load_session())

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        self.tab_main = ttk.Frame(self.notebook)
        self.tab_edit = ttk.Frame(self.notebook)
        self.tab_advice = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_main, text="1. Налаштування")
        self.notebook.add(self.tab_edit, text="2. Результати")
        self.notebook.add(self.tab_advice, text="3. Аналіз термінів")

        self.build_main_tab()
        self.build_edit_tab()
        self.build_advice_tab()

    def show_text_context_menu(self, event):
        menu = tk.Menu(event.widget, tearoff=0)
        menu.add_command(
            label="Вирізати", command=lambda: event.widget.event_generate("<<Cut>>")
        )
        menu.add_command(
            label="Копіювати", command=lambda: event.widget.event_generate("<<Copy>>")
        )
        menu.add_command(
            label="Вставити", command=lambda: event.widget.event_generate("<<Paste>>")
        )
        menu.add_separator()
        menu.add_command(
            label="Виділити все",
            command=lambda: event.widget.tag_add("sel", "1.0", "end"),
        )
        menu.tk_popup(event.x_root, event.y_root)

    def build_main_tab(self):
        sf = ttk.LabelFrame(
            self.tab_main, text="Дані здобувача та керівника", padding="5"
        )
        sf.pack(fill="x", padx=5, pady=2)
        ttk.Label(sf, text="Рік ради:").grid(row=0, column=0, sticky="w")
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Entry(sf, textvariable=self.year_var, width=10).grid(
            row=0, column=1, sticky="w", padx=2
        )
        years_frame = ttk.Frame(sf)
        years_frame.grid(row=0, column=2, sticky="w")
        ttk.Label(years_frame, text="Аналізувати останні:").pack(side="left")
        ttk.Spinbox(
            years_frame, from_=1, to=20, width=4, textvariable=self.years_back_var
        ).pack(side="left", padx=4)
        ttk.Label(years_frame, text="років").pack(side="left")

        ttk.Label(sf, text="Здобувач (ORCID / ПІБ):").grid(row=1, column=0, sticky="w")
        self.phd_id_var = tk.StringVar()
        ttk.Entry(sf, textvariable=self.phd_id_var, width=25).grid(
            row=1, column=1, sticky="w", padx=2
        )
        ttk.Button(sf, text="Отримати терміни", command=self.auto_fetch_keywords).grid(
            row=1, column=2, sticky="w", padx=2
        )

        ttk.Label(sf, text="Керівник (ORCID / GS / ПІБ):").grid(
            row=2, column=0, sticky="w"
        )
        self.super_id_var = tk.StringVar()
        ttk.Entry(sf, textvariable=self.super_id_var, width=25).grid(
            row=2, column=1, sticky="w", padx=2
        )
        self.deep_analysis_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            sf,
            text="Аналізувати анотації та співавторів (ORCID/OpenAlex)",
            variable=self.deep_analysis_var,
        ).grid(row=2, column=2, sticky="w", padx=2)
        self.deep_scholar_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            sf,
            text="Глибокий аналіз Scholar (повільніше)",
            variable=self.deep_scholar_var,
        ).grid(row=2, column=3, sticky="w")

        wa = ttk.Frame(self.tab_main)
        wa.pack(fill="both", expand=True, padx=5, pady=2)
        inf = ttk.LabelFrame(
            wa, text="Кандидати (ORCHID\Google Scholar через кому)", padding="5"
        )
        inf.pack(side="left", fill="both", expand=True, padx=(0, 2))
        self.candidates_text = tk.Text(inf, height=5)
        self.candidates_text.pack(fill="both", expand=True)
        self.candidates_text.insert("1.0", "")

        kwf = ttk.LabelFrame(wa, text="Ключові слова (через кому)", padding="5")
        kwf.pack(side="right", fill="both", expand=True, padx=(2, 0))
        self.keyword_text = tk.Text(kwf, height=3)
        self.keyword_text.pack(fill="both", expand=True)
        self.keyword_text.insert("1.0", "")
        self.keyword_text.bind("<KeyRelease>", self.update_keyword_preview)
        self.keyword_text.bind("<Button-3>", self.show_text_context_menu)
        self.candidates_text.bind("<Button-3>", self.show_text_context_menu)
        self.parsed_kw_label = ttk.Label(
            kwf, text="", foreground="#0056b3", wraplength=400
        )
        self.parsed_kw_label.pack(fill="x")

        bp = ttk.Frame(self.tab_main)
        bp.pack(fill="x", padx=5, pady=5)
        self.run_btn = ttk.Button(bp, text="Почати аналіз", command=self.start_analysis)
        self.run_btn.pack(side="left", fill="x", expand=True, ipady=3)
        ttk.Button(
            bp,
            text="Перевірка CAPTCHA",
            command=lambda: webbrowser.open(
                "https://scholar.google.com/scholar?q=test"
            ),
        ).pack(side="right", padx=2, ipady=3)

        lf = ttk.LabelFrame(self.tab_main, text="Журнал подій", padding="5")
        lf.pack(fill="both", expand=True, padx=5, pady=2)
        self.log_area = scrolledtext.ScrolledText(
            lf, wrap=tk.WORD, state="disabled", height=3, font=("Consolas", 9)
        )
        self.log_area.pack(fill="both", expand=True)

    def build_edit_tab(self):
        sumf = ttk.LabelFrame(self.tab_edit, text="Підсумок", padding="5")
        sumf.pack(fill="x", padx=5, pady=2)
        cols = ("cand_id", "name", "ids", "relevant", "conflict", "status")
        self.tree_sum = ttk.Treeview(sumf, columns=cols, show="headings", height=4)
        for c, t in zip(
            cols, ["ID", "Кандидат", "Джерела", "Статті 5р", "Конфлікт", "Статус"]
        ):
            self.tree_sum.heading(c, text=t)
        self.tree_sum.column("cand_id", width=0, stretch=tk.NO)
        self.tree_sum.pack(fill="x")
        self.tree_sum.tag_configure("pass", background="#d4edda")
        self.tree_sum.tag_configure("fail", background="#f8d7da")
        self.tree_sum.bind("<<TreeviewSelect>>", self.on_candidate_select)

        paf = ttk.LabelFrame(self.tab_edit, text="Список статей", padding="5")
        paf.pack(fill="both", expand=True, padx=5, pady=2)
        fp = ttk.Frame(paf)
        fp.pack(fill="x", pady=(0, 5))
        ttk.Button(fp, text="Всі", command=self.clear_candidate_filter).pack(
            side="left", padx=2
        )
        self.search_title_var = tk.StringVar()
        ent = ttk.Entry(fp, textvariable=self.search_title_var, width=25)
        ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda e: self.refresh_papers_table())
        self.filter_recent_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            fp,
            text="Відсікати старі",
            variable=self.filter_recent_var,
            command=self.refresh_papers_table,
        ).pack(side="left", padx=2)
        self.filter_score_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            fp,
            text="Тільки з балами",
            variable=self.filter_score_var,
            command=self.refresh_papers_table,
        ).pack(side="left", padx=2)
        ttk.Button(
            fp, text="Виключити слова", command=self.open_global_ban_keywords
        ).pack(side="left", padx=2)
        ttk.Button(fp, text="Додати статтю", command=self.open_add_manual_paper).pack(
            side="right", padx=2
        )
        ttk.Button(
            fp, text="Переіндексувати все", command=self.reindex_manual_papers
        ).pack(side="right", padx=2)

        pcols = ("uuid", "year", "recent", "score", "matches", "title", "source")
        self.tree_pap = ttk.Treeview(paf, columns=pcols, show="headings")
        for c, t in zip(
            pcols, ["UUID", "Рік", "Нова", "Бали", "Збіги", "Назва", "Джерело"]
        ):
            self.tree_pap.heading(c, text=t)

        self.tree_pap.column("uuid", width=0, minwidth=0, stretch=tk.NO)
        self.tree_pap.column(
            "year", width=50, minwidth=50, anchor="center", stretch=tk.NO
        )
        self.tree_pap.column(
            "recent", width=40, minwidth=40, anchor="center", stretch=tk.NO
        )
        self.tree_pap.column(
            "score", width=40, minwidth=40, anchor="center", stretch=tk.NO
        )
        self.tree_pap.column("matches", width=150, minwidth=100, stretch=tk.NO)
        self.tree_pap.column("title", minwidth=200, stretch=tk.YES)
        self.tree_pap.column(
            "source", width=100, minwidth=80, anchor="center", stretch=tk.NO
        )

        sb_y = ttk.Scrollbar(paf, orient="vertical", command=self.tree_pap.yview)
        sb_x = ttk.Scrollbar(paf, orient="horizontal", command=self.tree_pap.xview)
        self.tree_pap.configure(yscroll=sb_y.set, xscroll=sb_x.set)

        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.tree_pap.pack(fill="both", expand=True)
        self.tree_pap.bind("<<TreeviewSelect>>", self.on_paper_select)
        self.tree_pap.bind("<Double-Button-1>", lambda e: self.open_paper_details())

        self.pm = tk.Menu(self.root, tearoff=0)
        self.tree_pap.bind("<Button-3>", self.show_paper_context_menu)
        self.tree_pap.bind("<Control-c>", lambda e: self.copy_paper_title())

    def build_advice_tab(self):
        main_frame = ttk.Frame(self.tab_advice)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        left_panel = ttk.LabelFrame(main_frame, text="Кандидати", padding="5")
        left_panel.pack(side="left", fill="y", padx=(0, 5))

        list_frame = ttk.Frame(left_panel)
        list_frame.pack(fill="both", expand=True)

        self.advice_listbox = tk.Listbox(
            list_frame, selectmode=tk.MULTIPLE, width=35, exportselection=False
        )
        sb_l_y = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.advice_listbox.yview
        )
        sb_l_x = ttk.Scrollbar(
            list_frame, orient="horizontal", command=self.advice_listbox.xview
        )
        self.advice_listbox.config(yscrollcommand=sb_l_y.set, xscrollcommand=sb_l_x.set)

        self.advice_listbox.grid(row=0, column=0, sticky="nsew")
        sb_l_y.grid(row=0, column=1, sticky="ns")
        sb_l_x.grid(row=1, column=0, sticky="ew")

        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        mid_panel = ttk.Frame(main_frame)
        mid_panel.pack(side="left", fill="y", padx=2)

        ban_frame = ttk.LabelFrame(mid_panel, text="Виключення слів", padding="5")
        ban_frame.pack(fill="x", pady=(0, 5))

        ttk.Label(
            ban_frame,
            text="Слова, які не враховуються\nдля виділених кандидатів.",
            wraplength=200,
        ).pack(pady=2)
        ttk.Button(
            ban_frame, text="Список виключень", command=self.open_blacklist_window
        ).pack(fill="x", pady=2, ipady=3)

        ttk.Button(
            mid_panel, text="Аналізувати", command=self.generate_advice_strategy
        ).pack(fill="x", pady=5, ipady=3)
        ttk.Button(
            mid_panel, text="Зберегти звіт (.txt)", command=self.export_advice_report
        ).pack(fill="x", pady=2)

        sep = ttk.Separator(mid_panel, orient="horizontal")
        sep.pack(fill="x", pady=10)

        if AI_ADVISOR_AVAILABLE:
            self.ai_advisor_btn = ttk.Button(
                mid_panel,
                text="AI Консультант",
                command=self.launch_ai_advisor,
                state="disabled",
            )
            self.ai_advisor_btn.pack(fill="x", pady=2, ipady=5)
        else:
            ttk.Label(mid_panel, text="(AI недоступний)", foreground="gray").pack(
                pady=2
            )

        right_panel = ttk.LabelFrame(main_frame, text="Результати аналізу", padding="5")
        right_panel.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.advice_output = scrolledtext.ScrolledText(
            right_panel, wrap=tk.WORD, font=("Arial", 10)
        )
        self.advice_output.pack(fill="both", expand=True)
        self.advice_output.config(state="disabled")

    def log(self, msg):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.root.update()

    def clear_log(self):
        self.log_area.config(state="normal")
        self.log_area.delete("1.0", tk.END)
        self.log_area.config(state="disabled")
        self.root.update()

    def log_status(self, h, s):
        self.log_area.config(state="normal")
        self.log_area.delete("1.0", tk.END)
        self.log_area.insert(tk.END, f"{h}\n{'-' * 40}\n{s}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.root.update()

    def update_keyword_preview(self, e=None):
        raw = self.keyword_text.get("1.0", tk.END)
        self.target_keywords = [
            k.strip(string.punctuation + " \n").lower()
            for k in raw.split(",")
            if k.strip()
        ]
        self.parsed_kw_label.config(
            text=f"[Масив: {', '.join(self.target_keywords)}]"
            if self.target_keywords
            else ""
        )

    def _get_default_save_path(self):
        import os

        session_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "sessions"
        )
        os.makedirs(session_dir, exist_ok=True)
        year = self.year_var.get() or datetime.now().year
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        default_name = f"council_{year}_{timestamp}.acmp"
        return os.path.join(session_dir, default_name)

    def get_session_data(self, pin_for_encryption: str = None):
        data = {
            "version": 3,
            "saved_at": datetime.now().strftime("%Y-%m-%d_%H-%M-%S"),
            "tab1_inputs": {
                "council_year": self.year_var.get(),
                "years_back": self.years_back_var.get(),
                "phd_candidate_orcid": self.phd_id_var.get(),
                "supervisor_orcid": self.super_id_var.get(),
                "candidates_text": self.candidates_text.get("1.0", tk.END).strip(),
                "target_keywords_text": self.keyword_text.get("1.0", tk.END).strip(),
            },
            "analysis_state": {
                "cutoff_year": self.cutoff_year,
                "target_keywords": self.target_keywords,
                "global_banned_keywords": self.global_banned_keywords,
                "all_candidates": self.all_candidates,
                "all_papers": self.all_papers,
                "selected_cand_ids": [
                    self.advice_cid_map[i]
                    for i in self.advice_listbox.curselection()
                    if i < len(self.advice_cid_map)
                ],
            },
        }

        if self.ai_advisor_instance and hasattr(
            self.ai_advisor_instance, "get_state_for_session"
        ):
            ai_state = self.ai_advisor_instance.get_state_for_session(
                pin_for_encryption
            )
            if ai_state:
                data["ai_advisor"] = ai_state

        return data

    def load_session_data(self, data):
        if data.get("version", 1) >= 2:
            tab1 = data.get("tab1_inputs", {})
            self.year_var.set(tab1.get("council_year", datetime.now().year))
            self.years_back_var.set(tab1.get("years_back", 4))
            self.phd_id_var.set(tab1.get("phd_candidate_orcid", ""))
            self.super_id_var.set(tab1.get("supervisor_orcid", ""))
            self.candidates_text.delete("1.0", tk.END)
            self.candidates_text.insert("1.0", tab1.get("candidates_text", ""))
            self.keyword_text.delete("1.0", tk.END)
            self.keyword_text.insert("1.0", tab1.get("target_keywords_text", ""))
            self.update_keyword_preview()

        state = data.get("analysis_state", {})
        self.cutoff_year = state.get("cutoff_year", 2022)
        self.target_keywords = state.get("target_keywords", [])
        self.global_banned_keywords = state.get("global_banned_keywords", [])
        self.all_candidates = state.get("all_candidates", {})
        self.all_papers = state.get("all_papers", {})
        saved_selected_ids = set(state.get("selected_cand_ids", []))

        for i in self.tree_sum.get_children():
            self.tree_sum.delete(i)
        for i in self.tree_pap.get_children():
            self.tree_pap.delete(i)
        self.refresh_all_tables()
        if saved_selected_ids:
            new_cid_map = getattr(self, "advice_cid_map", [])
            for i, cid in enumerate(new_cid_map):
                if cid in saved_selected_ids:
                    self.advice_listbox.selection_set(i)
        self.notebook.select(1)
        self.log(f"Сесію завантажено: {data.get('saved_at', 'unknown')}")

        self._loaded_ai_state = data.get("ai_advisor")

    def restore_ai_advisor_if_loaded(self):
        if hasattr(self, "_loaded_ai_state") and self._loaded_ai_state:
            self._ai_restore_state = self._loaded_ai_state
            self._loaded_ai_state = None

            if AI_ADVISOR_AVAILABLE and hasattr(self, "ai_advisor_btn"):
                self.ai_advisor_btn.config(state="normal")

    def save_session(self):
        import os, zipfile, io

        if self._current_file_path and os.path.exists(self._current_file_path):
            path = self._current_file_path
        else:
            path = filedialog.asksaveasfilename(
                initialdir=os.path.dirname(self._get_default_save_path()),
                defaultextension=".acmp",
                filetypes=[("Academic Match Project", "*.acmp"), ("Всі файли", "*.*")],
                initialfile=os.path.basename(self._get_default_save_path()),
            )
            if not path:
                return

        pin_for_encryption = None
        if self.ai_advisor_instance and hasattr(
            self.ai_advisor_instance, "get_state_for_session"
        ):
            pin_dialog = tk.Toplevel(self.root)
            pin_dialog.title("PIN")
            pin_dialog.resizable(0, 0)
            pin_dialog.transient(self.root)
            pin_dialog.grab_set()

            pin_dialog.update_idletasks()
            x = (pin_dialog.winfo_screenwidth() // 2) - (
                pin_dialog.winfo_reqwidth() // 2
            )
            y = (pin_dialog.winfo_screenheight() // 2) - (
                pin_dialog.winfo_reqheight() // 2
            )
            pin_dialog.geometry(f"+{x}+{y}")

            frame = ttk.Frame(pin_dialog, padding="20")
            frame.pack()

            ttk.Label(
                frame,
                text="Введіть 4-значний PIN\nдля шифрування API ключів:",
                font=("Arial", 11),
            ).pack(pady=(0, 10))
            pin_var = tk.StringVar()
            pin_entry = ttk.Entry(
                frame, textvariable=pin_var, show="*", width=10, font=("Arial", 14)
            )
            pin_entry.pack(pady=(0, 10))
            pin_entry.focus()

            def on_pin_set():
                p = pin_var.get()
                if len(p) != 4 or not p.isdigit():
                    messagebox.showwarning("Помилка", "PIN має бути 4 цифри")
                    return
                pin_dialog.destroy()
                self._do_save_session_as_zip(path, p)

            ttk.Button(frame, text="Скасувати", command=pin_dialog.destroy).pack(
                side="left", padx=(0, 5)
            )
            ttk.Button(frame, text="Зберегти", command=on_pin_set).pack(side="left")
            pin_entry.bind("<Return>", lambda e: on_pin_set())
            return

        self._do_save_session_as_zip(path, None)

    def save_session_as(self):
        import os, zipfile, io

        path = filedialog.asksaveasfilename(
            initialdir=os.path.dirname(self._get_default_save_path()),
            defaultextension=".acmp",
            filetypes=[("Academic Match Project", "*.acmp"), ("Всі файли", "*.*")],
            initialfile=os.path.basename(self._get_default_save_path()),
        )
        if not path:
            return

        pin_for_encryption = None
        if self.ai_advisor_instance and hasattr(
            self.ai_advisor_instance, "get_state_for_session"
        ):
            pin_dialog = tk.Toplevel(self.root)
            pin_dialog.title("PIN")
            pin_dialog.resizable(0, 0)
            pin_dialog.transient(self.root)
            pin_dialog.grab_set()

            pin_dialog.update_idletasks()
            x = (pin_dialog.winfo_screenwidth() // 2) - (
                pin_dialog.winfo_reqwidth() // 2
            )
            y = (pin_dialog.winfo_screenheight() // 2) - (
                pin_dialog.winfo_reqheight() // 2
            )
            pin_dialog.geometry(f"+{x}+{y}")

            frame = ttk.Frame(pin_dialog, padding="20")
            frame.pack()

            ttk.Label(
                frame,
                text="Введіть 4-значний PIN\nдля шифрування API ключів:",
                font=("Arial", 11),
            ).pack(pady=(0, 10))
            pin_var = tk.StringVar()
            pin_entry = ttk.Entry(
                frame, textvariable=pin_var, show="*", width=10, font=("Arial", 14)
            )
            pin_entry.pack(pady=(0, 10))
            pin_entry.focus()

            def on_pin_set():
                p = pin_var.get()
                if len(p) != 4 or not p.isdigit():
                    messagebox.showwarning("Помилка", "PIN має бути 4 цифри")
                    return
                pin_dialog.destroy()
                self._do_save_session_as_zip(path, p)

            ttk.Button(frame, text="Скасувати", command=pin_dialog.destroy).pack(
                side="left", padx=(0, 5)
            )
            ttk.Button(frame, text="Зберегти", command=on_pin_set).pack(side="left")
            pin_entry.bind("<Return>", lambda e: on_pin_set())
            return

        self._do_save_session_as_zip(path, None)

    def _do_save_session_as_zip(self, path, pin_for_encryption):
        import json, zipfile, io

        try:
            data = self.get_session_data(pin_for_encryption)
            session_json = json.dumps(data, ensure_ascii=False, indent=2)

            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("session.json", session_json)

            self._current_file_path = path
            self.log(f"Сесію збережено: {path}")
            messagebox.showinfo("Збереження сесії", f"Сесію успішно збережено!")
        except Exception as e:
            self.log(f"Помилка збереження: {str(e)}")
            messagebox.showerror("Помилка", f"Не вдалося зберегти сесію: {str(e)}")

    def load_session(self):
        import os, json, zipfile

        path = filedialog.askopenfilename(
            initialdir=os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "sessions"
            ),
            filetypes=[
                ("Academic Match Project", "*.acmp"),
                ("JSON файли", "*.json"),
                ("Всі файли", "*.*"),
            ],
        )
        if not path:
            return
        try:
            if path.lower().endswith(".acmp"):
                with zipfile.ZipFile(path, "r") as zf:
                    with zf.open("session.json") as f:
                        data = json.load(f)
            else:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)

            ai_state = data.get("ai_advisor", {})
            api_key = ai_state.get("api_key", "")
            needs_pin = api_key.startswith("enc:") if api_key else False

            if needs_pin:
                pin_dialog = tk.Toplevel(self.root)
                pin_dialog.title("PIN")
                pin_dialog.resizable(0, 0)
                pin_dialog.transient(self.root)
                pin_dialog.grab_set()

                pin_dialog.update_idletasks()
                x = (pin_dialog.winfo_screenwidth() // 2) - (
                    pin_dialog.winfo_reqwidth() // 2
                )
                y = (pin_dialog.winfo_screenheight() // 2) - (
                    pin_dialog.winfo_reqheight() // 2
                )
                pin_dialog.geometry(f"+{x}+{y}")

                frame = ttk.Frame(pin_dialog, padding="20")
                frame.pack()

                ttk.Label(frame, text="Введіть PIN для розшифрування:").pack(
                    pady=(0, 10)
                )
                pin_var = tk.StringVar()
                pin_entry = ttk.Entry(
                    frame, textvariable=pin_var, show="*", width=10, font=("Arial", 14)
                )
                pin_entry.pack(pady=(0, 10))
                pin_entry.focus()

                def on_pin_decrypt():
                    pin_dialog.destroy()
                    self._decrypt_and_load_session(path, data, pin_var.get())

                ttk.Button(frame, text="Скасувати", command=pin_dialog.destroy).pack(
                    side="left", padx=(0, 5)
                )
                ttk.Button(frame, text="Розшифрувати", command=on_pin_decrypt).pack(
                    side="left"
                )
                pin_entry.bind("<Return>", lambda e: on_pin_decrypt())
            else:
                self.load_session_data(data)
                self.restore_ai_advisor_if_loaded()
                self._current_file_path = path
                messagebox.showinfo("Завантаження сесії", f"Сесію успішно завантажено!")
        except Exception as e:
            self.log(f"Помилка завантаження: {str(e)}")
            messagebox.showerror("Помилка", f"Не вдалося завантажити сесію: {str(e)}")

    def _decrypt_and_load_session(self, path, data, pin):
        import json

        ai_state = data.get("ai_advisor", {})
        api_key_encrypted = ai_state.get("api_key", "")

        if api_key_encrypted and api_key_encrypted.startswith("enc:"):
            pin_hash, api_key = decrypt_with_embedded_pin_hash(
                api_key_encrypted[4:], pin
            )
            if pin_hash:
                ai_state["api_key"] = api_key
            else:
                messagebox.showerror("Помилка", "Невірний PIN")
                return

        saved_keys_encrypted = ai_state.get("saved_api_keys", {})
        if saved_keys_encrypted:
            for provider_key, key_data in saved_keys_encrypted.items():
                encrypted_key = key_data.get("api_key", "")
                if encrypted_key.startswith("enc:"):
                    pk, decrypted = decrypt_with_embedded_pin_hash(
                        encrypted_key[4:], pin
                    )
                    if pk:
                        key_data["api_key"] = decrypted

        data["ai_advisor"] = ai_state

        self.load_session_data(data)
        self.restore_ai_advisor_if_loaded()
        self._current_file_path = path
        messagebox.showinfo("Завантаження сесії", f"Сесію успішно завантажено!")

    def auto_fetch_keywords(self):
        oid = self.phd_id_var.get().strip()
        if not oid:
            return
        self.log(f"Отримання термінів для {oid}...")
        threading.Thread(target=self._fetch_kw_thread, args=(oid,), daemon=True).start()

    def _fetch_kw_thread(self, oid):
        try:
            r = requests.get(
                f"https://api.openalex.org/works?filter=author.orcid:https://orcid.org/{oid}&per-page=50"
            )
            if r.status_code == 200:
                kws = [
                    t.get("display_name", "").lower()
                    for w in r.json().get("results", [])
                    for t in w.get("topics", [])
                ]
                if kws:
                    top = ", ".join([i[0] for i in Counter(kws).most_common(8)])
                    self.root.after(0, lambda: self.keyword_text.delete("1.0", tk.END))
                    self.root.after(0, lambda: self.keyword_text.insert("1.0", top))
                    self.root.after(0, self.update_keyword_preview)
        except:
            pass

    def recalculate_all_scores(self):
        self.cutoff_year = (
            int(self.year_var.get() or datetime.now().year) - self.years_back_var.get()
        )
        self.log(f"--- ОНОВЛЕННЯ БАЛІВ (Межа: {self.cutoff_year}) ---")
        for p in self.all_papers.values():
            cid = p["cand_id"]
            banned = self.all_candidates[cid].get("banned_keywords", [])
            sc, m = heuristic_score(
                p["title"],
                p.get("concepts", []),
                p.get("author_keywords", []),
                p.get("manual_keywords", ""),
                self.target_keywords,
                banned_keywords=banned,
                abstract=p.get("abstract", ""),
            )
            p.update(
                {
                    "score": sc,
                    "matched_details": ", ".join(m),
                    "recent": (p["year"] >= self.cutoff_year),
                }
            )
        self.refresh_all_tables()

    def reindex_manual_papers(self):
        self.cutoff_year = (
            int(self.year_var.get() or datetime.now().year) - self.years_back_var.get()
        )
        self.log(f"--- ПЕРЕІНДЕКСАЦІЯ (Межа: {self.cutoff_year}) ---")
        for p in self.all_papers.values():
            cid = p["cand_id"]
            banned = self.all_candidates[cid].get("banned_keywords", [])
            sc, m = heuristic_score(
                p["title"],
                p.get("concepts", []),
                p.get("author_keywords", []),
                p.get("manual_keywords", ""),
                self.target_keywords,
                banned_keywords=banned,
                abstract=p.get("abstract", ""),
            )
            p.update(
                {
                    "score": sc,
                    "matched_details": ", ".join(m),
                    "recent": (p["year"] >= self.cutoff_year),
                }
            )
        self.refresh_all_tables()

    def _parse_candidate_ids(self, line):
        """Parse ORCID and GS ID from a candidate line. Returns (orcid, gs_id)."""
        orcid = ""
        gs_id = ""
        parts = [p.strip() for p in line.split(",")]
        for p in parts:
            p_clean = p.replace("‑", "-").replace("−", "-")
            orcid_m = re.search(r"\b\d{4}-\d{4}-\d{4}-\d{3}[\dX]\b", p_clean)
            if orcid_m:
                orcid = orcid_m.group(0)
            elif "user=" in p_clean:
                gs_m = re.search(r"user=([\w-]{12})", p_clean)
                if gs_m:
                    gs_id = gs_m.group(1)
            elif re.match(r"^[\w-]{12}$", p_clean):
                gs_id = p_clean
            elif len(p_clean) > 5 and not orcid:
                gs_id = p_clean
        return orcid, gs_id

    def _find_existing_candidate(self, orcid, gs_id):
        """Find existing candidate by ORCID or GS ID. Returns cand_id or None."""
        if not orcid and not gs_id:
            return None
        for cid, c in self.all_candidates.items():
            ids_str = c.get("ids", "")
            if orcid and f"ORCID:{orcid}" in ids_str:
                return cid
            if gs_id and f"GS:{gs_id}" in ids_str:
                return cid
        return None

    def _ask_author_name(self, default_name="Невідомо"):
        """Ask user to enter author name. Returns entered name or default."""
        top = tk.Toplevel(self.root)
        top.title("Введіть ім'я автора")
        top.geometry("500x180")

        frame = ttk.Frame(top, padding="20")
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame, text="Ім'я автора не знайдено автоматично.", font="Helvetica 10 bold"
        ).pack(pady=(0, 5))
        ttk.Label(frame, text="Введіть ім'я автора вручну:").pack(
            anchor="w", pady=(0, 10)
        )

        name_var = tk.StringVar(
            value=default_name if default_name != "Невідомо" else ""
        )
        name_entry = ttk.Entry(frame, textvariable=name_var, width=60)
        name_entry.pack(pady=(0, 15), fill="x")
        name_entry.focus()

        result = {"name": default_name}

        def on_ok():
            result["name"] = name_var.get().strip() or default_name
            top.destroy()

        def on_cancel():
            result["name"] = default_name
            top.destroy()

        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        ttk.Button(btn_frame, text="Скасувати", command=on_cancel).pack(
            side="left", padx=10
        )
        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="left", padx=10)

        name_entry.bind("<Return>", lambda e: on_ok())
        name_entry.bind("<Escape>", lambda e: on_cancel())

        top.grab_set()
        self.root.wait_window(top)
        return result["name"]

    def _update_candidate_name(self, cand_id, new_name):
        """Update candidate name from manual entry dialog."""
        if cand_id in self.all_candidates:
            self.all_candidates[cand_id]["name"] = new_name
            self.refresh_all_tables()
            self.log(f"[{cand_id}] Ім'я оновлено: {new_name}")

    def _queue_name_request(self, cand_id, orcid, gs_id):
        """Queue a name entry request for processing after all fetching."""
        if not hasattr(self, "_name_request_queue"):
            self._name_request_queue = []
        self._name_request_queue.append(
            {"cand_id": cand_id, "orcid": orcid, "gs_id": gs_id}
        )

    def _process_name_queue(self):
        """Process queued name entry requests one at a time."""
        if not hasattr(self, "_name_request_queue") or not self._name_request_queue:
            return
        if hasattr(self, "_processing_name_queue") and self._processing_name_queue:
            return

        self._processing_name_queue = True
        self._process_next_name()

    def _process_next_name(self):
        """Process next name request in queue."""
        if not hasattr(self, "_name_request_queue") or not self._name_request_queue:
            self._processing_name_queue = False
            return

        request = self._name_request_queue.pop(0)
        cand_id = request["cand_id"]

        if (
            cand_id not in self.all_candidates
            or self.all_candidates[cand_id]["name"] != "Невідомо"
        ):
            self._process_next_name()
            return

        top = tk.Toplevel(self.root)
        top.title("Введіть ім'я автора")
        top.geometry("500x180")

        frame = ttk.Frame(top, padding="20")
        frame.pack(fill="both", expand=True)

        ids_info = []
        if request["orcid"]:
            ids_info.append(f"ORCID: {request['orcid']}")
        if request["gs_id"]:
            ids_info.append(f"GS: {request['gs_id']}")
        ids_text = ", ".join(ids_info) if ids_info else "Немає ID"

        ttk.Label(frame, text=f"Кандидат: {cand_id}", font="Helvetica 10 bold").pack(
            pady=(0, 5)
        )
        ttk.Label(frame, text=f"ID: {ids_text}").pack(anchor="w", pady=(0, 5))
        ttk.Label(frame, text="Введіть ім'я автора:", font="Helvetica 10").pack(
            anchor="w", pady=(5, 10)
        )

        name_var = tk.StringVar()
        name_entry = ttk.Entry(frame, textvariable=name_var, width=60)
        name_entry.pack(pady=(0, 15), fill="x")
        name_entry.focus()

        def on_next(ask_again=True):
            entered_name = name_var.get().strip()
            if entered_name:
                self._update_candidate_name(cand_id, entered_name)
            top.destroy()
            if ask_again and self._name_request_queue:
                self.root.after(100, self._process_next_name)
            else:
                self._processing_name_queue = False
                self._name_request_queue = []

        def on_skip():
            on_next(ask_again=True)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        ttk.Button(btn_frame, text="Пропустити", command=on_skip).pack(
            side="left", padx=10
        )
        ttk.Button(btn_frame, text="Далі", command=lambda: on_next()).pack(
            side="left", padx=10
        )

        name_entry.bind("<Return>", lambda e: on_next())
        name_entry.bind("<Escape>", lambda e: on_skip())

        top.grab_set()

    def _verify_author_match(self, orcid_name, gs_name, orcid, gs_id):
        """Verify ORCID and Google Scholar belong to same author. Returns (verified_name, warning)."""
        if not orcid_name or not gs_name:
            return (orcid_name or gs_name or "Невідомо", None)

        if orcid_name == "Невідомо" and gs_name == "Невідомо":
            return ("Невідомо", None)

        norm_orcid = orcid_name.lower().strip()
        norm_gs = gs_name.lower().strip()

        if norm_orcid == norm_gs:
            return (gs_name, None)

        words_orcid = set(norm_orcid.split())
        words_gs = set(norm_gs.split())
        common = words_orcid & words_gs

        if len(common) >= 2 and (
            len(words_orcid) <= 3 or len(common) >= len(words_orcid) * 0.5
        ):
            return (
                gs_name,
                f"УВАГА: Імена можуть відрізнятися (ORCID: '{orcid_name}', GS: '{gs_name}')",
            )

        return (
            gs_name,
            f"УВАГА: Імена не співпадають! ORCID: '{orcid_name}', Google Scholar: '{gs_name}'. Продовження...",
        )

    def start_analysis(self):
        lines = [
            l.strip()
            for l in self.candidates_text.get("1.0", tk.END).split("\n")
            if l.strip()
        ]
        if not lines or not self.target_keywords:
            return

        phd_id = self.phd_id_var.get().strip().lower()
        super_id = self.super_id_var.get().strip().lower()

        new_lines = []
        existing_info = {}
        has_existing = False

        for line in lines:
            orcid, gs_id = self._parse_candidate_ids(line)
            existing_cid = self._find_existing_candidate(orcid, gs_id)
            if existing_cid:
                has_existing = True
                sources_fetched = self.all_candidates[existing_cid].get(
                    "sources_fetched", []
                )
                existing_info[existing_cid] = {
                    "orcid": orcid,
                    "gs_id": gs_id,
                    "line": line,
                    "sources_fetched": sources_fetched,
                }
            else:
                new_lines.append(line)

        if has_existing and not new_lines:
            self.log("=== ВСІ ДАНІ ВЖЕ ЗАВАНТАЖЕНІ ===")
            self.log("Спробуйте інший список кандидатів або додайте нові ID.")
            self.run_btn.config(state="normal")
            return
        elif has_existing:
            self.log(
                f"=== ПРОДОВЖЕННЯ: {len(existing_info)} існуючих, {len(new_lines)} нових ==="
            )
            existing_papers_before = len(self.all_papers)
        else:
            self.log("=== НОВИЙ АНАЛІЗ ===")
            existing_info = {}
            self.all_candidates.clear()
            self.all_papers.clear()
            for i in self.tree_sum.get_children():
                self.tree_sum.delete(i)
            for i in self.tree_pap.get_children():
                self.tree_pap.delete(i)

        self.run_btn.config(state="disabled")
        self.clear_log()
        threading.Thread(
            target=self.run_algorithm,
            args=(lines, phd_id, super_id, existing_info),
            daemon=True,
        ).start()

    def run_algorithm(self, lines, phd_id, super_id, existing_info=None):
        self.cutoff_year = (
            int(self.year_var.get() or datetime.now().year) - self.years_back_var.get()
        )
        existing_info = existing_info or {}

        for idx, line in enumerate(lines):
            orcid, gs_id = self._parse_candidate_ids(line)
            existing_cid = self._find_existing_candidate(orcid, gs_id)
            is_existing = existing_cid is not None and existing_cid in existing_info

            if is_existing:
                cand_id = existing_cid
                info = existing_info[existing_cid]
                sources_fetched = info.get("sources_fetched", [])
                a_name = self.all_candidates[cand_id].get("name", "Невідомо")
                conflict = self.all_candidates[cand_id].get("conflict", "Немає")
                merged_local = {}
                doi_map = {}
                self.log(
                    f"[{cand_id}] ПРОДОВЖЕННЯ (вже має {len(self.all_candidates[cand_id]['papers_uuids'])} робіт)"
                )
            else:
                cand_id = f"cand_{idx}"
                sources_fetched = []
                parts = [p.strip() for p in line.split(",")]
                orcid = ""
                gs_id = ""

                for p in parts:
                    p_clean = p.replace("‑", "-").replace("−", "-")
                    orcid_m = re.search(r"\b\d{4}-\d{4}-\d{4}-\d{3}[\dX]\b", p_clean)
                    if orcid_m:
                        orcid = orcid_m.group(0)
                    elif "user=" in p_clean:
                        gs_m = re.search(r"user=([\w-]{12})", p_clean)
                        if gs_m:
                            gs_id = gs_m.group(1)
                    elif re.match(r"^[\w-]{12}$", p_clean):
                        gs_id = p_clean
                    elif len(p_clean) > 5 and not orcid:
                        gs_id = p_clean

                merged_local = {}
                doi_map = {}
                a_name = "Невідомо"
                conflict = "Немає"
                if super_id and (
                    super_id == orcid.lower() or super_id == gs_id.lower()
                ):
                    conflict = "Керівник"

            d_ids = []
            if orcid:
                d_ids.append(f"ORCID:{orcid}")
            if gs_id:
                d_ids.append(f"GS:{gs_id}")

            # 1. ORCID (skip if already fetched)
            if orcid and "ORCID" not in sources_fetched:
                self.log(f"[{cand_id}] Отримання робіт з ORCID...")
                try:
                    r = requests.get(
                        f"https://pub.orcid.org/v3.0/{orcid}/works",
                        headers={"Accept": "application/json"},
                        timeout=15,
                    )
                    if r.status_code == 200:
                        for g in r.json().get("group", []):
                            for s in g.get("work-summary", []):
                                t = s.get("title", {}).get("title", {}).get("value", "")
                                if not t:
                                    continue
                                y = 0
                                doi = ""
                                try:
                                    y = int(
                                        s.get("publication-date", {})
                                        .get("year", {})
                                        .get("value", "0")
                                    )
                                except:
                                    pass

                                ext_ids = s.get("external-ids") or {}
                                if isinstance(ext_ids, dict):
                                    for ext in ext_ids.get("external-id", []):
                                        if ext.get("external-id-type") == "doi":
                                            doi = (
                                                (ext.get("external-id-value") or "")
                                                .lower()
                                                .strip()
                                                .replace("https://doi.org/", "")
                                            )
                                            break

                                k = re.sub(r"\W+", "", t.lower())
                                p_data = {
                                    "title": t,
                                    "year": y,
                                    "doi": doi,
                                    "concepts": [],
                                    "author_keywords": [],
                                    "abstract": "",
                                    "source": "ORCID",
                                    "manual_keywords": "",
                                    "authors_full": [],
                                    "journal": "-",
                                    "url": s.get("url", {}).get("value", "")
                                    if s.get("url")
                                    else (f"https://doi.org/{doi}" if doi else ""),
                                }
                                merged_local[k] = p_data
                                if doi:
                                    doi_map[doi] = k
                    else:
                        self.log(f"   ! Помилка ORCID HTTP {r.status_code}")
                except Exception as e:
                    self.log(f"   ! Помилка ORCID: {str(e)}")

            # 2. OpenAlex (skip if already fetched)
            if orcid and "OpenAlex" not in sources_fetched:
                self.log(f"[{cand_id}] Доповнення через OpenAlex...")
                oa_h = {
                    "User-Agent": "AcademicMatch/1.0 (mailto:mon-phd-check@example.com)"
                }
                try:
                    an_fetch = get_author_info_openalex(orcid)
                    if an_fetch != "Невідомо":
                        a_name = an_fetch

                    r_l = requests.get(
                        f"https://api.openalex.org/works?filter=author.orcid:https://orcid.org/{orcid}&per-page=200",
                        headers=oa_h,
                        timeout=15,
                    )
                    if r_l.status_code == 200:
                        works = r_l.json().get("results", [])
                        self.log(f"   - Знайдено {len(works)} робіт")
                        for i, ws in enumerate(works):
                            w_title = ws.get("title") or ""
                            w_doi = (
                                (ws.get("doi") or "")
                                .lower()
                                .replace("https://doi.org/", "")
                            )
                            k_oa = re.sub(r"\W+", "", w_title.lower())

                            target_k = None
                            if w_doi and w_doi in doi_map:
                                target_k = doi_map[w_doi]
                            elif k_oa and k_oa in merged_local:
                                target_k = k_oa

                            if not self.deep_analysis_var.get():
                                pl = ws.get("primary_location") or {}
                                journal = (pl.get("source") or {}).get(
                                    "display_name", "-"
                                ) or "-"
                                meta = {
                                    "concepts": [
                                        c.get("display_name", "")
                                        for c in ws.get("topics", [])
                                    ],
                                    "author_keywords": [],
                                    "journal": journal,
                                }
                                if target_k:
                                    merged_local[target_k].update(meta)
                                else:
                                    merged_local[k_oa] = {
                                        "title": w_title,
                                        "year": ws.get("publication_year", 0),
                                        "doi": w_doi,
                                        "source": "OpenAlex",
                                        "manual_keywords": "",
                                        "abstract": "",
                                        "authors_full": [],
                                        **meta,
                                    }
                            else:
                                self.log_status(
                                    f"Деталі OpenAlex: {a_name}",
                                    f"Обробка {i + 1}/{len(works)}",
                                )
                                try:
                                    ab = decode_openalex_abstract(
                                        ws.get("abstract_inverted_index")
                                    )
                                    pl = ws.get("primary_location") or {}
                                    journal = (pl.get("source") or {}).get(
                                        "display_name", "-"
                                    ) or "-"
                                    meta = {
                                        "concepts": [
                                            c.get("display_name", "")
                                            for c in ws.get("topics", [])
                                        ]
                                        or [
                                            c.get("display_name", "")
                                            for c in ws.get("concepts", [])
                                        ],
                                        "author_keywords": [
                                            kw.get("display_name", "")
                                            for kw in ws.get("keywords", [])
                                        ],
                                        "abstract": ab,
                                        "journal": journal,
                                        "authors_full": [
                                            a.get("author", {}).get(
                                                "display_name", "Невідомо"
                                            )
                                            for a in ws.get("authorships", [])
                                        ],
                                        "url": ws.get("doi") or "",
                                    }
                                    if target_k:
                                        merged_local[target_k].update(meta)
                                        if (
                                            "OpenAlex"
                                            not in merged_local[target_k]["source"]
                                        ):
                                            merged_local[target_k]["source"] += " + OA"
                                    else:
                                        merged_local[k_oa] = {
                                            "title": w_title,
                                            "year": ws.get("publication_year", 0),
                                            "doi": w_doi,
                                            "source": "OpenAlex",
                                            "manual_keywords": "",
                                            **meta,
                                        }

                                    if phd_id:
                                        for auth in ws.get("authorships", []):
                                            if (
                                                phd_id
                                                in (
                                                    auth.get("author", {}).get("orcid")
                                                    or ""
                                                ).lower()
                                            ):
                                                conflict = "Співавтор"
                                except Exception as e:
                                    self.log(f"   ! Помилка OA Item: {str(e)}")
                    else:
                        self.log(f"   ! Помилка списку OA HTTP {r_l.status_code}")
                except Exception as e:
                    self.log(f"   ! Помилка запиту OA: {str(e)}")

            # 3. Scholar (skip if already fetched)
            gs_fetched_name = "Невідомо"
            if gs_id and "Scholar" not in sources_fetched:
                header = f"Scholar: {a_name if a_name != 'Невідомо' else gs_id}"
                self.log_status(header, "Отримання списку Scholar...")
                try:
                    aq = scholarly.search_author_id(gs_id)
                    ad = scholarly.fill(aq, sections=["publications"])
                    gs_fetched_name = ad.get("name", "Невідомо")
                    if a_name == "Невідомо":
                        a_name = gs_fetched_name
                    pubs = ad.get("publications", [])
                    for i, w in enumerate(pubs):
                        self.log_status(header, f"{i + 1}/{len(pubs)}")
                        if i > 0:
                            if self.deep_scholar_var.get():
                                time.sleep(random.uniform(10, 20))
                            else:
                                time.sleep(random.uniform(2, 4))
                        try:
                            bib = w.get("bib", {})
                            t = bib.get("title", "")
                            y = int(bib.get("pub_year", "0"))
                            ab = ""
                            if self.deep_scholar_var.get() and y >= (
                                self.cutoff_year - 1
                            ):
                                try:
                                    time.sleep(random.uniform(5, 10))
                                    wf = scholarly.fill(w)
                                    ab = wf.get("bib", {}).get("abstract", "")
                                except:
                                    pass
                            if phd_id and (
                                phd_id in bib.get("author", "").lower()
                                or a_name.lower() in bib.get("author", "").lower()
                            ):
                                conflict = "Співавтор"
                            k = re.sub(r"\W+", "", t.lower())
                            if k in merged_local:
                                merged_local[k]["source"] += " + GS"
                                if not merged_local[k].get("abstract"):
                                    merged_local[k]["abstract"] = ab
                            else:
                                merged_local[k] = {
                                    "title": t,
                                    "year": y,
                                    "concepts": [],
                                    "author_keywords": [],
                                    "abstract": ab,
                                    "source": "Scholar",
                                    "manual_keywords": "",
                                    "authors_full": [],
                                    "journal": "-",
                                    "url": w.get("pub_url", ""),
                                }
                        except:
                            continue
                except Exception as e:
                    self.log(f"Помилка Scholar: {str(e)}")

            # Verify ORCID and GS names match when both are present
            orcid_name_for_verify = None
            if "ORCID" not in sources_fetched and orcid:
                orcid_name_for_verify = a_name if a_name != "Невідомо" else None

            if (
                orcid_name_for_verify
                and gs_fetched_name
                and gs_fetched_name != "Невідомо"
            ):
                verified_name, warning = self._verify_author_match(
                    orcid_name_for_verify, gs_fetched_name, orcid, gs_id
                )
                if warning:
                    self.log(warning)
                if verified_name and verified_name != "Невідомо":
                    a_name = verified_name

            # Track candidates needing name entry (handled after all fetching)
            if a_name == "Невідомо" and (orcid or gs_id) and not is_existing:
                self.root.after(
                    0,
                    lambda cid=cand_id, oid=orcid, gid=gs_id: self._queue_name_request(
                        cid, oid, gid
                    ),
                )

            # Update sources_fetched for this candidate
            new_sources = []
            if orcid:
                new_sources.append("ORCID")
            if orcid:
                new_sources.append("OpenAlex")
            if gs_id:
                new_sources.append("Scholar")

            if is_existing:
                existing_sources = set(
                    self.all_candidates[cand_id].get("sources_fetched", [])
                )
                for s in new_sources:
                    existing_sources.add(s)
                self.all_candidates[cand_id]["sources_fetched"] = list(existing_sources)
                self.all_candidates[cand_id]["name"] = a_name
                self.all_candidates[cand_id]["ids"] = (
                    ", ".join(d_ids)
                    if d_ids
                    else self.all_candidates[cand_id].get("ids", "")
                )
                self.all_candidates[cand_id]["conflict"] = conflict
            else:
                self.all_candidates[cand_id] = {
                    "name": a_name,
                    "ids": ", ".join(d_ids),
                    "conflict": conflict,
                    "papers_uuids": [],
                    "banned_keywords": [],
                    "sources_fetched": new_sources,
                }

            for pid, pd_item in merged_local.items():
                u = str(uuid.uuid4())
                self.all_candidates[cand_id]["papers_uuids"].append(u)
                sc, m = heuristic_score(
                    pd_item["title"],
                    pd_item.get("concepts", []),
                    pd_item.get("author_keywords", []),
                    pd_item.get("manual_keywords", ""),
                    self.target_keywords,
                    abstract=pd_item.get("abstract", ""),
                )
                pd_item.update(
                    {
                        "score": sc,
                        "matched_details": ", ".join(m),
                        "recent": (pd_item["year"] >= self.cutoff_year),
                        "cand_id": cand_id,
                    }
                )
                self.all_papers[u] = pd_item

        self.root.after(0, self.notebook.select(1))
        self.root.after(0, self.refresh_all_tables)
        self.log("\nАналіз завершено")
        self.root.after(0, lambda: self.run_btn.config(state="normal"))
        self.root.after(500, self._process_name_queue)

    def on_candidate_select(self, e):
        sel = self.tree_sum.selection()
        if sel:
            self.current_cand_filter = self.tree_sum.item(sel[0])["values"][0]
            self.refresh_papers_table()

    def clear_candidate_filter(self):
        self.current_cand_filter = None
        self.tree_sum.selection_remove(self.tree_sum.selection())
        self.refresh_papers_table()

    def refresh_all_tables(self):
        selected = self.current_cand_filter
        [self.tree_sum.delete(i) for i in self.tree_sum.get_children()]
        item_to_sel = None
        for cid, c in self.all_candidates.items():
            rel = sum(
                1
                for u in c["papers_uuids"]
                if self.all_papers[u]["score"] > 0 and self.all_papers[u]["recent"]
            )
            passed = rel >= 3 and c["conflict"] == "Немає"
            status, tag = (
                ("Відповідає вимогам", "pass")
                if passed
                else (f"Не відповідає ({rel}/3)", "fail")
            )
            item = self.tree_sum.insert(
                "",
                tk.END,
                values=(cid, c["name"], c["ids"], rel, c["conflict"], status),
                tags=(tag,),
            )
            if cid == selected:
                item_to_sel = item
        if item_to_sel:
            self.tree_sum.selection_set(item_to_sel)
        self.refresh_papers_table()
        self.update_advice_authors_list()

    def refresh_papers_table(self):
        [self.tree_pap.delete(i) for i in self.tree_pap.get_children()]
        sq = self.search_title_var.get().strip().lower()
        f_rec = self.filter_recent_var.get()
        f_sc = self.filter_score_var.get()

        def norm_kw(txt):
            if not txt:
                return ""
            return (
                str(txt)
                .lower()
                .replace("'", "'")
                .replace("`", "'")
                .replace("'", "'")
                .strip(string.punctuation + " ")
            )

        sorted_p = sorted(
            self.all_papers.items(),
            key=lambda x: (x[1]["recent"], x[1]["score"]),
            reverse=True,
        )
        for u, p in sorted_p:
            if self.current_cand_filter and p["cand_id"] != self.current_cand_filter:
                continue
            if f_rec and not p["recent"]:
                continue
            if f_sc and p["score"] <= 0:
                continue
            paper_kws = [norm_kw(w) for w in p.get("author_keywords", [])] + [
                norm_kw(w) for w in p.get("manual_keywords", "").split(",") if w.strip()
            ]
            if any(
                norm_kw(b) in paper_kws
                for b in self.global_banned_keywords
                if b.strip()
            ):
                continue
            if sq:
                txt = (
                    p["title"]
                    + " "
                    + p["manual_keywords"]
                    + " "
                    + ",".join(p.get("concepts", []))
                ).lower()
                if sq not in txt:
                    continue
            self.tree_pap.insert(
                "",
                tk.END,
                values=(
                    u,
                    p["year"],
                    "Так" if p["recent"] else "Ні",
                    p["score"],
                    p["matched_details"],
                    p["title"],
                    p["source"],
                ),
            )

    def on_paper_select(self, e):
        sel = self.tree_pap.selection()
        if sel:
            self.selected_p_uuid = self.tree_pap.item(sel[0])["values"][0]

    def show_paper_context_menu(self, event):
        item = self.tree_pap.identify_row(event.y)
        if not item:
            return

        self.tree_pap.selection_set(item)
        self.tree_pap.focus(item)
        values = self.tree_pap.item(item, "values")
        self.selected_p_uuid = values[0]

        p = self.all_papers.get(self.selected_p_uuid)
        if not p:
            return

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Деталі", command=self.open_paper_details)
        menu.add_command(label="Копіювати назву", command=self.copy_paper_title)
        menu.add_command(
            label="Редагувати ключові слова", command=self.open_manual_tags_dialog
        )
        if p.get("source") == "Manual":
            menu.add_separator()
            menu.add_command(
                label="Видалити статтю", command=self.delete_selected_paper
            )

        menu.tk_popup(event.x_root, event.y_root)

    def open_manual_tags_dialog(self):
        if not hasattr(self, "selected_p_uuid"):
            return
        p = self.all_papers[self.selected_p_uuid]

        top = tk.Toplevel(self.root)
        top.title("Редагувати ключові слова")
        top.geometry("600x300")

        main_frame = ttk.Frame(top, padding="15")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="Стаття:", font=("Arial", 9, "bold")).pack(
            anchor="w"
        )
        ttk.Label(
            main_frame,
            text=p["title"][:80] + ("..." if len(p["title"]) > 80 else ""),
            wraplength=570,
            foreground="gray",
        ).pack(anchor="w", pady=(0, 10))

        kw_frame = ttk.LabelFrame(main_frame, text="Власні ключові слова", padding="10")
        kw_frame.pack(fill="both", expand=True, pady=(0, 10))

        ttk.Label(kw_frame, text="Ключові слова (через кому):").pack(anchor="w")
        mkw_box = tk.Text(kw_frame, height=6, width=70, wrap="word")
        mkw_box.pack(pady=5, fill="both", expand=True)
        mkw_box.insert("1.0", p["manual_keywords"])

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(0, 5))
        ttk.Button(btn_frame, text="Скасувати", command=top.destroy).pack(
            side="left", padx=10
        )
        ttk.Button(
            btn_frame,
            text="Зберегти",
            command=lambda: self.save_manual_tags(top, mkw_box, p),
        ).pack(side="left", padx=10)

    def save_manual_tags(self, top, mkw_box, p):
        mkw = mkw_box.get("1.0", tk.END).strip()
        p["manual_keywords"] = mkw
        cid = p["cand_id"]
        banned = self.all_candidates[cid].get("banned_keywords", [])
        sc, m = heuristic_score(
            p["title"],
            p.get("concepts", []),
            p.get("author_keywords", []),
            mkw,
            self.target_keywords,
            banned_keywords=banned,
            abstract=p.get("abstract", ""),
        )
        p.update({"score": sc, "matched_details": ", ".join(m)})
        self.refresh_all_tables()
        top.destroy()

    def delete_selected_paper(self):
        if not hasattr(self, "selected_p_uuid"):
            return
        u = self.selected_p_uuid
        p = self.all_papers[u]
        if p.get("source") != "Manual":
            return
        cid = p["cand_id"]
        self.all_candidates[cid]["papers_uuids"].remove(u)
        del self.all_papers[u]
        self.refresh_all_tables()

    def copy_paper_title(self, details_label=None):
        if not hasattr(self, "selected_p_uuid"):
            return
        p = self.all_papers[self.selected_p_uuid]
        self.root.clipboard_clear()
        self.root.clipboard_append(p["title"])
        if details_label:
            original = details_label.cget("fg")
            details_label.config(fg="#28a745")
            details_label.after(1500, lambda: details_label.config(fg=original))

    def open_paper_details(self):
        if not hasattr(self, "selected_p_uuid"):
            return
        p = self.all_papers[self.selected_p_uuid]
        top = tk.Toplevel(self.root)
        top.title("Деталі публікації")
        top.geometry("750x750")
        title_label = tk.Label(
            top,
            text=p["title"],
            wraplength=700,
            font=("Arial", 11, "bold"),
            justify="center",
        )
        title_label.pack(pady=10, padx=15)

        def copy_title():
            self.root.clipboard_clear()
            self.root.clipboard_append(p["title"])
            original = title_label.cget("fg")
            title_label.config(fg="#28a745")
            title_label.after(1500, lambda: title_label.config(fg=original))

        def smart_copy(e):
            try:
                selected = txt.get("sel.first", "sel.last")
                if selected:
                    txt.clipboard_clear()
                    txt.clipboard_append(selected)
                    return
            except:
                pass
            copy_title()

        title_label.bind("<Button-3>", lambda e: copy_title())

        txt = scrolledtext.ScrolledText(
            top, height=35, wrap=tk.WORD, font=("Arial", 10)
        )
        txt.pack(padx=15, fill="both", expand=True)
        c = f"АВТОР: {self.all_candidates[p['cand_id']]['name']}\nРІК: {p['year']} | БАЛИ: {p['score']}\nДЖЕРЕЛО: {p['source']} | ЖУРНАЛ: {p.get('journal', '-')}\n"
        c += f"ЗБІГИ: {p.get('matched_details', '-')}\n" + "-" * 60 + "\n"
        c += f"СПІВАВТОРИ: {', '.join(p.get('authors_full', []))}\n\n"
        c += f"КЛЮЧОВІ СЛОВА АВТОРА: {', '.join(p.get('author_keywords', []))}\n"
        c += f"КЛЮЧОВІ СЛОВА ШІ (OpenAlex): {', '.join(p.get('concepts', []))}\n\n"
        c += f"АНОТАЦІЯ:\n{p.get('abstract', 'Немає анотації.')}\n\n"
        if p["manual_keywords"]:
            c += f"ВЛАСНІ КЛЮЧОВІ СЛОВА: {p['manual_keywords']}\n"
        txt.insert("1.0", c)

        def show_txt_menu(e):
            m = tk.Menu(txt, tearoff=0)
            m.add_command(
                label="Копіювати", command=lambda: txt.event_generate("<<Copy>>")
            )
            m.add_command(
                label="Вирізати", command=lambda: txt.event_generate("<<Cut>>")
            )
            m.add_command(
                label="Вставити", command=lambda: txt.event_generate("<<Paste>>")
            )
            m.add_separator()
            m.add_command(
                label="Виділити все", command=lambda: txt.tag_add("sel", "1.0", "end")
            )
            m.tk_popup(e.x_root, e.y_root)

        txt.bind("<Button-3>", show_txt_menu)
        txt.bind("<Control-c>", smart_copy)
        top.bind("<Control-c>", smart_copy)

        ttk.Button(top, text="Копіювати назву", command=copy_title).pack(pady=5)
        ttk.Button(
            top,
            text="Відкрити в браузері",
            command=lambda: webbrowser.open(p["url"]) if p["url"] else None,
        ).pack(pady=5)

    def open_add_manual_paper(self):
        if not self.all_candidates:
            return

        top = tk.Toplevel(self.root)
        top.title("Додати статтю вручну")
        top.geometry("650x600")

        main_frame = ttk.Frame(top, padding="15")
        main_frame.pack(fill="both", expand=True)

        author_frame = ttk.LabelFrame(main_frame, text="Автор", padding="10")
        author_frame.pack(fill="x", pady=(0, 10))

        cand_var = tk.StringVar()
        cids_list = list(self.all_candidates.keys())
        if self.current_cand_filter and self.current_cand_filter in cids_list:
            default_idx = cids_list.index(self.current_cand_filter)
        else:
            default_idx = 0
        cb = ttk.Combobox(
            author_frame, textvariable=cand_var, state="readonly", width=60
        )
        cb["values"] = [self.all_candidates[cid]["name"] for cid in cids_list]
        cb.pack()
        cb.current(default_idx)

        info_frame = ttk.LabelFrame(main_frame, text="Основна інформація", padding="10")
        info_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(info_frame, text="Назва статті:").pack(anchor="w")
        title_box = tk.Text(info_frame, height=3, width=70)
        title_box.pack(pady=5)

        row_frame = ttk.Frame(info_frame)
        row_frame.pack(fill="x", pady=5)
        ttk.Label(row_frame, text="Рік:").pack(side="left")
        y_ent = ttk.Entry(row_frame, width=8)
        y_ent.pack(side="left", padx=(0, 20))
        ttk.Label(row_frame, text="Журнал:").pack(side="left")
        j_ent = ttk.Entry(row_frame, width=35)
        j_ent.pack(side="left", fill="x", expand=True)

        url_row_frame = ttk.Frame(info_frame)
        url_row_frame.pack(fill="x", pady=5)
        ttk.Label(url_row_frame, text="DOI/URL:").pack(side="left")
        url_ent = ttk.Entry(url_row_frame, width=50)
        url_ent.pack(side="left", fill="x", expand=True)

        kw_frame = ttk.LabelFrame(main_frame, text="Власні ключові слова", padding="10")
        kw_frame.pack(fill="both", expand=True, pady=(0, 10))

        ttk.Label(kw_frame, text="Ключові слова (через кому):").pack(anchor="w")
        mkw_box = tk.Text(kw_frame, height=4, width=70, wrap="word")
        mkw_box.pack(pady=5, fill="both", expand=True)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Скасувати", command=top.destroy).pack(
            side="left", padx=10
        )
        save_btn = ttk.Button(btn_frame, text="Зберегти")
        save_btn.pack(side="left", padx=10)

        def validate_and_update():
            title = title_box.get("1.0", tk.END).strip()
            year_text = y_ent.get().strip()
            year_valid = year_text.isdigit() and len(year_text) == 4
            save_btn.config(state="normal" if (title and year_valid) else "disabled")

        title_box.bind("<KeyRelease>", lambda e: validate_and_update())
        y_ent.bind("<KeyRelease>", lambda e: validate_and_update())

        def save():
            cid = cids_list[cb.current()]
            t = title_box.get("1.0", tk.END).strip()
            y = int(y_ent.get().strip())
            j = j_ent.get().strip()
            url = url_ent.get().strip()
            mkw = mkw_box.get("1.0", tk.END).strip()

            banned = self.all_candidates[cid].get("banned_keywords", [])
            p_d = {
                "title": t,
                "year": y,
                "journal": j or "-",
                "url": url,
                "concepts": [],
                "author_keywords": [],
                "manual_keywords": mkw,
                "abstract": "",
                "source": "Manual",
                "cand_id": cid,
            }
            sc, m = heuristic_score(
                t,
                [],
                [],
                mkw,
                self.target_keywords,
                banned_keywords=banned,
                abstract="",
            )
            p_d.update(
                {
                    "score": sc,
                    "matched_details": ", ".join(m),
                    "recent": (y >= self.cutoff_year),
                }
            )
            u = str(uuid.uuid4())
            self.all_papers[u] = p_d
            self.all_candidates[cid]["papers_uuids"].append(u)
            self.refresh_all_tables()
            top.destroy()

        save_btn.config(command=save)
        validate_and_update()

    # --- ВКЛАДКА ПОРАД (ADVICE TAB) ---

    def update_advice_authors_list(self):
        self.advice_listbox.delete(0, tk.END)
        self.advice_cid_map = []
        for cid, c in self.all_candidates.items():
            self.advice_listbox.insert(tk.END, c["name"])
            self.advice_cid_map.append(cid)

    def open_global_ban_keywords(self):
        top = tk.Toplevel(self.root)
        top.title("Виключення слів (глобально)")
        top.geometry("600x450")

        def norm_kw(txt):
            if not txt:
                return ""
            return (
                str(txt)
                .lower()
                .replace("'", "'")
                .replace("`", "'")
                .replace("'", "'")
                .strip(string.punctuation + " ")
            )

        kw_by_source = {"Author": {}, "OpenAlex": {}, "Manual": {}}
        kw_data = {}

        for u, p in self.all_papers.items():
            paper_title = p.get("title", "")[:60]

            for kw in p.get("author_keywords", []):
                n = norm_kw(kw)
                if n not in kw_data:
                    kw_data[n] = {
                        "word": kw,
                        "sources": {"Author": [], "OpenAlex": [], "Manual": []},
                    }
                kw_data[n]["sources"]["Author"].append(paper_title)
                if n not in kw_by_source["Author"]:
                    kw_by_source["Author"][n] = {"word": kw, "papers": []}
                kw_by_source["Author"][n]["papers"].append(paper_title)

            for kw in p.get("manual_keywords", "").split(","):
                kw = kw.strip()
                if kw:
                    n = norm_kw(kw)
                    if n not in kw_data:
                        kw_data[n] = {
                            "word": kw,
                            "sources": {"Author": [], "OpenAlex": [], "Manual": []},
                        }
                    kw_data[n]["sources"]["Manual"].append(paper_title)
                    if n not in kw_by_source["Manual"]:
                        kw_by_source["Manual"][n] = {"word": kw, "papers": []}
                    kw_by_source["Manual"][n]["papers"].append(paper_title)

            for kw in p.get("concepts", []):
                n = norm_kw(kw)
                if n not in kw_data:
                    kw_data[n] = {
                        "word": kw,
                        "sources": {"Author": [], "OpenAlex": [], "Manual": []},
                    }
                kw_data[n]["sources"]["OpenAlex"].append(paper_title)
                if n not in kw_by_source["OpenAlex"]:
                    kw_by_source["OpenAlex"][n] = {"word": kw, "papers": []}
                kw_by_source["OpenAlex"][n]["papers"].append(paper_title)

        for kw in self.global_banned_keywords:
            n = norm_kw(kw)
            if n not in kw_data:
                kw_data[n] = {
                    "word": kw,
                    "sources": {"Author": [], "OpenAlex": [], "Manual": []},
                }

        source_var = tk.StringVar(value="Всі")
        search_var = tk.StringVar()
        banned_norm = set(norm_kw(k) for k in self.global_banned_keywords)
        source_map = {
            "Всі": None,
            "Ключові слова": "Author",
            "OpenAlex концепти": "OpenAlex",
            "Manual": "Manual",
        }

        avail_listbox = None
        ban_listbox = None

        def move_items(src, dst, to_ban):
            nonlocal avail_listbox, ban_listbox
            sel = src.curselection()
            if not sel:
                return
            items = [src.get(i) for i in sel]
            for i in reversed(sel):
                src.delete(i)
            for item in items:
                dst.insert(tk.END, item)
                if to_ban:
                    if item not in self.global_banned_keywords:
                        self.global_banned_keywords.append(item)
                else:
                    if item in self.global_banned_keywords:
                        self.global_banned_keywords.remove(item)

        def populate_lists():
            nonlocal avail_listbox, ban_listbox
            avail_listbox.delete(0, tk.END)
            ban_listbox.delete(0, tk.END)
            search_text = search_var.get().lower()
            sel_source = source_map.get(source_var.get())

            if sel_source:
                src_kws = kw_by_source.get(sel_source, {})
                for n, data in src_kws.items():
                    if search_text and search_text not in data["word"].lower():
                        continue
                    if n in banned_norm:
                        ban_listbox.insert(tk.END, data["word"])
                    else:
                        avail_listbox.insert(tk.END, data["word"])
            else:
                for n, data in kw_data.items():
                    if search_text and search_text not in data["word"].lower():
                        continue
                    if n in banned_norm:
                        ban_listbox.insert(tk.END, data["word"])
                    else:
                        avail_listbox.insert(tk.END, data["word"])

        def on_source_change(*args):
            populate_lists()

        def on_search_change(*args):
            populate_lists()

        def on_source_change(*args):
            populate_lists()

        source_var.trace("w", on_source_change)
        search_var.trace("w", on_search_change)

        main_frame = ttk.Frame(top, padding="10")
        main_frame.pack(fill="both", expand=True)

        filter_frame = ttk.Frame(main_frame)
        filter_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(filter_frame, text="Джерело:").pack(side="left")
        source_combo = ttk.Combobox(
            filter_frame,
            textvariable=source_var,
            values=["Всі", "Ключові слова", "OpenAlex концепти", "Manual"],
            state="readonly",
            width=20,
        )
        source_combo.pack(side="left", padx=5)

        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(search_frame, text="Пошук:").pack(side="left")
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=40)
        search_entry.pack(side="left", padx=5)

        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True, pady=5)

        avail_frame = ttk.LabelFrame(content_frame, text="Доступні", padding="5")
        avail_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        ban_frame = ttk.LabelFrame(content_frame, text="Виключені", padding="5")
        ban_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))

        avail_listbox = tk.Listbox(avail_frame, selectmode=tk.EXTENDED)
        avail_listbox.pack(side="left", fill="both", expand=True)
        avail_sb = ttk.Scrollbar(
            avail_frame, orient="vertical", command=avail_listbox.yview
        )
        avail_listbox.config(yscrollcommand=avail_sb.set)
        avail_sb.pack(side="right", fill="y")

        ban_listbox = tk.Listbox(ban_frame, selectmode=tk.EXTENDED)
        ban_listbox.pack(side="left", fill="both", expand=True)
        ban_sb = ttk.Scrollbar(ban_frame, orient="vertical", command=ban_listbox.yview)
        ban_listbox.config(yscrollcommand=ban_sb.set)
        ban_sb.pack(side="right", fill="y")

        btn_mid_frame = ttk.Frame(content_frame)
        btn_mid_frame.pack(side="left", fill="y", padx=5)
        ttk.Button(
            btn_mid_frame,
            text="->",
            width=5,
            command=lambda: move_items(avail_listbox, ban_listbox, True),
        ).pack(pady=5)
        ttk.Button(
            btn_mid_frame,
            text="<-",
            width=5,
            command=lambda: move_items(ban_listbox, avail_listbox, False),
        ).pack(pady=5)

        context_menu = tk.Menu(top, tearoff=0)

        def show_context_menu(event):
            widget = event.widget
            if widget not in [avail_listbox, ban_listbox]:
                return
            idx = widget.nearest(event.y)
            if idx < 0:
                return
            word = widget.get(idx)
            word_norm = norm_kw(word)
            context_menu.delete(0, tk.END)
            if word_norm in kw_data:
                kw_info = kw_data[word_norm]
                context_menu.add_command(label=f"[{word}]", state="disabled")
                context_menu.add_separator()
                for source in ["Author", "OpenAlex", "Manual"]:
                    papers = kw_info["sources"][source]
                    if papers:
                        unique_papers = list(dict.fromkeys(papers))
                        context_menu.add_command(
                            label=f"  {source}: {len(unique_papers)}", state="disabled"
                        )
                        for pt in unique_papers[:3]:
                            context_menu.add_command(
                                label=f"    - {pt}...", state="disabled"
                            )
                        if len(unique_papers) > 3:
                            context_menu.add_command(
                                label=f"    +{len(unique_papers) - 3} more",
                                state="disabled",
                            )
            context_menu.post(event.x_root, event.y_root)

        avail_listbox.bind("<Button-3>", show_context_menu)
        ban_listbox.bind("<Button-3>", show_context_menu)

        populate_lists()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(
            btn_frame,
            text="Очистити все",
            command=lambda: (
                self.global_banned_keywords.clear(),
                ban_listbox.delete(0, tk.END),
                populate_lists(),
            ),
        ).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Закрити", command=top.destroy).pack(
            side="right", padx=5
        )
        ttk.Button(
            btn_frame,
            text="Зберегти та закрити",
            command=lambda: (self.refresh_papers_table(), top.destroy()),
        ).pack(side="right", padx=5)

    def open_blacklist_window(self):
        sel = self.advice_listbox.curselection()
        if not sel:
            messagebox.showwarning("Увага", "Оберіть хоча б одного кандидата.")
            return

        cids = [self.advice_cid_map[i] for i in sel]

        all_kws = set()
        banned_all = set()

        def norm_kw(txt):
            if not txt:
                return ""
            return (
                str(txt)
                .lower()
                .replace("'", "'")
                .replace("`", "'")
                .replace("'", "'")
                .strip(string.punctuation + " ")
            )

        for kw in self.global_banned_keywords:
            banned_all.add(norm_kw(kw))
        for cid in cids:
            banned_all.update(self.all_candidates[cid].get("banned_keywords", []))
            for pid in self.all_candidates[cid]["papers_uuids"]:
                p = self.all_papers[pid]
                if p["recent"]:
                    manual_kws = [
                        norm_kw(w)
                        for w in p.get("manual_keywords", "").split(",")
                        if w.strip()
                    ]
                    words = [
                        norm_kw(w) for w in p.get("author_keywords", [])
                    ] + manual_kws
                    all_kws.update([w for w in words if w])

        self.current_author_keywords = sorted(list(all_kws))
        self.current_banned_keywords = sorted(list(banned_all))
        self.blacklist_cids = cids

        top = tk.Toplevel(self.root)
        top.title("Виключення слів")
        top.geometry("600x500")

        search_frame = ttk.Frame(top)
        search_frame.pack(fill="x", padx=10, pady=10)
        ttk.Label(search_frame, text="Пошук:").pack(side="left")
        self.kw_search_var = tk.StringVar()
        self.kw_search_var.trace_add("write", lambda *args: self.refresh_kw_lists())
        ttk.Entry(search_frame, textvariable=self.kw_search_var).pack(
            side="left", fill="x", expand=True, padx=5
        )

        lists_frame = ttk.Frame(top)
        lists_frame.pack(fill="both", expand=True, padx=10, pady=5)

        avail_frame = ttk.Frame(lists_frame)
        avail_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(avail_frame, text="Доступні:").pack(anchor="w")
        self.avail_kw_listbox = tk.Listbox(
            avail_frame, height=15, width=25, selectmode=tk.EXTENDED
        )
        self.avail_kw_listbox.pack(side="left", fill="both", expand=True)
        sb_a = ttk.Scrollbar(
            avail_frame, orient="vertical", command=self.avail_kw_listbox.yview
        )
        self.avail_kw_listbox.config(yscrollcommand=sb_a.set)
        sb_a.pack(side="right", fill="y")
        self.avail_kw_listbox.bind("<Double-1>", self.ban_selected_kw)

        btn_frame = ttk.Frame(lists_frame)
        btn_frame.pack(side="left", fill="y", padx=10)
        ttk.Button(
            btn_frame, text="Виключити", command=self.ban_selected_kw, width=11
        ).pack(pady=(30, 5))
        ttk.Button(
            btn_frame, text="Повернути", command=self.unban_selected_kw, width=11
        ).pack(pady=5)

        banned_frame = ttk.Frame(lists_frame)
        banned_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(banned_frame, text="Виключені:").pack(anchor="w")
        self.banned_kw_listbox = tk.Listbox(
            banned_frame, height=15, width=25, selectmode=tk.EXTENDED
        )
        self.banned_kw_listbox.pack(side="left", fill="both", expand=True)
        sb_b = ttk.Scrollbar(
            banned_frame, orient="vertical", command=self.banned_kw_listbox.yview
        )
        self.banned_kw_listbox.config(yscrollcommand=sb_b.set)
        sb_b.pack(side="right", fill="y")
        self.banned_kw_listbox.bind("<Double-1>", self.unban_selected_kw)

        btn_save = ttk.Button(
            top, text="Зберегти", command=lambda: self.save_and_close_blacklist(top)
        )
        btn_save.pack(fill="x", padx=10, pady=10, ipady=5)

        self.refresh_kw_lists()

    def refresh_kw_lists(self):
        if not hasattr(self, "current_author_keywords"):
            return
        self.avail_kw_listbox.delete(0, tk.END)
        self.banned_kw_listbox.delete(0, tk.END)

        search_q = self.kw_search_var.get().strip().lower()
        banned_set = set(self.current_banned_keywords)

        for kw in self.current_author_keywords:
            if kw not in banned_set:
                if not search_q or search_q in kw:
                    self.avail_kw_listbox.insert(tk.END, kw)

        for kw in self.current_banned_keywords:
            if not search_q or search_q in kw:
                self.banned_kw_listbox.insert(tk.END, kw)

    def ban_selected_kw(self, e=None):
        sel = self.avail_kw_listbox.curselection()
        if not sel:
            return
        kws_to_ban = [self.avail_kw_listbox.get(i) for i in sel]
        for kw in kws_to_ban:
            if kw not in self.current_banned_keywords:
                self.current_banned_keywords.append(kw)
        self.refresh_kw_lists()

    def unban_selected_kw(self, e=None):
        sel = self.banned_kw_listbox.curselection()
        if not sel:
            return
        kws_to_unban = [self.banned_kw_listbox.get(i) for i in sel]
        for kw in kws_to_unban:
            if kw in self.current_banned_keywords:
                self.current_banned_keywords.remove(kw)
        self.refresh_kw_lists()

    def save_and_close_blacklist(self, top):
        def norm_kw(txt):
            if not txt:
                return ""
            return (
                str(txt)
                .lower()
                .replace("'", "'")
                .replace("`", "'")
                .replace("'", "'")
                .strip(string.punctuation + " ")
            )

        for cid in self.blacklist_cids:
            self.all_candidates[cid]["banned_keywords"] = [
                norm_kw(k) for k in self.current_banned_keywords
            ]
        self.recalculate_all_scores()
        top.destroy()
        messagebox.showinfo("Збережено", "Слова виключено з аналізу.")

    def generate_advice_strategy(self):
        sel = self.advice_listbox.curselection()
        if not sel:
            messagebox.showwarning("Увага", "Оберіть хоча б одного кандидата.")
            return

        cids = [self.advice_cid_map[i] for i in sel]

        def norm_kw(txt):
            if not txt:
                return ""
            return (
                str(txt)
                .lower()
                .replace("'", "'")
                .replace("`", "'")
                .replace("'", "'")
                .strip(string.punctuation + " ")
            )

        banned_set = set()
        for b in self.global_banned_keywords:
            banned_set.add(norm_kw(b))
        for cid in cids:
            for b in self.all_candidates[cid].get("banned_keywords", []):
                banned_set.add(norm_kw(b))

        all_kw_by_author = {}
        for cid in cids:
            auth_kws = []
            for pid in self.all_candidates[cid]["papers_uuids"]:
                p = self.all_papers[pid]
                if not p["recent"]:
                    continue

                manual_kws = [
                    norm_kw(w)
                    for w in p.get("manual_keywords", "").split(",")
                    if w.strip()
                ]
                matched_kws = []
                md = p.get("matched_details", "")
                if md:
                    for part in md.split(","):
                        part = part.strip()
                        if part.startswith("'") and "' (" in part:
                            kw = part.split("' (")[0].strip("'")
                            if kw:
                                matched_kws.append(norm_kw(kw))
                words = (
                    [norm_kw(w) for w in p.get("author_keywords", [])]
                    + manual_kws
                    + matched_kws
                )
                words = [w for w in words if w and w not in banned_set]
                auth_kws.extend(words)
            all_kw_by_author[cid] = auth_kws

        self.advice_output.config(state="normal")
        self.advice_output.delete("1.0", tk.END)

        report = f"Аналіз термінів за останні 5 років\n"
        report += "=" * 60 + "\n\n"

        aggregated_all = []
        for cid in cids:
            name = self.all_candidates[cid]["name"]
            kws = all_kw_by_author[cid]
            aggregated_all.extend(kws)
            top = Counter(kws).most_common(10)

            report += f"Найчастіші терміни: {name}\n"
            if not top:
                report += "   Недостатньо даних.\n"
            for word, count in top:
                report += f"   - {word} ({count} разів)\n"
            report += "\n"

        if len(cids) > 1:
            report += "Спільні теми\n"
            report += "-" * 60 + "\n"
            sets = [set(all_kw_by_author[cid]) for cid in cids]
            intersection = set.intersection(*sets) if sets else set()
            if intersection:
                report += f"Можливі теми для співавторства:\n   > {', '.join(list(intersection)[:15])}\n"
            else:
                report += "Спільних тем не знайдено.\n"
            report += "\n"

        report += "Відсутні ключові слова\n"
        report += "-" * 60 + "\n"
        matched = []
        missing = []
        if not self.target_keywords:
            report += "Ключові слова не задані.\n"
        else:
            matched = []
            missing = []
            for target_kw in self.target_keywords:
                target_norm = norm_kw(target_kw)
                pat = rf"(?u)(?<!\w){re.escape(target_norm)}(?!\w)"
                found = False
                for word in aggregated_all:
                    if re.search(pat, word):
                        found = True
                        break
                if found:
                    matched.append(target_kw)
                else:
                    missing.append(target_kw)

            report += f"Використано ключових слів: {len(matched)}\n"
            if matched:
                report += f"   > {', '.join(matched)}\n"

            report += f"\nНе використано ключових слів: {len(missing)}\n"
            if missing:
                report += f"   > {', '.join(missing)}\n"

            report += "\nПоради:\n"
            top_overall = [w for w, c in Counter(aggregated_all).most_common(3)]
            if top_overall and missing:
                report += f"Щоб підвищити відповідність, поєднайте частий напрям\n"
                report += (
                    f"('{top_overall[0]}') з відсутньою темою ('{list(missing)[0]}').\n"
                )
            else:
                report += "Продовжуйте публікувати в основній тематиці."

        self.advice_output.insert("1.0", report)
        self.advice_output.config(state="disabled")

        if AI_ADVISOR_AVAILABLE and hasattr(self, "ai_advisor_btn"):
            self.ai_advisor_btn.config(state="normal")

    def export_advice_report(self):
        content = self.advice_output.get("1.0", tk.END).strip()
        if not content:
            return

        f_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("Markdown", "*.md")],
        )
        if not f_path:
            return
        with open(f_path, "w", encoding="utf-8") as f:
            f.write(content)
        messagebox.showinfo("Збереження", "Успішно збережено!")

    def launch_ai_advisor(self, restore_state=None):
        if not AI_ADVISOR_AVAILABLE:
            messagebox.showerror("Помилка", "Модуль AI Консультанта недоступний")
            return

        if not self.all_candidates:
            messagebox.showwarning("Увага", "Спочатку проведіть аналіз кандидатів")
            return

        if restore_state is None and hasattr(self, "_ai_restore_state"):
            restore_state = self._ai_restore_state
            self._ai_restore_state = None

        if self.ai_advisor_instance is not None:
            try:
                self.ai_advisor_instance.show_window()
                return
            except:
                self.ai_advisor_instance = None

        selected_indices = self.advice_listbox.curselection()
        if selected_indices:
            selected_cand_ids = [self.advice_cid_map[i] for i in selected_indices]
        else:
            selected_cand_ids = list(self.all_candidates.keys())

        try:
            self.ai_advisor_instance = launch_ai_advisor(
                parent_window=self.root,
                candidates=self.all_candidates,
                papers=self.all_papers,
                target_keywords=self.target_keywords,
                cutoff_year=self.cutoff_year,
                global_banned=self.global_banned_keywords,
                selected_cand_ids=selected_cand_ids,
                restore_state=restore_state,
            )
        except Exception as e:
            messagebox.showerror(
                "Помилка AI", f"Не вдалося запустити AI Консультанта:\n{str(e)}"
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = MonCouncilProApp(root)
    root.mainloop()
