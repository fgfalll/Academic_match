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
            "Referer": "https://scholar.google.com/"
        }
        if hasattr(scholarly, 'nav'):
            scholarly.nav.session.headers.update(headers)
    except Exception as e:
        print(f"Помилка налаштування: {e}")

setup_scholarly()


# --- ДОПОМІЖНІ ФУНКЦІЇ ---

def decode_openalex_abstract(inverted_index):
    """Декодує анотацію з формату inverted index OpenAlex."""
    if not inverted_index: return ""
    try:
        word_index = []
        for word, locations in inverted_index.items():
            for loc in locations: word_index.append((loc, word))
        word_index.sort(key=lambda x: x[0])
        return " ".join(word for index, word in word_index)
    except: return ""


def get_author_info_openalex(orcid):
    """Отримує офіційне ім'я автора через OpenAlex Author API."""
    headers = {'User-Agent': 'AcademicMatch/1.0 (mailto:mon-phd-check@example.com)'}
    url = f"https://api.openalex.org/authors/https://orcid.org/{orcid}"
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.json().get('display_name', 'Невідомо')
    except: pass
    return "Невідомо"


def heuristic_score(title, concepts, author_keywords, manual_keywords, target_keywords, banned_keywords=[], abstract=""):
    """Алгоритм оцінки релевантності за назвою, ключовими словами та анотацією з урахуванням чорного списку."""
    score = 0
    def norm(txt):
        return str(txt).lower().replace("'", "'").replace("`", "'").replace("'", "'") if txt else ""
        
    t_l = norm(title)
    ab_l = norm(abstract)
    c_l = [norm(c) for c in concepts]
    ak_l = [norm(k) for k in author_keywords]
    mk_l = norm(manual_keywords)
    matched = []

    banned_set = set([norm(b).strip() for b in banned_keywords if b.strip()])

    for kw in target_keywords:
        kw = norm(kw).strip(string.punctuation + " ")
        if not kw or kw in banned_set: continue
        pat = rf'(?u)(?<!\w){re.escape(kw)}(?!\w)'

        found_in_title = False
        if re.search(pat, t_l):
            score += 5; matched.append(f"'{kw}' (Назва:+5)")
            found_in_title = True

        combined_kw = ak_l + [mk_l] if mk_l else ak_l
        found_kw = False
        for kw_src in combined_kw:
            if re.search(pat, kw_src):
                score += 4; matched.append(f"'{kw}' (Ключове слово:+4)")
                found_kw = True; break
        if found_kw: continue

        found_c = False
        for c in c_l:
            if re.search(pat, c):
                score += 3; matched.append(f"'{kw}' (Напрям:+3)")
                found_c = True; break
        if found_c: continue

        if not found_in_title and ab_l and re.search(pat, ab_l):
            score += 2; matched.append(f"'{kw}' (Анотація:+2)")

    return score, list(set(matched))


class MonCouncilProApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Аналітика разової вченої ради (КМУ №44)")
        self.root.geometry("1200x900")
        self.all_candidates = {}; self.all_papers = {}
        self.cutoff_year = 2022; self.target_keywords = []; self.current_cand_filter = None
        self.current_author_keywords = []
        self.current_banned_keywords = []
        self.create_widgets(); self.update_keyword_preview()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root); self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_main = ttk.Frame(self.notebook)
        self.tab_edit = ttk.Frame(self.notebook)
        self.tab_advice = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_main, text="1. Налаштування")
        self.notebook.add(self.tab_edit, text="2. Результати")
        self.notebook.add(self.tab_advice, text="3. Аналіз термінів")
        
        self.build_main_tab()
        self.build_edit_tab()
        self.build_advice_tab()

    def build_main_tab(self):
        sf = ttk.LabelFrame(self.tab_main, text="Дані здобувача та керівника", padding="10"); sf.pack(fill="x", padx=10, pady=5)
        ttk.Label(sf, text="Рік ради:").grid(row=0, column=0, sticky="w")
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Entry(sf, textvariable=self.year_var, width=10).grid(row=0, column=1, sticky="w", padx=5)
        
        ttk.Label(sf, text="Здобувач (ORCID / ПІБ):").grid(row=1, column=0, sticky="w")
        self.phd_id_var = tk.StringVar()
        ttk.Entry(sf, textvariable=self.phd_id_var, width=25).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Button(sf, text="Отримати терміни", command=self.auto_fetch_keywords).grid(row=1, column=2, sticky="w", padx=10)

        ttk.Label(sf, text="Керівник (ORCID / GS / ПІБ):").grid(row=2, column=0, sticky="w")
        self.super_id_var = tk.StringVar()
        ttk.Entry(sf, textvariable=self.super_id_var, width=25).grid(row=2, column=1, sticky="w", padx=5)
        self.deep_analysis_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(sf, text="Аналізувати анотації та співавторів", variable=self.deep_analysis_var).grid(row=2, column=2, sticky="w", padx=10)

        wa = ttk.Frame(self.tab_main); wa.pack(fill="both", expand=True, padx=10, pady=5)
        inf = ttk.LabelFrame(wa, text="Кандидати (ORCHID\Google Scholar через кому)", padding="10"); inf.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.candidates_text = tk.Text(inf, height=5); self.candidates_text.pack(fill="both", expand=True)
        self.candidates_text.insert("1.0", "")

        kwf = ttk.LabelFrame(wa, text="Ключові слова (через кому)", padding="10"); kwf.pack(side="right", fill="both", expand=True, padx=(5, 0))
        self.keyword_text = tk.Text(kwf, height=3); self.keyword_text.pack(fill="both", expand=True, pady=(0, 5))
        self.keyword_text.insert("1.0", "")
        self.keyword_text.bind("<KeyRelease>", self.update_keyword_preview)
        self.parsed_kw_label = ttk.Label(kwf, text="", foreground="#0056b3", wraplength=400); self.parsed_kw_label.pack(fill="x")

        bp = ttk.Frame(self.tab_main); bp.pack(fill="x", padx=10, pady=10)
        self.run_btn = ttk.Button(bp, text="Почати аналіз", command=self.start_analysis); self.run_btn.pack(side="left", fill="x", expand=True, ipady=5)
        ttk.Button(bp, text="Перевірка CAPTCHA", command=lambda: webbrowser.open("https://scholar.google.com/scholar?q=test")).pack(side="right", padx=5, ipady=5)

        lf = ttk.LabelFrame(self.tab_main, text="Журнал подій", padding="10"); lf.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_area = scrolledtext.ScrolledText(lf, wrap=tk.WORD, state='disabled', height=6, font=("Consolas", 9)); self.log_area.pack(fill="both", expand=True)

    def build_edit_tab(self):
        sumf = ttk.LabelFrame(self.tab_edit, text="Підсумок", padding="10"); sumf.pack(fill="x", padx=10, pady=5)
        cols = ("cand_id", "name", "ids", "relevant", "conflict", "status")
        self.tree_sum = ttk.Treeview(sumf, columns=cols, show="headings", height=4)
        for c, t in zip(cols, ["ID", "Кандидат", "Джерела", "Статті 5р", "Конфлікт", "Статус"]): self.tree_sum.heading(c, text=t)
        self.tree_sum.column("cand_id", width=0, stretch=tk.NO); self.tree_sum.pack(fill="x")
        self.tree_sum.tag_configure("pass", background="#d4edda"); self.tree_sum.tag_configure("fail", background="#f8d7da")
        self.tree_sum.bind("<<TreeviewSelect>>", self.on_candidate_select)

        paf = ttk.LabelFrame(self.tab_edit, text="Список статей", padding="10"); paf.pack(fill="both", expand=True, padx=10, pady=5)
        fp = ttk.Frame(paf); fp.pack(fill="x", pady=(0, 10))
        ttk.Button(fp, text="Всі", command=self.clear_candidate_filter).pack(side="left", padx=5)
        self.search_title_var = tk.StringVar()
        ent = ttk.Entry(fp, textvariable=self.search_title_var, width=25); ent.pack(side="left"); ent.bind("<KeyRelease>", lambda e: self.refresh_papers_table())
        self.filter_recent_var = tk.BooleanVar(value=True); ttk.Checkbutton(fp, text="Відсікати старі", variable=self.filter_recent_var, command=self.refresh_papers_table).pack(side="left", padx=10)
        self.filter_score_var = tk.BooleanVar(value=False); ttk.Checkbutton(fp, text="Тільки з балами", variable=self.filter_score_var, command=self.refresh_papers_table).pack(side="left", padx=5)
        ttk.Button(fp, text="Додати статтю", command=self.open_add_manual_paper).pack(side="right", padx=5)

        pcols = ("uuid", "year", "recent", "score", "matches", "title", "source")
        self.tree_pap = ttk.Treeview(paf, columns=pcols, show="headings")
        for c, t in zip(pcols, ["UUID", "Рік", "Нова", "Бали", "Збіги", "Назва", "Джерело"]): self.tree_pap.heading(c, text=t)
        
        self.tree_pap.column("uuid", width=0, minwidth=0, stretch=tk.NO)
        self.tree_pap.column("year", width=50, minwidth=50, anchor="center", stretch=tk.NO)
        self.tree_pap.column("recent", width=40, minwidth=40, anchor="center", stretch=tk.NO)
        self.tree_pap.column("score", width=40, minwidth=40, anchor="center", stretch=tk.NO)
        self.tree_pap.column("matches", width=150, minwidth=100, stretch=tk.NO)
        self.tree_pap.column("title", minwidth=200, stretch=tk.YES)
        self.tree_pap.column("source", width=100, minwidth=80, anchor="center", stretch=tk.NO)

        sb_y = ttk.Scrollbar(paf, orient="vertical", command=self.tree_pap.yview)
        sb_x = ttk.Scrollbar(paf, orient="horizontal", command=self.tree_pap.xview)
        self.tree_pap.configure(yscroll=sb_y.set, xscroll=sb_x.set)
        
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.tree_pap.pack(fill="both", expand=True)
        self.tree_pap.bind("<<TreeviewSelect>>", self.on_paper_select)
        
        self.pm = tk.Menu(self.root, tearoff=0)
        self.pm.add_command(label="Деталі", command=self.open_paper_details)
        self.pm.add_command(label="Редагувати ключові слова", command=self.open_manual_tags_dialog)
        self.tree_pap.bind("<Button-3>", lambda e: self.pm.tk_popup(e.x_root, e.y_root))

    def build_advice_tab(self):
        main_frame = ttk.Frame(self.tab_advice)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        left_panel = ttk.LabelFrame(main_frame, text="Кандидати", padding="5")
        left_panel.pack(side="left", fill="y", padx=(0, 10))

        list_frame = ttk.Frame(left_panel)
        list_frame.pack(fill="both", expand=True)

        self.advice_listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, width=35, exportselection=False)
        sb_l_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.advice_listbox.yview)
        sb_l_x = ttk.Scrollbar(list_frame, orient="horizontal", command=self.advice_listbox.xview)
        self.advice_listbox.config(yscrollcommand=sb_l_y.set, xscrollcommand=sb_l_x.set)
        
        self.advice_listbox.grid(row=0, column=0, sticky="nsew")
        sb_l_y.grid(row=0, column=1, sticky="ns")
        sb_l_x.grid(row=1, column=0, sticky="ew")
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        mid_panel = ttk.Frame(main_frame)
        mid_panel.pack(side="left", fill="y", padx=5)

        ban_frame = ttk.LabelFrame(mid_panel, text="Виключення слів", padding="5")
        ban_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(ban_frame, text="Слова, які не враховуються\nдля виділених кандидатів.", wraplength=200).pack(pady=5)
        ttk.Button(ban_frame, text="Список виключень", command=self.open_blacklist_window).pack(fill="x", pady=5, ipady=3)

        ttk.Button(mid_panel, text="Аналізувати", command=self.generate_advice_strategy).pack(fill="x", pady=10, ipady=5)
        ttk.Button(mid_panel, text="Зберегти звіт (.txt)", command=self.export_advice_report).pack(fill="x", pady=5)

        right_panel = ttk.LabelFrame(main_frame, text="Результати аналізу", padding="5")
        right_panel.pack(side="left", fill="both", expand=True, padx=(10, 0))
        
        self.advice_output = scrolledtext.ScrolledText(right_panel, wrap=tk.WORD, font=("Arial", 10))
        self.advice_output.pack(fill="both", expand=True)
        self.advice_output.config(state="disabled")

    def log(self, msg):
        self.log_area.config(state='normal'); self.log_area.insert(tk.END, msg + "\n"); self.log_area.see(tk.END); self.log_area.config(state='disabled'); self.root.update()

    def clear_log(self):
        self.log_area.config(state='normal'); self.log_area.delete("1.0", tk.END); self.log_area.config(state='disabled'); self.root.update()

    def log_status(self, h, s):
        self.log_area.config(state='normal'); self.log_area.delete("1.0", tk.END); self.log_area.insert(tk.END, f"{h}\n{'-'*40}\n{s}\n"); self.log_area.see(tk.END); self.log_area.config(state='disabled'); self.root.update()

    def update_keyword_preview(self, e=None):
        raw = self.keyword_text.get("1.0", tk.END)
        self.target_keywords = [k.strip(string.punctuation + " \n").lower() for k in raw.split(',') if k.strip()]
        self.parsed_kw_label.config(text=f"[Масив: {', '.join(self.target_keywords)}]" if self.target_keywords else "")

    def auto_fetch_keywords(self):
        oid = self.phd_id_var.get().strip()
        if not oid: return
        self.log(f"Отримання термінів для {oid}..."); threading.Thread(target=self._fetch_kw_thread, args=(oid,), daemon=True).start()

    def _fetch_kw_thread(self, oid):
        try:
            r = requests.get(f"https://api.openalex.org/works?filter=author.orcid:https://orcid.org/{oid}&per-page=50")
            if r.status_code == 200:
                kws = [t.get('display_name', '').lower() for w in r.json().get('results', []) for t in w.get('topics', [])]
                if kws:
                    top = ", ".join([i[0] for i in Counter(kws).most_common(8)])
                    self.root.after(0, lambda: self.keyword_text.delete("1.0", tk.END))
                    self.root.after(0, lambda: self.keyword_text.insert("1.0", top))
                    self.root.after(0, self.update_keyword_preview)
        except: pass

    def recalculate_all_scores(self):
        self.cutoff_year = int(self.year_var.get() or datetime.now().year) - 4
        self.log(f"--- ОНОВЛЕННЯ БАЛІВ (Межа: {self.cutoff_year}) ---")
        for p in self.all_papers.values():
            cid = p['cand_id']
            banned = self.all_candidates[cid].get('banned_keywords', [])
            sc, m = heuristic_score(p['title'], p.get('concepts', []), p.get('author_keywords', []), p.get('manual_keywords', ''), self.target_keywords, banned_keywords=banned, abstract=p.get('abstract', ''))
            p.update({'score': sc, 'matched_details': ", ".join(m), 'recent': (p['year'] >= self.cutoff_year)})
        self.refresh_all_tables()

    def start_analysis(self):
        lines = [l.strip() for l in self.candidates_text.get("1.0", tk.END).split('\n') if l.strip()]
        if not lines or not self.target_keywords: return
        self.all_candidates.clear(); self.all_papers.clear()
        for i in self.tree_sum.get_children(): self.tree_sum.delete(i)
        for i in self.tree_pap.get_children(): self.tree_pap.delete(i)
        self.run_btn.config(state='disabled'); self.clear_log()
        threading.Thread(target=self.run_algorithm, args=(lines, self.phd_id_var.get().strip().lower(), self.super_id_var.get().strip().lower()), daemon=True).start()

    def run_algorithm(self, lines, phd_id, super_id):
        self.cutoff_year = int(self.year_var.get() or datetime.now().year) - 4
        for idx, line in enumerate(lines):
            cand_id = f"cand_{idx}"
            parts = [p.strip() for p in line.split(',')]
            orcid = ""
            gs_id = ""
            
            for p in parts:
                p_clean = p.replace('‑', '-').replace('−', '-') # Нормалізація дефісів
                
                orcid_m = re.search(r'\b\d{4}-\d{4}-\d{4}-\d{3}[\dX]\b', p_clean)
                if orcid_m:
                    orcid = orcid_m.group(0)
                elif 'user=' in p_clean:
                    gs_m = re.search(r'user=([\w-]{12})', p_clean)
                    if gs_m: gs_id = gs_m.group(1)
                elif re.match(r'^[\w-]{12}$', p_clean):
                    gs_id = p_clean
                elif len(p_clean) > 5 and not orcid:
                    gs_id = p_clean # Резервний варіант
            
            d_ids = []; doi_map = {}; merged_local = {}
            if orcid: d_ids.append(f"ORCID:{orcid}")
            if gs_id: d_ids.append(f"GS:{gs_id}")
            a_name = "Невідомо"
            
            # Перевірка на керівника: тільки якщо вказано super_id І він збігається з orcid/gs_id поточного кандидата
            conflict = "Немає"
            if super_id and (super_id == orcid.lower() or super_id == gs_id.lower()): 
                conflict = "Керівник"

            # 1. ORCID
            if orcid:
                self.log(f"[{cand_id}] Отримання робіт з ORCID...")
                try:
                    r = requests.get(f"https://pub.orcid.org/v3.0/{orcid}/works", headers={'Accept': 'application/json'}, timeout=15)
                    if r.status_code == 200:
                        for g in r.json().get('group', []):
                            for s in g.get('work-summary', []):
                                t = s.get('title', {}).get('title', {}).get('value', '')
                                if not t: continue
                                y = 0; doi = ""
                                try: y = int(s.get('publication-date', {}).get('year', {}).get('value', '0'))
                                except: pass
                                
                                ext_ids = s.get('external-ids') or {}
                                if isinstance(ext_ids, dict):
                                    for ext in ext_ids.get('external-id', []):
                                        if ext.get('external-id-type') == 'doi':
                                            doi = (ext.get('external-id-value') or '').lower().strip().replace('https://doi.org/', '')
                                            break
                                            
                                k = re.sub(r'\W+', '', t.lower())
                                p_data = {'title': t, 'year': y, 'doi': doi, 'concepts': [], 'author_keywords': [], 'abstract': '', 'source': 'ORCID', 'manual_keywords': '', 'authors_full': [], 'journal': '-', 'url': s.get('url', {}).get('value', '') if s.get('url') else (f"https://doi.org/{doi}" if doi else '')}
                                merged_local[k] = p_data
                                if doi: doi_map[doi] = k
                    else: self.log(f"   ! Помилка ORCID HTTP {r.status_code}")
                except Exception as e: self.log(f"   ! Помилка ORCID: {str(e)}")

            # 2. OpenAlex
            if orcid:
                self.log(f"[{cand_id}] Доповнення через OpenAlex...")
                oa_h = {'User-Agent': 'AcademicMatch/1.0 (mailto:mon-phd-check@example.com)'}
                try:
                    an_fetch = get_author_info_openalex(orcid)
                    if an_fetch != "Невідомо": a_name = an_fetch

                    r_l = requests.get(f"https://api.openalex.org/works?filter=author.orcid:https://orcid.org/{orcid}&per-page=200", headers=oa_h, timeout=15)
                    if r_l.status_code == 200:
                        works = r_l.json().get('results', [])
                        self.log(f"   - Знайдено {len(works)} робіт")
                        for i, ws in enumerate(works):
                            w_title = ws.get('title') or ''
                            w_doi = (ws.get('doi') or '').lower().replace('https://doi.org/', '')
                            k_oa = re.sub(r'\W+', '', w_title.lower())

                            target_k = None
                            if w_doi and w_doi in doi_map: target_k = doi_map[w_doi]
                            elif k_oa and k_oa in merged_local: target_k = k_oa

                            if not self.deep_analysis_var.get():
                                pl = ws.get('primary_location') or {}
                                journal = (pl.get('source') or {}).get('display_name', '-') or '-'
                                meta = {'concepts': [c.get('display_name', '') for c in ws.get('topics', [])], 'author_keywords': [], 'journal': journal}
                                if target_k: merged_local[target_k].update(meta)
                                else: merged_local[k_oa] = {'title': w_title, 'year': ws.get('publication_year', 0), 'doi': w_doi, 'source': 'OpenAlex', 'manual_keywords': '', 'abstract': '', 'authors_full': [], **meta}
                            else:
                                self.log_status(f"Деталі OpenAlex: {a_name}", f"Обробка {i+1}/{len(works)}")
                                try:
                                    ab = decode_openalex_abstract(ws.get('abstract_inverted_index'))
                                    pl = ws.get('primary_location') or {}
                                    journal = (pl.get('source') or {}).get('display_name', '-') or '-'
                                    meta = {
                                        'concepts': [c.get('display_name', '') for c in ws.get('topics', [])] or [c.get('display_name', '') for c in ws.get('concepts', [])], 
                                        'author_keywords': [kw.get('display_name', '') for kw in ws.get('keywords', [])], 
                                        'abstract': ab, 
                                        'journal': journal, 
                                        'authors_full': [a.get('author', {}).get('display_name', 'Невідомо') for a in ws.get('authorships', [])], 
                                        'url': ws.get('doi') or ''
                                    }
                                    if target_k: 
                                        merged_local[target_k].update(meta)
                                        if 'OpenAlex' not in merged_local[target_k]['source']: merged_local[target_k]['source'] += " + OA"
                                    else: 
                                        merged_local[k_oa] = {'title': w_title, 'year': ws.get('publication_year', 0), 'doi': w_doi, 'source': 'OpenAlex', 'manual_keywords': '', **meta}

                                    if phd_id:
                                        for auth in ws.get('authorships', []):
                                            if phd_id in (auth.get('author', {}).get('orcid') or "").lower(): conflict = "Співавтор"
                                except Exception as e: self.log(f"   ! Помилка OA Item: {str(e)}")
                    else: self.log(f"   ! Помилка списку OA HTTP {r_l.status_code}")
                except Exception as e: self.log(f"   ! Помилка запиту OA: {str(e)}")

            # 3. Scholar
            if gs_id:
                header = f"Scholar: {a_name if a_name != 'Невідомо' else gs_id}"
                self.log_status(header, "Отримання списку Scholar...")
                try:
                    aq = scholarly.search_author_id(gs_id); ad = scholarly.fill(aq, sections=['publications'])
                    if a_name == "Невідомо": a_name = ad.get('name', 'Невідомо')
                    pubs = ad.get('publications', []); interests = ad.get('interests', [])
                    for i, w in enumerate(pubs):
                        self.log_status(header, f"{i+1}/{len(pubs)}")
                        if i > 0: time.sleep(random.uniform(15, 25) if i % 5 == 0 else random.uniform(5, 10))
                        try:
                            bib = w.get('bib', {}); t = bib.get('title', ''); y = int(bib.get('pub_year', '0')); ab = ""
                            if self.deep_analysis_var.get() and y >= (self.cutoff_year - 1):
                                try: time.sleep(random.uniform(3, 7)); wf = scholarly.fill(w); ab = wf.get('bib', {}).get('abstract', '')
                                except: pass
                            if phd_id and (phd_id in bib.get('author', '').lower() or a_name.lower() in bib.get('author', '').lower()): conflict = "Співавтор"
                            k = re.sub(r'\W+', '', t.lower())
                            if k in merged_local:
                                merged_local[k]['source'] += " + GS"
                                if not merged_local[k].get('abstract'): merged_local[k]['abstract'] = ab
                            else: merged_local[k] = {'title': t, 'year': y, 'concepts': interests, 'author_keywords': [], 'abstract': ab, 'source': 'Scholar', 'manual_keywords': '', 'authors_full': [], 'journal': '-', 'url': w.get('pub_url', '')}
                        except: continue
                except Exception as e: self.log(f"Помилка Scholar: {str(e)}")

            self.all_candidates[cand_id] = {'name': a_name, 'ids': ", ".join(d_ids), 'conflict': conflict, 'papers_uuids': [], 'banned_keywords': []}
            for pid, pd_item in merged_local.items():
                u = str(uuid.uuid4()); self.all_candidates[cand_id]['papers_uuids'].append(u)
                sc, m = heuristic_score(pd_item['title'], pd_item.get('concepts', []), pd_item.get('author_keywords', []), pd_item.get('manual_keywords', ''), self.target_keywords, abstract=pd_item.get('abstract', ''))
                pd_item.update({'score': sc, 'matched_details': ", ".join(m), 'recent': (pd_item['year'] >= self.cutoff_year), 'cand_id': cand_id})
                self.all_papers[u] = pd_item

        self.root.after(0, self.refresh_all_tables); self.log("\nАналіз завершено"); self.root.after(0, lambda: self.run_btn.config(state='normal'))

    def on_candidate_select(self, e):
        sel = self.tree_sum.selection()
        if sel: self.current_cand_filter = self.tree_sum.item(sel[0])['values'][0]; self.refresh_papers_table()

    def clear_candidate_filter(self):
        self.current_cand_filter = None; self.tree_sum.selection_remove(self.tree_sum.selection()); self.refresh_papers_table()

    def refresh_all_tables(self):
        selected = self.current_cand_filter; [self.tree_sum.delete(i) for i in self.tree_sum.get_children()]
        item_to_sel = None
        for cid, c in self.all_candidates.items():
            rel = sum(1 for u in c['papers_uuids'] if self.all_papers[u]['score'] > 0 and self.all_papers[u]['recent'])
            passed = (rel >= 3 and c['conflict'] == "Немає")
            status, tag = ("Відповідає вимогам", "pass") if passed else (f"Не відповідає ({rel}/3)", "fail")
            item = self.tree_sum.insert("", tk.END, values=(cid, c['name'], c['ids'], rel, c['conflict'], status), tags=(tag,))
            if cid == selected: item_to_sel = item
        if item_to_sel: self.tree_sum.selection_set(item_to_sel)
        self.refresh_papers_table()
        self.update_advice_authors_list()

    def refresh_papers_table(self):
        [self.tree_pap.delete(i) for i in self.tree_pap.get_children()]
        sq = self.search_title_var.get().strip().lower()
        f_rec = self.filter_recent_var.get(); f_sc = self.filter_score_var.get()
        sorted_p = sorted(self.all_papers.items(), key=lambda x: (x[1]['recent'], x[1]['score']), reverse=True)
        for u, p in sorted_p:
            if self.current_cand_filter and p['cand_id'] != self.current_cand_filter: continue
            if f_rec and not p['recent']: continue
            if f_sc and p['score'] <= 0: continue
            if sq:
                txt = (p['title'] + " " + p['manual_keywords'] + " " + ",".join(p.get('concepts', []))).lower()
                if sq not in txt: continue
            self.tree_pap.insert("", tk.END, values=(u, p['year'], "Так" if p['recent'] else "Ні", p['score'], p['matched_details'], p['title'], p['source']))

    def on_paper_select(self, e):
        sel = self.tree_pap.selection()
        if sel: self.selected_p_uuid = self.tree_pap.item(sel[0])['values'][0]

    def open_manual_tags_dialog(self):
        if not hasattr(self, 'selected_p_uuid'): return
        p = self.all_papers[self.selected_p_uuid]
        res = simpledialog.askstring("Редагувати ключові слова", f"Введіть ключові слова:\n{p['title'][:60]}...", initialvalue=p['manual_keywords'], parent=self.root)
        if res is not None:
            p['manual_keywords'] = res.strip()
            cid = p['cand_id']
            banned = self.all_candidates[cid].get('banned_keywords', [])
            sc, m = heuristic_score(p['title'], p.get('concepts', []), p.get('author_keywords', []), p['manual_keywords'], self.target_keywords, banned_keywords=banned, abstract=p.get('abstract', ''))
            p.update({'score': sc, 'matched_details': ", ".join(m)}); self.refresh_all_tables()

    def open_paper_details(self):
        if not hasattr(self, 'selected_p_uuid'): return
        p = self.all_papers[self.selected_p_uuid]
        top = tk.Toplevel(self.root); top.title("Деталі публікації"); top.geometry("750x750")
        tk.Label(top, text=p['title'], wraplength=700, font=("Arial", 11, "bold"), justify="left").pack(pady=10, padx=15)
        txt = scrolledtext.ScrolledText(top, height=35, wrap=tk.WORD, font=("Arial", 10))
        txt.pack(padx=15, fill="both", expand=True)
        c = f"АВТОР: {self.all_candidates[p['cand_id']]['name']}\nРІК: {p['year']} | БАЛИ: {p['score']}\nДЖЕРЕЛО: {p['source']} | ЖУРНАЛ: {p.get('journal','-')}\n"
        c += f"ЗБІГИ: {p.get('matched_details', '-')}\n" + "-"*60 + "\n"
        c += f"СПІВАВТОРИ: {', '.join(p.get('authors_full', []))}\n\n"
        c += f"КЛЮЧОВІ СЛОВА АВТОРА: {', '.join(p.get('author_keywords', []))}\n"
        c += f"КЛЮЧОВІ СЛОВА ШІ (OpenAlex): {', '.join(p.get('concepts', []))}\n\n"
        c += f"АНОТАЦІЯ:\n{p.get('abstract', 'Немає анотації.')}\n\n"
        if p['manual_keywords']: c += f"ВЛАСНІ КЛЮЧОВІ СЛОВА: {p['manual_keywords']}\n"
        txt.insert("1.0", c); txt.config(state="disabled")
        ttk.Button(top, text="Відкрити в браузері", command=lambda: webbrowser.open(p['url']) if p['url'] else None).pack(pady=10)

    def open_add_manual_paper(self):
        if not self.all_candidates: return
        
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
        cb = ttk.Combobox(author_frame, textvariable=cand_var, state="readonly", width=60)
        cb['values'] = [self.all_candidates[cid]['name'] for cid in cids_list]
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
        ttk.Button(btn_frame, text="Скасувати", command=top.destroy).pack(side="left", padx=10)
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
            
            banned = self.all_candidates[cid].get('banned_keywords', [])
            p_d = {
                'title': t, 'year': y, 'journal': j or '-', 'url': url,
                'concepts': [], 'author_keywords': [],
                'manual_keywords': mkw, 'abstract': '',
                'source': 'Manual', 'cand_id': cid
            }
            sc, m = heuristic_score(t, [], [], mkw, self.target_keywords, banned_keywords=banned, abstract='')
            p_d.update({'score': sc, 'matched_details': ", ".join(m), 'recent': (y >= self.cutoff_year)})
            u = str(uuid.uuid4())
            self.all_papers[u] = p_d
            self.all_candidates[cid]['papers_uuids'].append(u)
            self.refresh_all_tables()
            top.destroy()
        
        save_btn.config(command=save)
        validate_and_update()

    # --- ВКЛАДКА ПОРАД (ADVICE TAB) ---

    def update_advice_authors_list(self):
        self.advice_listbox.delete(0, tk.END)
        self.advice_cid_map = []
        for cid, c in self.all_candidates.items():
            self.advice_listbox.insert(tk.END, c['name'])
            self.advice_cid_map.append(cid)

    def open_blacklist_window(self):
        sel = self.advice_listbox.curselection()
        if not sel:
            messagebox.showwarning("Увага", "Оберіть хоча б одного кандидата.")
            return
            
        cids = [self.advice_cid_map[i] for i in sel]
        
        all_kws = set()
        banned_all = set()
        for cid in cids:
            banned_all.update(self.all_candidates[cid].get('banned_keywords', []))
            for pid in self.all_candidates[cid]['papers_uuids']:
                p = self.all_papers[pid]
                if p['recent']:
                    words = [w.strip().lower() for w in p.get('concepts', []) + p.get('author_keywords', [])]
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
        ttk.Entry(search_frame, textvariable=self.kw_search_var).pack(side="left", fill="x", expand=True, padx=5)

        lists_frame = ttk.Frame(top)
        lists_frame.pack(fill="both", expand=True, padx=10, pady=5)

        avail_frame = ttk.Frame(lists_frame)
        avail_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(avail_frame, text="Доступні:").pack(anchor="w")
        self.avail_kw_listbox = tk.Listbox(avail_frame, height=15, width=25, selectmode=tk.EXTENDED)
        self.avail_kw_listbox.pack(side="left", fill="both", expand=True)
        sb_a = ttk.Scrollbar(avail_frame, orient="vertical", command=self.avail_kw_listbox.yview)
        self.avail_kw_listbox.config(yscrollcommand=sb_a.set); sb_a.pack(side="right", fill="y")
        self.avail_kw_listbox.bind("<Double-1>", self.ban_selected_kw)

        btn_frame = ttk.Frame(lists_frame)
        btn_frame.pack(side="left", fill="y", padx=10)
        ttk.Button(btn_frame, text="Виключити", command=self.ban_selected_kw, width=11).pack(pady=(30, 5))
        ttk.Button(btn_frame, text="Повернути", command=self.unban_selected_kw, width=11).pack(pady=5)

        banned_frame = ttk.Frame(lists_frame)
        banned_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(banned_frame, text="Виключені:").pack(anchor="w")
        self.banned_kw_listbox = tk.Listbox(banned_frame, height=15, width=25, selectmode=tk.EXTENDED)
        self.banned_kw_listbox.pack(side="left", fill="both", expand=True)
        sb_b = ttk.Scrollbar(banned_frame, orient="vertical", command=self.banned_kw_listbox.yview)
        self.banned_kw_listbox.config(yscrollcommand=sb_b.set); sb_b.pack(side="right", fill="y")
        self.banned_kw_listbox.bind("<Double-1>", self.unban_selected_kw)
        
        btn_save = ttk.Button(top, text="Зберегти", command=lambda: self.save_and_close_blacklist(top))
        btn_save.pack(fill="x", padx=10, pady=10, ipady=5)
        
        self.refresh_kw_lists()

    def refresh_kw_lists(self):
        if not hasattr(self, 'current_author_keywords'): return
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
        if not sel: return
        kws_to_ban = [self.avail_kw_listbox.get(i) for i in sel]
        for kw in kws_to_ban:
            if kw not in self.current_banned_keywords:
                self.current_banned_keywords.append(kw)
        self.refresh_kw_lists()

    def unban_selected_kw(self, e=None):
        sel = self.banned_kw_listbox.curselection()
        if not sel: return
        kws_to_unban = [self.banned_kw_listbox.get(i) for i in sel]
        for kw in kws_to_unban:
            if kw in self.current_banned_keywords:
                self.current_banned_keywords.remove(kw)
        self.refresh_kw_lists()

    def save_and_close_blacklist(self, top):
        for cid in self.blacklist_cids:
            self.all_candidates[cid]['banned_keywords'] = list(self.current_banned_keywords)
        self.recalculate_all_scores()
        top.destroy()
        messagebox.showinfo("Збережено", "Слова виключено з аналізу.")

    def generate_advice_strategy(self):
        sel = self.advice_listbox.curselection()
        if not sel:
            messagebox.showwarning("Увага", "Оберіть хоча б одного кандидата.")
            return
        
        cids = [self.advice_cid_map[i] for i in sel]
        banned_set = set()
        for cid in cids:
            banned_set.update(self.all_candidates[cid].get('banned_keywords', []))
        
        all_kw_by_author = {}
        for cid in cids:
            auth_kws = []
            for pid in self.all_candidates[cid]['papers_uuids']:
                p = self.all_papers[pid]
                if not p['recent']: continue
                
                words = [w.strip().lower() for w in p.get('concepts', []) + p.get('author_keywords', [])]
                words = [w for w in words if w and w not in banned_set]
                auth_kws.extend(words)
            all_kw_by_author[cid] = auth_kws

        self.advice_output.config(state="normal")
        self.advice_output.delete("1.0", tk.END)
        
        report = f"Аналіз термінів за останні 5 років\n"
        report += "="*60 + "\n\n"
        
        aggregated_all = []
        for cid in cids:
            name = self.all_candidates[cid]['name']
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
            report += "-"*60 + "\n"
            sets = [set(all_kw_by_author[cid]) for cid in cids]
            intersection = set.intersection(*sets) if sets else set()
            if intersection:
                report += f"Можливі теми для співавторства:\n   > {', '.join(list(intersection)[:15])}\n"
            else:
                report += "Спільних тем не знайдено.\n"
            report += "\n"
        
        report += "Відсутні цільові терміни\n"
        report += "-"*60 + "\n"
        if not self.target_keywords:
            report += "Цільові терміни не задані.\n"
        else:
            all_used_set = set(aggregated_all)
            target_set = set(self.target_keywords)
            missing = target_set - all_used_set
            matched = target_set.intersection(all_used_set)
            
            report += f"Використано цільових тем: {len(matched)}\n"
            if matched: report += f"   > {', '.join(matched)}\n"
            
            report += f"\nНе використано цільових тем: {len(missing)}\n"
            if missing: report += f"   > {', '.join(missing)}\n"
            
            report += "\nПоради:\n"
            top_overall = [w for w, c in Counter(aggregated_all).most_common(3)]
            if top_overall and missing:
                report += f"Щоб підвищити відповідність, поєднайте частий напрям\n"
                report += f"('{top_overall[0]}') з відсутньою темою ('{list(missing)[0]}').\n"
            else:
                report += "Продовжуйте публікувати в основній тематиці."
                
        self.advice_output.insert("1.0", report)
        self.advice_output.config(state="disabled")

    def export_advice_report(self):
        content = self.advice_output.get("1.0", tk.END).strip()
        if not content:
            return
        
        f_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text file", "*.txt"), ("Markdown", "*.md")])
        if not f_path: return
        with open(f_path, 'w', encoding='utf-8') as f:
            f.write(content)
        messagebox.showinfo("Збереження", "Успішно збережено!")

if __name__ == "__main__":
    root = tk.Tk(); app = MonCouncilProApp(root); root.mainloop()