import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import json
import re
import os
import requests
from datetime import datetime
from typing import Optional, List, Dict, Any, Tuple
from dataclasses import dataclass, field
from collections import Counter

import markdown
import tkinterweb
import litellm

from crypto_utils import (
    encrypt_with_pin,
    decrypt_with_pin,
    encrypt_with_embedded_pin_hash,
    decrypt_with_embedded_pin_hash,
)

TAVILY_API_KEY = os.environ.get("TAVILY_API_KEY", "")


def web_search(query: str, num_results: int = 5) -> Dict[str, Any]:
    result = {"query": query, "source": "duckduckgo", "results": [], "error": None}

    try:
        from ddgs import DDGS

        with DDGS() as ddgs:
            ddg_results = list(ddgs.text(query, max_results=num_results))
            for r in ddg_results:
                result["results"].append(
                    {
                        "title": r.get("title", ""),
                        "url": r.get("href", ""),
                        "snippet": (
                            r.get("body", "")[:200] + "..."
                            if r.get("body") and len(r.get("body", "")) > 200
                            else r.get("body", "") or ""
                        ),
                    }
                )
        if not result["results"]:
            result["error"] = "Не вдалося знайтити результати"
    except Exception as e:
        result["error"] = str(e)

    if TAVILY_API_KEY and (result["error"] or not result["results"]):
        tavily_result = _tavily_search(query, num_results)
        if tavily_result["results"]:
            return tavily_result

    return result


def fetch_url_content(url: str, max_chars: int = 5000) -> dict:
    result = {"url": url, "content": "", "error": None}

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()

        from bs4 import BeautifulSoup

        soup = BeautifulSoup(response.text, "html.parser")

        for script in soup(["script", "style", "nav", "footer", "header", "aside"]):
            script.decompose()

        text = soup.get_text(separator="\n", strip=True)
        if not text:
            text = ""
        else:
            lines = [line for line in text.split("\n") if line.strip()]
            text = "\n".join(lines)

        if text and len(text) > max_chars:
            text = text[:max_chars] + f"\n... [Truncated {len(text) - max_chars} chars]"

        result["content"] = text
    except Exception as e:
        result["error"] = str(e)

    return result


def _tavily_search(query: str, num_results: int = 5) -> Dict[str, Any]:
    result = {"query": query, "source": "tavily", "results": [], "error": None}

    try:
        response = requests.post(
            "https://api.tavily.com/search",
            json={
                "api_key": TAVILY_API_KEY,
                "query": query,
                "max_results": num_results,
                "include_answer": True,
                "include_raw_content": False,
            },
            timeout=15,
        )
        response.raise_for_status()
        data = response.json()

        results = []
        for item in data.get("results", [])[:num_results]:
            results.append(
                {
                    "title": item.get("title", ""),
                    "url": item.get("url", ""),
                    "snippet": (
                        item.get("content", "")[:200] + "..."
                        if item.get("content") and len(item.get("content", "")) > 200
                        else item.get("content", "") or ""
                    ),
                }
            )

        result["results"] = results
        if data.get("answer"):
            result["answer"] = data["answer"]

    except Exception as e:
        result["error"] = str(e)

    return result


OPENALEX_MAILTO = os.environ.get("OPENALEX_MAILTO", "academic-match@example.com")
OPENALEX_BASE = "https://api.openalex.org"
OPENALEX_PER_PAGE = 5


def _openalex_get(endpoint: str, params: dict) -> dict:
    params["mailto"] = OPENALEX_MAILTO
    params["per_page"] = OPENALEX_PER_PAGE
    try:
        resp = requests.get(f"{OPENALEX_BASE}/{endpoint}", params=params, timeout=15)
        if resp.status_code == 400:
            return {"error": f"Bad request (400): {resp.text[:200]}"}
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        return {"error": str(e)}


def search_works(query: str, year: int = None, sort: str = None) -> dict:
    select = (
        "id,title,publication_year,authorships,cited_by_count,abstract_inverted_index"
    )
    params = {"search": query, "select": select}
    if year:
        params["filter"] = f"publication_year:{year}"
    if sort:
        params["sort"] = sort
    return _openalex_get("works", params)


def search_concepts(query: str) -> dict:
    params = {"search": query, "select": "id,display_name,level,description"}
    return _openalex_get("concepts", params)


def search_authors(query: str, field: str = None) -> dict:
    params = {
        "search": query,
        "select": "id,display_name,orcid,cited_by_count,h-index,works_count,topics",
    }
    return _openalex_get("authors", params)


SCHOLAR_DELAY_BASIC = 3
SCHOLAR_DELAY_DEEP = 15
_last_scholarly_call = 0

_scholar_cache: Dict[str, Tuple[Any, float]] = {}
SCHOLAR_CACHE_TTL = 1800

TOOL_DISPLAY_NAMES = {
    "get_candidate_data": "Завантажую дані кандидата",
    "compare_candidates": "Порівнюю кандидатів",
    "web_search": "Шукаю в інтернеті",
    "fetch_page": "Завантажую сторінку",
    "scholar_search": "Шукаю в Google Scholar",
    "openalex_search": "Шукаю в OpenAlex",
    "manage_banned_keywords": "Оновлюю фільтри",
}


def _scholar_rate_limit(delay_type: str = "basic"):
    global _last_scholarly_call
    import time
    import random

    min_delay = SCHOLAR_DELAY_DEEP if delay_type == "deep" else SCHOLAR_DELAY_BASIC
    elapsed = time.time() - _last_scholarly_call
    if elapsed < min_delay:
        time.sleep(min_delay - elapsed + random.uniform(1, 3))
    _last_scholarly_call = time.time()


def search_google_scholar(
    query: str, max_results: int = 10, fetch_details: bool = False
) -> dict:
    import time

    cache_key = f"search:{query}:{max_results}:{fetch_details}"
    now = time.time()

    if cache_key in _scholar_cache:
        cached_result, cached_time = _scholar_cache[cache_key]
        if now - cached_time < SCHOLAR_CACHE_TTL:
            cached_result["_cached"] = True
            return cached_result

    result = {"query": query, "results": [], "error": None}

    try:
        from scholarly import scholarly

        _scholar_rate_limit("basic" if not fetch_details else "deep")

        search_results = scholarly.search_pubs(query)
        count = 0

        for pub in search_results:
            if count >= max_results:
                break

            bib = pub.get("bib", {})
            title = bib.get("title", "N/A")
            year = bib.get("pub_year", "N/A")
            abstract = bib.get("abstract", "")
            citation_count = pub.get("num_citations", 0)
            authors = bib.get("author", [])
            if isinstance(authors, str):
                authors = [a.strip() for a in authors.split(",")]

            paper_info = {
                "title": title,
                "year": year,
                "authors": authors,
                "citation_count": citation_count,
                "abstract": abstract[:500] + "..."
                if abstract and len(abstract) > 500
                else abstract,
            }

            if fetch_details and year and str(year).isdigit():
                _scholar_rate_limit("deep")
                try:
                    filled = scholarly.fill(pub)
                    paper_info["abstract"] = filled.get("bib", {}).get(
                        "abstract", paper_info["abstract"]
                    )
                    paper_info["journal"] = filled.get("bib", {}).get("venue", "")
                    paper_info["url"] = filled.get("pub_url", "")
                except Exception:
                    pass

            result["results"].append(paper_info)
            count += 1

        _scholar_cache[cache_key] = (result.copy(), now)

    except Exception as e:
        result["error"] = str(e)

    return result


def search_google_scholar_author(
    author_name: str, max_results: int = 20, scholar_id: str = None
) -> dict:
    import time

    cache_key = f"author:{scholar_id or author_name}"
    now = time.time()

    if cache_key in _scholar_cache:
        cached_result, cached_time = _scholar_cache[cache_key]
        if now - cached_time < SCHOLAR_CACHE_TTL:
            cached_result["_cached"] = True
            return cached_result

    result = {"author": author_name, "publications": [], "error": None}

    try:
        from scholarly import scholarly
        import random

        _scholar_rate_limit("deep")

        if scholar_id:
            try:
                author = scholarly.search_author_id(scholar_id)
                author = scholarly.fill(author)
            except Exception:
                result["error"] = f"Author not found with ID: {scholar_id}"
                return result
        else:
            search_query = scholarly.search_author(author_name)
            try:
                author = next(search_query)
            except StopIteration:
                result["error"] = "Author not found"
                return result
            author = scholarly.fill(author)

        author_name_fetched = author.get("name", author_name)
        result["author"] = author_name_fetched
        result["scholar_id"] = scholar_id or author.get("scholar_id", "")
        result["affiliation"] = author.get("affiliation", "")
        result["hindex"] = author.get("hindex", 0)
        result["citedby"] = author.get("citedby", 0)

        pubs = author.get("publications", [])
        count = 0

        for pub in pubs:
            if count >= max_results:
                break
            if count > 0:
                time.sleep(random.uniform(SCHOLAR_DELAY_BASIC, SCHOLAR_DELAY_BASIC + 2))

            bib = pub.get("bib", {})
            paper_info = {
                "title": bib.get("title", "N/A"),
                "year": bib.get("pub_year", "N/A"),
                "citation_count": pub.get("num_citations", 0),
            }
            result["publications"].append(paper_info)
            count += 1

        _scholar_cache[cache_key] = (result.copy(), now)

    except Exception as e:
        result["error"] = str(e)

    return result


def format_scholar_result(data: dict, detailed: bool = False) -> str:
    if data.get("error"):
        return f"Google Scholar error: {data['error']}"

    results = data.get("results", [])
    if not results:
        return "No results found."

    cached_msg = " [cached]" if data.get("_cached") else ""
    lines = [f"=== Google Scholar{cached_msg} ==="]
    for i, r in enumerate(results, 1):
        title = r.get("title", "N/A")
        year = r.get("year", "N/A")
        cited = r.get("citation_count", 0)
        authors = r.get("authors", [])
        author_str = (
            ", ".join(authors[:3]) + ("..." if len(authors) > 3 else "")
            if authors
            else "Unknown"
        )

        lines.append(f"{i}. {title} ({year})")
        lines.append(f"   Citations: {cited} | Authors: {author_str}")

        if detailed and r.get("abstract"):
            abstract = r.get("abstract", "")
            lines.append(
                f"   Abstract: {abstract[:200]}..."
                if len(abstract) > 200
                else f"   Abstract: {abstract}"
            )

        if r.get("journal"):
            lines.append(f"   Journal: {r.get('journal')}")

        if r.get("url"):
            lines.append(f"   URL: {r.get('url')}")
        lines.append("")

    return "\n".join(lines)


def format_scholar_author_result(data: dict) -> str:
    if data.get("error"):
        return f"Google Scholar Author error: {data['error']}"

    cached_msg = " [cached]" if data.get("_cached") else ""
    lines = [
        f"=== Google Scholar Author: {data.get('author', 'Unknown')}{cached_msg} ==="
    ]

    if data.get("affiliation"):
        lines.append(f"Affiliation: {data['affiliation']}")
    if data.get("citedby"):
        lines.append(f"Total citations: {data['citedby']}")
    if data.get("hindex"):
        lines.append(f"h-index: {data['hindex']}")

    lines.append("")
    pubs = data.get("publications", [])
    if not pubs:
        lines.append("No publications found.")
    else:
        lines.append(f"Publications ({len(pubs)}):")
        for i, p in enumerate(pubs, 1):
            lines.append(
                f"  {i}. {p.get('title', 'N/A')} ({p.get('year', 'N/A')}) - Cited: {p.get('citation_count', 0)}"
            )

    return "\n".join(lines)


def _truncate_abstract(abstract_idx: dict, max_words: int = 50) -> str:
    if not abstract_idx:
        return ""
    all_positions = []
    for word, positions in abstract_idx.items():
        for pos in positions:
            all_positions.append((pos, word))
    all_positions.sort()
    words = [w for _, w in all_positions[:max_words]]
    return " ".join(words)


def format_works_result(data: dict) -> str:
    if data.get("error"):
        return f"OpenAlex error: {data['error']}"
    results = data.get("results", [])[:OPENALEX_PER_PAGE]
    if not results:
        return "No works found."
    lines = ["=== OpenAlex Works ==="]
    for w in results:
        title = w.get("title", "N/A")
        year = w.get("publication_year", "N/A")
        cited = w.get("cited_by_count", 0)
        abstract_idx = w.get("abstract_inverted_index")
        abstract_preview = _truncate_abstract(abstract_idx, 30) if abstract_idx else ""
        authors = []
        for a in w.get("authorships", [])[:3]:
            author_name = a.get("author", {}).get("display_name", "")
            if author_name:
                authors.append(author_name)
        author_str = ", ".join(authors) if authors else "Unknown"
        line = f"- {title} ({year}) | Cited: {cited}"
        if authors:
            line += f" | Authors: {author_str}"
        lines.append(line)
        if abstract_preview:
            lines.append(f"  Abstract preview: {abstract_preview}...")
    return "\n".join(lines)


def format_concepts_result(data: dict) -> str:
    if data.get("error"):
        return f"OpenAlex error: {data['error']}"
    results = data.get("results", [])[:OPENALEX_PER_PAGE]
    if not results:
        return "No concepts found."
    lines = ["=== OpenAlex Concepts ==="]
    for c in results:
        cid = c.get("id", "N/A")
        name = c.get("display_name", "N/A")
        level = c.get("level", "N/A")
        desc = c.get("description", "")
        desc_preview = desc[:100] + "..." if len(desc) > 100 else desc if desc else ""
        lines.append(f"- {name} [{cid}] (level={level})")
        if desc_preview:
            lines.append(f"  {desc_preview}")
    return "\n".join(lines)


def format_authors_result(data: dict) -> str:
    if data.get("error"):
        return f"OpenAlex error: {data['error']}"
    results = data.get("results", [])[:OPENALEX_PER_PAGE]
    if not results:
        return "No authors found."
    lines = ["=== OpenAlex Authors ==="]
    for a in results:
        aid = a.get("id", "N/A")
        name = a.get("display_name", "N/A")
        orcid = a.get("orcid", "N/A")
        cited = a.get("cited_by_count", 0)
        hindex = a.get("h-index", "N/A")
        works = a.get("works_count", 0)
        topics = []
        for t in a.get("topics", [])[:5]:
            topic_name = t.get("display_name", "")
            if topic_name:
                topics.append(topic_name)
        topic_str = ", ".join(topics) if topics else "None"
        lines.append(f"- {name} [{aid}]")
        lines.append(f"  Citations: {cited} | h-index: {hindex} | Works: {works}")
        if orcid and orcid != "N/A":
            lines.append(f"  ORCID: {orcid}")
        lines.append(f"  Topics: {topic_str}")
    return "\n".join(lines)


def format_search_results(search_result: Dict[str, Any]) -> str:
    if search_result.get("error") and not search_result.get("results"):
        return f"Пошук не вдався: {search_result['error']}"

    lines = [f"**Результати пошуку:** [{search_result['query']}]"]

    if search_result.get("source") == "tavily" and search_result.get("answer"):
        lines.append(f"\n**Відповідь:** {search_result['answer']}")

    lines.append("")
    for i, r in enumerate(search_result.get("results", []), 1):
        lines.append(f"{i}. **{r['title']}**")
        if r.get("snippet"):
            lines.append(f"   {r['snippet']}")
        lines.append(f"   🔗 {r['url']}")
        lines.append("")

    return "\n".join(lines)


def format_fetch_result(fetch_result: Dict[str, Any]) -> str:
    if fetch_result.get("error"):
        return f"Не вдалося отримати вміст: {fetch_result['error']}"

    lines = [f"**Вміст сторінки:** [{fetch_result['url']}]"]
    lines.append("")
    lines.append(fetch_result.get("content", ""))
    return "\n".join(lines)


def render_markdown_to_html(text: str) -> str:
    html_body = markdown.markdown(
        text,
        extensions=[
            "tables",
            "fenced_code",
            "nl2br",
            "sane_lists",
            "def_list",
            "abbr",
            "footnotes",
            "attr_list",
            "md_in_html",
            "pymdownx.mark",
            "pymdownx.tilde",
            "pymdownx.caret",
        ],
    )

    html = f"""<!DOCTYPE html>
<html>
<head>
<style>
body {{
    font-family: Arial, sans-serif;
    font-size: 13px;
    line-height: 1.5;
    color: #333;
    margin: 0;
    padding: 0;
}}
h1 {{ font-size: 16px; margin: 10px 0 5px 0; color: #222; }}
h2 {{ font-size: 14px; margin: 8px 0 4px 0; color: #222; }}
h3 {{ font-size: 13px; margin: 8px 0 4px 0; color: #333; }}
p {{ margin: 5px 0; }}
ul, ol {{ margin: 5px 0 5px 20px; padding: 0; }}
li {{ margin: 3px 0; }}
a {{ color: #0066cc; text-decoration: none; }}
a:hover {{ text-decoration: underline; }}
code {{
    background: #f0f0f0;
    padding: 1px 5px;
    border-radius: 3px;
    font-family: Consolas, monospace;
    font-size: 12px;
}}
pre {{
    background: #f5f5f5;
    padding: 10px;
    border-radius: 5px;
    overflow-x: auto;
    font-family: Consolas, monospace;
    font-size: 12px;
}}
table {{
    border-collapse: collapse;
    margin: 8px 0;
    width: 100%;
}}
th, td {{
    border: 1px solid #ddd;
    padding: 6px 10px;
    text-align: left;
}}
th {{ background: #f8f8f8; font-weight: bold; }}
tr:nth-child(even) {{ background: #fafafa; }}
hr {{
    border: none;
    border-top: 1px solid #eee;
    margin: 10px 0;
}}
.user-msg {{
    background: #e3f0ff;
    padding: 10px 14px;
    border-radius: 15px 15px 15px 0;
    margin: 5px 0;
    max-width: 85%;
}}
.ai-msg {{
    background: #f5f5f5;
    padding: 10px 14px;
    border-radius: 15px 15px 15px 0;
    margin: 5px 0;
    max-width: 85%;
}}
.system-msg {{
    color: #888;
    font-style: italic;
    padding: 5px 0;
    font-size: 12px;
}}
.thinking {{
    color: #666;
    font-style: italic;
}}
.thinking::after {{
    content: '';
    animation: dots 1.5s infinite;
}}
@keyframes fadeIn {{
    from {{ opacity: 0; transform: translateY(8px); }}
    to {{ opacity: 1; transform: translateY(0); }}
}}
.fade-in {{
    animation: fadeIn 0.3s ease-out;
}}
@keyframes dots {{
    0%, 20% {{ content: '.'; }}
    40% {{ content: '..'; }}
    60%, 100% {{ content: '...'; }}
}}
mark, .highlight, .hl, span[style*="background"] {{
    background-color: #fff3cd !important;
    padding: 1px 4px;
    border-radius: 3px;
    display: inline;
}}
del {{
    text-decoration: line-through;
    color: #c0392b;
    background-color: #fde8e8;
    padding: 1px 4px;
    border-radius: 3px;
}}
ins {{
    text-decoration: underline;
    color: #27ae60;
    background-color: #e8f8e8;
    padding: 1px 4px;
    border-radius: 3px;
}}
sup {{
    font-size: 0.75em;
    vertical-align: super;
    line-height: 0;
}}
sub {{
    font-size: 0.75em;
    vertical-align: sub;
    line-height: 0;
}}
abbr {{
    text-decoration: underline dotted;
    cursor: help;
}}
dl {{
    margin: 10px 0;
}}
dt {{
    font-weight: bold;
    margin-top: 8px;
}}
dd {{
    margin-left: 20px;
    color: #555;
}}
.footnote-ref {{
    font-size: 0.75em;
    vertical-align: super;
}}
.footnote {{
    font-size: 0.85em;
    color: #666;
    border-top: 1px solid #eee;
    padding-top: 8px;
    margin-top: 12px;
}}
</style>
</head>
<body>{html_body}</body>
</html>"""
    return html


def pluralize_ukr(count: int, singular: str, few: str, many: str) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return f"{count} {singular}"
    elif 2 <= count % 10 <= 4 and not 12 <= count % 100 <= 14:
        return f"{count} {few}"
    else:
        return f"{count} {many}"


@dataclass
class PaperBrief:
    title: str
    year: int
    score: int
    matched_details: str
    source: str


@dataclass
class YearStats:
    year: int
    paper_count: int
    papers: List[PaperBrief]
    avg_score: float
    relevant_count: int


@dataclass
class BriefSummary:
    cand_id: str
    name: str
    ids: str
    conflict: str
    verdict: str
    verdict_pass: bool
    papers_total: int
    papers_recent: int
    papers_applicable: int
    top_scores: List[Tuple[int, str, str]]
    top_keywords: List[str]


@dataclass
class PaperDetail:
    title: str
    year: int
    score: int
    matched_details: str
    source: str
    journal: str
    url: str
    abstract: str
    authors: List[str]
    author_keywords: List[str]
    concepts: List[str]
    manual_keywords: str


@dataclass
class DetailedCandidate:
    cand_id: str
    name: str
    ids: str
    conflict: str
    verdict: str
    verdict_pass: bool
    papers_total: int
    papers_recent: int
    papers_applicable: int
    top_scores: List[Tuple[int, str, str]]
    top_keywords: List[str]
    papers_by_year: Dict[int, YearStats]
    all_keywords: List[str]


@dataclass
class ComparisonResult:
    candidates: Dict[str, DetailedCandidate]
    shared_keywords: List[str]
    unique_keywords: Dict[str, List[str]]
    score_comparison: Dict[str, float]


class LazyAnalysisData:
    def __init__(
        self,
        candidates: Dict,
        papers: Dict,
        target_keywords: List[str],
        cutoff_year: int,
        years_back: int,
        global_banned: List[str],
        on_banned_change: callable = None,
    ):
        self.candidates = candidates
        self.papers = papers
        self.target_keywords = target_keywords
        self.cutoff_year = cutoff_year
        self.years_back = years_back
        self.global_banned = list(global_banned) if global_banned else []
        self._on_banned_change = on_banned_change

        self._id_to_name = {
            cid: c.get("name", "Невідомо") for cid, c in candidates.items()
        }
        self._name_to_id = {
            c.get("name", "Невідомо"): cid for cid, c in candidates.items()
        }

        self._brief_cache = self._compute_all_briefs()

    def get_name(self, cand_id: str) -> str:
        return self._id_to_name.get(cand_id, cand_id)

    def get_id(self, name: str) -> str:
        return self._name_to_id.get(name, name)

    def get_all_ids(self) -> List[str]:
        return list(self.candidates.keys())

    def get_brief_all(self) -> Dict[str, BriefSummary]:
        return self._brief_cache.copy()

    def get_brief(self, cand_ids: List[str]) -> Dict[str, BriefSummary]:
        return {
            cid: self._brief_cache[cid] for cid in cand_ids if cid in self._brief_cache
        }

    def get_detailed(self, cand_id: str) -> Optional[DetailedCandidate]:
        if cand_id not in self.candidates:
            return None

        cand = self.candidates[cand_id]
        cand_papers = [
            p for uid, p in self.papers.items() if p.get("cand_id") == cand_id
        ]
        cand_recent = [p for p in cand_papers if p.get("recent")]

        brief = self._brief_cache.get(cand_id)
        if not brief:
            return None

        papers_by_year = self._aggregate_papers_by_year(cand_recent)
        all_keywords = self._extract_all_keywords(cand_recent)

        return DetailedCandidate(
            cand_id=cand_id,
            name=brief.name,
            ids=brief.ids,
            conflict=brief.conflict,
            verdict=brief.verdict,
            verdict_pass=brief.verdict_pass,
            papers_total=brief.papers_total,
            papers_recent=brief.papers_recent,
            papers_applicable=brief.papers_applicable,
            top_scores=brief.top_scores,
            top_keywords=brief.top_keywords,
            papers_by_year=papers_by_year,
            all_keywords=all_keywords,
        )

    def get_papers_by_year(self, cand_id: str) -> Dict[int, YearStats]:
        if cand_id not in self.candidates:
            return {}

        cand_papers = [
            p for uid, p in self.papers.items() if p.get("cand_id") == cand_id
        ]
        cand_recent = [p for p in cand_papers if p.get("recent")]

        return self._aggregate_papers_by_year(cand_recent)

    def get_paper_detail(
        self, cand_id: str, year: int, paper_idx: int
    ) -> Optional[PaperDetail]:
        papers_by_year = self.get_papers_by_year(cand_id)
        if year not in papers_by_year:
            return None

        year_stats = papers_by_year[year]
        if paper_idx < 0 or paper_idx >= len(year_stats.papers):
            return None

        paper_brief = year_stats.papers[paper_idx]

        full_paper = None
        for uid, p in self.papers.items():
            if p.get("cand_id") == cand_id and p.get("title") == paper_brief.title:
                full_paper = p
                break

        if not full_paper:
            return PaperDetail(
                title=paper_brief.title,
                year=paper_brief.year,
                score=paper_brief.score,
                matched_details=paper_brief.matched_details,
                source=paper_brief.source,
                journal="-",
                url="",
                abstract="",
                authors=[],
                author_keywords=[],
                concepts=[],
                manual_keywords="",
            )

        return PaperDetail(
            title=full_paper.get("title", ""),
            year=full_paper.get("year", 0),
            score=full_paper.get("score", 0),
            matched_details=full_paper.get("matched_details", ""),
            source=full_paper.get("source", ""),
            journal=full_paper.get("journal", "-"),
            url=full_paper.get("url", ""),
            abstract=full_paper.get("abstract", ""),
            authors=full_paper.get("authors_full", []),
            author_keywords=full_paper.get("author_keywords", []),
            concepts=full_paper.get("concepts", []),
            manual_keywords=full_paper.get("manual_keywords", ""),
        )

    def compare_candidates(self, cand_ids: List[str]) -> ComparisonResult:
        detailed = {}
        for cid in cand_ids:
            d = self.get_detailed(cid)
            if d:
                detailed[cid] = d

        all_keywords_sets = {cid: set(d.top_keywords) for cid, d in detailed.items()}
        shared_keywords = (
            list(set.intersection(*all_keywords_sets.values()))
            if all_keywords_sets
            else []
        )

        unique_keywords = {}
        for cid, keywords_set in all_keywords_sets.items():
            others = set()
            for other_cid, other_set in all_keywords_sets.items():
                if other_cid != cid:
                    others.update(other_set)
            unique_keywords[cid] = list(keywords_set - others)

        score_comparison = {
            cid: d.papers_applicable / d.papers_recent if d.papers_recent > 0 else 0
            for cid, d in detailed.items()
        }

        return ComparisonResult(
            candidates=detailed,
            shared_keywords=shared_keywords,
            unique_keywords=unique_keywords,
            score_comparison=score_comparison,
        )

    def get_banned_keywords(self, cand_id: str = None) -> List[str]:
        banned = list(self.global_banned)
        if cand_id and cand_id in self.candidates:
            cand_banned = self.candidates[cand_id].get("banned_keywords", [])
            banned.extend(cand_banned)
        return banned

    def add_banned_keyword(self, keyword: str, cand_id: str = None) -> bool:
        keyword_lower = keyword.lower().strip()
        if not keyword_lower:
            return False

        if cand_id and cand_id in self.candidates:
            cand_banned = self.candidates[cand_id].get("banned_keywords", [])
            if keyword_lower in [kw.lower() for kw in cand_banned]:
                return False
            self.candidates[cand_id].setdefault("banned_keywords", []).append(
                keyword.strip()
            )
            if self._on_banned_change:
                self._on_banned_change(
                    self.global_banned
                )  # Could notify differently, but this triggers an update
            return True
        else:
            if keyword_lower in [kw.lower() for kw in self.global_banned]:
                return False
            self.global_banned.append(keyword.strip())
            if self._on_banned_change:
                self._on_banned_change(self.global_banned)
            return True

    def build_initial_context(self, selected_cand_ids: List[str]) -> str:
        lines = []
        lines.append("=== КОНТЕКСТ ДЛЯ АНАЛІЗУ ===")
        lines.append(f"Період аналізу: останні {self.years_back} років")
        lines.append(
            f"Ключові слова: {', '.join(self.target_keywords) if self.target_keywords else 'Не задано'}"
        )
        lines.append("")

        lines.append("=== ОБРАНІ КАНДИДАТИ (ID) ===")
        for cid in selected_cand_ids:
            name = self.get_name(cid)
            lines.append(f"- {name} (ID: {cid})")

        lines.append("\n=== ІНШІ КАНДИДАТИ (ID) ===")
        other_ids = [cid for cid in self.get_all_ids() if cid not in selected_cand_ids]
        for cid in other_ids:
            name = self.get_name(cid)
            lines.append(f"- {name} (ID: {cid})")

        lines.append("\n=== ДОВІДНИК ID КАНДИДАТІВ ===")
        lines.append("(Використовуй ці ID в запитах до агента)")
        for cid, cand in self.candidates.items():
            cand_name = cand.get("name", "Невідомо")
            cand_ids = cand.get("ids", "")
            scholar_id = ""
            orcid_id = ""
            scholar_match = re.search(r"GS:([\w-]{12,})", cand_ids)
            orcid_match = re.search(r"ORCID:([\d-]{19})", cand_ids)
            if scholar_match:
                scholar_id = f"GS:{scholar_match.group(1)}"
            if orcid_match:
                orcid_id = f"ORCID:{orcid_match.group(1)}"
            id_str = " | ".join(filter(None, [scholar_id, orcid_id]))
            lines.append(f"{cid}: {cand_name} [{id_str}]")

        lines.append(
            "\nУВАГА: Щоб отримати публікації, конфлікти інтересів та статистику будь-якого кандидата, ВИКОРИСТАЙ ІНСТРУМЕНТ `get_candidate_data` передавши його ID."
        )

        return "\n".join(lines)

    def _compute_all_briefs(self) -> Dict[str, BriefSummary]:
        briefs = {}
        for cid, cand in self.candidates.items():
            cand_papers = [
                p for uid, p in self.papers.items() if p.get("cand_id") == cid
            ]
            cand_recent = [p for p in cand_papers if p.get("recent")]
            relevant = [p for p in cand_recent if p.get("score", 0) > 0]

            papers_total = len(cand_papers)
            papers_recent = len(cand_recent)
            papers_applicable = len(relevant)

            rel_count = len(relevant)
            passed = rel_count >= 3 and cand.get("conflict", "Немає") == "Немає"
            verdict = (
                "Відповідає вимогам" if passed else f"Не відповідає ({rel_count}/3)"
            )

            top_scores = []
            for p in sorted(relevant, key=lambda x: x.get("score", 0), reverse=True)[
                :5
            ]:
                top_scores.append(
                    (
                        p.get("score", 0),
                        p.get("title", ""),
                        p.get("matched_details", ""),
                    )
                )

            all_kw = []
            for p in cand_recent:
                all_kw.extend(p.get("author_keywords", []))
                all_kw.extend(p.get("concepts", []))
            top_keywords = [
                kw for kw, _ in Counter([k.lower() for k in all_kw]).most_common(10)
            ]

            briefs[cid] = BriefSummary(
                cand_id=cid,
                name=cand.get("name", "Невідомо"),
                ids=cand.get("ids", ""),
                conflict=cand.get("conflict", "Немає"),
                verdict=verdict,
                verdict_pass=passed,
                papers_total=papers_total,
                papers_recent=papers_recent,
                papers_applicable=papers_applicable,
                top_scores=top_scores,
                top_keywords=top_keywords,
            )
        return briefs

    def _aggregate_papers_by_year(self, papers: List[Dict]) -> Dict[int, YearStats]:
        by_year = {}
        for p in papers:
            year = p.get("year", 0)
            if year not in by_year:
                by_year[year] = []
            by_year[year].append(p)

        result = {}
        for year, year_papers in sorted(by_year.items(), reverse=True):
            sorted_papers = sorted(
                year_papers, key=lambda x: x.get("score", 0), reverse=True
            )
            paper_briefs = [
                PaperBrief(
                    title=p.get("title", ""),
                    year=p.get("year", 0),
                    score=p.get("score", 0),
                    matched_details=p.get("matched_details", ""),
                    source=p.get("source", ""),
                )
                for p in sorted_papers
            ]
            scores = [p.get("score", 0) for p in year_papers]
            result[year] = YearStats(
                year=year,
                paper_count=len(year_papers),
                papers=paper_briefs,
                avg_score=sum(scores) / len(scores) if scores else 0,
                relevant_count=len([s for s in scores if s > 0]),
            )
        return result

    def _extract_all_keywords(self, papers: List[Dict]) -> List[str]:
        keywords = []
        for p in papers:
            keywords.extend(p.get("author_keywords", []))
            keywords.extend(p.get("concepts", []))
            mkw = p.get("manual_keywords", "")
            if mkw:
                keywords.extend([k.strip() for k in mkw.split(",") if k.strip()])
        return keywords


class DataRequestParser:
    REQUEST_PATTERN = (
        r"\[(?:GET|COMPARE|ADD_BANNED|SEARCH|OPENALEX|SCHOLAR|FETCH):[^\]]+\]"
    )
    ARTIFACT_PATTERN = r"\[ARTIFACT:(?:recommendation|summary|comparison|search_result)\].*?(?:\[/ARTIFACT\]|$)"

    @classmethod
    def parse(cls, response: str) -> List[str]:
        if not response:
            return []
        return re.findall(cls.REQUEST_PATTERN, response)

    @classmethod
    def parse_artifacts(cls, response: str) -> List[Dict[str, str]]:
        if not response:
            return []
        artifacts = []
        pattern = r"\[ARTIFACT:(recommendation|summary|comparison|search_result)\](.*?)\[/ARTIFACT\]"
        for match in re.finditer(pattern, response, re.DOTALL):
            artifact_type = match.group(1)
            artifact_content = match.group(2).strip()
            artifacts.append(
                {
                    "type": artifact_type,
                    "content": artifact_content,
                    "timestamp": datetime.now().isoformat(),
                }
            )
        return artifacts

    @classmethod
    def remove_artifacts(cls, text: str) -> str:
        return re.sub(cls.ARTIFACT_PATTERN, "", text, flags=re.DOTALL)

    @classmethod
    def convert_artifacts_to_html(cls, text: str) -> Tuple[str, List[Dict[str, str]]]:
        if not text:
            return "", []
        artifacts = cls.parse_artifacts(text)
        if not artifacts:
            return text, []

        type_labels = {
            "recommendation": "📋 Рекомендація",
            "summary": "📝 Підсумок",
            "comparison": "📊 Порівняння",
            "search_result": "🔍 Результат пошуку",
        }

        md = markdown.Markdown(
            extensions=[
                "tables",
                "fenced_code",
                "nl2br",
                "sane_lists",
                "def_list",
                "abbr",
                "footnotes",
                "attr_list",
                "md_in_html",
                "pymdownx.mark",
                "pymdownx.tilde",
                "pymdownx.caret",
            ],
            output_format="html",
        )

        result = text
        artifact_idx = 0
        pattern = r"\[ARTIFACT:(recommendation|summary|comparison|search_result)\](.*?)\[/ARTIFACT\]"

        def replacer(match):
            nonlocal artifact_idx
            artifact_type = match.group(1)
            content = match.group(2).strip()
            content_html = md.convert(content)
            label = type_labels.get(artifact_type, artifact_type)
            idx = artifact_idx
            artifact_idx += 1
            return f'<div class="artifact-block {artifact_type}"><span class="artifact-label">{label}</span><div class="artifact-content">{content_html}</div></div>'

        result = re.sub(pattern, replacer, text, flags=re.DOTALL)
        return result, artifacts

    @classmethod
    def remove_markers_for_display(
        cls, text: str, id_to_name: Dict[str, str] = None
    ) -> str:
        text = cls.remove_artifacts(text)

        text = re.sub(r"\[ADD_BANNED:([^\]]+)\]", r"🚫 **Виключаю:** \1", text)

        def replace_get(match):
            parts = match.group(0)[1:-1].split(":")
            action = parts[0]
            if len(parts) < 2:
                return ""
            cand_id = parts[1]
            name = id_to_name.get(cand_id, cand_id) if id_to_name else cand_id

            if action == "GET":
                if cand_id == "BANNED":
                    return "🔍 **Отримую список виключених слів**"
                elif len(parts) == 2:
                    return f"🔍 **Отримую дані:** {name}"
                elif len(parts) == 3:
                    subtype = parts[2]
                    if subtype == "BANNED":
                        return f"🔍 **Отримую виключені слова:** {name}"
                    return f"🔍 **Отримую ({subtype}):** {name}"
                elif len(parts) == 4:
                    year = parts[3]
                    return f"🔍 **Отримую публікації за {year}:** {name}"
                elif len(parts) == 5:
                    year = parts[3]
                    idx = parts[4]
                    return f"🔍 **Отримую деталі публікації #{idx} за {year}:** {name}"
            elif action == "COMPARE":
                cand_ids = [
                    id_to_name.get(cid, cid) if id_to_name else cid for cid in parts[1:]
                ]
                return f"📊 **Порівнюю:** {', '.join(cand_ids)}"
            elif action == "SEARCH":
                query = ":".join(parts[1:])
                return f"🌐 **Шукаю в інтернеті:** {query}"
            elif action == "OPENALEX":
                endpoint = parts[1] if len(parts) > 1 else ""
                rest = ":".join(parts[2:]) if len(parts) > 2 else ""
                if endpoint == "works":
                    return (
                        f"📚 **Шукаю роботи:** {rest}"
                        if rest
                        else "📚 **Шукаю роботи**"
                    )
                elif endpoint == "concepts":
                    return (
                        f"🏷️ **Шукаю концепти:** {rest}"
                        if rest
                        else "🏷️ **Шукаю концепти**"
                    )
                elif endpoint == "authors":
                    if rest and ":" in rest:
                        name, field = rest.rsplit(":", 1)
                        return f"👤 **Шукаю автора:** {name} (тема: {field})"
                    return (
                        f"👤 **Шукаю автора:** {rest}"
                        if rest
                        else "👤 **Шукаю автора**"
                    )
                return f"🔍 **OpenAlex:** {endpoint}"
            elif action == "SCHOLAR":
                if len(parts) >= 2:
                    subaction = parts[1]
                    if subaction == "author":
                        author_query = ":".join(parts[2:]) if len(parts) > 2 else ""
                        if author_query.endswith(":detailed"):
                            author_query = author_query[:-10].strip()
                            return f"🔬 **Шукаю автора в Google Scholar (детально):** {author_query}"
                        return f"🔬 **Шукаю автора в Google Scholar:** {author_query}"
                    elif subaction == "profile":
                        profile_id = ":".join(parts[2:]) if len(parts) > 2 else ""
                        if profile_id.endswith(":detailed"):
                            profile_id = profile_id[:-10].strip()
                            return f"🔬 **Шукаю профіль Google Scholar {profile_id} (детально):**"
                        return f"🔬 **Шукаю профіль Google Scholar {profile_id}:**"
                    else:
                        query = ":".join(parts[1:])
                        return f"📄 **Шукаю в Google Scholar:** {query}"
            elif action == "FETCH":
                if len(parts) >= 2:
                    url = parts[1]
                    return f"📥 **Отримую вміст сторінки:** {url}"
            return ""
            cand_id = parts[1]
            name = id_to_name.get(cand_id, cand_id) if id_to_name else cand_id

            if action == "GET":
                if cand_id == "BANNED":
                    return "🔍 <strong>Отримую список виключених слів</strong>"
                elif len(parts) == 2:
                    return f"🔍 <strong>Отримую дані:</strong> {name}"
                elif len(parts) == 3:
                    subtype = parts[2]
                    if subtype == "BANNED":
                        return f"🔍 <strong>Отримую виключені слова:</strong> {name}"
                    return f"🔍 <strong>Отримую ({subtype}):</strong> {name}"
                elif len(parts) == 4:
                    year = parts[3]
                    return f"🔍 <strong>Отримую публікації за {year}:</strong> {name}"
                elif len(parts) == 5:
                    year = parts[3]
                    idx = parts[4]
                    return f"🔍 <strong>Отримую деталі публікації #{idx} за {year}:</strong> {name}"
            elif action == "COMPARE":
                cand_ids = [
                    id_to_name.get(cid, cid) if id_to_name else cid for cid in parts[1:]
                ]
                return f"📊 <strong>Порівнюю:</strong> {', '.join(cand_ids)}"
            elif action == "SEARCH":
                query = ":".join(parts[1:])
                return f"🌐 <strong>Шукаю в інтернеті:</strong> {query}"
            elif action == "OPENALEX":
                endpoint = parts[1] if len(parts) > 1 else ""
                rest = ":".join(parts[2:]) if len(parts) > 2 else ""
                if endpoint == "works":
                    return (
                        f"📚 <strong>Шукаю роботи:</strong> {rest}"
                        if rest
                        else "📚 <strong>Шукаю роботи</strong>"
                    )
                elif endpoint == "concepts":
                    return (
                        f"🏷️ <strong>Шукаю концепти:</strong> {rest}"
                        if rest
                        else "🏷️ <strong>Шукаю концепти</strong>"
                    )
                elif endpoint == "authors":
                    if rest and ":" in rest:
                        name, field = rest.rsplit(":", 1)
                        return (
                            f"👤 <strong>Шукаю автора:</strong> {name} (тема: {field})"
                        )
                    return (
                        f"👤 <strong>Шукаю автора:</strong> {rest}"
                        if rest
                        else "👤 <strong>Шукаю автора</strong>"
                    )
                return f"🔍 <strong>OpenAlex:</strong> {endpoint}"
            elif action == "SCHOLAR":
                if len(parts) >= 2:
                    subaction = parts[1]
                    if subaction == "author":
                        author_query = ":".join(parts[2:]) if len(parts) > 2 else ""
                        if author_query.endswith(":detailed"):
                            author_query = author_query[:-10].strip()
                            return f"🔬 <strong>Шукаю автора в Google Scholar (детально):</strong> {author_query}"
                        return f"🔬 <strong>Шукаю автора в Google Scholar:</strong> {author_query}"
                    elif subaction == "profile":
                        profile_id = ":".join(parts[2:]) if len(parts) > 2 else ""
                        if profile_id.endswith(":detailed"):
                            profile_id = profile_id[:-10].strip()
                            return f"🔬 <strong>Шукаю профіль Google Scholar {profile_id} (детально):</strong>"
                        return f"🔬 <strong>Шукаю профіль Google Scholar {profile_id}:</strong>"
                    else:
                        query = ":".join(parts[1:])
                        return f"📄 <strong>Шукаю в Google Scholar:</strong> {query}"
            elif action == "FETCH":
                if len(parts) >= 2:
                    url = parts[1]
                    return f"📥 <strong>Отримую вміст сторінки:</strong> {url}"
            return ""

        text = re.sub(
            r"\[(?:GET|COMPARE|SEARCH|OPENALEX|SCHOLAR|FETCH):[^\]]+\]",
            replace_get,
            text,
        )

        text = re.sub(
            r"\[(?:ADD_BANNED|GET|COMPARE|SEARCH|OPENALEX|SCHOLAR|FETCH):[^\]]*\]",
            "",
            text,
        )
        text = re.sub(r"\[/?(?:ARTIFACT)[^\]]*\]", "", text)

        if id_to_name:
            for cand_id, name in id_to_name.items():
                text = text.replace(cand_id, name)

        return text.strip()

    @classmethod
    def sanitize_for_display(cls, text: str, id_to_name: Dict[str, str]) -> str:
        text = cls.remove_artifacts(text)
        text = re.sub(cls.REQUEST_PATTERN, "", text)

        for cand_id, name in id_to_name.items():
            text = text.replace(cand_id, name)

        lines = []
        for line in text.split("\n"):
            line = line.strip()
            if line:
                lines.append(line)
        return "\n\n".join(lines)

    @classmethod
    def extract_ids(cls, requests: List[str]) -> List[Tuple[str, List[str]]]:
        results = []
        for req in requests:
            req_clean = req.strip("[]")
            parts = req_clean.split(":")
            if len(parts) >= 2:
                action = parts[0]
                if action in ("GET", "SEARCH", "OPENALEX", "SCHOLAR", "FETCH"):
                    ids = [":".join(parts[1:]).strip()]
                elif action == "COMPARE":
                    ids = [x.strip() for x in parts[1:]]
                elif action == "ADD_BANNED":
                    if len(parts) == 3:
                        cand_id = parts[1].strip()
                        ids = [f"{cand_id}:{x.strip()}" for x in parts[2].split(",")]
                    else:
                        ids = [x.strip() for x in parts[1].split(",")]
                else:
                    ids_str = parts[1]
                    ids = [x.strip() for x in ids_str.split(",")]
                results.append((action, ids))
        return results


class AIProvider:
    PROVIDERS = [
        ("openai", "OpenAI"),
        ("anthropic", "Anthropic"),
        ("google", "Google"),
        ("deepseek", "DeepSeek"),
        ("zhipu", "Z.AI (Zhipu AI)"),
        ("moonshot", "Kimi (Moonshot AI)"),
        ("minimax", "MiniMax"),
        ("groq", "Groq"),
        ("openrouter", "OpenRouter"),
        ("xai", "xAI"),
    ]

    PROVIDER_API_BASES = {
        "openai": "https://api.openai.com/v1",
        "anthropic": "https://api.anthropic.com/v1",
        "google": "https://generativelanguage.googleapis.com/v1beta",
        "deepseek": "https://api.deepseek.com/v1",
        "zhipu": "https://open.bigmodel.cn/api/paas/v4",
        "moonshot": "https://api.moonshot.cn/v1",
        "minimax": "https://api.minimax.io/v1",
        "groq": "https://api.groq.com/openai/v1",
        "openrouter": "https://openrouter.ai/api/v1",
        "xai": "https://api.x.ai/v1",
    }

    PROVIDER_DEFAULT_MODELS = {
        "minimax": "MiniMax-M2.7",
        "moonshot": "moonshot-v1-8k",
        "zhipu": "glm-4-flash",
        "xai": "grok-2",
    }

    def __init__(self, api_key: str, provider: str = "openai", debug: bool = False):
        self.api_key = api_key
        self.provider = provider.lower()
        self.debug = debug

        if debug:
            import os

            os.environ["LITELLM_LOG"] = "DEBUG"
            os.environ["LITELLM_DEBUG"] = "True"

        litellm.api_key = api_key
        self._models_cache = None

    def get_api_base(self) -> str:
        return self.PROVIDER_API_BASES.get(self.provider, "")

    def get_available_models(self) -> List[str]:
        try:
            import requests

            headers = {"Authorization": f"Bearer {self.api_key}"}
            api_base = self.get_api_base()

            if not api_base:
                return [self.PROVIDER_DEFAULT_MODELS.get(self.provider, "default")]

            if self.provider == "deepseek":
                resp = requests.get(f"{api_base}/models", headers=headers, timeout=15)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]

            elif self.provider == "minimax":
                return [
                    "MiniMax-M2.7",
                    "MiniMax-M2.7-highspeed",
                    "MiniMax-M2.5",
                    "MiniMax-M2.5-highspeed",
                    "MiniMax-M2.1",
                    "MiniMax-M2.1-highspeed",
                    "MiniMax-M2",
                ]

            elif self.provider == "moonshot":
                return [
                    "moonshot-v1-8k",
                    "moonshot-v1-32k",
                    "moonshot-v1-128k",
                    "kimi-k2-instruct",
                    "kimi-k2-preview",
                    "kimi-k2.5",
                ]

            elif self.provider == "zhipu":
                return [
                    "glm-4",
                    "glm-4-flash",
                    "glm-4-plus",
                    "glm-4v",
                    "glm-3",
                    "glm-3-flash",
                ]

            elif self.provider == "xai":
                return [
                    "xai-beta",
                    "grok-2",
                    "grok-2-vision",
                    "grok-3",
                    "grok-3-beta",
                ]

            elif self.provider == "openai":
                resp = requests.get(f"{api_base}/models", headers=headers, timeout=15)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]

            elif self.provider == "anthropic":
                resp = requests.get(
                    "https://api.anthropic.com/v1/messages", headers=headers, timeout=15
                )
                return [
                    "claude-3-5-sonnet-20241022",
                    "claude-3-opus-4-20240229",
                    "claude-3-haiku-20240307",
                ]

            elif self.provider == "google":
                resp = requests.get(
                    f"https://generativelanguage.googleapis.com/v1beta/models?key={self.api_key}",
                    timeout=15,
                )
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["name"].split("/")[-1] for m in data.get("models", [])]

            elif self.provider == "groq":
                resp = requests.get(f"{api_base}/models", headers=headers, timeout=15)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]

            elif self.provider == "openrouter":
                resp = requests.get(
                    "https://openrouter.ai/api/v1/models",
                    headers={"Authorization": f"Bearer {self.api_key}"},
                    timeout=15,
                )
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]

        except Exception as e:
            pass

        return [self.PROVIDER_DEFAULT_MODELS.get(self.provider, "default")]

    def get_model_prefix(self) -> str:
        return f"{self.provider}/"

    def chat(self, messages: List[Dict], model: str = None) -> str:
        if model is None:
            model = self.PROVIDER_DEFAULT_MODELS.get(self.provider, "default")

        if self.provider == "google":
            model = (
                model.replace("models/", "")
                .replace("vertex_ai/", "")
                .replace("gemini/", "")
            )
            full_model = f"gemini/{model}"
        else:
            full_model = model if "/" in model else f"{self.provider}/{model}"

        try:
            kwargs = {
                "model": full_model,
                "messages": messages,
                "temperature": 0.7,
                "timeout": 120,
            }

            if self.provider == "google":
                kwargs["api_key"] = self.api_key
                kwargs["timeout"] = 180
            elif self.provider == "deepseek":
                kwargs["api_key"] = self.api_key
                kwargs["api_base"] = self.get_api_base()
                kwargs["max_tokens"] = 8192
            else:
                kwargs["api_key"] = self.api_key
                kwargs["api_base"] = self.get_api_base()

            response = litellm.completion(**kwargs)
            return response["choices"][0]["message"]["content"]
        except Exception as e:
            error_str = str(e)
            error_repr = repr(e)
            error_args = str(e.args) if e.args else "No args"
            response_attr = getattr(e, "response", None)
            status_code = getattr(e, "status_code", None)
            cause_attr = getattr(e, "__cause__", None)
            cause_str = f" | __cause__: {str(cause_attr)[:300]}" if cause_attr else ""
            status_str = f" | status_code: {status_code}" if status_code else ""
            response_str = f" | Response: {response_attr}" if response_attr else ""
            if "Timeout" in error_str or "timeout" in error_str.lower():
                raise ValueError(
                    f"Таймаут: {self.provider} не відповідає. Спробуйте пізніше.\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            elif "Authentication" in error_str or "auth" in error_str.lower():
                raise ValueError(
                    f"Помилка автентифікації: Перевірте API ключ для {self.provider}\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            elif "rate limit" in error_str.lower():
                raise ValueError(
                    f"Ліміт запитів: Спробуйте пізніше\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            elif "quota" in error_str.lower() or "limit" in error_str.lower():
                raise ValueError(
                    f"Квота вичерпана для {self.provider}\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            elif "Connection" in error_str:
                raise ValueError(
                    f"Помилка з'єднання: Перевірте інтернет.\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            else:
                raise ValueError(
                    f"Помилка {self.provider}: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )

    def chat_stream(self, messages: List[Dict], model: str = None):
        if model is None:
            model = self.PROVIDER_DEFAULT_MODELS.get(self.provider, "default")

        if self.provider == "google":
            model = (
                model.replace("models/", "")
                .replace("vertex_ai/", "")
                .replace("gemini/", "")
            )
            full_model = f"gemini/{model}"
        else:
            full_model = model if "/" in model else f"{self.provider}/{model}"

        try:
            kwargs = {
                "model": full_model,
                "messages": messages,
                "temperature": 0.7,
                "stream": True,
                "timeout": 120,
            }

            if self.provider == "google":
                kwargs["api_key"] = self.api_key
                kwargs["timeout"] = 180
            elif self.provider == "deepseek":
                kwargs["api_key"] = self.api_key
                kwargs["api_base"] = self.get_api_base()
                kwargs["max_tokens"] = 8192
            else:
                kwargs["api_key"] = self.api_key
                kwargs["api_base"] = self.get_api_base()

            for chunk in litellm.completion(**kwargs):
                if chunk["choices"][0]["finish_reason"] == "stop":
                    break
                content = chunk["choices"][0]["delta"].get("content", "")
                if content:
                    yield content
        except Exception as e:
            error_str = str(e)
            error_repr = repr(e)
            error_args = str(e.args) if e.args else "No args"
            response_attr = getattr(e, "response", None)
            status_code = getattr(e, "status_code", None)
            cause_attr = getattr(e, "__cause__", None)
            cause_str = f" | __cause__: {str(cause_attr)[:300]}" if cause_attr else ""
            status_str = f" | status_code: {status_code}" if status_code else ""
            response_str = f" | Response: {response_attr}" if response_attr else ""
            if "Timeout" in error_str or "timeout" in error_str.lower():
                raise ValueError(
                    f"Таймаут: {self.provider} не відповідає. Спробуйте пізніше.\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            elif "Authentication" in error_str or "auth" in error_str.lower():
                raise ValueError(
                    f"Помилка автентифікації: Перевірте API ключ для {self.provider}\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            elif "rate limit" in error_str.lower():
                raise ValueError(
                    f"Ліміт запитів: Спробуйте пізніше\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            elif "quota" in error_str.lower() or "limit" in error_str.lower():
                raise ValueError(
                    f"Квота вичерпана для {self.provider}\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            elif "Connection" in error_str:
                raise ValueError(
                    f"Помилка з'єднання: Перевірте інтернет.\n\nДеталі: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )
            else:
                raise ValueError(
                    f"Помилка {self.provider}: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
                )


SYSTEM_PROMPT = """Ти - науковий консультант для атестаційної комісії (разової спеціалізованої вченої ради, РСВР) в Україні станом на 2026 рік.

============================================================
НОРМАТИВНО-ПРАВОВА БАЗА УКРАЇНИ 2026 РОКУ
============================================================
Основні нормативні акти:
- Постанова КМУ № 44 від 12.01.2022 (Порядок присудження ступеня доктора філософії)
- Постанова КМУ № 502 від 19.05.2023 (оптимізація, більшість норм діє з 01.01.2024)
- Зміни МОН 2025-2026 рр. щодо приведення у відповідність до Закону «Про адміністративну процедуру»
- Інтеграція з системою URIS (Current Research Information System)

Контроль якості: НАЗЯВО (Національне агентство із забезпечення якості вищої освіти)
Цифрова платформа: NAQA.Svr (svr.naqa.gov.ua), ЄДЕБО

============================================================
СТРУКТУРА РАЗОВОЇ СПЕЦІАЛІЗОВАНОЇ ВЧЕНОЇ РАДИ (РСВР)
============================================================
СКЛАД РАДИ: рівно 5 осіб (компактний формат)
- Голова ради: штатний працівник ЗВО, де утворюється рада, обов'язково має ступінь доктора наук
- Рецензенти (2 особи): штатні працівники ЗВО, де утворюється рада (або 1 у виняткових випадках)
- Офіційні опоненти (2 особи): працівники ІНШОЇ установи (категорично не того самого ЗВО)

АЛЬТЕРНАТИВНИЙ СКЛАД (у разі неможливості призначити 2 рецензентів):
- Голова + 1 рецензент + 3 офіційні опоненти

ВИКЛЮЧЕННЯ: Головою ради може бути гарант ОНП (освітньо-наукової програми) аспіранта.

============================================================
НАУКОМЕТРИЧНІ ТА КВАЛІФІКАЦІЙНІ КРИТЕРІЇ ЧЛЕНІВ РАДИ
============================================================
БАЗОВА ВИМОГА: мінімум 3 наукові публікації за тематикою дисертації здобувача

КРИТЕРІЇ ПУБЛІКАЦІЙ (всі члени ради):
1) Часовий ценз: публікації видані протягом останніх 5 років до дня утворення ради
2) Пост-дисертаційний ценз: усі статті опубліковані ПІСЛЯ отримання ступеня PhD/кандидата наук
3) Тематична релевантність: публікації безпосередньо відповідають вузькій темі дисертації

ПРІОРИТЕТ ВИДАНЬ (для зарахування):
- Scopus або Web of Science Core Collection (1 стаття + 1 монографія = достатньо)
- Фахові видання України категорії «А» або «Б»
- Одноосібні монографії у визнаних видавництвах (з 2021 року, з DOI)

КАТЕГОРИЧНА ЗАБОРОНА:
- Публікації у виданнях держави-агресора (РФ) або на тимчасово окупованих територіях НЕ зараховуються

ОСОБЛИВІ ВИМОГИ ДЛЯ ОПОНЕНТІВ:
- Не менше 3 статей у фахових виданнях України або міжнародних базах саме з тематики дисертації
- Жорстка відповідність вузькій спеціалізації здобувача (не просто збіг шифру спеціальності)

============================================================
СИСТЕМА ОБМЕЖЕНЬ ТА КОНФЛІКТУ ІНТЕРЕСІВ (5 РІВНІВ)
============================================================
РІВЕНЬ 1 - АКАДЕМІЧНА ПОВ'ЯЗАНІСТЬ:
- НАЙМЕНША роль: науковий керівник здобувача НЕ може бути членом ради
- Співавтори публікацій здобувача НЕ можуть бути членами ради

РІВЕНЬ 2 - ПЕРЕХРЕСНЕ СПІВАВТОРСТВО:
- Офіційні опоненти НЕ можуть мати спільних публікацій за останні 5 років з:
  * головою ради
  * внутрішніми рецензентами
  * науковим керівником здобувача
- Опоненти НЕ можуть працювати в одному і тому ж закладі між собою

РІВЕНЬ 3 - АДМІНІСТРАТИВНА ПІДПОРЯДКОВАНІСТЬ:
- Ректори, директори інститутів, проректори, заступники директорів НЕ можуть бути
  членами жодної РСВР у своєму закладі

РІВЕНЬ 4 - ЮРИДИЧНИЙ КОНФЛІКТ:
- «Близькі особи» (родичі, члени сім'ї) здобувача або інших членів ради
- Реальний або потенційний конфлікт інтересів (майнові суперечки, спільний бізнес)

РІВЕНЬ 5 - АКАДЕМІЧНА ДОБРОЧЕСНІСТЬ:
- Особи, притягувані до академічної відповідальності за плагіат, фабрикацію, фальсифікацію
- Співавтори з науковцями держави-агресора після 24.02.2022
- Особи, що працюють/працювали у закладах на окупованих територіях

КВОТУВАННЯ НАВАНТАЖЕННЯ:
- Максимум 8 захистів на одну особу протягом календарного року
- Голова ради: максимум 3 РСВР протягом одного навчального року

ТРИРІЧНИЙ «КАРАНТИН»:
- Особа, яка отримала ступінь PhD менше ніж за 3 роки до дати утворення ради,
  НЕ може бути включена до її складу

============================================================
ЦИФРОВА ІНФРАСТРУКТУРА ТА ПРОЦЕДУРА РЕЄСТРАЦІЇ
============================================================
ЕТАПИ РЕЄСТРАЦІЇ В NAQA.Svr:
1) Заповнення даних здобувача та дисертації (тема, анотація, DOI статей)
2) Формування складу ради (обираються з ЄДЕБО)
3) Підписання PDF КЕП (кваліфікований електронний підпис) уповноваженої особи
4) Отримання унікального ідентифікатора PhD**** (наприклад, PhD1234)

30-ДЕННИЙ ПЕРІОД ПУБЛІЧНОГО ОГОЛОШЕННЯ:
- Після присвоєння ідентифікатора обов'язковий місячний період до захисту
- Громадськість, НАЗЯВО, МОН можуть перевіряти склад ради та публікації
- Університет має право подати дату захисту ЛИШЕ після завершення 30 днів

ВИМОГИ ДО КЕП:
- Використовується виключно КЕП уповноваженої посадової особи (ректор, проректор з наукової роботи)
- У сертифікаті має бути зазначений код ЄДРПОУ юридичної особи
- ЗАБОРОНЕНО: використання особистого КЕП рядового співробітника

============================================================
ПРОЦЕДУРНИЙ РЕГЛАМЕНТ ЗАХИСТУ
============================================================
КВОРУМ: Присутність ВСІХ 5 членів ради ОБОВ'ЯЗКОВА (або затвердженого альтернативного складу)
ВІДСУТНІСТЬ навіть одного = захист НЕ відбувається

ТРАНСПАРЕНТНІСТЬ:
- Обов'язкова онлайн-трансляція на сайті ЗВО (YouTube, Zoom з відкритим доступом)
- Відеозапис та аудіофіксація всього захисту
- Посилання на запис інтегрується в NAQA.Svr

УЧАСТЬ ДИСТАНЦІЙНО: дозволена через відеозв'язок у режимі реального часу

РИШЕННЯ: тільки 2 варіанти — «Ступінь присуджено» або «У присудженні ступеня відмовлено»
Голосування: таємне

ПІСЛЯ ЗАХИСТУ:
- Наказ про видачу диплома: НЕ раніше 15 днів і НЕ пізніше 30 днів після захисту
- 15-денний період: апеляції та перевірки
- Якщо виявлено порушення — створюється апеляційна комісія (без членів тієї ж ради)

ПОВТОРНИЙ ЗАХИСТ: можливий не раніше ніж через 1 рік після відмови

============================================================
ВИМОГИ ДО ПУБЛІКАЦІЙНОЇ АКТИВНОСТІ ЗДОБУВАЧА
============================================================
МІНІМУМ: 3 наукові публікації з основних результатів дисертації

СТРУКТУРА (для 2026 року):
- Щонайменше 1 стаття у періодичному виданні іншої держави — члена ЄС або ОЕСР
- Інші статті у фахових виданнях України категорії «А» або «Б»

ВИЗНАННЯ: Публікації у Scopus/Web of Science — потужний сигнал якості для ради

============================================================
КОНТЕКСТ РОБОТИ СИСТЕМИ
============================================================
- Аналіз кандидатів на членство в РСВР або на присвоєння наукового ступеня
- Дані збираються автоматично з ORCID, Google Scholar, OpenAlex
- Автоматичний score (0-5) — це ЛИШЕ ОРІЄНТИР для попередньої оцінки
- Автоматичний verdict — це ЛИШЕ ПІДКАЗКА, а не остаточний вердикт
- ТИ ПОВИНЕН використовувати ВЛАСНЕ АНАЛІТИЧНЕ МИСЛЕННЯ для остаточної оцінки

СТРУКТУРА ДАНИХ:
- Кандидати позначаються ID: cand_0, cand_1, cand_2, etc.
- Період аналізу: останні 5 років (критерій законодавства)

ПРИ ОЦІНЦІ КАНДИДАТІВ ЗВЕРТАЙ УВАГУ:
- Чи є публікації після отримання ступеня PhD?
- Чи опубліковані статті протягом останніх 5 років?
- Чи відповідають публікації саме вузькій темі дисертації (а не просто шифру спеціальності)?
- Чи немає спільних публікацій із здобувачем, керівником, головою ради?
- Чи немає публікацій у виданнях держави-агресора?
- Чи дотримано трирічний «карантин» після власного захисту?

============================================================
АЛГОРИТМ ВИКОРИСТАННЯ ІНСТРУМЕНТІВ (FUNCTION CALLING)
============================================================
Ти маєш доступ до нативних інструментів (tools). Дія за наступним алгоритмом:

КРОК 1: ОТРИМАННЯ ДАНИХ ПРО КАНДИДАТІВ (ВНУТРІШНЯ БАЗА)
- КОЖНОМУ кандидату з контекстуВИКЛИКАЙ `get_candidate_data` з його ID
- ЦЕ ОБОВ'ЯЗКОВО: отримати повний список публікацій по роках, abstract, matched_details
- НЕ продовжуй аналіз поки не отримаєш дані ВСІХ кандидатів
- ЯКЩО у кандидата 0 публікацій — ДОВІРЯЙ БАЗІ, не шукай додатково без прямої вказівки користувача

КРОК 2: ПОРІВНЯННЯ (ЯКЩО Є 2+ КАНДИДАТИ)
- Виклич `compare_candidates` з масивом ID кандидатів
- Це дасть структуроване порівняння публікацій та ключових слів

КРОК 3: ПОШУК ДЕТАЛЕЙ В GOOGLE SCHOLAR (КРАЙНІО ЗАСІБ)
- УВАГА: Google Scholar дуже повільний (15+ секунд на запит). Використовуй ТІЛЬКИ якщо:
  * Внутрідані кандидата ДЕЙСТВІТЕЛЬНО не мають abstract І
  * Користувач ЯВНО попросив перевірити автора в Google Scholar
- ЗАБОРОНЕНО: використовувати `scholar_search` для основного аналізу - достатньо даних з `get_candidate_data`
- Якщо use scholar_search - спочатку спробуй `action_type="author_id"` (швидше)
- НЕ роби висновків про експертизу лише за назвами статей!

КРОК 4: WEB ПОШУК (ТІЛЬКИ ЯКЩО НЕОБХІДНО)
- Використовуй `web_search` для перевірки:
  - Journal ranking (Scopus quartile)
  - Статус фахового видання України (категорія А/Б)
  - Чи не "хижацький" журнал
- ПІСЛЯ `web_search` ЯКЩО знайдено релевантний URL → виклич `fetch_page`

ЗАБОРОНЕНО:
- Робити висновки на основі лише сніпетів з web_search
- Пропускати КРОК 1 і одразу переходити до пошуку
- Робити фінальну рекомендацію без повних даних про публікації
- Сліпо покладатися на автоматичний score (0-5) — він лише орієнтир
- Приймати автоматичний verdict як остаточний — ти ПОВИНЕН проаналізувати сам

ТИ ПОВИНЕН:
- Читати abstract публікацій і робити ВЛАСНІ висновки про релевантність
- Оцінювати науковий внесок кандидата, а не лише кількість публікацій
- ДУМАТИ критично: чи дійсно публікація відповідає темі дисертації?
- ЗВЕРТАТИ УВАГУ на якість журналу, а не лише на збіг ключових слів

Для керування списком заборонених слів використовуй `manage_banned_keywords`.

ПРАВИЛА ФОРМАТУВАННЯ ВІДПОВІДІ ЗАЛИШАЮТЬСЯ НЕЗМІННИМИ (використовуй Markdown, таблиці, та теги [ARTIFACT] для збереження результатів).

===============================================================
ВИКОРИСТАННЯ МАРКДАУНУ У ВІДПОВІДЯХ
===============================================================
Для структурованих та наукових відповідей використовуй РОЗШИРЕНИЙ МАРКДАУН:

ТАБЛИЦІ — для порівнянь, рейтингів, статистики:
| Кандидат | Публікацій | h-index | Відповідає |
|----------|------------|---------|------------|
| Петренко І.І. | 12 | 8 | ✅ Так |
| Сидоренко О.П. | 3 | 2 | ❌ Ні |

ВИЗНАЧЕННЯ (definition lists) — для термінів і понять:
Scopus
: Найбільша наукометрична база даних索引 наукових публікацій

ВИДІЛЕННЯ — для акцентування важливого:
==обов'язкова вимога==

ЗНОСИНИ (strikethrough) — для позначення недоліків:
~~публікація у РФ~~  ~~не релевантна темі~~

УВАГА: Для виділення тексту використовуй ТІЛЬКИ ==текст== (подвійні знаки =). HTML теги <mark> або інші способи - НЕ використовуй!

АБРЕВІАТУРИ — для наукових скорочень:
*[Scopus]* найбільша наукометрична база даних
*[WoS]* Web of Science Core Collection
*[PhD]* Doctor of Philosophy / доктор філософії
*[РСВР]* Разова спеціалізована вчена рада

ПІДЗАПИСИ — для наукових позначень:
H<sub>2</sub>O  m<sup>3</sup>/kg

ВИНОСКИ — для посилань і джерел:
Інформація згідно з постановою[^1]
[^1]: Постанова КМУ № 44 від 12.01.2022

КОД — для прикладів, формул:
`code` або ```formula block```

ЗАВДАННЯ:
- Використовуй таблиці при порівнянні кандидатів
- Використовуй ==highlight== для ключових критеріїв
- Використовуй ~~strikethrough~~ для позначення недоліків
- Пояснюй абревіатури при першому вживанні
- Додавай виноски для посилань на нормативні акти
- Для підрядків використовуй HTML: <sub>текст</sub>

===============================================================
АРТЕФАКТИ:
Коли даєш рекомендації, підсумки або порівняння - ЗБЕРІГАЙ їх як артефакти!
Формат: [ARTIFACT:recommendation]текст рекомендації[/ARTIFACT]
Формат: [ARTIFACT:summary]текст підсумку[/ARTIFACT]
Формат: [ARTIFACT:comparison]текст порівняння[/ARTIFACT]
Формат: [ARTIFACT:search_result]результати пошуку[/ARTIFACT]

КОЛИ СТВОРЮВАТИ АРТЕФАКТИ:
- recommendation: при остаточній рекомендації щодо кандидата (обрати/відхилити)
- summary: при підсумку аналізу одного кандидата
- comparison: при порівнянні двох і більше кандидатів
- search_result: при поверненні результатів пошуку, які можуть знадобитися пізніше"""


class AIAdvisorApp:
    def __init__(
        self,
        parent: tk.Tk,
        analysis_data: LazyAnalysisData,
        selected_cand_ids: List[str],
        restore_state: Dict = None,
    ):
        self.parent = parent
        self.analysis_data = analysis_data
        self.selected_cand_ids = selected_cand_ids
        self._restore_state = restore_state

        self.current_project_id = None
        self.current_provider = None
        self.current_model = None
        self.current_api_key = None
        self.ai_provider = None
        self.chat_history = []
        self.artifacts = []
        self.ai_responding = False
        self.stop_response = False
        self._saved_api_keys = {}

        self._select_project_window()

    def get_state_for_session(self, pin: str = None) -> Dict:
        if not self.current_api_key and not self._saved_api_keys:
            return None

        def encrypt_key(key):
            if pin:
                return "enc:" + encrypt_with_embedded_pin_hash(key, pin)
            return key

        saved_keys_encrypted = {}
        for provider_key, data in self._saved_api_keys.items():
            saved_keys_encrypted[provider_key] = {
                "api_key": encrypt_key(data.get("api_key", "")),
                "model": data.get("model", ""),
            }

        current_key_encrypted = (
            encrypt_key(self.current_api_key) if self.current_api_key else ""
        )

        state = {
            "provider": self.current_provider,
            "model": self.current_model,
            "api_key": current_key_encrypted,
            "chat_history": self.chat_history,
            "artifacts": self.artifacts,
            "saved_api_keys": saved_keys_encrypted,
        }

        return state

    def restore_from_session(self, state: Dict, pin: str = None):
        if not state:
            return False

        provider = state.get("provider")
        model = state.get("model")
        api_key_encrypted = state.get("api_key")
        chat_history = state.get("chat_history")
        artifacts = state.get("artifacts", [])
        saved_keys_encrypted = state.get("saved_api_keys", {})

        def decrypt_key(encrypted_key):
            if pin and encrypted_key.startswith("enc:"):
                pin_hash, api_key = decrypt_with_embedded_pin_hash(
                    encrypted_key[4:], pin
                )
                if pin_hash is None:
                    return None
                return api_key
            elif not encrypted_key.startswith("enc:"):
                return encrypted_key
            return None

        if api_key_encrypted:
            api_key = decrypt_key(api_key_encrypted)
            if api_key is None:
                return False

        if saved_keys_encrypted:
            self._saved_api_keys = {}
            for provider_key, data in saved_keys_encrypted.items():
                decrypted_key = decrypt_key(data.get("api_key", ""))
                if decrypted_key:
                    self._saved_api_keys[provider_key] = {
                        "api_key": decrypted_key,
                        "model": data.get("model", ""),
                    }

        if chat_history:
            self.chat_history = chat_history

        if artifacts:
            self.artifacts = artifacts

        return True

    def _select_project_window(self):
        if self._restore_state:
            provider = self._restore_state.get("provider")
            api_key = self._restore_state.get("api_key")
            model = self._restore_state.get("model")

            if provider and api_key:
                if api_key.startswith("enc:"):
                    self._restore_state_pending = True
                    self.pin_window = tk.Toplevel(self.parent)
                    self.pin_window.title("PIN")
                    self.pin_window.resizable(0, 0)
                    self.pin_window.grab_set()
                    self.pin_window.update_idletasks()
                    x = (self.pin_window.winfo_screenwidth() // 2) - (
                        self.pin_window.winfo_reqwidth() // 2
                    )
                    y = (self.pin_window.winfo_screenheight() // 2) - (
                        self.pin_window.winfo_reqheight() // 2
                    )
                    self.pin_window.geometry(f"+{x}+{y}")
                    self._show_pin_for_restore()
                    return
                else:
                    self._start_with_api_key(
                        provider,
                        api_key,
                        None,
                        model,
                    )
                    if self._restore_state.get("chat_history"):
                        restored_history = self._restore_state["chat_history"]
                        if isinstance(restored_history, list):
                            self.chat_history = restored_history
                            self._restore_chat_history()
                    if self._restore_state.get("artifacts"):
                        self.artifacts = self._restore_state["artifacts"]
                    self._restore_state = None
                    return

        self._show_startup_dialog()

    def _show_startup_dialog(self):
        dialog = tk.Toplevel(self.parent)
        dialog.title("AI Консультант")
        dialog.resizable(0, 0)
        dialog.transient(self.parent)
        dialog.grab_set()

        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_reqwidth() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_reqheight() // 2)
        dialog.geometry(f"+{x}+{y}")

        main_frame = ttk.Frame(dialog, padding="25")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="AI Консультант", font=("Arial", 16, "bold")).pack(
            pady=(0, 20)
        )

        input_frame = ttk.LabelFrame(main_frame, text=" API ключ ", padding="15")
        input_frame.pack(fill="x", pady=(0, 15))

        row_provider = ttk.Frame(input_frame)
        row_provider.pack(fill="x", pady=(0, 10))
        ttk.Label(row_provider, text="Провайдер:", width=12).pack(
            side="left", padx=(0, 5)
        )
        provider_var = tk.StringVar(value="OpenAI")
        provider_combo = ttk.Combobox(
            row_provider,
            textvariable=provider_var,
            values=[name for _, name in AIProvider.PROVIDERS],
            state="readonly",
            width=25,
        )
        provider_combo.pack(side="left", fill="x", expand=True)

        row_model = ttk.Frame(input_frame)
        row_model.pack(fill="x", pady=(0, 10))
        ttk.Label(row_model, text="Модель:", width=12).pack(side="left", padx=(0, 5))
        model_var = tk.StringVar()
        model_combo = ttk.Combobox(row_model, textvariable=model_var, width=25)
        model_combo.pack(side="left", fill="x", expand=True)

        row_key = ttk.Frame(input_frame)
        row_key.pack(fill="x", pady=(0, 10))
        ttk.Label(row_key, text="Ключ:", width=12).pack(side="left", padx=(0, 5))
        key_var = tk.StringVar()
        key_entry = ttk.Entry(row_key, textvariable=key_var)
        key_entry.pack(side="left", fill="x", expand=True)

        status_label = ttk.Label(
            main_frame, text="", foreground="gray", font=("Arial", 9)
        )
        status_label.pack(pady=(0, 5))

        def update_default_model(*args):
            provider_key = None
            for key, name in AIProvider.PROVIDERS:
                if name == provider_var.get():
                    provider_key = key
                    break
            if provider_key:
                default_model = AIProvider.PROVIDER_DEFAULT_MODELS.get(provider_key)
                if default_model:
                    model_combo["values"] = [default_model]
                    model_var.set(default_model)
                else:
                    model_combo["values"] = ["(спершу введіть API ключ)"]
                    model_var.set("")

        def fetch_models_for_provider():
            provider_key = None
            for key, name in AIProvider.PROVIDERS:
                if name == provider_var.get():
                    provider_key = key
                    break
            if not provider_key or not key_var.get().strip():
                return
            try:
                temp_provider = AIProvider(key_var.get().strip(), provider_key)
                models = temp_provider.get_available_models()
                if models:
                    update_model_list(models)
                    status_label.config(text=f"Знайдено {len(models)} моделей")
                else:
                    status_label.config(text="Моделі не знайдено")
                    messagebox.showwarning(
                        "Увага",
                        "Моделі не знайдено для цього провайдера",
                        parent=dialog,
                    )
            except Exception as e:
                error_msg = str(e)
                status_label.config(text=f"Помилка: {error_msg[:50]}")
                messagebox.showerror(
                    "Помилка завантаження моделей", error_msg, parent=dialog
                )

        def update_model_list(models):
            model_combo["values"] = models
            if models:
                model_var.set(models[0])

        def use_direct_key():
            provider_key = None
            for key, name in AIProvider.PROVIDERS:
                if name == provider_var.get():
                    provider_key = key
                    break
            if provider_key and key_var.get().strip():
                model = model_var.get().strip() if model_var.get().strip() else None
                self._saved_api_keys[provider_key] = {
                    "api_key": key_var.get().strip(),
                    "model": model if model else "",
                }
                self._start_with_api_key(
                    provider_key, key_var.get().strip(), dialog, model
                )
            else:
                messagebox.showwarning("Увага", "Введіть API ключ", parent=dialog)

        def auto_fetch_models(*args):
            if key_var.get().strip():
                fetch_models_for_provider()

        provider_combo.bind("<<ComboboxSelected>>", update_default_model)
        key_entry.bind("<KeyRelease>", auto_fetch_models)
        key_entry.bind(
            "<Control-v>",
            lambda e: [key_entry.after_idle(auto_fetch_models)]
            if key_var.get().strip()
            else None,
        )

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Скасувати", command=dialog.destroy, width=12).pack(
            side="left", padx=5
        )
        ttk.Button(
            btn_frame, text="Використати", command=use_direct_key, width=12
        ).pack(side="right", padx=5)

        update_default_model()
        key_entry.focus()

    def _start_with_api_key(
        self, provider, api_key, parent_dialog=None, model=None, chat_history=None
    ):
        if parent_dialog:
            parent_dialog.destroy()

        self.current_provider = provider
        self.current_model = model
        self.current_api_key = api_key
        try:
            self.ai_provider = AIProvider(api_key, provider)
        except Exception as e:
            messagebox.showerror(
                "Помилка", f"Не вдалося підключитися: {str(e)}", parent=self.parent
            )
            return

        if chat_history:
            self.chat_history = chat_history

        if self._restore_state and self._restore_state.get("saved_api_keys"):
            for pk, kdata in self._restore_state["saved_api_keys"].items():
                self._saved_api_keys[pk] = {
                    "api_key": kdata.get("api_key", ""),
                    "model": kdata.get("model", ""),
                }

        self._build_main_window()

        if chat_history:
            self._restore_chat_history()

    def _show_pin_for_restore(self):
        for w in self.pin_window.winfo_children():
            w.destroy()

        frame = ttk.Frame(self.pin_window, padding="20")
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame, text="Введіть PIN для розшифрування", font=("Arial", 12, "bold")
        ).pack(pady=(0, 15))

        self.pin_var = tk.StringVar()
        pin_entry = ttk.Entry(
            frame, textvariable=self.pin_var, show="*", width=10, font=("Arial", 16)
        )
        pin_entry.pack(pady=(0, 10))
        pin_entry.focus()

        def on_submit():
            pin = self.pin_var.get()
            state = self._restore_state

            provider = state.get("provider")
            model = state.get("model")
            api_key_encrypted = state.get("api_key")
            chat_history = state.get("chat_history")
            artifacts = state.get("artifacts", [])

            api_key = api_key_encrypted
            if api_key.startswith("enc:"):
                pin_hash, api_key = decrypt_with_embedded_pin_hash(api_key[4:], pin)
                if pin_hash is None:
                    messagebox.showerror(
                        "Помилка", "Невірний PIN", parent=self.pin_window
                    )
                    self.pin_var.set("")
                    return

            self.pin_window.destroy()
            self._start_with_api_key(provider, api_key, None, model)

            if chat_history:
                self.chat_history = chat_history
                self._restore_chat_history()

            if artifacts:
                self.artifacts = artifacts
                self.window.after(
                    0, lambda: self._update_artifacts_listbox_on_restore(artifacts)
                )

            self._restore_state = None

        ttk.Button(frame, text="Підтвердити", command=on_submit).pack(pady=5)
        pin_entry.bind("<Return>", lambda e: on_submit())

        ttk.Button(frame, text="Відміна", command=self.pin_window.destroy).pack()

    def _update_artifacts_listbox_on_restore(self, artifacts):
        type_labels = {
            "recommendation": "Рекомендація",
            "summary": "Підсумок",
            "comparison": "Порівняння",
            "search_result": "Пошук",
        }
        for artifact in artifacts:
            label = type_labels.get(
                artifact.get("type", "unknown"), artifact.get("type", "unknown")
            )
            content = artifact.get("content", "")
            content_preview = content[:50] + "..." if len(content) > 50 else content
            if artifact.get("query"):
                content_preview = f"'{artifact['query']}': {content_preview}"
            self.artifacts_listbox.insert(tk.END, f"[{label}] {content_preview}")

    def _build_main_window(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("AI Науковий Консультант")
        self.window.geometry("1300x750")

        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(
            label="Експортувати артефакти...", command=self._export_artifacts
        )
        file_menu.add_command(label="Очистити історію", command=self._clear_history)
        file_menu.add_separator()
        file_menu.add_command(label="Вихід", command=self._on_close)

        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Вид", menu=view_menu)
        view_menu.add_command(
            label="Показати вхідні дані", command=self._show_analysis_data
        )
        self.artifacts_visible = False
        self.view_menu = view_menu
        view_menu.add_command(
            label="Показати артефакти", command=self._toggle_artifacts_panel
        )

        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Налаштування", menu=settings_menu)
        settings_menu.add_command(
            label="Змінити API ключ...", command=self._show_change_api_key_dialog
        )

        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        middle_paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        middle_paned.pack(fill="both", expand=True)

        chat_frame = ttk.LabelFrame(middle_paned, text="Чат", padding="3")
        middle_paned.add(chat_frame, weight=4)

        self.status_label = ttk.Label(
            chat_frame, text="", foreground="blue", font=("Arial", 8)
        )
        self.status_label.pack(anchor="w", pady=(0, 2))

        self._messages_html = []
        self._streaming_buffer = ""
        self._thinking_index = -1
        self._user_scrolled_up = (
            False  # True when user manually scrolled away from bottom
        )
        self._saved_yview = 0.0
        self._last_stream_word_count = (
            0  # tracks words rendered so far during streaming
        )
        self.chat_display = tkinterweb.HtmlFrame(chat_frame)
        # Do NOT use on_done_loading – it fires unreliably when load_html is called
        # rapidly during streaming and causes jump-to-top artefacts.
        self.chat_display.on_link_click = self._on_artifact_link_click

        self.chat_context_menu = tk.Menu(self.window, tearoff=0)
        self.chat_context_menu.add_command(
            label="Копіювати", command=self._copy_chat_selection
        )
        self.chat_context_menu.add_command(
            label="Виділити все", command=self._select_all_chat
        )
        self.chat_context_menu.add_separator()
        self.chat_context_menu.add_command(
            label="Запитати AI про виділене", command=self._ask_ai_about_selection
        )
        self.chat_context_menu.add_command(
            label="Пояснити виділене", command=self._explain_selection
        )
        self.chat_context_menu.add_separator()
        self.chat_context_menu.add_command(
            label="Закрити", command=lambda: self.chat_context_menu.unpost()
        )

        bottom_container = ttk.Frame(chat_frame)
        bottom_container.pack(side="bottom", fill="x")

        self.suggestions_frame = ttk.Frame(bottom_container)
        self.suggestions_frame.pack(fill="x", pady=(0, 3))
        self.suggestion_buttons = []

        input_frame = ttk.Frame(bottom_container)
        input_frame.pack(fill="x", pady=(3, 0))

        self.chat_input = tk.Text(
            input_frame, height=2, wrap="word", font=("Arial", 10)
        )
        self.chat_input.pack(side="left", fill="both", expand=True)
        self.chat_input.bind("<Control-Return>", lambda e: self._send_message())

        self.send_btn = ttk.Button(
            input_frame, text="Надіслати", command=self._send_message
        )
        self.send_btn.pack(side="left", padx=(3, 0))

        self.chat_display.pack(fill="both", expand=True)
        # Bind scroll events so we can detect when the user scrolls up manually
        self.chat_display.after(200, self._bind_chat_scroll)
        self.chat_display.bind("<Button-3>", self._show_chat_context_menu)
        self.chat_display.bind("<Escape>", lambda e: self._do_load_html())

        right_frame = ttk.Frame(middle_paned)
        middle_paned.add(right_frame, weight=2)

        self.artifacts_frame = ttk.LabelFrame(
            right_frame, text="Артефакти", padding="5"
        )
        self.artifacts_frame.pack(fill="both", expand=True)
        self.artifacts_frame.pack_forget()

        scrollbar = ttk.Scrollbar(self.artifacts_frame)
        scrollbar.pack(side="right", fill="y")

        self.artifacts_listbox = tk.Listbox(
            self.artifacts_frame,
            font=("Arial", 10),
            yscrollcommand=scrollbar.set,
            activestyle="none",
        )
        self.artifacts_listbox.pack(fill="both", expand=True, padx=(5, 0))
        self.artifacts_listbox.bind("<Double-Button-1>", self._on_artifact_click)
        scrollbar.config(command=self.artifacts_listbox.yview)

        self._add_welcome_message()
        self._generate_suggestions()
        self.window.update_idletasks()
        self.window.deiconify()
        self.window.lift()
        self.window.update()

    def _restore_chat_history(self):
        self._messages_html = []
        self._thinking_index = -1
        if not isinstance(self.artifacts, list):
            self.artifacts = []
        restored_artifacts = []
        for msg in self.chat_history:
            if msg["role"] == "user":
                self._append_html_message(msg["content"], "user")
            elif msg["role"] == "assistant":
                response = msg["content"]
                response, artifacts = DataRequestParser.convert_artifacts_to_html(
                    response
                )
                if artifacts:
                    restored_artifacts.extend(artifacts)
                display_text = self._strip_markers_for_display(response)
                html_content = self._markdown_to_html(display_text)
                msg_html = f'<div class="ai-msg">{html_content}</div>'
                self._messages_html.append(msg_html)
        if restored_artifacts:
            self.artifacts.extend(restored_artifacts)
            self.window.after(
                0, lambda a=restored_artifacts: self._update_artifacts_listbox(a)
            )
        self._generate_suggestions()
        self._do_load_html()

    def _update_status(self, msg: str):
        self.status_label.config(text=msg)
        self.window.update()

    def _add_welcome_message(self):
        welcome = """Вітаю! Я ваш AI науковий консультант.

Я можу допомогти вам з:
- Аналізом публікаційної активності кандидатів
- Виявленням сильних і слабких сторін
- Порівнянням кандидатів
- Наданням рекомендацій для атестації
- Генерацією звітів

Оберіть питання зі списку пропозицій або задайте своє."""

        self._append_message(welcome, "system")

    def _append_chat(self, tag: str, message: str):
        self._append_message(message, tag)

    def _markdown_to_display_text(self, text: str) -> str:
        text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
        text = re.sub(r"\*(.+?)\*", r"\1", text)
        text = re.sub(r"__(.+?)__", r"\1", text)
        text = re.sub(r"_(.+?)_", r"\1", text)
        text = re.sub(r"`(.+?)`", r"\1", text)

        lines = []
        for line in text.split("\n"):
            if line.startswith("# "):
                lines.append("\n" + line[2:])
            elif line.startswith("## "):
                lines.append("\n" + line[3:])
            elif line.startswith("### "):
                lines.append("\n" + line[4:])
            elif line.startswith("- ") or line.startswith("* "):
                lines.append("  • " + line[2:])
            elif re.match(r"^\d+\.\s", line):
                num = re.match(r"^(\d+)\.\s", line).group(1)
                lines.append("  " + num + ". " + line[len(num) + 2 :])
            else:
                lines.append(line)

        return "\n".join(lines)

    def _append_message(self, content: str, msg_type: str = "ai"):
        html_content = self._markdown_to_html(content)
        if msg_type == "user":
            msg_html = f'<div class="user-msg">Ви: {html_content}</div>'
        elif msg_type == "ai":
            msg_html = f'<div class="ai-msg">{html_content}</div>'
        else:
            msg_html = f'<div class="system-msg">{html_content}</div>'
        self._messages_html.append(msg_html)
        # A new message always re-enables autoscroll so the latest content
        # is shown, regardless of where the user had manually scrolled to.
        self._user_scrolled_up = False
        self._update_html_display()

    def _append_html_message(self, content: str, msg_type: str = "ai"):
        self._append_message(content, msg_type)

    def _send_message(self):
        if self.ai_responding:
            self._stop_ai_response()
            return

        msg = self.chat_input.get("1.0", tk.END).strip()
        if not msg:
            return

        self.chat_input.delete("1.0", tk.END)
        self._append_chat("user", msg)

        if not isinstance(self.chat_history, list):
            self.chat_history = []
        self.chat_history.append({"role": "user", "content": msg})

        self.ai_responding = True
        self.stop_response = False
        self._update_send_button()
        threading.Thread(target=self._get_ai_response, args=(msg,), daemon=True).start()

    def _stop_ai_response(self):
        self.stop_response = True
        self._append_chat("system", "[Сеанс перервано користувачем]")

    def _update_send_button(self):
        if self.ai_responding:
            self.send_btn.config(text="Стоп", style="Stop.TButton")
        else:
            self.send_btn.config(text="Надіслати", style="TButton")

    def _get_tools_schema(self) -> List[Dict]:
        tools = [
            {
                "type": "function",
                "function": {
                    "name": "get_candidate_data",
                    "description": "Отримати детальні дані, статистику та список публікацій кандидата з внутрішньої бази",
                    "strict": False,
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "cand_id": {
                                "type": "string",
                                "description": "ID кандидата (наприклад, cand_0)",
                            }
                        },
                        "required": ["cand_id"],
                    },
                },
            },
            {
                "type": "function",
                "function": {
                    "name": "compare_candidates",
                    "description": "Порівняти двох або більше кандидатів за їх науковими профілями",
                    "strict": False,
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "cand_ids": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Список ID кандидатів (наприклад, ['cand_0', 'cand_1'])",
                            }
                        },
                        "required": ["cand_ids"],
                    },
                },
            },
            {
                "type": "function",
                "function": {
                    "name": "web_search",
                    "description": "Шукати інформацію в інтернеті. Повертає короткі сніпети.",
                    "strict": False,
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "query": {
                                "type": "string",
                                "description": "Пошуковий запит",
                            },
                            "num_results": {
                                "type": "integer",
                                "description": "Кількість результатів (за замовчуванням 5)",
                            },
                        },
                        "required": ["query"],
                    },
                },
            },
            {
                "type": "function",
                "function": {
                    "name": "fetch_page",
                    "description": "Отримати повний текстовий вміст веб-сторінки за URL. ВИКОРИСТОВУВАТИ ОБОВ'ЯЗКОВО після web_search, якщо знайдено релевантний лінк.",
                    "strict": False,
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "url": {
                                "type": "string",
                                "description": "URL адреса сторінки",
                            }
                        },
                        "required": ["url"],
                    },
                },
            },
            {
                "type": "function",
                "function": {
                    "name": "scholar_search",
                    "description": "Шукати публікації або профілі авторів у Google Scholar. ВАЖЛИВО: для отримання abstract публікацій обов'язково став detailed=True",
                    "strict": False,
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "action_type": {
                                "type": "string",
                                "enum": ["search_query", "author_name", "author_id"],
                                "description": "search_query = шукати статті, author_name = профіль автора, author_id = профіль за GS ID",
                            },
                            "query": {
                                "type": "string",
                                "description": "Текст запиту, ім'я автора або Google Scholar ID (наприклад, m1Lx2fYAAAAJ)",
                            },
                            "detailed": {
                                "type": "boolean",
                                "description": "ОБОВ'ЯЗКОВО=true для отримання abstract публікацій. Без abstract неможливо оцінити релевантність!",
                            },
                        },
                        "required": ["action_type", "query"],
                    },
                },
            },
            {
                "type": "function",
                "function": {
                    "name": "openalex_search",
                    "description": "Шукати публікації, концепти або авторів в OpenAlex",
                    "strict": False,
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "entity_type": {
                                "type": "string",
                                "enum": ["works", "concepts", "authors"],
                                "description": "Тип сутності для пошуку",
                            },
                            "query": {
                                "type": "string",
                                "description": "Пошуковий запит",
                            },
                        },
                        "required": ["entity_type", "query"],
                    },
                },
            },
            {
                "type": "function",
                "function": {
                    "name": "manage_banned_keywords",
                    "description": "Керувати списком заборонених ключових слів",
                    "strict": False,
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "action": {
                                "type": "string",
                                "enum": ["get", "add"],
                                "description": "Дія: отримати список або додати слово",
                            },
                            "keyword": {
                                "type": "string",
                                "description": "Ключове слово для додавання",
                            },
                            "cand_id": {
                                "type": "string",
                                "description": "ID кандидата (якщо слово для конкретного кандидата)",
                            },
                        },
                        "required": ["action"],
                    },
                },
            },
        ]
        return tools

    def _execute_tool_call(self, tool_name: str, arguments: dict) -> str:
        try:
            if tool_name == "get_candidate_data":
                cand_id = arguments.get("cand_id")
                detailed = self.analysis_data.get_detailed(cand_id)
                if not detailed:
                    return f"Помилка: кандидата з ID {cand_id} не знайдено."

                result = self._format_detailed_candidate(detailed) + "\n\n"
                papers = self.analysis_data.get_papers_by_year(cand_id)
                if papers:
                    result += self._format_papers_by_year(papers)
                return result

            elif tool_name == "compare_candidates":
                ids = arguments.get("cand_ids", [])
                comparison = self.analysis_data.compare_candidates(ids)
                return self._format_comparison(comparison)

            elif tool_name == "web_search":
                query = arguments.get("query")
                num = arguments.get("num_results", 5)
                res = web_search(query, num)

                artifact_content = format_search_results(res)
                self.artifacts.append(
                    {
                        "type": "search_result",
                        "content": artifact_content,
                        "query": query,
                        "timestamp": datetime.now().isoformat(),
                    }
                )
                self.window.after(
                    0,
                    lambda a={
                        "type": "search_result",
                        "content": artifact_content,
                        "query": query,
                    }: self._update_artifacts_listbox([a]),
                )

                return artifact_content

            elif tool_name == "fetch_page":
                url = arguments.get("url")
                res = fetch_url_content(url)
                return format_fetch_result(res)

            elif tool_name == "scholar_search":
                action = arguments.get("action_type")
                query = arguments.get("query")
                detailed = arguments.get("detailed", False)

                if action == "search_query":
                    res = search_google_scholar(query, fetch_details=detailed)
                    return format_scholar_result(res, detailed=detailed)
                elif action == "author_name":
                    res = search_google_scholar_author(
                        author_name=query, max_results=20
                    )
                    return format_scholar_author_result(res)
                elif action == "author_id":
                    res = search_google_scholar_author(
                        author_name="", max_results=20, scholar_id=query
                    )
                    return format_scholar_author_result(res)

            elif tool_name == "openalex_search":
                entity_type = arguments.get("entity_type")
                query = arguments.get("query")

                if entity_type == "works":
                    res = search_works(query)
                    return format_works_result(res)
                elif entity_type == "concepts":
                    res = search_concepts(query)
                    return format_concepts_result(res)
                elif entity_type == "authors":
                    res = search_authors(query)
                    return format_authors_result(res)

            elif tool_name == "manage_banned_keywords":
                action = arguments.get("action")

                if action == "get":
                    cand_id = arguments.get("cand_id")
                    banned = self.analysis_data.get_banned_keywords(cand_id)
                    return f"Виключені ключові слова ({len(banned)}): {', '.join(banned) if banned else 'немає'}"
                elif action == "add":
                    keyword = arguments.get("keyword", "")
                    cand_id = arguments.get("cand_id")
                    if cand_id:
                        success = self.analysis_data.add_banned_keyword(
                            keyword, cand_id
                        )
                        name = self.analysis_data.get_name(cand_id)
                        return (
                            f"Додано '{keyword}' до виключень кандидата {name}"
                            if success
                            else f"'{keyword}' вже є у виключеннях кандидата {name}"
                        )
                    else:
                        success = self.analysis_data.add_banned_keyword(keyword)
                        return (
                            f"Додано '{keyword}' до загальних виключень"
                            if success
                            else f"'{keyword}' вже є у загальних виключеннях"
                        )

            return f"Помилка: Інструмент {tool_name} не знайдено."
        except Exception as e:
            return f"Помилка виконання інструменту {tool_name}: {str(e)}"

    def _process_data_requests(self, requests: List[str]) -> Dict[str, Any]:
        results = {}

        parsed = DataRequestParser.extract_ids(requests)

        for action, ids in parsed:
            if action == "GET":
                for cand_id in ids:
                    if cand_id == "BANNED":
                        banned = self.analysis_data.get_banned_keywords()
                        results["GET:BANNED"] = (
                            f"Виключені ключові слова ({len(banned)}): {', '.join(banned) if banned else 'немає'}"
                        )
                        continue

                    parts = cand_id.split(":")
                    cand_id_clean = parts[0]

                    if len(parts) == 1:
                        detailed = self.analysis_data.get_detailed(cand_id_clean)
                        if detailed:
                            results[f"GET:{cand_id}"] = self._format_detailed_candidate(
                                detailed
                            )

                    elif len(parts) == 2:
                        subtype = parts[1]
                        if subtype == "summary":
                            brief = self.analysis_data.get_brief([cand_id_clean]).get(
                                cand_id_clean
                            )
                            if brief:
                                results[f"GET:{cand_id}"] = self._format_brief_summary(
                                    brief
                                )
                        elif subtype == "papers":
                            papers = self.analysis_data.get_papers_by_year(
                                cand_id_clean
                            )
                            if papers:
                                results[f"GET:{cand_id}"] = self._format_papers_by_year(
                                    papers
                                )
                        elif subtype == "BANNED":
                            banned = self.analysis_data.get_banned_keywords(
                                cand_id_clean
                            )
                            results[f"GET:{cand_id}"] = (
                                f"Виключені ключові слова ({len(banned)}): {', '.join(banned) if banned else 'немає'}"
                            )

                    elif len(parts) == 3:
                        year = int(parts[2])
                        papers = self.analysis_data.get_papers_by_year(cand_id_clean)
                        if year in papers:
                            results[f"GET:{cand_id}"] = self._format_year_stats(
                                papers[year]
                            )

                    elif len(parts) == 4:
                        year = int(parts[2])
                        idx = int(parts[3])
                        paper = self.analysis_data.get_paper_detail(
                            cand_id_clean, year, idx
                        )
                        if paper:
                            results[f"GET:{cand_id}"] = self._format_paper_detail(paper)

            elif action == "COMPARE":
                comparison = self.analysis_data.compare_candidates(ids)
                formatted = self._format_comparison(comparison)
                results[f"COMPARE:{','.join(ids)}"] = formatted
                self.artifacts.append(
                    {
                        "type": "comparison",
                        "content": formatted,
                        "candidates": ",".join(ids),
                        "timestamp": datetime.now().isoformat(),
                    }
                )
                self.window.after(
                    0,
                    lambda a={
                        "type": "comparison",
                        "content": formatted,
                        "candidates": ",".join(ids),
                    }: self._update_artifacts_listbox([a]),
                )

            elif action == "ADD_BANNED":
                for item in ids:
                    if ":" in item:
                        cand_id, keyword = item.split(":", 1)
                        success = self.analysis_data.add_banned_keyword(
                            keyword, cand_id
                        )
                        name = self.analysis_data.get_name(cand_id)
                        results[f"ADD_BANNED:{item}"] = (
                            f"Додано '{keyword}' до виключень кандидата {name}"
                            if success
                            else f"'{keyword}' вже є у виключеннях кандидата {name}"
                        )
                    else:
                        keyword = item
                        success = self.analysis_data.add_banned_keyword(keyword)
                        results[f"ADD_BANNED:{keyword}"] = (
                            f"Додано '{keyword}' до загальних виключень"
                            if success
                            else f"'{keyword}' вже є у загальних виключеннях"
                        )

            elif action == "SEARCH":
                for query in ids:
                    num_results = 5
                    if ":" in query:
                        parts = query.rsplit(":", 1)
                        if parts[1].isdigit():
                            query = parts[0]
                            num_results = int(parts[1])

                    search_result = web_search(query, num_results)
                    results[f"SEARCH:{query}"] = format_search_results(search_result)

                    artifact_content = format_search_results(search_result)
                    self.artifacts.append(
                        {
                            "type": "search_result",
                            "content": artifact_content,
                            "query": query,
                            "source": search_result.get("source", "duckduckgo"),
                            "timestamp": datetime.now().isoformat(),
                        }
                    )
                    self.window.after(
                        0,
                        lambda a={
                            "type": "search_result",
                            "content": artifact_content,
                            "query": query,
                        }: self._update_artifacts_listbox([a]),
                    )

            elif action == "OPENALEX":
                for query in ids:
                    parts = query.split(":", 1)
                    endpoint = parts[0]
                    rest = parts[1] if len(parts) > 1 else ""

                    if endpoint == "works":
                        subparts = rest.rsplit(":", 2)
                        search_term = subparts[0]
                        year = (
                            int(subparts[1])
                            if len(subparts) > 1 and subparts[1].isdigit()
                            else None
                        )
                        sort = (
                            subparts[2]
                            if len(subparts) > 2
                            else "publication_year:desc"
                        )
                        result = search_works(search_term, year, sort)
                        formatted = format_works_result(result)
                    elif endpoint == "concepts":
                        result = search_concepts(rest)
                        formatted = format_concepts_result(result)
                    elif endpoint == "authors":
                        subparts = rest.rsplit(":", 1)
                        search_term = subparts[0]
                        field = subparts[1] if len(subparts) > 1 else None
                        result = search_authors(search_term, field)
                        formatted = format_authors_result(result)
                    else:
                        formatted = f"Unknown endpoint: {endpoint}"

                    results[f"OPENALEX:{query}"] = formatted

            elif action == "SCHOLAR":
                for query in ids:
                    parts = query.split(":", 1)
                    subaction = parts[0] if len(parts) > 0 else ""
                    rest = parts[1] if len(parts) > 1 else ""

                    if subaction == "author":
                        author_query = rest.replace(":detailed", "").strip()
                        fetch_details = "detailed" in rest
                        result = search_google_scholar_author(
                            author_query, max_results=20
                        )
                        formatted = format_scholar_author_result(result)
                    elif subaction == "profile":
                        profile_id = rest.replace(":detailed", "").strip()
                        fetch_details = "detailed" in rest
                        result = search_google_scholar_author(
                            author_name="", max_results=20, scholar_id=profile_id
                        )
                        formatted = format_scholar_author_result(result)
                    else:
                        result = search_google_scholar(query, max_results=10)
                        formatted = format_scholar_result(result, detailed=False)

                    results[f"SCHOLAR:{query}"] = formatted

            elif action == "FETCH":
                for query in ids:
                    url = query.strip()
                    if url:
                        result = fetch_url_content(url)
                        formatted = format_fetch_result(result)
                        results[f"FETCH:{url}"] = formatted

        return results

    def _format_brief_summary(self, brief: BriefSummary) -> str:
        return f"""{brief.name}
Verdict: {brief.verdict}
Конфлікт: {brief.conflict}
Публікацій всього: {brief.papers_total}
Останні роки: {brief.papers_recent}
Придатних публікацій: {brief.papers_applicable}
Top ключові слова: {", ".join(brief.top_keywords[:8]) if brief.top_keywords else "Немає"}"""

    def _format_detailed_candidate(self, cand: DetailedCandidate) -> str:
        lines = []
        lines.append(f"=== {cand.name} ===")
        lines.append(f"Verdict: {cand.verdict}")
        lines.append(f"Конфлікт: {cand.conflict}")
        lines.append(f"Публікацій всього: {cand.papers_total}")
        lines.append(f"Останні роки: {cand.papers_recent}")
        lines.append(f"Придатних публікацій: {cand.papers_applicable}")

        if cand.papers_by_year:
            lines.append("\nПублікації по роках:")
            for year, year_stats in sorted(cand.papers_by_year.items(), reverse=True):
                lines.append(
                    f"  {year}: {pluralize_ukr(year_stats.paper_count, 'публікація', 'публікації', 'публікацій')}, avg_score={year_stats.avg_score:.1f}, relevant={year_stats.relevant_count}"
                )

        if cand.all_keywords:
            lines.append(f"\nВсі ключові слова: {', '.join(cand.all_keywords[:20])}")

        return "\n".join(lines)

    def _format_papers_by_year(self, papers_by_year: Dict[int, YearStats]) -> str:
        lines = ["Публікації по роках:"]
        for year, stats in sorted(papers_by_year.items(), reverse=True):
            lines.append(
                f"\n{year}: {pluralize_ukr(stats.paper_count, 'публікація', 'публікації', 'публікацій')}"
            )
            lines.append(
                f"  avg_score: {stats.avg_score:.1f}, relevant: {stats.relevant_count}"
            )
            for p in stats.papers[:3]:
                lines.append(f"  [{p.score}] {p.title}")
        return "\n".join(lines)

    def _format_year_stats(self, stats: YearStats) -> str:
        lines = [f"=== {stats.year} ==="]
        lines.append(f"Всього публікацій: {stats.paper_count}")
        lines.append(f"Avg score: {stats.avg_score:.1f}")
        lines.append(f"Relevant: {stats.relevant_count}")
        lines.append("\nПублікації:")
        for i, p in enumerate(stats.papers):
            lines.append(f"  {i}. [{p.score}] {p.title}")
            lines.append(f"     Збіги: {p.matched_details}")
        return "\n".join(lines)

    def _format_paper_detail(self, paper: PaperDetail) -> str:
        lines = []
        lines.append(f"=== {paper.title} ===")
        lines.append(f"Рік: {paper.year}")
        lines.append(f"Score: {paper.score}")
        lines.append(f"Збіги: {paper.matched_details}")
        lines.append(f"Джерело: {paper.source}")
        lines.append(f"Журнал: {paper.journal}")
        if paper.authors:
            lines.append(f"Автори: {', '.join(paper.authors[:5])}")
        if paper.author_keywords:
            lines.append(f"Ключові слова автора: {', '.join(paper.author_keywords)}")
        if paper.concepts:
            lines.append(f"Концепти: {', '.join(paper.concepts)}")
        if paper.abstract:
            lines.append(f"\nАнотація:\n{paper.abstract[:500]}...")
        if paper.url:
            lines.append(f"\nURL: {paper.url}")
        return "\n".join(lines)

    def _format_comparison(self, comp: ComparisonResult) -> str:
        lines = ["=== ПОРІВНЯННЯ КАНДИДАТІВ ==="]

        for cid, cand in comp.candidates.items():
            lines.append(f"\n--- {cand.name} ---")
            lines.append(
                f"Придатних публікацій: {cand.papers_applicable}/{cand.papers_recent}"
            )
            lines.append(
                f"Ключові слова: {', '.join(cand.top_keywords[:6]) if cand.top_keywords else 'Немає'}"
            )

        if comp.shared_keywords:
            lines.append(
                f"\nСпільні ключові слова: {', '.join(comp.shared_keywords[:10])}"
            )

        for cid, unique in comp.unique_keywords.items():
            name = self.analysis_data.get_name(cid)
            if unique:
                lines.append(f"\nУнікальні для {name}: {', '.join(unique[:5])}")

        return "\n".join(lines)

    def _get_ai_response(self, user_message: str):
        try:
            messages = [{"role": "system", "content": SYSTEM_PROMPT}]
            initial_context = self.analysis_data.build_initial_context(
                self.selected_cand_ids
            )
            context_prompt = f"КОНТЕКСТ:\n{initial_context}\n---\nКористувач запитує: {user_message}\n---"

            messages.extend(self.chat_history[-8:])
            messages.append({"role": "user", "content": context_prompt})

            self._streaming_buffer = ""
            tools = self._get_tools_schema()

            max_loops = 10
            loop_count = 0
            final_response = ""
            recent_content = ""

            self.window.after(0, lambda: self._start_streaming())
            self.window.after(0, lambda: self._show_thinking("Думаємо..."))

            while loop_count < max_loops:
                if self.stop_response:
                    break

                model = (
                    self.current_model
                    if self.current_model
                    and self.current_model not in ("", "(спершу введіть API ключ)")
                    else AIProvider.PROVIDER_DEFAULT_MODELS.get(self.current_provider)
                )

                if not model:
                    self.window.after(
                        0,
                        lambda: self._append_chat(
                            "system",
                            "Помилка: модель не вибрано. Оберіть модель у налаштуваннях.",
                        ),
                    )
                    self.ai_responding = False
                    self._update_send_button()
                    return

                if self.ai_provider.provider == "google":
                    model = (
                        model.replace("models/", "")
                        .replace("vertex_ai/", "")
                        .replace("gemini/", "")
                    )
                    full_model = f"gemini/{model}"
                else:
                    full_model = (
                        model
                        if "/" in model
                        else f"{self.ai_provider.provider}/{model}"
                    )

                kwargs = {
                    "model": full_model,
                    "messages": messages,
                    "temperature": 0.5,
                    "tools": tools,
                    "tool_choice": "auto",
                    "timeout": 120,
                }

                if self.ai_provider.provider == "google":
                    kwargs["api_key"] = self.ai_provider.api_key
                    kwargs["timeout"] = 180
                elif self.ai_provider.provider == "deepseek":
                    kwargs["api_key"] = self.ai_provider.api_key
                    kwargs["api_base"] = self.ai_provider.get_api_base()
                    kwargs["max_tokens"] = 8192
                else:
                    kwargs["api_key"] = self.ai_provider.api_key
                    kwargs["api_base"] = self.ai_provider.get_api_base()

                response = litellm.completion(**kwargs)
                message = response.choices[0].message

                if hasattr(message, "tool_calls") and message.tool_calls:
                    if message.content:
                        if recent_content:
                            recent_content += "\n" + message.content
                        else:
                            recent_content = message.content
                        messages.append(
                            {
                                "role": "assistant",
                                "content": message.content,
                            }
                        )
                    msg_dict = message.model_dump()
                    content_val = msg_dict.get("content")
                    if self.ai_provider.provider == "deepseek":
                        if content_val is None or content_val == []:
                            msg_dict["content"] = ""
                    messages.append(msg_dict)

                    for tool_call in message.tool_calls:
                        if self.stop_response:
                            break

                        tool_name = tool_call.function.name
                        try:
                            args = json.loads(tool_call.function.arguments)
                        except:
                            args = {}

                        display_name = TOOL_DISPLAY_NAMES.get(tool_name, tool_name)
                        self.window.after(
                            0,
                            lambda dn=display_name: self._show_thinking(
                                f"Використовую: {dn}..."
                            ),
                        )

                        result = self._execute_tool_call(tool_name, args)

                        if result is None:
                            result = f"[Помилка виконання: {display_name} - порожня відповідь]"
                        elif str(result).strip() == "":
                            result = f"[{display_name} повернув порожній результат]"

                        tool_id = getattr(tool_call, "id", None) or f"call_{loop_count}"
                        messages.append(
                            {
                                "role": "tool",
                                "tool_call_id": tool_id,
                                "name": tool_name,
                                "content": str(result),
                            }
                        )

                    loop_count += 1
                    continue

                final_response = message.content or ""
                if not final_response:
                    finish_reason = (
                        response.choices[0].finish_reason
                        if hasattr(response.choices[0], "finish_reason")
                        else "unknown"
                    )
                    debug_info = f"Empty response. finish_reason={finish_reason}, message keys={dir(message)}"
                    self.window.after(
                        0, lambda d=debug_info: self._append_chat("system", d)
                    )
                    final_response = (
                        f"[Пуста відповідь від моделі. finish_reason={finish_reason}]"
                    )
                break

            if not final_response and recent_content:
                final_response = recent_content

            if self.stop_response:
                self.ai_responding = False
                self._update_send_button()
                return

            self.window.after(0, lambda: self._hide_thinking())

            chunk_size = 4
            words = final_response.split(" ")
            for i in range(0, len(words), chunk_size):
                if self.stop_response:
                    break
                chunk = " ".join(words[i : i + chunk_size]) + " "
                self.window.after(0, lambda c=chunk: self._append_streaming_chunk(c))
                import time

                time.sleep(0.02)

            self.chat_history.append({"role": "assistant", "content": final_response})

            html_ready_response, artifacts = (
                DataRequestParser.convert_artifacts_to_html(final_response)
            )
            if artifacts:
                if not isinstance(self.artifacts, list):
                    self.artifacts = []
                self.artifacts.extend(artifacts)
                self.window.after(
                    0, lambda a=artifacts: self._update_artifacts_listbox(a)
                )

            self.window.after(
                0, lambda r=html_ready_response: self._finalize_streaming_message(r)
            )
            self.window.after(0, self._generate_suggestions)

            self.ai_responding = False
            self._update_send_button()

        except Exception as e:
            self.window.after(0, lambda: self._hide_thinking())
            error_str = str(e)
            error_repr = repr(e)
            error_args = str(e.args) if e.args else "No args"
            response_attr = getattr(e, "response", None)
            status_code = getattr(e, "status_code", None)
            message_attr = getattr(e, "message", None)
            cause_attr = getattr(e, "__cause__", None)
            cause_str = f" | __cause__: {str(cause_attr)[:300]}" if cause_attr else ""
            status_str = f" | status_code: {status_code}" if status_code else ""
            response_str = f" | Response: {response_attr}" if response_attr else ""
            if "Timeout" in error_str or "timeout" in error_str.lower():
                error_msg = f"⚠️ Таймаут з'єднання: {self.current_provider} не відповідає. Спробуйте пізніше або змініть провайдера."
            elif "Connection" in error_str:
                error_msg = f"⚠️ Помилка з'єднання: Перевірте інтернет-з'єднання."
            else:
                error_msg = f"Помилка: {error_str[:1000]}\n\nrepr={error_repr[:500]}\n\nargs={error_args[:500]}{cause_str}{status_str}{response_str}"
            self.window.after(0, lambda: self._append_chat("system", error_msg))
            self.ai_responding = False
            self._update_send_button()

    def _show_thinking(self, msg="Думаємо..."):
        thinking_html = f'<div class="system-msg" style="color: #666; font-style: italic;">{msg}</div>'
        if (
            hasattr(self, "_thinking_index")
            and self._thinking_index >= 0
            and self._thinking_index < len(self._messages_html)
        ):
            self._messages_html[self._thinking_index] = thinking_html
        else:
            self._messages_html.append(thinking_html)
            self._thinking_index = len(self._messages_html) - 1
        # Always force a visual update — even during streaming tool-use phases
        self._do_load_html()

    def _hide_thinking(self):
        if (
            hasattr(self, "_thinking_index")
            and self._thinking_index >= 0
            and self._thinking_index < len(self._messages_html)
        ):
            self._messages_html.pop(self._thinking_index)
            self._thinking_index = -1
        # Always force a visual update — even during streaming tool-use phases
        self._do_load_html()

    def _markdown_to_html(self, text: str) -> str:
        artifact_pattern = r'<div class="artifact-block[^"]*">.*?</div></div>'
        artifacts = []

        def preserve_artifact(match):
            artifacts.append(match.group(0))
            return f"\x00ARTIFACT_PLACEHOLDER_{len(artifacts) - 1}\x00"

        text_with_placeholders = re.sub(
            artifact_pattern, preserve_artifact, text, flags=re.DOTALL
        )

        md = markdown.Markdown(
            extensions=[
                "tables",
                "fenced_code",
                "nl2br",
                "sane_lists",
                "def_list",
                "abbr",
                "footnotes",
                "attr_list",
                "md_in_html",
                "pymdownx.mark",
                "pymdownx.tilde",
                "pymdownx.caret",
            ],
            output_format="html",
        )
        html_body = md.convert(text_with_placeholders)

        for i, artifact_html in enumerate(artifacts):
            html_body = html_body.replace(
                f"\x00ARTIFACT_PLACEHOLDER_{i}\x00", artifact_html
            )

        return html_body

    def _update_html_display(self, force: bool = False):
        """Reload the HtmlFrame. Streaming chunks go through _append_streaming_chunk
        which has its own word-count throttle; all other callers (append_message,
        show/hide thinking) are infrequent and always need to render."""
        self._do_load_html()

    def _do_load_html(self):
        # Save scroll position if the user has manually scrolled up
        if self._user_scrolled_up:
            try:
                y = self.chat_display._html.yview()
                self._saved_yview = y[0]
            except Exception:
                pass
        full_html = self._build_full_html()
        self.chat_display.load_html(full_html)
        # Scroll after a short delay so tkinterweb has finished layout
        self.chat_display.after(80, self._do_scroll)

    def _build_full_html(self) -> str:
        css = """body {
    font-family: Arial, sans-serif;
    font-size: 13px;
    line-height: 1.5;
    color: #333;
    margin: 0;
    padding: 10px;
    background: white;
}
h1 { font-size: 16px; margin: 10px 0 5px 0; color: #222; }
h2 { font-size: 14px; margin: 8px 0 4px 0; color: #222; }
h3 { font-size: 13px; margin: 8px 0 4px 0; color: #333; }
p { margin: 5px 0; }
ul, ol { margin: 5px 0 5px 20px; padding: 0; }
li { margin: 3px 0; }
a { color: #0066cc; text-decoration: none; }
a:hover { text-decoration: underline; }
.artifact-link { 
    color: #d35400; 
    background: #fdf2e9; 
    padding: 2px 8px; 
    border-radius: 4px; 
    font-size: 12px;
    cursor: pointer;
}
.artifact-link:hover { 
    background: #fdebd0; 
    text-decoration: none;
}
.artifact-block {
    background: #fff8f0;
    border: 1px solid #e8d4c4;
    border-radius: 6px;
    padding: 12px;
    margin: 10px 0;
}
.artifact-block.recommendation {
    background: #f0fff0;
    border-left: 4px solid #27ae60;
}
.artifact-block.recommendation:hover {
    background: #e8f8e8;
}
.artifact-block.summary {
    background: #f0f0ff;
    border-left: 4px solid #2980b9;
}
.artifact-block.summary:hover {
    background: #e8e8ff;
}
.artifact-block.comparison {
    background: #fff0f0;
    border-left: 4px solid #c0392b;
}
.artifact-block.comparison:hover {
    background: #ffe8e8;
}
.artifact-block.search_result {
    background: #fff8f0;
    border-left: 4px solid #d35400;
}
.artifact-block.search_result:hover {
    background: #fef5eb;
}
.artifact-label {
    font-weight: bold;
    font-size: 13px;
    display: block;
    margin-bottom: 8px;
}
.recommendation .artifact-label { color: #27ae60; }
.summary .artifact-label { color: #2980b9; }
.comparison .artifact-label { color: #c0392b; }
.search_result .artifact-label { color: #d35400; }
.artifact-content {
    color: #333;
    font-size: 13px;
    line-height: 1.5;
}
code {
    background: #f0f0f0;
    padding: 1px 5px;
    border-radius: 3px;
    font-family: Consolas, monospace;
    font-size: 12px;
}
pre {
    background: #f5f5f5;
    padding: 10px;
    border-radius: 5px;
    overflow-x: auto;
    font-family: Consolas, monospace;
    font-size: 12px;
}
table { border-collapse: collapse; margin: 8px 0; width: 100%; }
th, td { border: 1px solid #ddd; padding: 6px 10px; }
th { background: #f8f8f8; }
.user-msg { background: #e3f0ff; padding: 10px 14px; border-radius: 15px 15px 0 15px; margin: 8px 0; max-width: 85%; margin-left: auto; clear: both; }
.ai-msg { background: #f5f5f5; padding: 10px 14px; border-radius: 15px 15px 15px 0; margin: 8px 0; max-width: 85%; clear: both; }
.system-msg { color: #888; font-style: italic; padding: 5px 0; font-size: 12px; clear: both; }
mark, .highlight, .hl, span[style*="background"] {
    background-color: #fff3cd !important;
    padding: 1px 4px;
    border-radius: 3px;
}
mark, .highlight, .hl, span[style*="background"] {
    background-color: #fff3cd !important;
    padding: 1px 4px;
    border-radius: 3px;
    display: inline;
}
del {
    text-decoration: line-through;
    color: #c0392b;
    background-color: #fde8e8;
    padding: 1px 4px;
    border-radius: 3px;
}
sup {
    font-size: 0.75em;
    vertical-align: super;
    line-height: 0;
}
sub {
    font-size: 0.75em;
    vertical-align: sub;
    line-height: 0;
}
abbr {
    text-decoration: underline dotted;
    cursor: help;
}
dl {
    margin: 10px 0;
}
dt {
    font-weight: bold;
    margin-top: 8px;
}
dd {
    margin-left: 20px;
    color: #555;
}
.footnote-ref {
    font-size: 0.75em;
    vertical-align: super;
}
.footnote {
    font-size: 0.85em;
    color: #666;
    border-top: 1px solid #eee;
    padding-top: 8px;
    margin-top: 12px;
}
"""
        return f"""<!DOCTYPE html>
<html>
<head>
<style>{css}</style>
</head>
<body>
{"".join(self._messages_html)}
</body>
</html>"""

    def _update_artifacts_listbox(self, new_artifacts: List[Dict]):
        type_labels = {
            "recommendation": "Рекомендація",
            "summary": "Підсумок",
            "comparison": "Порівняння",
            "search_result": "Пошук",
        }
        for artifact in new_artifacts:
            label = type_labels.get(
                artifact.get("type", "unknown"), artifact.get("type", "unknown")
            )
            content_preview = (
                artifact.get("content", "")[:50] + "..."
                if artifact.get("content") and len(artifact.get("content", "")) > 50
                else artifact.get("content", "") or ""
            )
            if artifact.get("query"):
                content_preview = f"'{artifact['query']}': {content_preview}"
            self.artifacts_listbox.insert(tk.END, f"[{label}] {content_preview}")

    def _show_chat_context_menu(self, event):
        self.chat_context_menu.tk_popup(event.x_root, event.y_root)

    def _copy_chat_selection(self):
        selected = self.chat_display.get_selection()
        if selected:
            self.window.clipboard_clear()
            self.window.clipboard_append(selected)

    def _select_all_chat(self):
        self.chat_display.select_all()

    def _ask_ai_about_selection(self):
        selected = self.chat_display.get_selection()
        if selected:
            self.chat_input.delete("1.0", tk.END)
            self.chat_input.insert("1.0", f"Розкажи більше про: {selected}")
            self._send_message()

    def _explain_selection(self):
        selected = self.chat_display.get_selection()
        if selected:
            self.chat_input.delete("1.0", tk.END)
            self.chat_input.insert("1.0", f"Поясни: {selected}")
            self._send_message()

    def _strip_markers_for_display(self, text: str) -> str:
        id_to_name = None
        if hasattr(self, "analysis_data") and self.analysis_data:
            id_to_name = self.analysis_data._id_to_name
        return DataRequestParser.remove_markers_for_display(text, id_to_name)

    # Number of new words that must accumulate before the streaming display refreshes.
    # Higher = less frequent redraws, lower = more "typing" feel but risks jiggle.
    _STREAM_WORDS_PER_REFRESH = 25

    def _append_streaming_chunk(self, chunk: str):
        self._streaming_buffer += chunk
        word_count = len(self._streaming_buffer.split())

        display_text = self._strip_markers_for_display(self._streaming_buffer)
        html_content = self._markdown_to_html(display_text)
        msg_index = getattr(self, "_streaming_msg_index", -1)
        if msg_index >= 0 and msg_index < len(self._messages_html):
            self._messages_html[msg_index] = f'<div class="ai-msg">{html_content}</div>'
        else:
            self._messages_html[-1] = f'<div class="ai-msg">{html_content}</div>'

        # Update status label with live word count
        try:
            self.status_label.config(text=f"АI відповідає... {word_count} слів")
        except Exception:
            pass

        # Buffered visual update: only reload the HtmlFrame every N words.
        # This gives a smooth "typing" effect without the per-token jiggle.
        if word_count - self._last_stream_word_count >= self._STREAM_WORDS_PER_REFRESH:
            self._last_stream_word_count = word_count
            self._do_load_html()
            self.chat_display.after(80, self._do_scroll)

    def _start_streaming(self):
        """Called on the main thread when AI streaming begins."""
        self._last_stream_word_count = 0
        self._append_message("", "ai")
        self._streaming_msg_index = len(self._messages_html) - 1
        self._hide_thinking()

    def _bind_chat_scroll(self):
        """Bind mouse-wheel events to detect when the user manually scrolls up."""
        inner = getattr(self.chat_display, "_html", None)
        targets = [self.chat_display]
        if inner is not None:
            targets.append(inner)
        for widget in targets:
            try:
                widget.bind("<MouseWheel>", self._on_chat_mousewheel, add="+")
                widget.bind(
                    "<Button-4>", self._on_chat_scroll_up, add="+"
                )  # Linux scroll-up
                widget.bind(
                    "<Button-5>", self._on_chat_scroll_down, add="+"
                )  # Linux scroll-down
            except Exception:
                pass

    def _on_chat_mousewheel(self, event):
        """Windows/macOS: negative delta = scroll up."""
        if event.delta < 0:
            self._on_chat_scroll_down(event)
        else:
            self._on_chat_scroll_up(event)

    def _on_chat_scroll_up(self, event=None):
        """User scrolled toward the top — disable autoscroll."""
        self._user_scrolled_up = True

    def _on_chat_scroll_down(self, event=None):
        """User scrolled down — re-enable autoscroll if they reach the bottom."""
        try:
            y = self.chat_display._html.yview()
            if y[1] >= 0.99:
                self._user_scrolled_up = False
        except Exception:
            pass

    def _do_scroll(self):
        """Execute the actual scroll — called via after() so the HTML has rendered."""
        if self._user_scrolled_up:
            saved_y = self._saved_yview
            try:
                self.chat_display._html.yview_moveto(saved_y)
            except Exception:
                try:
                    self.chat_display.yview_moveto(saved_y)
                except Exception:
                    pass
        else:
            try:
                self.chat_display._html.yview_moveto(1.0)
            except Exception:
                try:
                    self.chat_display.yview_moveto(1.0)
                except Exception:
                    pass

    def _finalize_streaming_message(self, final_response: str):
        display_text = self._strip_markers_for_display(final_response)
        html_content = self._markdown_to_html(display_text)
        msg_index = getattr(self, "_streaming_msg_index", -1)
        if msg_index >= 0 and msg_index < len(self._messages_html):
            self._messages_html[msg_index] = f'<div class="ai-msg">{html_content}</div>'
        else:
            self._messages_html[-1] = f'<div class="ai-msg">{html_content}</div>'
        self._streaming_buffer = ""
        self._last_stream_word_count = 0
        # Clear the status label
        try:
            self.status_label.config(text="")
        except Exception:
            pass
        # One final load with the complete response, then scroll after layout
        self._do_load_html()
        self.chat_display.after(250, self._do_scroll)

    def _generate_suggestions(self):
        self.suggestions = [
            "Хто найкращий кандидат?",
            "Проблеми публікацій?",
            "Порівняйте кандидатів",
            "Рекомендації",
            "Динаміка публікацій",
            "Ознаки наукометрії?",
            "Відповідність ключовим",
        ]

        for btn in self.suggestion_buttons:
            btn.destroy()
        self.suggestion_buttons.clear()

        for i, s in enumerate(self.suggestions):
            short_texts = [
                "Хто найкращий кандидат і чому?",
                "Які основні проблеми з публікаціями?",
                "Порівняйте кандидатів за ключовими словами",
                "Які рекомендації для покращення?",
                "Проаналізуйте динаміку публікацій",
                "Чи є ознаки наукометрії?",
                "Оцініть відповідність ключовим словам",
            ]
            btn = tk.Button(
                self.suggestions_frame,
                text=s,
                wraplen=80,
                font=("Arial", 8),
                bg="#e8e8e8",
                relief="groove",
                cursor="hand2",
                command=lambda idx=i, texts=short_texts: self._on_suggestion_click(
                    idx, texts[idx]
                ),
            )
            btn.pack(side="left", padx=2, pady=2, ipadx=5, ipady=2)
            self.suggestion_buttons.append(btn)

    def _on_suggestion_click(self, idx, full_text):
        self.chat_input.delete("1.0", tk.END)
        self.chat_input.insert("1.0", full_text)
        self._send_message()

    def _show_analysis_data(self):
        dialog = tk.Toplevel(self.window)
        dialog.title("Вхідні дані аналізу")

        text = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, font=("Consolas", 9))
        text.pack(fill="both", expand=True, padx=5, pady=5)

        context = self.analysis_data.build_initial_context(self.selected_cand_ids)
        text.insert("1.0", context)
        text.config(state="disabled")

    def _refresh_context(self):
        self.analysis_data._brief_cache = self.analysis_data._compute_all_briefs()
        self.context_text.config(state="normal")
        self.context_text.delete("1.0", tk.END)
        initial_context = self.analysis_data.build_initial_context(
            self.selected_cand_ids
        )
        self.context_text.insert("1.0", initial_context)
        self.context_text.config(state="disabled")

    def _toggle_artifacts_panel(self):
        if self.artifacts_visible:
            self.artifacts_frame.pack_forget()
            self.artifacts_visible = False
            self.view_menu.entryconfigure(0, label="Показати артефакти")
        else:
            self.artifacts_frame.pack(fill="both", expand=True)
            self.artifacts_frame.config(width=350)
            self.artifacts_visible = True
            self.view_menu.entryconfigure(0, label="Сховати артефакти")

    def _on_artifact_link_click(self, url):
        import re

        if not url:
            return True
        match = re.search(r"artifact://(\d+)", url)
        if not match:
            return True
        try:
            idx = int(match.group(1))
        except (ValueError, IndexError):
            return True
        if not self.artifacts or idx >= len(self.artifacts):
            return True
        self._show_artifact_dialog(idx)
        return True

    def _show_artifact_dialog(self, idx):
        if not self.artifacts or idx >= len(self.artifacts):
            return
        artifact = self.artifacts[idx]
        dialog = tk.Toplevel(self.window)
        type_labels = {
            "recommendation": "Рекомендація",
            "summary": "Підсумок",
            "comparison": "Порівняння",
            "search_result": "Результат пошуку",
        }
        label = type_labels.get(
            artifact.get("type", "unknown"), artifact.get("type", "unknown")
        )
        dialog.title(f"Артефакт: {label}")
        dialog.geometry("700x500")
        text = scrolledtext.ScrolledText(dialog, wrap="word", font=("Arial", 10))
        text.pack(fill="both", expand=True, padx=5, pady=5)
        content = artifact.get("content", "")
        content = re.sub(r"\*(.+?)\*", r"\1", content)
        content = re.sub(r"\*\*(.+?)\*\*", r"\1", content)
        if artifact.get("candidates"):
            content = f"Кандидати: {artifact['candidates']}\n\n{content}"
        if artifact.get("query"):
            content = f"Запит: {artifact['query']}\n\n{content}"
        if artifact.get("source"):
            content = f"{content}\n\nДжерело: {artifact['source']}"
        text.insert("1.0", content if content else "(порожній артефакт)")
        text.config(state="disabled")

    def _on_artifact_click(self, event=None):
        sel = self.artifacts_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if not self.artifacts or idx >= len(self.artifacts):
            return
        self._show_artifact_dialog(idx)

    def _export_artifacts(self):
        if not self.artifacts:
            messagebox.showinfo(
                "Експорт", "Немає артефактів для експорту", parent=self.window
            )
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON", "*.json"), ("Text", "*.txt")]
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.artifacts, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Експорт", "Артефакти збережено!", parent=self.window)

    def _show_change_api_key_dialog(self):
        dialog = tk.Toplevel(self.window)
        dialog.title("Налаштування AI")
        dialog.geometry("650x400")
        dialog.transient(self.window)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(
            main_frame, text="Керування API ключами", font=("Arial", 14, "bold")
        ).pack(anchor="w", pady=(0, 10))

        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True)

        left_frame = ttk.Frame(paned)
        right_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        paned.add(right_frame, weight=2)

        ttk.Label(left_frame, text="Збережені ключі", font=("Arial", 11, "bold")).pack(
            anchor="w", pady=(0, 5)
        )

        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill="both", expand=True)

        providers_listbox = tk.Listbox(list_frame, font=("Arial", 10), height=10)
        providers_listbox.pack(side="left", fill="both", expand=True)

        list_scroll = ttk.Scrollbar(
            list_frame, orient="vertical", command=providers_listbox.yview
        )
        providers_listbox.config(yscrollcommand=list_scroll.set)
        list_scroll.pack(side="right", fill="y")

        def mask_key(key):
            if not key or len(key) < 8:
                return "***"
            return key[:4] + "..." + key[-4:]

        def populate_providers_list():
            providers_listbox.delete(0, tk.END)
            saved_providers = [
                (pk, pn)
                for pk, pn in AIProvider.PROVIDERS
                if pk in self._saved_api_keys
            ]

            if not saved_providers:
                providers_listbox.insert(0, "(немає)")
                providers_listbox.itemconfig(0, fg="gray")
            else:
                for provider_key, provider_name in saved_providers:
                    saved = self._saved_api_keys[provider_key]
                    key_mask = mask_key(saved.get("api_key", ""))
                    model = saved.get("model", "")
                    model_short = (
                        (model[:12] + "...")
                        if model and len(model) > 12
                        else (model or "-")
                    )
                    providers_listbox.insert(tk.END, f"{provider_name} [{model_short}]")

        populate_providers_list()

        ttk.Label(right_frame, text="Редагування", font=("Arial", 11, "bold")).pack(
            anchor="w", pady=(0, 10)
        )

        detail_frame = ttk.LabelFrame(right_frame, text=" Провайдер ", padding="10")
        detail_frame.pack(fill="x", pady=(0, 10))
        detail_frame.pack_forget()

        selected_provider_key = tk.StringVar()
        selected_provider_name = tk.Label(
            detail_frame, text="Оберіть зі списку", font=("Arial", 10)
        )
        selected_provider_name.pack(anchor="w")

        model_frame = ttk.Frame(detail_frame)
        model_frame.pack(fill="x", pady=(5, 5))
        ttk.Label(model_frame, text="Модель:", width=10).pack(side="left")
        model_var = tk.StringVar()
        model_combo = ttk.Combobox(model_frame, textvariable=model_var, width=30)
        model_combo.pack(side="left", fill="x", expand=True)

        key_var = tk.StringVar()
        show_key_var = tk.BooleanVar(value=False)

        key_frame = ttk.Frame(detail_frame)
        key_frame.pack(fill="x", pady=(5, 5))
        ttk.Label(key_frame, text="API ключ:", width=10).pack(side="left")
        key_entry = ttk.Entry(key_frame, textvariable=key_var, width=30, show="*")
        key_entry.pack(side="left", fill="x", expand=True)
        ttk.Checkbutton(key_frame, text="Показати", variable=show_key_var).pack(
            side="left", padx=(5, 0)
        )

        def on_show_toggle():
            if show_key_var.get():
                key_entry.config(show="")
            else:
                key_entry.config(show="*")

        show_key_var.trace_add("write", lambda *a: on_show_toggle())

        def on_paste():
            key_entry.config(show="")
            show_key_var.set(True)
            dialog.after(
                2000,
                lambda: (
                    key_entry.config(show="*"),
                    show_key_var.set(False),
                    on_key_change(),
                ),
            )

        key_entry.bind("<Control-v>", lambda e: (on_paste(), None))

        status_label = tk.Label(detail_frame, text="", fg="gray", font=("Arial", 9))
        status_label.pack(anchor="w", pady=(5, 0))

        def on_key_change(*args):
            key = key_var.get().strip()
            provider_key = selected_provider_key.get()
            if len(key) >= 10 and provider_key:
                status_label.config(text="Завантажую моделі...")
                dialog.update()
                try:
                    temp_provider = AIProvider(key, provider_key)
                    models = temp_provider.get_available_models()
                    if models:
                        model_combo["values"] = models
                        if not model_var.get() or model_var.get() not in models:
                            model_var.set(models[0])
                        status_label.config(
                            text=f"Знайдено {len(models)} моделей", fg="green"
                        )
                    else:
                        status_label.config(text="Моделі не знайдено")
                except Exception as e:
                    status_label.config(text=f"Помилка: {str(e)[:50]}")
            elif len(key) >= 10:
                status_label.config(text="Оберіть провайдера")
            else:
                status_label.config(text="")

        key_entry.bind("<KeyRelease>", lambda e: on_key_change())
        key_var.trace_add("write", on_key_change)

        def on_provider_select(event):
            selection = providers_listbox.curselection()
            if not selection:
                return
            idx = selection[0]
            saved_providers = [
                (pk, pn)
                for pk, pn in AIProvider.PROVIDERS
                if pk in self._saved_api_keys
            ]
            if idx >= len(saved_providers):
                return
            provider_key, provider_name = saved_providers[idx]
            selected_provider_key.set(provider_key)
            selected_provider_name.config(text=f"{provider_name}")

            saved = self._saved_api_keys[provider_key]
            key_var.set(saved.get("api_key", ""))
            show_key_var.set(False)
            model_var.set(saved.get("model", ""))

            detail_frame.pack(fill="x", pady=(0, 10))

            key = saved.get("api_key", "")
            if key and len(key) >= 10:
                status_label.config(text="Завантажую моделі...")
                dialog.update()
                try:
                    temp_provider = AIProvider(key, provider_key)
                    models = temp_provider.get_available_models()
                    if models:
                        model_combo["values"] = models
                        if not model_var.get() or model_var.get() not in models:
                            model_var.set(models[0])
                        status_label.config(
                            text=f"Знайдено {len(models)} моделей", fg="green"
                        )
                    else:
                        status_label.config(text="Моделі не знайдено")
                except Exception as e:
                    status_label.config(text=f"Помилка: {str(e)[:50]}")
            else:
                model_combo["values"] = []

        providers_listbox.bind("<<ListboxSelect>>", on_provider_select)

        def save_provider():
            provider_key = selected_provider_key.get()
            if not provider_key:
                status_label.config(text="Оберіть провайдера")
                return
            key = key_var.get().strip()
            model = model_var.get().strip()
            if not key:
                status_label.config(text="Введіть API ключ")
                return
            try:
                AIProvider(key, provider_key)
            except Exception as e:
                status_label.config(text=f"Помилка: {str(e)[:50]}")
                return
            self._saved_api_keys[provider_key] = {"api_key": key, "model": model}
            status_label.config(text="Збережено!", fg="green")
            populate_providers_list()

        def use_provider():
            provider_key = selected_provider_key.get()
            if not provider_key:
                status_label.config(text="Оберіть провайдера")
                return
            if provider_key not in self._saved_api_keys:
                status_label.config(text="Спершу збережіть ключ")
                return
            saved = self._saved_api_keys[provider_key]
            key = saved.get("api_key", "")
            model = saved.get("model", "")
            if not key:
                status_label.config(text="Ключ відсутній")
                return
            try:
                self.ai_provider = AIProvider(key, provider_key)
                self.current_provider = provider_key
                self.current_model = model if model else None
                self.current_api_key = key
                provider_name = dict(AIProvider.PROVIDERS).get(
                    provider_key, provider_key
                )
                messagebox.showinfo(
                    "Успіх", f"Переключено на {provider_name}", parent=dialog
                )
                dialog.destroy()
            except Exception as e:
                status_label.config(text=f"Помилка: {str(e)[:50]}")

        def delete_provider():
            provider_key = selected_provider_key.get()
            if not provider_key:
                return
            if provider_key in self._saved_api_keys:
                del self._saved_api_keys[provider_key]
                key_var.set("")
                model_var.set("")
                model_combo["values"] = []
                populate_providers_list()
                selected_provider_key.set("")
                selected_provider_name.config(text="Оберіть зі списку")
                detail_frame.pack_forget()
                status_label.config(text="Видалено")

        edit_btn_frame = ttk.Frame(detail_frame)
        edit_btn_frame.pack(fill="x", pady=(5, 0))
        ttk.Button(
            edit_btn_frame, text="Зберегти", command=save_provider, width=10
        ).pack(side="left", padx=2)
        ttk.Button(
            edit_btn_frame, text="Видалити", command=delete_provider, width=10
        ).pack(side="left", padx=2)
        ttk.Button(
            edit_btn_frame, text="Використати", command=use_provider, width=10
        ).pack(side="left", padx=2)

        def add_new_provider():
            add_dialog = tk.Toplevel(dialog)
            add_dialog.title("Додати")
            add_dialog.geometry("280x120")
            add_dialog.transient(dialog)
            add_dialog.grab_set()
            add_dialog.update_idletasks()
            x = (add_dialog.winfo_screenwidth() // 2) - (
                add_dialog.winfo_reqwidth() // 2
            )
            y = (add_dialog.winfo_screenheight() // 2) - (
                add_dialog.winfo_reqheight() // 2
            )
            add_dialog.geometry(f"+{x}+{y}")

            ttk.Label(add_dialog, text="Провайдер:", font=("Arial", 11)).pack(pady=5)

            add_var = tk.StringVar()
            available = [
                name
                for key, name in AIProvider.PROVIDERS
                if key not in self._saved_api_keys
            ]

            if not available:
                ttk.Label(add_dialog, text="Всі додані", foreground="gray").pack(
                    pady=15
                )
                ttk.Button(add_dialog, text="OK", command=add_dialog.destroy).pack()
                return

            add_combo = ttk.Combobox(
                add_dialog,
                textvariable=add_var,
                values=available,
                state="readonly",
                width=22,
            )
            add_combo.pack(pady=5)
            add_combo.current(0)

            def do_add():
                selected = add_var.get()
                provider_key = dict((v, k) for k, v in AIProvider.PROVIDERS)[selected]
                selected_provider_key.set(provider_key)
                selected_provider_name.config(text=f"{selected}")
                key_var.set("")
                model_var.set("")
                model_combo["values"] = []
                show_key_var.set(False)
                detail_frame.pack(fill="x", pady=(0, 10))
                add_dialog.destroy()

            ttk.Button(add_dialog, text="Додати", command=do_add).pack(pady=5)

        bottom_btn_frame = ttk.Frame(main_frame)
        bottom_btn_frame.pack(fill="x", pady=(5, 0))
        ttk.Button(
            bottom_btn_frame, text="Додати +", command=add_new_provider, width=12
        ).pack(side="left", padx=2)
        ttk.Button(
            bottom_btn_frame, text="Закрити", command=dialog.destroy, width=12
        ).pack(side="right", padx=2)
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_reqwidth() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_reqheight() // 2)
        dialog.geometry(f"+{x}+{y}")

    def _clear_history(self):
        if messagebox.askyesno(
            "Очистити", "Очистити історію чату?", parent=self.window
        ):
            self.chat_history = []
            self._messages_html = []
            self._thinking_index = -1
            self._add_welcome_message()

    def _on_close(self):
        self.window.withdraw()

    def show_window(self):
        self.window.deiconify()
        self.window.lift()


def launch_ai_advisor(
    parent_window,
    candidates: Dict,
    papers: Dict,
    target_keywords: List[str],
    cutoff_year: int,
    global_banned: List[str],
    selected_cand_ids: List[str] = None,
    restore_state: Dict = None,
):
    years_back = 4

    data = LazyAnalysisData(
        candidates=candidates,
        papers=papers,
        target_keywords=target_keywords,
        cutoff_year=cutoff_year,
        years_back=years_back,
        global_banned=global_banned,
    )

    if selected_cand_ids is None:
        selected_cand_ids = list(candidates.keys())

    return AIAdvisorApp(parent_window, data, selected_cand_ids, restore_state)
