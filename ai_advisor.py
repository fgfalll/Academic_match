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


def render_markdown_to_html(text: str) -> str:
    html_body = markdown.markdown(
        text, extensions=["tables", "fenced_code", "nl2br", "sane_lists"]
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
        selected_briefs = self.get_brief(selected_cand_ids)
        other_ids = [cid for cid in self.get_all_ids() if cid not in selected_cand_ids]
        other_briefs = self.get_brief(other_ids)

        lines = []
        lines.append("=== КОНТЕКСТ ДЛЯ АНАЛІЗУ ===")
        lines.append(f"Період аналізу: останні {self.years_back} років")
        lines.append(
            f"Ключові слова: {', '.join(self.target_keywords) if self.target_keywords else 'Не задано'}"
        )
        lines.append("")

        if selected_briefs:
            lines.append("=== ОБРАНІ КАНДИДАТИ ===")
            for cid, brief in sorted(selected_briefs.items(), key=lambda x: x[1].name):
                lines.append(f"\n--- {brief.name} ({cid}) ---")
                lines.append(f"Verdict: {brief.verdict}")
                lines.append(f"Конфлікт: {brief.conflict}")
                lines.append(f"Публікацій всього: {brief.papers_total}")
                lines.append(f"Останні роки: {brief.papers_recent}")
                lines.append(f"Придатних публікацій: {brief.papers_applicable}")

                if brief.top_scores:
                    lines.append("Top публікації:")
                    for score, title, matched in brief.top_scores[:3]:
                        lines.append(f'  [{score}] "{title}" - {matched}')

                if brief.top_keywords:
                    lines.append(f"Ключові слова: {', '.join(brief.top_keywords[:8])}")

        if other_briefs:
            lines.append("\n=== ІНШІ КАНДИДАТИ ===")
            for cid, brief in sorted(other_briefs.items(), key=lambda x: x[1].name):
                lines.append(
                    f"{brief.name} - {brief.verdict} - {pluralize_ukr(brief.papers_recent, 'публікація', 'публікації', 'публікацій')}"
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
    REQUEST_PATTERN = r"\[(?:GET|COMPARE|ADD_BANNED|SEARCH|OPENALEX):[^\]]+\]"
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
            return ""

        text = re.sub(r"\[(?:GET|COMPARE|SEARCH|OPENALEX):[^\]]+\]", replace_get, text)

        text = re.sub(
            r"\[(?:ADD_BANNED|GET|COMPARE|SEARCH|OPENALEX):[^\]]*\]", "", text
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
                if action in ("GET", "SEARCH", "OPENALEX"):
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
        "openai": "gpt-4o",
        "anthropic": "claude-3-5-sonnet-20241022",
        "google": "gemini-1.5-pro",
        "deepseek": "deepseek-chat",
        "zhipu": "glm-4",
        "moonshot": "moonshot-v1-8k",
        "minimax": "MiniMax-M2.1",
        "groq": "llama-3.3-70b-versatile",
        "openrouter": "openrouter/auto",
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
                resp = requests.get(f"{api_base}/models", headers=headers, timeout=15)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]

            elif self.provider == "moonshot":
                resp = requests.get(f"{api_base}/models", headers=headers, timeout=15)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]

            elif self.provider == "zhipu":
                resp = requests.get(f"{api_base}/models", headers=headers, timeout=15)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]

            elif self.provider == "xai":
                resp = requests.get(f"{api_base}/models", headers=headers, timeout=15)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]

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

        full_model = model if "/" in model else f"{self.provider}/{model}"

        try:
            kwargs = {
                "model": full_model,
                "messages": messages,
                "temperature": 0.7,
            }

            if self.provider == "google":
                kwargs["api_key"] = self.api_key
            else:
                kwargs["api_key"] = self.api_key
                kwargs["api_base"] = self.get_api_base()

            response = litellm.completion(**kwargs)
            return response["choices"][0]["message"]["content"]
        except Exception as e:
            error_str = str(e)
            if "Authentication" in error_str or "auth" in error_str.lower():
                raise ValueError(
                    f"Помилка автентифікації: Перевірте API ключ для {self.provider}\n\nДеталі: {error_str[:300]}"
                )
            elif "rate limit" in error_str.lower():
                raise ValueError(
                    f"Ліміт запитів: Спробуйте пізніше\n\nДеталі: {error_str[:300]}"
                )
            elif "quota" in error_str.lower() or "limit" in error_str.lower():
                raise ValueError(
                    f"Квота вичерпана для {self.provider}\n\nДеталі: {error_str[:300]}"
                )
            else:
                raise ValueError(f"Помилка {self.provider}: {error_str[:300]}")

    def chat_stream(self, messages: List[Dict], model: str = None):
        if model is None:
            model = self.PROVIDER_DEFAULT_MODELS.get(self.provider, "default")

        full_model = model if "/" in model else f"{self.provider}/{model}"

        try:
            kwargs = {
                "model": full_model,
                "messages": messages,
                "temperature": 0.7,
                "stream": True,
            }

            if self.provider == "google":
                kwargs["api_key"] = self.api_key
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
            if "Authentication" in error_str or "auth" in error_str.lower():
                raise ValueError(
                    f"Помилка автентифікації: Перевірте API ключ для {self.provider}\n\nДеталі: {error_str[:300]}"
                )
            elif "rate limit" in error_str.lower():
                raise ValueError(
                    f"Ліміт запитів: Спробуйте пізніше\n\nДеталі: {error_str[:300]}"
                )
            elif "quota" in error_str.lower() or "limit" in error_str.lower():
                raise ValueError(
                    f"Квота вичерпана для {self.provider}\n\nДеталі: {error_str[:300]}"
                )
            else:
                raise ValueError(f"Помилка {self.provider}: {error_str[:300]}")


SYSTEM_PROMPT = """Ти - науковий консультант для атестаційної комісії (разової спеціалізованої ради).

КОНТЕКСТ РОБОТИ:
- Аналіз кандидатів на присвоєння наукового ступеня
- Дані збираються автоматично з ORCID, Google Scholar, OpenAlex
- Релевантність публікацій оцінюється за ключовими словами (score 0-5)
- Придатна публікація: score > 0

СТРУКТУРА ДАНИХ:
- Кандидати позначаються ID: cand_0, cand_1, cand_2, etc.
- Період аналізу: останні роки (обычно 4)
- Verdict: "Відповідає вимогам" якщо ≥3 придатних публікацій і немає конфлікту

ЗАПИТ ДАНИХ (тільки коли реально потрібно):
[GET:cand_0] - повні дані кандидата (публікації по роках, всі ключові слова)
[GET:cand_0:summary] - короткий підсумок кандидата
[GET:cand_0:papers] - публікації агреговані по роках
[GET:cand_0:papers:2024] - публікації за конкретний рік
[GET:cand_0:paper:2024:0] - деталі конкретної публікації (рік, індекс)
[COMPARE:cand_0:cand_1] - порівняння двох кандидатів
[GET:BANNED] - отримати список загальних виключених ключових слів
[GET:cand_0:BANNED] - отримати список виключених ключових слів конкретного кандидата
[ADD_BANNED:слово] - додати ключове слово до загальних виключень
[ADD_BANNED:cand_0:слово] - додати ключове слово до виключень конкретного кандидата

ПОШУК В ІНТЕРНЕТІ (ВИКОРИСТОВУЙ АКТИВНО!):
Використовуй [SEARCH:запит] коли потрібно:
- Знайти актуальні наукові статті з теми
- Перевірити journal ranking або impact factor
- Знайти інформацію про науковця чи установу
- Отримати дані про цитування
- Знайти recent research не в твоїх даних
- Перевірити якість журналу де опублікована стаття

OPENALEX API (використовуй для детальних даних про авторів):
[OPENALEX:works:search term] - search scientific papers on OpenAlex
[OPENALEX:concepts:field name] - get OpenAlex concept ID for a field
[OPENALEX:authors:author name] - find author ID and detailed citation metrics

Приклади:
[SEARCH:CCUS carbon capture storage latest research 2024]
[SEARCH:journal impact factor petroleum science]
[SEARCH:Dr. Petrov ORCID publication list]

ПРАВИЛА:
- НЕ пиши [GET...], [COMPARE...] або [ADD_BANNED...] в повідомленнях користувачу - вони для внутреннього використання
- НЕ пиши [SEARCH...] в повідомленнях користувачу - вони для внутреннього використання
- ВІДПОВІДАЙ природною українською мовою
- Звертайся до кандидатів за іменем (Петренко І.І., не cand_0)
- Будь об'єктивним та конструктивним
- Вказуй конкретні проблеми та рекомендації
- Після відповіді пропонуй можливі наступні кроки
- Якщо рекомендуєш виключити якесь слово загалом - використай [ADD_BANNED:слово], якщо лише для кандидата - [ADD_BANNED:cand_0:слово]
- Якщо потрібна актуальна інформація - використай [SEARCH:запит]

АРТЕФАКТИ:
Коли даєш рекомендації, підсумки або порівняння - ЗБЕРІГАЙ їх як артефакти!
Формат: [ARTIFACT:recommendation]текст рекомендації[/ARTIFACT]
Формат: [ARTIFACT:summary]текст підсумку[/ARTIFACT]
Формат: [ARTIFACT:comparison]текст порівняння[/ARTIFACT]
Формат: [ARTIFACT:search_result]результати пошуку[/ARTIFACT]"""


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

        self._select_project_window()

    def get_state_for_session(self, pin: str = None) -> Dict:
        if not self.current_api_key:
            return None

        api_key_to_store = self.current_api_key
        chat_history_to_store = self.chat_history

        if pin:
            api_key_to_store = "enc:" + encrypt_with_embedded_pin_hash(
                self.current_api_key, pin
            )
            chat_history_json = json.dumps(self.chat_history, ensure_ascii=False)
            chat_history_to_store = "enc:" + encrypt_with_embedded_pin_hash(
                chat_history_json, pin
            )

        state = {
            "provider": self.current_provider,
            "model": self.current_model,
            "api_key": api_key_to_store,
            "chat_history": chat_history_to_store,
            "artifacts": self.artifacts,
        }

        return state

    def restore_from_session(self, state: Dict, pin: str = None):
        if not state:
            return False

        provider = state.get("provider")
        model = state.get("model")
        api_key_encrypted = state.get("api_key")
        chat_history_encrypted = state.get("chat_history")
        artifacts = state.get("artifacts", [])

        if api_key_encrypted:
            if pin and api_key_encrypted.startswith("enc:"):
                pin_hash, api_key = decrypt_with_embedded_pin_hash(
                    api_key_encrypted[4:], pin
                )
                if pin_hash is None:
                    return False
            elif not api_key_encrypted.startswith("enc:"):
                api_key = api_key_encrypted
            else:
                return False

        if chat_history_encrypted:
            if pin and chat_history_encrypted.startswith("enc:"):
                try:
                    _, chat_history_json = decrypt_with_embedded_pin_hash(
                        chat_history_encrypted[4:], pin
                    )
                    if chat_history_json:
                        self.chat_history = json.loads(chat_history_json)
                except:
                    pass

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
                        self.chat_history = self._restore_state["chat_history"]
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
                default_model = AIProvider.PROVIDER_DEFAULT_MODELS.get(
                    provider_key, "default"
                )
                model_combo["values"] = [default_model]
                model_var.set(default_model)

        def fetch_models_for_provider():
            provider_key = None
            for key, name in AIProvider.PROVIDERS:
                if name == provider_var.get():
                    provider_key = key
                    break
            if not provider_key or not key_var.get().strip():
                status_label.config(text="Введіть API ключ для отримання моделей")
                return
            status_label.config(text="Завантаження моделей...")
            try:
                temp_provider = AIProvider(key_var.get().strip(), provider_key)
                models = temp_provider.get_available_models()
                if models:
                    model_combo.after(0, lambda m=models: update_model_list(m))
                    status_label.config(text=f"Знайдено {len(models)} моделей")
                else:
                    status_label.config(text="Моделі не знайдено")
            except Exception as e:
                status_label.config(text=f"Помилка: {str(e)[:50]}")

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
                self._start_with_api_key(
                    provider_key, key_var.get().strip(), dialog, model
                )
            else:
                messagebox.showwarning("Увага", "Введіть API ключ")

        provider_combo.bind("<<ComboboxSelected>>", update_default_model)
        key_entry.bind("<KeyRelease>", lambda e: status_label.config(text=""))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x")
        ttk.Button(
            btn_frame,
            text="Оновити моделі",
            command=fetch_models_for_provider,
            width=15,
        ).pack(side="left", padx=5)
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
            messagebox.showerror("Помилка", f"Не вдалося підключитися: {str(e)}")
            return

        if chat_history:
            self.chat_history = chat_history

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
            chat_history_encrypted = state.get("chat_history")
            artifacts = state.get("artifacts", [])

            api_key = api_key_encrypted
            if api_key.startswith("enc:"):
                pin_hash, api_key = decrypt_with_embedded_pin_hash(api_key[4:], pin)
                if pin_hash is None:
                    messagebox.showerror("Помилка", "Невірний PIN")
                    self.pin_var.set("")
                    return

            self.pin_window.destroy()
            self._start_with_api_key(provider, api_key, None, model)

            if chat_history_encrypted:
                if chat_history_encrypted.startswith("enc:"):
                    try:
                        _, chat_history_json = decrypt_with_embedded_pin_hash(
                            chat_history_encrypted[4:], pin
                        )
                        if chat_history_json:
                            self.chat_history = json.loads(chat_history_json)
                    except:
                        pass

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
        for msg in self.chat_history:
            if msg["role"] == "user":
                self._append_html_message(msg["content"], "user")
            elif msg["role"] == "assistant":
                self._append_html_message(msg["content"], "ai")
        self._generate_suggestions()

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

            id_to_name = self.analysis_data._id_to_name
            initial_context = self.analysis_data.build_initial_context(
                self.selected_cand_ids
            )

            context_prompt = f"""КОНТЕКСТ:
{initial_context}

---
Користувач запитує: {user_message}
---
"""
            messages.append({"role": "user", "content": context_prompt})
            messages.extend(self.chat_history[-8:])

            self.window.after(0, lambda: self._show_thinking())

            self._streaming_buffer = ""
            self.window.after(0, lambda: self._start_streaming())

            full_response = []

            for chunk in self.ai_provider.chat_stream(messages):
                if self.stop_response:
                    break
                full_response.append(chunk)
                self.window.after(0, lambda c=chunk: self._append_streaming_chunk(c))

            if self.stop_response:
                self.ai_responding = False
                self._update_send_button()
                return

            response = "".join(full_response)
            self.chat_history.append({"role": "assistant", "content": response})

            self.window.after(0, lambda: self._hide_thinking())

            artifacts = DataRequestParser.parse_artifacts(response)
            if artifacts:
                self.artifacts.extend(artifacts)
                self.window.after(
                    0, lambda a=artifacts: self._update_artifacts_listbox(a)
                )
                response = DataRequestParser.remove_artifacts(response)

            requests = DataRequestParser.parse(response)

            if not requests:
                self.window.after(
                    0, lambda r=response: self._finalize_streaming_message(r)
                )
            else:
                parsed = DataRequestParser.extract_ids(requests)
                request_names = []
                for action, ids in parsed:
                    if action == "ADD_BANNED":
                        request_names.append(f"додати виключення: {ids[0]}")
                    elif action == "GET" and ids[0] == "BANNED":
                        request_names.append("список виключень")
                    else:
                        for i in ids:
                            parts = i.split(":")
                            cid = parts[0]
                            name = self.analysis_data.get_name(cid)
                            request_names.append(name)

                self.window.after(
                    0,
                    lambda names=request_names: self._show_thinking(
                        f"Думаємо... Запит даних: {', '.join(names)}"
                    ),
                )

                results = self._process_data_requests(requests)

                for req_id, result in results.items():
                    self.window.after(
                        0,
                        lambda r=req_id, res=result: self._show_thinking(
                            f"Думаємо... Отримано: {r}"
                        ),
                    )

                continuation_prompt = "Отримані дані:\n"
                for req, result in results.items():
                    continuation_prompt += f"\n=== Результат [{req}] ===\n{result}\n"

                continuation_prompt += "\nПродовж відповідь враховуючи ці дані."

                messages.append({"role": "user", "content": continuation_prompt})

                self.window.after(
                    0, lambda: self._show_thinking("Думаємо... Аналізуємо дані")
                )

                full_response2 = []
                for chunk in self.ai_provider.chat_stream(messages):
                    if self.stop_response:
                        break
                    full_response2.append(chunk)
                    self.window.after(
                        0, lambda c=chunk: self._append_streaming_chunk(c)
                    )

                if self.stop_response:
                    self.ai_responding = False
                    self._update_send_button()
                    return

                response2 = "".join(full_response2)
                self.chat_history.append({"role": "assistant", "content": response2})

                self.window.after(0, lambda: self._hide_thinking())

                artifacts2 = DataRequestParser.parse_artifacts(response2)
                if artifacts2:
                    self.artifacts.extend(artifacts2)
                    self.window.after(
                        0, lambda a=artifacts2: self._update_artifacts_listbox(a)
                    )

                combined_response = response + "\n\n" + response2
                self.window.after(
                    0,
                    lambda r=combined_response: self._finalize_streaming_message(r),
                )

            self.window.after(0, self._generate_suggestions)
            self.ai_responding = False
            self._update_send_button()

        except Exception as e:
            self.window.after(0, lambda: self._hide_thinking())
            error_msg = f"Помилка: {str(e)}"
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
        md = markdown.Markdown(
            extensions=["tables", "fenced_code", "nl2br", "sane_lists"],
            output_format="html",
        )
        html_body = md.convert(text)
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

        # Always keep _messages_html in sync so any forced render shows latest content
        display_text = self._strip_markers_for_display(self._streaming_buffer)
        html_content = self._markdown_to_html(display_text)
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
        self._append_message("", "ai")  # clean HTML load + scroll to bottom

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

    def _on_artifact_click(self, event=None):
        sel = self.artifacts_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self.artifacts):
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

    def _export_artifacts(self):
        if not self.artifacts:
            messagebox.showinfo("Експорт", "Немає артефактів для експорту")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON", "*.json"), ("Text", "*.txt")]
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.artifacts, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Експорт", "Артефакти збережено!")

    def _show_change_api_key_dialog(self):
        dialog = tk.Toplevel(self.window)
        dialog.title("Зміна API ключа")
        dialog.resizable(0, 0)
        dialog.transient(self.window)
        dialog.grab_set()

        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_reqwidth() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_reqheight() // 2)
        dialog.geometry(f"+{x}+{y}")

        main_frame = ttk.Frame(dialog, padding="25")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="Зміна API ключа", font=("Arial", 16, "bold")).pack(
            pady=(0, 20)
        )

        current_provider = self.current_provider or "OpenAI"
        current_model = self.current_model or ""
        provider_names = [name for _, name in AIProvider.PROVIDERS]
        current_provider_name = dict(AIProvider.PROVIDERS).get(
            current_provider, current_provider
        )

        current_frame = ttk.LabelFrame(main_frame, text=" Поточний ключ ", padding="15")
        current_frame.pack(fill="x", pady=(0, 15))
        ttk.Label(current_frame, text=f"Провайдер: {current_provider_name}").pack(
            anchor="w"
        )
        ttk.Label(
            current_frame,
            text=f"Модель: {current_model if current_model else '(default)'}",
        ).pack(anchor="w")

        input_frame = ttk.LabelFrame(main_frame, text=" Новий API ключ ", padding="15")
        input_frame.pack(fill="x", pady=(0, 15))

        row_provider = ttk.Frame(input_frame)
        row_provider.pack(fill="x", pady=(0, 10))
        ttk.Label(row_provider, text="Провайдер:", width=12).pack(
            side="left", padx=(0, 5)
        )
        provider_var = tk.StringVar(value=current_provider_name)
        provider_combo = ttk.Combobox(
            row_provider,
            textvariable=provider_var,
            values=provider_names,
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
                default_model = AIProvider.PROVIDER_DEFAULT_MODELS.get(
                    provider_key, "default"
                )
                model_combo["values"] = [default_model]
                model_var.set(default_model)

        def fetch_models_for_provider():
            provider_key = None
            for key, name in AIProvider.PROVIDERS:
                if name == provider_var.get():
                    provider_key = key
                    break
            if not provider_key or not key_var.get().strip():
                status_label.config(text="Введіть API ключ для отримання моделей")
                return
            status_label.config(text="Завантаження моделей...")
            try:
                temp_provider = AIProvider(key_var.get().strip(), provider_key)
                models = temp_provider.get_available_models()
                if models:
                    model_combo.after(0, lambda m=models: update_model_list(m))
                    status_label.config(text=f"Знайдено {len(models)} моделей")
                else:
                    status_label.config(text="Моделі не знайдено")
            except Exception as e:
                status_label.config(text=f"Помилка: {str(e)[:50]}")

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
                try:
                    new_provider = AIProvider(key_var.get().strip(), provider_key)
                    self.current_provider = provider_key
                    self.current_model = model
                    self.current_api_key = key_var.get().strip()
                    self.ai_provider = new_provider
                    messagebox.showinfo("Успіх", "API ключ змінено!")
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror(
                        "Помилка", f"Не вдалося підключитися: {str(e)}"
                    )
            else:
                messagebox.showwarning("Увага", "Введіть API ключ")

        provider_combo.bind("<<ComboboxSelected>>", update_default_model)
        key_entry.bind("<KeyRelease>", lambda e: status_label.config(text=""))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x")
        ttk.Button(
            btn_frame,
            text="Оновити моделі",
            command=fetch_models_for_provider,
            width=15,
        ).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Скасувати", command=dialog.destroy, width=12).pack(
            side="left", padx=5
        )
        ttk.Button(btn_frame, text="Зберегти", command=use_direct_key, width=12).pack(
            side="right", padx=5
        )

        update_default_model()
        key_entry.focus()

    def _clear_history(self):
        if messagebox.askyesno("Очистити", "Очистити історію чату?"):
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
