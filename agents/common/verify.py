# -*- coding: utf-8 -*-
"""
agents.common.verify
外部权威反向核验门禁（Zero-Fake Gate v2）。

旧门禁只检查“本地 PDF 有 %PDF- 头”和“key 在我们自己写的 Zotero 里”，属于自证。
本模块对 manifest 中每条记录向 Crossref / OpenAlex / Europe PMC 反查，并（可选）用 PDF 首页文本防错配。

判定：
  verified    — DOI 可解析 且 标题相似度≥0.90 且 第一作者姓氏匹配 且 |年份差|≤1 且 非撤稿
  suspicious  — DOI 可解析但有一项不一致（标题/作者/年份），保留但标注，builder 打 [待核]
  unverified  — DOI 不存在 / 无 DOI 且任何源都查不到标题
  retracted   — 任一权威源标记撤稿 → 必须剔除

对外接口：
  verify_record(rec, fetch_json=..., pdf_first_page_text=None) -> dict(status, checks, ...)
  verify_manifest(theme_dir, purge=False) -> (manifest, summary)
"""

import difflib
import json
import os
import re
import unicodedata
import urllib.parse

from .http_client import get_json
from .enrich import normalize_doi, fetch_crossref, fetch_openalex, fetch_europepmc

TITLE_SIM_THRESHOLD = 0.90
YEAR_TOLERANCE = 1


def _norm_title(t):
    t = unicodedata.normalize("NFKD", t or "")
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", t.lower())
    return re.sub(r"\s+", " ", t).strip()


def title_similarity(a, b):
    na, nb = _norm_title(a), _norm_title(b)
    if not na or not nb:
        return 0.0
    if na == nb or na in nb or nb in na:
        return 1.0
    return difflib.SequenceMatcher(None, na, nb).ratio()


def _surname(author):
    """从 'Given Family' / 'Family, Given' / '王丽丽' / {'family':..} 中取姓氏（小写、去变音）。"""
    if isinstance(author, dict):
        fam = author.get("family") or author.get("lastName") or ""
    else:
        s = (author or "").strip()
        if not s:
            return ""
        if re.fullmatch(r"[\u4e00-\u9fff]+", s):
            return s[0]
        if "," in s:
            fam = s.split(",", 1)[0]
        else:
            fam = s.split()[-1]
    fam = unicodedata.normalize("NFKD", fam)
    fam = "".join(ch for ch in fam if not unicodedata.combining(ch))
    return re.sub(r"[^a-z\u4e00-\u9fff]", "", fam.lower())


def _first_author(rec):
    a = rec.get("authors") or []
    if isinstance(a, str):
        a = [x.strip() for x in a.split(",") if x.strip()]
    return a[0] if a else ""


def _to_int_year(y):
    try:
        return int(str(y)[:4])
    except (TypeError, ValueError):
        return None


def _search_by_title(title, fetch_json):
    """无 DOI 时按标题在 Crossref 查（仅取第一条），返回 (doi, crossref_record) 或 (None, None)。"""
    if not title:
        return None, None
    q = urllib.parse.quote(title[:200])
    status, data = fetch_json(f"https://api.crossref.org/works?query.title={q}&rows=3&select=DOI,title,author,issued,container-title",
                              rate_key="crossref", min_interval=0.05)
    if status != 200 or not data:
        return None, None
    for it in (data.get("message") or {}).get("items") or []:
        t = (it.get("title") or [""])[0]
        if title_similarity(title, t) >= TITLE_SIM_THRESHOLD:
            return normalize_doi(it.get("DOI")), it
    return None, None


def verify_record(rec, fetch_json=get_json, pdf_first_page_text=None):
    doi = normalize_doi(rec.get("doi"))
    title = rec.get("title", "")
    checks = {}
    result = {"doi": doi, "checks": checks, "status": "unverified", "reasons": []}

    if not doi:
        found_doi, _ = _search_by_title(title, fetch_json)
        if found_doi:
            doi = found_doi
            rec["doi"] = doi
            checks["doi_recovered_by_title"] = True
        else:
            result["reasons"].append("无 DOI 且 Crossref 标题检索无匹配")
            return result

    cr = fetch_crossref(doi, fetch_json)
    oa = fetch_openalex(doi, fetch_json)
    ep = fetch_europepmc(doi, fetch_json)
    checks["crossref_status"] = (cr or {}).get("_status")
    checks["openalex_status"] = (oa or {}).get("_status")
    checks["europepmc_status"] = (ep or {}).get("_status")
    cr = cr if cr and cr.get("_status") == 200 else None
    oa = oa if oa and oa.get("_status") == 200 else None
    ep = ep if ep and ep.get("_status") == 200 else None

    authority = cr or oa or ep
    if not authority:
        result["reasons"].append(f"DOI {doi} 在 Crossref / OpenAlex / Europe PMC 均无法解析")
        return result
    checks["resolved_by"] = "crossref" if cr else "openalex" if oa else "europepmc"

    # 撤稿
    retracted = bool((cr or {}).get("is_retracted") or (oa or {}).get("is_retracted"))
    checks["is_retracted"] = retracted
    if retracted:
        result["status"] = "retracted"
        result["reasons"].append("权威源标记为撤稿")
        return result

    # 标题
    auth_title = authority.get("title", "")
    sim = title_similarity(title, auth_title)
    checks["title_similarity"] = round(sim, 3)
    checks["authority_title"] = auth_title

    # 第一作者
    auth_authors = authority.get("authors") or []
    auth_first = _surname(auth_authors[0]) if auth_authors else ""
    mine_first = _surname(_first_author(rec))
    checks["first_author_match"] = bool(auth_first and mine_first and (auth_first == mine_first or auth_first in mine_first or mine_first in auth_first))
    checks["authority_first_author"] = auth_first
    checks["manifest_first_author"] = mine_first

    # 年份
    ay, my = _to_int_year(authority.get("year")), _to_int_year(rec.get("year"))
    checks["year_delta"] = (abs(ay - my) if ay and my else None)
    checks["authority_year"] = ay

    # 期刊
    checks["journal_present"] = bool(authority.get("journal"))

    # PDF 首页防错配（可选）
    if pdf_first_page_text is not None:
        txt = pdf_first_page_text.lower()
        checks["pdf_mentions_doi"] = doi in txt
        checks["pdf_mentions_title"] = _norm_title(title)[:40] in _norm_title(pdf_first_page_text)
        checks["pdf_matches"] = checks["pdf_mentions_doi"] or checks["pdf_mentions_title"]
        blocked = re.search(r"access denied|purchase (this )?article|institutional login|sign in to continue", txt)
        checks["pdf_is_paywall_page"] = bool(blocked)

    problems = []
    if sim < TITLE_SIM_THRESHOLD:
        problems.append(f"标题相似度 {sim:.2f} < {TITLE_SIM_THRESHOLD}")
    if not checks["first_author_match"]:
        problems.append(f"第一作者不匹配 ({mine_first!r} vs {auth_first!r})")
    if checks["year_delta"] is not None and checks["year_delta"] > YEAR_TOLERANCE:
        problems.append(f"年份差 {checks['year_delta']} > {YEAR_TOLERANCE}")
    if pdf_first_page_text is not None and (not checks.get("pdf_matches") or checks.get("pdf_is_paywall_page")):
        problems.append("PDF 首页与元数据不符或为付费墙页")

    result["reasons"] = problems
    result["status"] = "verified" if not problems else "suspicious"
    return result


def pdf_first_page_text(pdf_path, max_chars=6000):
    """只读第 1 页文本用于防错配；PyMuPDF 不可用时返回 None（跳过该项检查）。"""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return None
    if not pdf_path or not os.path.exists(pdf_path):
        return None
    try:
        with fitz.open(pdf_path) as d:
            if d.page_count == 0:
                return ""
            return d[0].get_text("text")[:max_chars]
    except Exception:
        return None


def verify_manifest(theme_dir, purge=False, check_pdf=True, verbose=True):
    """核验整个 manifest。purge=True 时把 retracted / unverified 条目移出 manifest（写入 manifest.rejected.json）。"""
    mf = os.path.join(theme_dir, "manifest.json")
    with open(mf, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    kept, rejected = [], []
    summary = {"verified": 0, "suspicious": 0, "unverified": 0, "retracted": 0}
    for i, rec in enumerate(manifest, 1):
        pdf_txt = None
        if check_pdf:
            p = rec.get("local_pdf") or (os.path.join(theme_dir, rec["pdf_filename"]) if rec.get("pdf_filename") else "")
            pdf_txt = pdf_first_page_text(p)
        v = verify_record(rec, pdf_first_page_text=pdf_txt)
        rec["verification"] = {"status": v["status"], "reasons": v["reasons"], "checks": v["checks"]}
        summary[v["status"]] += 1
        if verbose:
            mark = {"verified": "✅", "suspicious": "⚠️", "unverified": "❌", "retracted": "🛑"}[v["status"]]
            print(f"{mark} [{i}/{len(manifest)}] {v['status']:<11} {rec.get('doi') or '(no doi)'} — {rec.get('title', '')[:60]}")
            for r in v["reasons"]:
                print(f"      · {r}")
        if purge and v["status"] in ("unverified", "retracted"):
            rejected.append(rec)
        else:
            kept.append(rec)

    if purge and rejected:
        with open(os.path.join(theme_dir, "manifest.rejected.json"), "w", encoding="utf-8") as f:
            json.dump(rejected, f, ensure_ascii=False, indent=2)
        manifest = kept
    with open(mf, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    if verbose:
        print(f"\n核验汇总: {summary}" + (f"，已移出 {len(rejected)} 条 → manifest.rejected.json" if purge and rejected else ""))
    return manifest, summary


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("usage: python -m agents.common.verify <theme_dir> [--purge] [--no-pdf]")
        sys.exit(1)
    verify_manifest(sys.argv[1], purge="--purge" in sys.argv, check_pdf="--no-pdf" not in sys.argv)
