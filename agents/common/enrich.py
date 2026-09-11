# -*- coding: utf-8 -*-
"""
agents.common.enrich
多源元数据聚合：不读 PDF，只用权威 API 把每篇文献的写作素材“加厚”。

来源与字段：
  Crossref      → title / authors / container-title / volume / issue / page / year / type / abstract(JATS) / is_retracted(update-to) / references_count
  OpenAlex      → type / is_retracted / cited_by_count / concepts / keywords / abstract(inverted index) / oa pdf_url / pmcid
  Europe PMC    → abstract / mesh / pmcid / language / pubType / citedByCount
  Europe PMC fullTextXML (JATS, 仅 PMC OA) → results_excerpt / conclusion_excerpt / key_sentences（含数字的句子）
  Semantic Scholar → tldr（无 key 时 429 则跳过）

对外接口：
  enrich_record(rec, fetch_json=get_json, fetch_text=get_text) -> rec（原地补字段并返回）
  enrich_manifest(theme_dir) -> manifest
"""

import json
import os
import re
import urllib.parse
import xml.etree.ElementTree as ET

from .http_client import get_json, get_text

S2_API_KEY = os.environ.get("S2_API_KEY", "")

NUM_SENT_RE = re.compile(
    r"[^.。!?]*?(?:\d+(?:\.\d+)?\s*(?:%|nm|Å|kcal|kJ|℃|°C|K\b|ns\b|ps\b|μs|ms\b|fold|倍|mg|μg|µg|mM|μM|µM|nM|kDa|mmHg|p\s*[<=>]|R²|r\s*=|n\s*=|±))[^.。!?]*[.。!?]",
    re.IGNORECASE,
)

RETRACTION_TYPES = {"retraction", "withdrawal", "removal"}


# --------------------------------------------------------------------------
# 工具
# --------------------------------------------------------------------------
def normalize_doi(doi):
    if not doi:
        return ""
    d = str(doi).strip().lower()
    d = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", d)
    d = re.sub(r"^doi:\s*", "", d)
    return d


def strip_tags(s):
    if not s:
        return ""
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def reconstruct_openalex_abstract(inv):
    if not inv or not isinstance(inv, dict):
        return ""
    pos = []
    for w, ps in inv.items():
        if isinstance(ps, list):
            pos.extend((p, w) for p in ps)
    pos.sort()
    return " ".join(w for _, w in pos)


def _sentences(text):
    text = re.sub(r"\s+", " ", text or "")
    return [s.strip() for s in re.split(r"(?<=[.。!?])\s+", text) if s.strip()]


def extract_key_sentences(text, limit=12):
    """抽取含数值/统计量的句子，这是 Antigravity 引注数值时唯一允许的来源。"""
    out, seen = [], set()
    for s in _sentences(text):
        if NUM_SENT_RE.search(s) and 30 <= len(s) <= 400:
            k = s.lower()
            if k not in seen:
                seen.add(k)
                out.append(s)
        if len(out) >= limit:
            break
    return out


# --------------------------------------------------------------------------
# 各源
# --------------------------------------------------------------------------
def fetch_crossref(doi, fetch_json=get_json):
    if not doi:
        return None
    status, data = fetch_json(f"https://api.crossref.org/works/{urllib.parse.quote(doi, safe='')}",
                              rate_key="crossref", min_interval=0.05)
    if status != 200 or not data:
        return {"_status": status}
    m = data.get("message", {})
    authors = []
    for a in m.get("author", []) or []:
        if a.get("family"):
            authors.append({"family": a["family"], "given": a.get("given", ""), "orcid": a.get("ORCID", "")})
        elif a.get("name"):
            authors.append({"family": a["name"], "given": "", "literal": True})
    year = None
    for k in ("published-print", "published-online", "issued", "created"):
        dp = (m.get(k) or {}).get("date-parts") or [[None]]
        if dp and dp[0] and dp[0][0]:
            year = dp[0][0]
            break
    retracted = any((u.get("type") or "").lower() in RETRACTION_TYPES for u in m.get("update-to", []) or [])
    return {
        "_status": 200,
        "title": (m.get("title") or [""])[0],
        "authors": authors,
        "journal": (m.get("container-title") or [""])[0],
        "journal_short": (m.get("short-container-title") or [""])[0],
        "issn": m.get("ISSN") or [],
        "publisher": m.get("publisher", ""),
        "volume": m.get("volume", ""),
        "issue": m.get("issue", ""),
        "pages": m.get("page", ""),
        "year": year,
        "type": m.get("type", ""),
        "abstract": strip_tags(m.get("abstract", "")),
        "is_retracted": retracted,
        "references_count": m.get("reference-count", 0),
        "cited_by_count": m.get("is-referenced-by-count", 0),
        "license": ((m.get("license") or [{}])[0]).get("URL", ""),
    }


def fetch_openalex(doi, fetch_json=get_json):
    if not doi:
        return None
    sel = "id,doi,title,type,publication_year,is_retracted,cited_by_count,primary_location,best_oa_location,authorships,concepts,keywords,ids,abstract_inverted_index,referenced_works_count,language"
    status, data = fetch_json(f"https://api.openalex.org/works/doi:{urllib.parse.quote(doi, safe='')}?select={sel}&mailto=academic_pipeline@research.org",
                              rate_key="openalex", min_interval=0.1)
    if status != 200 or not data:
        return {"_status": status}
    src = (data.get("primary_location") or {}).get("source") or {}
    concepts = [c["display_name"] for c in (data.get("concepts") or []) if c.get("score", 0) >= 0.4][:10]
    kws = [k["display_name"] for k in (data.get("keywords") or [])][:10]
    return {
        "_status": 200,
        "openalex_id": data.get("id", ""),
        "title": data.get("title", ""),
        "type": data.get("type", ""),
        "year": data.get("publication_year"),
        "is_retracted": bool(data.get("is_retracted")),
        "cited_by_count": data.get("cited_by_count", 0),
        "journal": src.get("display_name", ""),
        "journal_in_doaj": bool(src.get("is_in_doaj")),
        "pdf_url": (data.get("best_oa_location") or {}).get("pdf_url") or "",
        "pmcid": (data.get("ids") or {}).get("pmcid", ""),
        "pmid": (data.get("ids") or {}).get("pmid", ""),
        "abstract": reconstruct_openalex_abstract(data.get("abstract_inverted_index")),
        "concepts": concepts,
        "keywords": kws,
        "authors": [a.get("author", {}).get("display_name", "") for a in data.get("authorships") or [] if a.get("author")],
        "language": data.get("language", ""),
        "references_count": data.get("referenced_works_count", 0),
    }


def fetch_europepmc(doi, fetch_json=get_json):
    if not doi:
        return None
    q = urllib.parse.quote(f'DOI:"{doi}"')
    status, data = fetch_json(f"https://www.ebi.ac.uk/europepmc/webservices/rest/search?query={q}&format=json&resultType=core&pageSize=1",
                              rate_key="epmc", min_interval=0.1)
    if status != 200 or not data:
        return {"_status": status}
    res = (data.get("resultList") or {}).get("result") or []
    if not res:
        return {"_status": 404}
    r = res[0]
    ji = r.get("journalInfo") or {}
    mesh = [m.get("descriptorName") for m in ((r.get("meshHeadingList") or {}).get("meshHeading") or []) if m.get("descriptorName")]
    pubtypes = (r.get("pubTypeList") or {}).get("pubType") or []
    return {
        "_status": 200,
        "pmid": r.get("pmid", ""),
        "pmcid": r.get("pmcid", ""),
        "title": (r.get("title") or "").rstrip("."),
        "abstract": strip_tags(r.get("abstractText", "")),
        "journal": (ji.get("journal") or {}).get("title", ""),
        "volume": ji.get("volume", ""),
        "issue": ji.get("issue", ""),
        "pages": r.get("pageInfo", ""),
        "year": ji.get("yearOfPublication") or (int(r["pubYear"]) if str(r.get("pubYear", "")).isdigit() else None),
        "mesh": mesh,
        "pub_types": pubtypes,
        "is_review": any("review" in p.lower() for p in pubtypes),
        "language": r.get("language", ""),
        "cited_by_count": r.get("citedByCount", 0),
        "in_pmc": r.get("inPMC") == "Y",
        "has_fulltext_xml": bool(r.get("pmcid")) and r.get("inEPMC") == "Y",
    }


def _sec_text(sec):
    parts = []
    for p in sec.iter("p"):
        t = "".join(p.itertext())
        t = re.sub(r"\s+", " ", t).strip()
        if t:
            parts.append(t)
    return " ".join(parts)


def _find_sections(root, want):
    """按 sec-type 或 <title> 匹配 JATS 章节；want 为小写关键词元组。"""
    hits = []
    for sec in root.iter("sec"):
        st = (sec.get("sec-type") or "").lower()
        title_el = sec.find("title")
        title = ("".join(title_el.itertext()) if title_el is not None else "").strip().lower()
        if any(w in st for w in want) or any(title.startswith(w) or title == w for w in want):
            hits.append(sec)
    return hits


def fetch_jats_excerpts(pmcid, fetch_text=get_text, max_chars=1200):
    """从 Europe PMC fullTextXML 抽 Results / Conclusion 段落 + 含数值句子。无 PMC 全文则返回 None。"""
    if not pmcid:
        return None
    status, xml = fetch_text(f"https://www.ebi.ac.uk/europepmc/webservices/rest/{pmcid}/fullTextXML",
                             rate_key="epmc", min_interval=0.1, timeout=30)
    if status != 200 or not xml.strip():
        return {"_status": status}
    try:
        root = ET.fromstring(xml.encode("utf-8"))
    except ET.ParseError:
        return {"_status": "parse_error"}
    # 去掉命名空间前缀，简化匹配
    for el in root.iter():
        if "}" in el.tag:
            el.tag = el.tag.split("}", 1)[1]

    body = root.find(".//body")
    if body is None:
        return {"_status": "no_body"}

    def grab(want, take_first_n_paras=2):
        secs = _find_sections(body, want)
        out = []
        for sec in secs:
            # 直接子段落优先；PLOS/eLife 常把 Results 拆成嵌套子节，此时取后代段落
            paras = [re.sub(r"\s+", " ", "".join(p.itertext())).strip() for p in sec.findall("p")]
            paras = [p for p in paras if p]
            if not paras:
                paras = [re.sub(r"\s+", " ", "".join(p.itertext())).strip() for p in sec.iter("p")]
                paras = [p for p in paras if p]
            out.extend(paras[:take_first_n_paras])
        text = " ".join(out)
        return text[:max_chars]

    results = grab(("results", "result", "findings"), take_first_n_paras=2)
    conclusion = grab(("conclusion", "conclusions", "concluding", "summary"), take_first_n_paras=3)
    discussion = grab(("discussion",), take_first_n_paras=1) if not conclusion else ""

    full_body_text = _sec_text(body)
    key_sents = extract_key_sentences(" ".join([results, conclusion, discussion, full_body_text]), limit=12)

    return {
        "_status": 200,
        "results_excerpt": results,
        "conclusion_excerpt": conclusion or discussion,
        "key_sentences": key_sents,
        "body_words": len(full_body_text.split()),
    }


def fetch_s2_tldr(doi, fetch_json=get_json):
    if not doi:
        return None
    headers = {"x-api-key": S2_API_KEY} if S2_API_KEY else None
    status, data = fetch_json(f"https://api.semanticscholar.org/graph/v1/paper/DOI:{urllib.parse.quote(doi, safe='')}?fields=tldr,citationCount,publicationTypes",
                              headers=headers, rate_key="s2", min_interval=1.1 if not S2_API_KEY else 0.1, retries=1)
    if status != 200 or not data:
        return {"_status": status}
    return {
        "_status": 200,
        "tldr": ((data.get("tldr") or {}).get("text") or "").strip(),
        "cited_by_count": data.get("citationCount", 0),
        "pub_types": data.get("publicationTypes") or [],
    }


# --------------------------------------------------------------------------
# 聚合
# --------------------------------------------------------------------------
def _first(*vals):
    for v in vals:
        if v:
            return v
    return ""


def classify_type(cr, oa, ep):
    t_oa = (oa or {}).get("type", "")
    t_cr = (cr or {}).get("type", "")
    if (ep or {}).get("is_review") or t_oa == "review":
        return "review"
    if t_oa in ("preprint",) or t_cr in ("posted-content",):
        return "preprint"
    if t_oa in ("article",) or t_cr in ("journal-article",):
        return "article"
    return t_oa or t_cr or "unknown"


def enrich_record(rec, fetch_json=get_json, fetch_text=get_text, want_fulltext=True):
    """原地补齐 rec 的写作素材字段。rec 至少需要 doi（或 title）。"""
    doi = normalize_doi(rec.get("doi"))
    rec["doi"] = doi
    sources = {}

    cr = fetch_crossref(doi, fetch_json) if doi else None
    oa = fetch_openalex(doi, fetch_json) if doi else None
    ep = fetch_europepmc(doi, fetch_json) if doi else None
    sources["crossref"] = (cr or {}).get("_status")
    sources["openalex"] = (oa or {}).get("_status")
    sources["europepmc"] = (ep or {}).get("_status")
    cr = cr if cr and cr.get("_status") == 200 else None
    oa = oa if oa and oa.get("_status") == 200 else None
    ep = ep if ep and ep.get("_status") == 200 else None

    # 权威字段以 Crossref 为准，其余回退
    if cr:
        rec["title"] = cr["title"] or rec.get("title", "")
        if cr["authors"]:
            rec["authors"] = [(f"{a['given']} {a['family']}".strip() if not a.get("literal") else a["family"]) for a in cr["authors"]]
            rec["authors_structured"] = cr["authors"]
        rec["journal"] = _first(cr["journal"], rec.get("journal"))
        rec["journal_short"] = cr["journal_short"]
        rec["issn"] = cr["issn"]
        rec["publisher"] = cr["publisher"]
        rec["volume"] = _first(cr["volume"], rec.get("volume"))
        rec["issue"] = _first(cr["issue"], rec.get("issue"))
        rec["pages"] = _first(cr["pages"], rec.get("pages"))
        if cr["year"]:
            rec["year"] = cr["year"]
    if ep:
        rec["pmid"] = _first(ep["pmid"], rec.get("pmid"))
        rec["pmcid"] = _first(ep["pmcid"], rec.get("pmcid"))
        rec["mesh"] = ep["mesh"]
        rec["pub_types"] = ep["pub_types"]
        rec["volume"] = _first(rec.get("volume"), ep["volume"])
        rec["issue"] = _first(rec.get("issue"), ep["issue"])
        rec["pages"] = _first(rec.get("pages"), ep["pages"])
        if not rec.get("year") and ep["year"]:
            rec["year"] = ep["year"]
    if oa:
        rec["openalex_id"] = oa["openalex_id"]
        rec["concepts"] = oa["concepts"]
        rec["keywords"] = oa["keywords"]
        rec["pmcid"] = _first(rec.get("pmcid"), oa["pmcid"].replace("https://www.ncbi.nlm.nih.gov/pmc/articles/", "").strip("/"))
        rec["pdf_url"] = _first(rec.get("pdf_url"), oa["pdf_url"])
        rec["journal_in_doaj"] = oa["journal_in_doaj"]
        if not rec.get("authors") and oa["authors"]:
            rec["authors"] = oa["authors"]

    rec["type"] = classify_type(cr, oa, ep)
    rec["is_retracted"] = bool((cr or {}).get("is_retracted") or (oa or {}).get("is_retracted"))
    rec["cited_by_count"] = max((cr or {}).get("cited_by_count", 0), (oa or {}).get("cited_by_count", 0), (ep or {}).get("cited_by_count", 0))
    rec["references_count"] = max((cr or {}).get("references_count", 0), (oa or {}).get("references_count", 0))

    # 摘要回退链：Europe PMC（最干净）→ Crossref → OpenAlex → 原有
    old_abs = rec.get("abstract", "")
    if old_abs and re.match(r"^(Indexed scholarly article|Scholarly (article|research) published)", old_abs):
        old_abs = ""
    rec["abstract"] = _first((ep or {}).get("abstract"), (cr or {}).get("abstract"), (oa or {}).get("abstract"), old_abs)
    rec["abstract_source"] = "europepmc" if (ep or {}).get("abstract") else "crossref" if (cr or {}).get("abstract") else "openalex" if (oa or {}).get("abstract") else ("original" if old_abs else "")

    # JATS 全文摘录（仅 PMC OA）
    rec["results_excerpt"] = ""
    rec["conclusion_excerpt"] = ""
    rec["key_sentences"] = []
    if want_fulltext and rec.get("pmcid"):
        pm = rec["pmcid"] if str(rec["pmcid"]).upper().startswith("PMC") else f"PMC{rec['pmcid']}"
        jx = fetch_jats_excerpts(pm, fetch_text)
        sources["jats"] = (jx or {}).get("_status")
        if jx and jx.get("_status") == 200:
            rec["results_excerpt"] = jx["results_excerpt"]
            rec["conclusion_excerpt"] = jx["conclusion_excerpt"]
            rec["key_sentences"] = jx["key_sentences"]
            rec["fulltext_words"] = jx["body_words"]
    if not rec["key_sentences"]:
        rec["key_sentences"] = extract_key_sentences(rec.get("abstract", ""), limit=6)

    # TL;DR
    s2 = fetch_s2_tldr(doi, fetch_json) if doi else None
    sources["s2"] = (s2 or {}).get("_status")
    rec["tldr"] = (s2 or {}).get("tldr", "") if s2 and s2.get("_status") == 200 else ""
    if not rec["tldr"]:
        sents = _sentences(rec.get("abstract", ""))
        rec["tldr"] = sents[-1] if sents else ""
        rec["tldr_source"] = "abstract_last_sentence" if sents else ""
    else:
        rec["tldr_source"] = "semantic_scholar"

    rec.setdefault("lang", "zh" if re.search(r"[\u4e00-\u9fff]", rec.get("title", "")) else "en")
    rec["enrich_sources"] = sources
    return rec


def enrich_manifest(theme_dir, want_fulltext=True, verbose=True):
    mf = os.path.join(theme_dir, "manifest.json")
    with open(mf, "r", encoding="utf-8") as f:
        manifest = json.load(f)
    for i, rec in enumerate(manifest, 1):
        if verbose:
            print(f"[{i}/{len(manifest)}] enrich {rec.get('doi') or rec.get('title', '')[:50]}")
        try:
            enrich_record(rec, want_fulltext=want_fulltext)
        except Exception as e:  # 单条失败不阻断
            rec["enrich_error"] = str(e)
            if verbose:
                print(f"   ⚠️ {e}")
        if verbose:
            print(f"   type={rec.get('type')} abs={rec.get('abstract_source')} jats={rec.get('enrich_sources', {}).get('jats')} key_sents={len(rec.get('key_sentences', []))}")
    with open(mf, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return manifest


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("usage: python -m agents.common.enrich <theme_dir> [--no-fulltext]")
        sys.exit(1)
    enrich_manifest(sys.argv[1], want_fulltext="--no-fulltext" not in sys.argv)
