# -*- coding: utf-8 -*-
"""
离线测试：用真实 API 响应快照（tests/fixtures/*.json）注入 fetcher，
覆盖 enrich / verify / digest 全链路。快照来自 2026-09-11 的真实调用。
"""
import json, os, sys, tempfile, shutil
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from agents.common import enrich, verify, digest

FX = os.path.join(ROOT, "tests", "fixtures")

def _load(name):
    with open(os.path.join(FX, name), encoding="utf-8") as f:
        return f.read()

# 按 (host, doi/pmcid 片段) 路由到 fixture；未命中 → 404
import urllib.parse
def _route(url):
    u = urllib.parse.unquote(url).lower()
    if "pcbi.1009079" in u:
        if "crossref" in u: return "json", "crossref_pcbi1009079.json"
        if "openalex" in u: return "json", "openalex_pcbi1009079.json"
        if "europepmc" in u and "search" in u: return "json", "epmc_pcbi1009079.json"
    if "pmc8224937/fulltextxml" in u: return "text", "jats_PMC8224937.xml"
    if "s0140-6736(97)11096-0" in u:
        if "openalex" in u: return "json", "openalex_wakefield.json"
        if "crossref" in u: return "json", "crossref_wakefield.json"
    return None, None

def fake_json(url, **kw):
    kind, fn = _route(url)
    return (200, json.loads(_load(fn))) if kind == "json" else (404, None)

def fake_text(url, **kw):
    kind, fn = _route(url)
    return (200, _load(fn)) if kind == "text" else (404, "")

def test_enrich_real_plos():
    rec = {"doi": "https://doi.org/10.1371/journal.pcbi.1009079", "title": "Mechanism of collagen folding propagation studied by Molecular Dynamics simulations",
           "authors": ["Morteza Rad-Malekshahi"], "year": "2021", "abstract": "Indexed scholarly article in OpenAlex database.", "zotero_key": "7WUA8PLB"}
    enrich.enrich_record(rec, fetch_json=fake_json, fetch_text=fake_text)
    assert rec["doi"] == "10.1371/journal.pcbi.1009079"
    assert rec["authors"][0] == "Julian Hartmann", rec["authors"]        # Crossref 覆盖了 manifest 里错误的作者
    assert rec["journal"] == "PLoS Computational Biology"
    assert rec["volume"] == "17" and rec["issue"] == "6" and rec["pages"] == "e1009079"
    assert rec["type"] == "article" and rec["is_retracted"] is False
    assert rec["abstract_source"] == "europepmc" and "triple helical" in rec["abstract"]
    assert rec["pmcid"] == "PMC8224937"
    assert "Protein Folding" in rec["mesh"]
    assert rec["results_excerpt"] and rec["conclusion_excerpt"], rec["enrich_sources"]
    assert any("75 ns" in s for s in rec["key_sentences"]), rec["key_sentences"][:3]
    assert rec["tldr"]  # 无 S2 时回退摘要末句
    print("  enrich ok:", rec["type"], rec["abstract_source"], f"{len(rec['key_sentences'])} facts, jats={rec['enrich_sources']['jats']}")

def test_verify_states():
    # 1) 真实文献 + 元数据一致 → verified
    good = {"doi": "10.1371/journal.pcbi.1009079", "title": "Mechanism of collagen folding propagation studied by Molecular Dynamics simulations",
            "authors": ["Julian Hartmann", "Martin Zacharias"], "year": 2021}
    v = verify.verify_record(good, fetch_json=fake_json)
    assert v["status"] == "verified", v
    # 2) 真实 DOI 但作者/年份被篡改（init_collagen_manifest 里的情形） → suspicious
    bad_meta = {"doi": "10.1371/journal.pcbi.1009079", "title": "Mechanism of collagen folding propagation studied by Molecular Dynamics simulations",
                "authors": ["Morteza Rad-Malekshahi"], "year": 2021}
    v = verify.verify_record(bad_meta, fetch_json=fake_json)
    assert v["status"] == "suspicious" and any("第一作者" in r for r in v["reasons"]), v
    # 3) 伪造知网 DOI（Crossref 真实返回 Resource not found） → unverified
    fake = {"doi": "10.13995/j.cnki.11-1802/ts.033104", "title": "鱼皮胶原蛋白肽-钙螯合物结合位点分子对接及热变性模拟", "authors": ["刘海燕"], "year": 2023}
    v = verify.verify_record(fake, fetch_json=fake_json)
    assert v["status"] == "unverified", v
    # 4) 撤稿 → retracted
    wf = {"doi": "10.1016/S0140-6736(97)11096-0", "title": "Ileal-lymphoid-nodular hyperplasia, non-specific colitis, and pervasive developmental disorder in children", "authors": ["AJ Wakefield"], "year": 1998}
    v = verify.verify_record(wf, fetch_json=fake_json)
    assert v["status"] == "retracted", v
    # 5) PDF 首页为付费墙 → suspicious
    v = verify.verify_record(dict(good), fetch_json=fake_json, pdf_first_page_text="Access Denied. Purchase this article to continue.")
    assert v["status"] == "suspicious" and any("PDF" in r for r in v["reasons"]), v
    # 6) PDF 首页含 DOI → 通过
    v = verify.verify_record(dict(good), fetch_json=fake_json, pdf_first_page_text="PLOS Comput Biol ... https://doi.org/10.1371/journal.pcbi.1009079 ...")
    assert v["status"] == "verified", v
    print("  verify ok: verified / suspicious / unverified / retracted / paywall 全部判定正确")

def test_digest_end_to_end():
    d = tempfile.mkdtemp()
    try:
        recs = [
            {"doi": "10.1371/journal.pcbi.1009079", "title": "Mechanism of collagen folding propagation studied by Molecular Dynamics simulations",
             "authors": ["Julian Hartmann", "Martin Zacharias"], "year": 2021, "zotero_key": "7WUA8PLB"},
            {"doi": "10.13995/j.cnki.11-1802/ts.033104", "title": "鱼皮胶原蛋白肽-钙螯合物结合位点分子对接及热变性模拟", "authors": ["刘海燕"], "year": 2023, "zotero_key": "6LIUHY64", "lang": "zh"},
            {"doi": "10.1016/S0140-6736(97)11096-0", "title": "Ileal-lymphoid-nodular hyperplasia, non-specific colitis, and pervasive developmental disorder in children", "authors": ["AJ Wakefield"], "year": 1998, "zotero_key": "WAKE1998"},
        ]
        for r in recs:
            enrich.enrich_record(r, fetch_json=fake_json, fetch_text=fake_text)
            v = verify.verify_record(r, fetch_json=fake_json)
            r["verification"] = {"status": v["status"], "reasons": v["reasons"], "checks": v["checks"]}
        json.dump(recs, open(os.path.join(d, "manifest.json"), "w", encoding="utf-8"), ensure_ascii=False)
        out = digest.write_digest(d)
        files = sorted(os.listdir(out))
        assert files == ["7WUA8PLB.md", "INDEX.md", "literature_matrix.md"], files   # 伪造+撤稿被排除
        card = open(os.path.join(out, "7WUA8PLB.md"), encoding="utf-8").read()
        assert len(card) <= digest.MAX_CARD_CHARS + 200, len(card)
        assert "Quotable facts" in card and "75 ns" in card
        idx = open(os.path.join(out, "INDEX.md"), encoding="utf-8").read()
        assert "unverified" in idx and "`7WUA8PLB`" in idx and "6LIUHY64" not in idx
        print(f"  digest ok: card {len(card)} chars, INDEX {len(idx)} chars")
    finally:
        shutil.rmtree(d)

if __name__ == "__main__":
    test_enrich_real_plos(); test_verify_states(); test_digest_end_to_end()
    print("\n🎉 all enrich/verify/digest tests passed")
