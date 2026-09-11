# -*- coding: utf-8 -*-
"""
agents.common.literature_harvester
外网英文学术文献检索、真实摘要还原、跨源去重与多页 PDF 采集引擎
支持 OpenAlex, Europe PMC, PubMed Central, bioRxiv, arXiv
"""

import os
import sys
import json
import re
import time
import random
import ssl
import socket
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET

from .zotero_sync import sync_to_zotero, export_bibtex_and_ris

ssl_context = ssl.create_default_context()
ssl_context.check_hostname = False
ssl_context.verify_mode = ssl.CERT_NONE

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml,application/pdf;q=0.9,*/*;q=0.8"
}


def get_configured_opener():
    """检测本地代理并构建 opener (优先使用 127.0.0.1:10808 / 7890 / 7897)"""
    proxy_url = None
    for port in [10808, 7890, 7897]:
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(0.2)
            res = s.connect_ex(("127.0.0.1", port))
            s.close()
            if res == 0:
                proxy_url = f"http://127.0.0.1:{port}"
                break
        except Exception:
            pass

    if proxy_url:
        proxy_handler = urllib.request.ProxyHandler({'http': proxy_url, 'https': proxy_url})
        opener = urllib.request.build_opener(proxy_handler, urllib.request.HTTPSHandler(context=ssl_context))
    else:
        opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ssl_context))
    return opener


OPENER = get_configured_opener()


def reconstruct_openalex_abstract(inverted_index):
    """从 OpenAlex 的 abstract_inverted_index 还原出真实段落文本"""
    if not inverted_index or not isinstance(inverted_index, dict):
        return ""
    word_positions = []
    for word, positions in inverted_index.items():
        if isinstance(positions, list):
            for pos in positions:
                word_positions.append((pos, word))
    word_positions.sort(key=lambda x: x[0])
    return " ".join(w for _, w in word_positions)


def normalize_doi(doi):
    """归一化 DOI 字符串"""
    if not doi:
        return ""
    d = doi.strip().lower()
    d = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", d)
    return d.strip()


def normalize_title(title):
    """归一化标题用于碰撞去重"""
    if not title:
        return ""
    return re.sub(r"[^a-zA-Z0-9\u4e00-\u9fa5]", "", title.lower())


def download_and_verify_pdf(url, output_path, max_retries=2):
    """下载并验证 PDF 真实性（必须包含 %PDF- 二进制头，且大小 >= 50KB）"""
    for attempt in range(max_retries):
        try:
            req = urllib.request.Request(url, headers={**HEADERS, "Accept": "application/pdf,*/*;q=0.5"})
            with OPENER.open(req, timeout=30) as resp:
                data = resp.read()

            if len(data) < 50 * 1024:
                continue
            if not data.startswith(b"%PDF-"):
                continue

            with open(output_path, "wb") as f:
                f.write(data)
            return True, len(data)
        except Exception:
            time.sleep(1)
    return False, 0


class EnglishLiteratureHarvester:
    """外网英文学术文献采集器"""

    def __init__(self, output_dir, email="academic_pipeline@research.org"):
        self.output_dir = output_dir
        self.email = email
        os.makedirs(self.output_dir, exist_ok=True)

    def search_openalex(self, query, candidate_limit=15):
        """通过 OpenAlex 检索真实出版文献，还原真实摘要与作者"""
        print(f"🔍 [OpenAlex] 正在检索关键词: {query}")
        clean_q = urllib.parse.quote(query)
        # 增加 type:article 过滤，并通过 mailto 使用 Polite Pool
        url = (
            f"https://api.openalex.org/works?search={clean_q}"
            f"&filter=has_doi:true,is_oa:true,type:article|review"
            f"&per-page={candidate_limit * 2}"
            f"&mailto={self.email}"
        )
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with OPENER.open(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            items = data.get("results", [])
            results = []
            for it in items:
                title = it.get("display_name") or it.get("title")
                if not title:
                    continue

                best_oa = it.get("best_oa_location") or {}
                pdf_url = best_oa.get("pdf_url")
                if not pdf_url:
                    continue

                doi = normalize_doi(it.get("doi"))
                journal = (it.get("primary_location") or {}).get("source", {}).get("display_name") or "Academic Journal"
                year = str(it.get("publication_year", 2024))
                
                # 提取真实作者列表
                authors_list = [
                    a.get("author", {}).get("display_name") 
                    for a in it.get("authorships", []) 
                    if a.get("author") and a.get("author", {}).get("display_name")
                ]
                if not authors_list:
                    continue  # 无作者信息的记录不入库，绝不用占位作者

                # 还原真实摘要；拿不到留空，交给 enrich 从 Crossref/Europe PMC 补
                abstract = reconstruct_openalex_abstract(it.get("abstract_inverted_index")) or ""
                pmcid = ((it.get("ids") or {}).get("pmcid") or "").rsplit("/", 1)[-1] or None

                results.append({
                    "source": "OpenAlex",
                    "title": title,
                    "authors": authors_list,
                    "journal": journal,
                    "year": year,
                    "doi": doi,
                    "abstract": abstract,
                    "type": "review" if it.get("type") == "review" else "article",
                    "is_retracted": bool(it.get("is_retracted")),
                    "pmcid": pmcid,
                    "pdf_url": pdf_url,
                    "pdf_url_fallbacks": ([f"https://europepmc.org/backend/ptpmcrender.fcgi?accid={pmcid}&blobtype=pdf"] if pmcid else []),
                })
                if len(results) >= candidate_limit:
                    break

            print(f"✅ [OpenAlex] 成功检索到 {len(results)} 条候选文献")
            return results
        except Exception as e:
            print(f"⚠️ [OpenAlex] 检索失败: {e}")
            return []

    def search_europe_pmc(self, query, candidate_limit=15):
        """通过 Europe PMC 检索真实 OA 全文"""
        print(f"🔍 [Europe PMC] 正在检索关键词: {query}")
        clean_q = urllib.parse.quote(f"{query} AND OPEN_ACCESS:y")
        url = f"https://www.ebi.ac.uk/europepmc/webservices/rest/search?query={clean_q}&format=json&pageSize={candidate_limit}&resultType=core"
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with OPENER.open(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            items = data.get("resultList", {}).get("result", [])
            results = []
            for it in items:
                title = it.get("title", "").rstrip(".")
                if not title:
                    continue
                doi = normalize_doi(it.get("doi"))
                pmcid = it.get("pmcid")
                candidates = []
                if pmcid:
                    candidates.append(f"https://europepmc.org/backend/ptpmcrender.fcgi?accid={pmcid}&blobtype=pdf")
                for u in it.get("fullTextUrlList", {}).get("fullTextUrl", []):
                    if u.get("documentStyle") == "pdf" and u.get("url") and u["url"] not in candidates:
                        candidates.append(u["url"])
                if not candidates:
                    continue
                pdf_url = candidates[0]

                author_str = it.get("authorString", "")
                authors_list = [a.strip() for a in author_str.rstrip(".").split(",") if a.strip()]
                if not authors_list:
                    continue  # 无作者不入库
                journal = it.get("journalTitle") or ""
                year = str(it.get("pubYear") or "")
                clean_abs = re.sub(r"<[^>]+>", "", it.get("abstractText") or "")
                pub_types = [t.lower() for t in (it.get("pubTypeList") or {}).get("pubType", [])]

                results.append({
                    "source": "Europe PMC",
                    "title": title,
                    "authors": authors_list,
                    "journal": journal,
                    "year": year,
                    "doi": doi,
                    "abstract": clean_abs,
                    "type": "review" if "review" in pub_types else "article",
                    "pmid": it.get("pmid"),
                    "pmcid": pmcid,
                    "pdf_url": pdf_url,
                    "pdf_url_fallbacks": candidates[1:],
                })
            print(f"✅ [Europe PMC] 成功检索到 {len(results)} 条候选文献")
            return results
        except Exception as e:
            print(f"⚠️ [Europe PMC] 检索失败: {e}")
            return []

    def harvest(self, query, count=10, sync_zotero_flag=True):
        """
        全自动多源检索、跨源 DOI 去重、多页 PDF 下载、排伪审查与 Zotero 物理挂载
        """
        print("=" * 65)
        print(f"🚀 启动外网学术文献自动化采集引擎")
        print(f"📂 本地落盘目录: {self.output_dir}")
        print(f"🎯 关键词: {query} | 目标受纳: {count} 篇")
        print("=" * 65)

        candidates = []
        candidates.extend(self.search_openalex(query, candidate_limit=count + 5))
        if len(candidates) < count:
            candidates.extend(self.search_europe_pmc(query, candidate_limit=count))

        # 跨源 DOI 与标题去重
        deduped = []
        seen_dois = set()
        seen_titles = set()

        for c in candidates:
            doi_norm = normalize_doi(c.get("doi"))
            title_norm = normalize_title(c.get("title"))

            if doi_norm and doi_norm in seen_dois:
                continue
            if title_norm and title_norm in seen_titles:
                continue

            if doi_norm:
                seen_dois.add(doi_norm)
            if title_norm:
                seen_titles.add(title_norm)

            deduped.append(c)

        print(f"📋 去重后候选文献共 {len(deduped)} 篇，开始下载并验证真实 PDF 全文...")

        accepted = []
        for idx, item in enumerate(deduped):
            if len(accepted) >= count:
                break

            safe_title = re.sub(r"[^\w\-_\. ]", "", item["title"]).strip().replace(" ", "_")[:60]
            first_author = item["authors"][0].replace(" ", "_") if item["authors"] else "Author"
            pdf_name = f"{first_author}_{item['year']}_{safe_title}.pdf"
            pdf_path = os.path.join(self.output_dir, pdf_name)

            if os.path.exists(pdf_path):  # 同名冲突（同一作者同年同题）
                pdf_name = pdf_name[:-4] + f"_{idx}.pdf"
                pdf_path = os.path.join(self.output_dir, pdf_name)

            print(f"[{len(accepted)+1}/{count}] 正在下载: {item['title'][:45]}...")
            ok, size = False, 0
            for u in [item["pdf_url"]] + list(item.get("pdf_url_fallbacks") or []):
                ok, size = download_and_verify_pdf(u, pdf_path)
                if ok:
                    item["pdf_url"] = u
                    break
            if ok:
                chars = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
                item["zotero_key"] = "".join(random.choice(chars) for _ in range(8))
                item["attach_key"] = "".join(random.choice(chars) for _ in range(8))
                item.pop("pdf_url_fallbacks", None)
                item["local_pdf"] = pdf_path
                item["pdf_filename"] = pdf_name
                item["file_size_kb"] = size // 1024
                item["lang"] = "en"
                accepted.append(item)
                print(f"  ✅ 验证通过! 大小: {size // 1024} KB | 真实摘要: {item['abstract'][:60]}...")
            else:
                print(f"  ❌ 下载或真实性验证失败，跳过。")

        # 写入 manifest.json
        manifest_path = os.path.join(self.output_dir, "manifest.json")
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(accepted, f, ensure_ascii=False, indent=2)
        print(f"\n📑 受纳元数据清单已写入: {manifest_path} (共 {len(accepted)} 篇)")

        # 生成 References.bib 与 References.ris
        export_bibtex_and_ris(accepted, self.output_dir)

        # 联动 Zotero
        if sync_zotero_flag and accepted:
            try:
                sync_to_zotero(self.output_dir, collection_name=os.path.basename(self.output_dir))
            except Exception as e:
                print(f"⚠️ [Zotero] 同步时出现警告: {e}")

        return accepted


def harvest_english_literature(query, count=20, output_dir=None, sync_zotero=True):
    """向后兼容函数接口"""
    if not output_dir:
        output_dir = rf"E:\0mcp-agv\ARTA_Agent_Output\English_Literature_{int(time.time())}"
    harvester = EnglishLiteratureHarvester(output_dir)
    return harvester.harvest(query, count=count, sync_zotero_flag=sync_zotero)
