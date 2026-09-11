#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
english_downloader.py
English Academic Agent 核心文献检索与下载入口
基于 agents.common.literature_harvester 共享模块构建
"""

import os
import sys
import argparse

WORKSPACE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if WORKSPACE_ROOT not in sys.path:
    sys.path.insert(0, WORKSPACE_ROOT)

from agents.common.literature_harvester import (
    EnglishLiteratureHarvester,
    harvest_english_literature,
    reconstruct_openalex_abstract,
    normalize_doi,
    normalize_title,
    download_and_verify_pdf
)
from agents.common.zotero_sync import sync_to_zotero, export_bibtex_and_ris

__all__ = [
    "EnglishLiteratureHarvester",
    "harvest_english_literature",
    "reconstruct_openalex_abstract",
    "normalize_doi",
    "normalize_title",
    "download_and_verify_pdf",
    "sync_to_zotero",
    "export_bibtex_and_ris"
]

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="外网真实多页英文学术文献采集与 Zotero 挂载工具")
    parser.add_argument("--query", default="antimicrobial peptides machine learning metagenomics", help="检索关键词")
    parser.add_argument("--count", type=int, default=10, help="计划下载篇数")
    parser.add_argument("--out-dir", default=None, help="落盘目录")
    parser.add_argument("--sync-zotero", action="store_true", help="是否自动挂载到本地 Zotero 数据库")
    args = parser.parse_args()

    items = harvest_english_literature(args.query, count=args.count, output_dir=args.out_dir, sync_zotero=args.sync_zotero)
    print(f"🎉 成功完成采集，共受纳 {len(items)} 篇学术文献。")
