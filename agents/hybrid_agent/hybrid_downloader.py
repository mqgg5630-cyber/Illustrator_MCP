#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hybrid_downloader.py - 中英双轨混合文献下载与 Zotero 挂载引擎
1. 规范化管理 CNKI 中文核心与外网 SCI 顶刊文献；
2. 严格遵循零 C 盘占用规则：PDF 实体落盘于 E:/ozotero/storage 与主题目录；
3. 向本地 E:/ozotero/zotero.sqlite 建立实体挂载条目与 📎 附件；
4. 导出双语 manifest.json、References.bib 与 References.ris。
基于 agents.common.zotero_sync 共享引擎构建。
"""

import os
import sys
import json
import tempfile

# 强制重定向临时目录至 E 盘
SCRATCH_DIR = os.environ.get("HUB_SCRATCH_DIR", r"E:\0mcp-agv\scratch")
try:
    if not os.path.isabs(SCRATCH_DIR):
        raise OSError("scratch dir must be absolute on this platform")
    os.makedirs(SCRATCH_DIR, exist_ok=True)
    os.environ["TEMP"] = SCRATCH_DIR
    os.environ["TMP"] = SCRATCH_DIR
    tempfile.tempdir = SCRATCH_DIR
except OSError:
    pass

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

WORKSPACE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if WORKSPACE_ROOT not in sys.path:
    sys.path.insert(0, WORKSPACE_ROOT)

from agents.common.zotero_sync import sync_to_zotero, export_bibtex_and_ris, ZOTERO_SQLITE, ZOTERO_STORAGE

def sync_hybrid_to_zotero(topic_dir, collection_name=None):
    """将混合文献条目与本地 Zotero 数据库建立物理附件关联"""
    return sync_to_zotero(topic_dir, collection_name=collection_name, zotero_db_path=ZOTERO_SQLITE, zotero_storage_dir=ZOTERO_STORAGE)

__all__ = [
    "sync_hybrid_to_zotero",
    "sync_to_zotero",
    "export_bibtex_and_ris",
    "ZOTERO_SQLITE",
    "ZOTERO_STORAGE"
]

if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = sys.argv[1]
        col = sys.argv[2] if len(sys.argv) > 2 else None
        sync_hybrid_to_zotero(target, col)
    else:
        print("用法: python hybrid_downloader.py <topic_dir> [collection_name]")
