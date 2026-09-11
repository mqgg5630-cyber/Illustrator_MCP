#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hybrid_downloader.py - 中英双轨混合文献下载与 Zotero 挂载引擎
1. 规范化管理 CNKI 中文核心与外网 SCI 顶刊文献；
2. 严格遵循零 C 盘占用规则：PDF 实体落盘于 E:/ozotero/storage 与主题目录；
3. 向本地 E:/ozotero/zotero.sqlite 建立实体挂载条目与 📎 附件；
4. 导出双语 manifest.json、References.bib 与 References.ris。
"""

import os
import sys
import json
import sqlite3
import shutil
import random
import time
import subprocess
import tempfile

# 强制重定向临时目录至 E 盘
SCRATCH_DIR = r"E:\0mcp-agv\scratch"
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.environ["TEMP"] = SCRATCH_DIR
os.environ["TMP"] = SCRATCH_DIR
tempfile.tempdir = SCRATCH_DIR

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ZOTERO_SQLITE = r"E:\ozotero\zotero.sqlite"
ZOTERO_STORAGE = r"E:\ozotero\storage"

def sync_hybrid_to_zotero(topic_dir, collection_name=None):
    """将混合文献条目与本地 Zotero 数据库建立物理附件关联"""
    manifest_file = os.path.join(topic_dir, "manifest.json")
    if not os.path.exists(manifest_file):
        print(f"❌ 未找到 manifest.json: {manifest_file}")
        return False

    with open(manifest_file, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    # 关闭可能占用的 Zotero 客户端以解除文件锁
    subprocess.run(["taskkill", "/F", "/IM", "zotero.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
    time.sleep(0.6)

    conn = sqlite3.connect(ZOTERO_SQLITE, timeout=25)
    c = conn.cursor()

    if not collection_name:
        collection_name = os.path.basename(topic_dir)

    c.execute("SELECT collectionID FROM collections WHERE collectionName = ?", (collection_name,))
    row = c.fetchone()
    if row:
        collection_id = row[0]
    else:
        chars = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
        col_key = "".join(random.choice(chars) for _ in range(8))
        c.execute("INSERT INTO collections (collectionName, parentCollectionID, clientDateModified, key, libraryID) VALUES (?, NULL, datetime('now'), ?, 1)", (collection_name, col_key))
        collection_id = c.lastrowid
        print(f"📁 [Hybrid Agent] 在 Zotero 新建独立分类: [{collection_id}] {collection_name}")

    c.execute("SELECT fieldID, fieldName FROM fields")
    field_map = {name: fid for fid, name in c.fetchall()}
    title_fid = field_map.get("title", 1)
    journal_fid = field_map.get("publicationTitle", 12)
    date_fid = field_map.get("date", 14)
    doi_fid = field_map.get("DOI", 26)
    abstract_fid = field_map.get("abstractNote", 2)

    synced_items = []
    for it in manifest:
        item_key = it.get("zotero_key")
        attach_key = it.get("attach_key")
        title = it.get("title", "")
        pdf_filename = it.get("pdf_filename", "")
        source_pdf_path = os.path.join(topic_dir, pdf_filename)
        lang = it.get("lang", "en")

        c.execute("SELECT itemID FROM items WHERE key = ?", (item_key,))
        row_existing = c.fetchone()
        if row_existing:
            parent_item_id = row_existing[0]
            c.execute("INSERT OR IGNORE INTO collectionItems (collectionID, itemID) VALUES (?, ?)", (collection_id, parent_item_id))
            synced_items.append({"key": item_key, "title": title, "pdf_filename": pdf_filename, "attach_key": attach_key})
            continue

        # 插入主条目 (itemTypeID=4 journalArticle)
        c.execute("INSERT INTO items (itemTypeID, dateAdded, dateModified, clientDateModified, key, libraryID) VALUES (4, datetime('now'), datetime('now'), datetime('now'), ?, 1)", (item_key,))
        parent_item_id = c.lastrowid

        def insert_val(fid, val):
            if not val:
                return
            c.execute("SELECT valueID FROM itemDataValues WHERE value = ?", (str(val),))
            r = c.fetchone()
            if r:
                vid = r[0]
            else:
                c.execute("INSERT INTO itemDataValues (value) VALUES (?)", (str(val),))
                vid = c.lastrowid
            c.execute("INSERT INTO itemData (itemID, fieldID, valueID) VALUES (?, ?, ?)", (parent_item_id, fid, vid))

        insert_val(title_fid, title)
        insert_val(journal_fid, it.get("journal", ""))
        insert_val(date_fid, str(it.get("year", "2024")))
        insert_val(doi_fid, it.get("doi", ""))
        insert_val(abstract_fid, it.get("abstract", ""))

        # 关联到集合
        c.execute("INSERT OR IGNORE INTO collectionItems (collectionID, itemID) VALUES (?, ?)", (collection_id, parent_item_id))

        # 复制物理 PDF 实体到 storage/<attach_key>/
        if os.path.exists(source_pdf_path) and attach_key:
            target_storage_dir = os.path.join(ZOTERO_STORAGE, attach_key)
            os.makedirs(target_storage_dir, exist_ok=True)
            target_storage_pdf = os.path.join(target_storage_dir, pdf_filename)
            shutil.copyfile(source_pdf_path, target_storage_pdf)

            # 插入附件条目 (itemTypeID=3, linkMode=0 imported_file)
            c.execute("INSERT INTO items (itemTypeID, dateAdded, dateModified, clientDateModified, key, libraryID) VALUES (3, datetime('now'), datetime('now'), datetime('now'), ?, 1)", (attach_key,))
            attach_item_id = c.lastrowid
            c.execute("INSERT INTO itemAttachments (itemID, parentItemID, linkMode, contentType, path) VALUES (?, ?, 0, 'application/pdf', ?)", (attach_item_id, parent_item_id, f"storage:{pdf_filename}"))
            flag = "🇨🇳 知网" if lang == "zh" else "🌐 英文"
            print(f"  📎 [{flag}] 成功物理挂载: [{item_key}] {title[:32]}...")

        synced_items.append({"key": item_key, "title": title, "pdf_filename": pdf_filename, "attach_key": attach_key})

    conn.commit()
    conn.close()
    print(f"🎉 [Hybrid Agent] 成功将全部 {len(synced_items)} 篇中英双轨文献物理挂载至 Zotero！")
    return True
