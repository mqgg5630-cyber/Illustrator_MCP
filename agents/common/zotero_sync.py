# -*- coding: utf-8 -*-
"""
agents.common.zotero_sync
Zotero 本地 SQLite 活体写入、物理附件挂载与 BibTeX/RIS 文献库导出引擎
严格遵循零 C 盘占用规范，全流程落盘于 E 盘。
"""

import os
import sys
import json
import sqlite3
import random
import shutil
import time
import re

ZOTERO_SQLITE = r"E:\ozotero\zotero.sqlite"
ZOTERO_STORAGE = r"E:\ozotero\storage"

COMPOUND_CHINESE_SURNAMES = {
    "欧阳", "太史", "端木", "上官", "司马", "东方", "独孤", "南宫",
    "万俟", "闻人", "夏侯", "诸葛", "尉迟", "公羊", "赫连", "澹台",
    "皇甫", "宗政", "濮阳", "淳于", "单于", "太叔", "申屠", "公孙",
    "仲孙", "轩辕", "令狐", "钟离", "宇文", "长孙", "慕容", "鲜于",
    "闾丘", "司徒", "司空", "亓官", "司寇", "仉督", "子车", "颛孙"
}


def parse_author_name(author_str):
    """
    解析作者姓名，返回 (lastName, firstName, fieldMode)
    fieldMode=0: 包含姓和名
    fieldMode=1: 单一字段（机构名或无法拆分的单一名称）
    """
    if not author_str:
        return "Unknown", "", 1
    
    author_str = author_str.strip()
    
    # 0. 机构名 / 协作组 / 联合体识别
    consortium_keywords = {
        "consortium", "group", "team", "committee", "collaboration", 
        "initiative", "association", "organization", "society", 
        "alliance", "center", "centre", "laboratory", "project"
    }
    lower_words = set(re.findall(r"\b[a-zA-Z]+\b", author_str.lower()))
    if author_str.lower().startswith("the ") or lower_words.intersection(consortium_keywords):
        return author_str, "", 1

    # 1. 含有逗号: "Last, First"
    if "," in author_str:
        parts = author_str.split(",", 1)
        return parts[0].strip(), parts[1].strip(), 0
    
    # 2. 纯中文姓名
    if re.fullmatch(r"[\u4e00-\u9fa5]+", author_str):
        if len(author_str) >= 3 and author_str[:2] in COMPOUND_CHINESE_SURNAMES:
            return author_str[:2], author_str[2:], 0
        elif len(author_str) >= 2:
            return author_str[0], author_str[1:], 0
        else:
            return author_str, "", 1

    # 3. 英文姓名含空格: "First Middle Last"
    if " " in author_str:
        parts = author_str.split()
        return parts[-1].strip(), " ".join(parts[:-1]).strip(), 0

    # 4. 单一单词 / 机构名
    return author_str, "", 1


def export_bibtex_and_ris(manifest, topic_dir):
    """根据 manifest 导出标准的 References.bib 和 References.ris"""
    bib_path = os.path.join(topic_dir, "References.bib")
    ris_path = os.path.join(topic_dir, "References.ris")

    bib_entries = []
    ris_entries = []

    for it in manifest:
        key = it.get("zotero_key") or it.get("key") or "REF"
        title = it.get("title", "Untitled")
        authors = it.get("authors") or []
        if isinstance(authors, str):
            authors = [a.strip() for a in authors.split(",") if a.strip()]
        journal = it.get("journal", "")
        year = str(it.get("year", "2024"))
        doi = it.get("doi", "")
        abstract = it.get("abstract", "")

        # BibTeX
        bib_authors = " and ".join(authors) if authors else "Unknown"
        bib_lines = [
            f"@article{{{key},",
            f"  title = {{{{{title}}}}},",
            f"  author = {{{bib_authors}}},",
            f"  journal = {{{journal}}},",
            f"  year = {{{year}}},",
        ]
        if doi:
            bib_lines.append(f"  doi = {{{doi}}},")
        if abstract:
            clean_abs = abstract.replace("{", "").replace("}", "").replace("\n", " ")
            bib_lines.append(f"  abstract = {{{clean_abs[:800]}}},")
        bib_lines.append("}\n")
        bib_entries.append("\n".join(bib_lines))

        # RIS
        ris_lines = [
            "TY  - JOUR",
            f"ID  - {key}",
            f"TI  - {title}",
            f"JO  - {journal}",
            f"PY  - {year}",
        ]
        for a in authors:
            ris_lines.append(f"AU  - {a}")
        if doi:
            ris_lines.append(f"DO  - {doi}")
        if abstract:
            ris_lines.append(f"AB  - {abstract[:800]}")
        ris_lines.append("ER  -\n")
        ris_entries.append("\n".join(ris_lines))

    with open(bib_path, "w", encoding="utf-8") as f:
        f.write("\n".join(bib_entries))
    with open(ris_path, "w", encoding="utf-8") as f:
        f.write("\n".join(ris_entries))

    print(f"📄 [BibTeX/RIS] 成功生成 References.bib 与 References.ris")


def sync_to_zotero(topic_dir, collection_name=None, zotero_db_path=ZOTERO_SQLITE, zotero_storage_dir=ZOTERO_STORAGE):
    """
    将 topic_dir 下的文献与 manifest.json 写入 Zotero SQLite：
    1. 插入 items 主条目 (journalArticle)
    2. 插入元数据 fields (title, publicationTitle, date, DOI, abstractNote)
    3. 插入作者表 creators 与关联表 itemCreators (author)
    4. 复制 PDF 实体到 storage/<attach_key>/ 并插入 attachment 条目
    5. 建立 collection 关联
    6. 导出 References.bib 和 References.ris
    """
    manifest_file = os.path.join(topic_dir, "manifest.json")
    if not os.path.exists(manifest_file):
        print(f"⚠️ [Zotero Sync] 未找到 manifest.json: {topic_dir}")
        return False

    with open(manifest_file, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    if not os.path.exists(zotero_db_path):
        print(f"⚠️ [Zotero Sync] Zotero 数据库不存在: {zotero_db_path}")
        return False

    # 导出 BibTeX 与 RIS
    export_bibtex_and_ris(manifest, topic_dir)

    print(f"🔗 [Zotero Sync] 连接 SQLite: {zotero_db_path}")
    conn = sqlite3.connect(zotero_db_path, timeout=30.0)
    try:
        synced_items = _sync_transaction(conn, manifest, topic_dir, collection_name, zotero_storage_dir)
        conn.commit()
        print(f"🎉 [Zotero Sync] 事务提交成功，共同步 {len(synced_items)} 篇文献与附件。")
        return True
    except sqlite3.OperationalError as e:
        conn.rollback()
        if "locked" in str(e).lower():
            print(f"⚠️ [Zotero Sync] 数据库被锁定 (可能 Zotero 客户端正在运行)。请关闭 Zotero 后重试。")
        else:
            print(f"❌ [Zotero Sync] SQLite 错误: {e}")
        raise
    except Exception as e:
        conn.rollback()
        print(f"❌ [Zotero Sync] 写入异常，事务已回滚: {e}")
        raise
    finally:
        conn.close()


def _sync_transaction(conn, manifest, topic_dir, collection_name, zotero_storage_dir):
    c = conn.cursor()

    if not collection_name:
        collection_name = os.path.basename(topic_dir)

    # 获取动态 itemTypeID 与 creatorTypeID
    c.execute("SELECT itemTypeID FROM itemTypes WHERE typeName = 'journalArticle'")
    r = c.fetchone()
    journal_type_id = r[0] if r else 22

    c.execute("SELECT itemTypeID FROM itemTypes WHERE typeName = 'preprint'")
    r = c.fetchone()
    preprint_type_id = r[0] if r else journal_type_id

    c.execute("SELECT itemTypeID FROM itemTypes WHERE typeName = 'attachment'")
    r = c.fetchone()
    attach_type_id = r[0] if r else 3

    c.execute("SELECT creatorTypeID FROM creatorTypes WHERE creatorType = 'author'")
    r = c.fetchone()
    author_type_id = r[0] if r else 10

    # 建立或获取集合
    c.execute("SELECT collectionID FROM collections WHERE collectionName = ?", (collection_name,))
    row = c.fetchone()
    if row:
        collection_id = row[0]
    else:
        chars = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
        col_key = "".join(random.choice(chars) for _ in range(8))
        c.execute("INSERT INTO collections (collectionName, parentCollectionID, clientDateModified, key, libraryID) VALUES (?, NULL, datetime('now'), ?, 1)", (collection_name, col_key))
        collection_id = c.lastrowid
        print(f"📁 [Zotero Sync] 新建文献分类: [{collection_id}] {collection_name}")

    # 字段映射
    c.execute("SELECT fieldID, fieldName FROM fields")
    field_map = {name: fid for fid, name in c.fetchall()}
    title_fid = field_map.get("title", 1)
    journal_fid = field_map.get("publicationTitle", 12)
    date_fid = field_map.get("date", 14)
    doi_fid = field_map.get("DOI", 26)
    abstract_fid = field_map.get("abstractNote", 2)
    volume_fid = field_map.get("volume")
    issue_fid = field_map.get("issue")
    pages_fid = field_map.get("pages")
    lang_fid = field_map.get("language")
    extra_fid = field_map.get("extra")

    synced_items = []
    chars = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"

    for it in manifest:
        item_key = it.get("zotero_key") or it.get("key")
        if not item_key:
            item_key = "".join(random.choice(chars) for _ in range(8))
            it["zotero_key"] = item_key

        attach_key = it.get("attach_key")
        if not attach_key:
            attach_key = "".join(random.choice(chars) for _ in range(8))
            it["attach_key"] = attach_key

        title = it.get("title", "")
        pdf_filename = it.get("pdf_filename", "") or os.path.basename(it.get("local_pdf", ""))
        source_pdf_path = os.path.join(topic_dir, pdf_filename) if pdf_filename else ""

        # 检查是否已存在主条目
        c.execute("SELECT itemID FROM items WHERE key = ?", (item_key,))
        row_existing = c.fetchone()
        if row_existing:
            parent_item_id = row_existing[0]
            c.execute("INSERT OR IGNORE INTO collectionItems (collectionID, itemID) VALUES (?, ?)", (collection_id, parent_item_id))
            synced_items.append({"key": item_key, "title": title, "parent_item_id": parent_item_id})
            continue

        # 插入主条目
        c.execute("""
            INSERT INTO items (itemTypeID, dateAdded, dateModified, clientDateModified, key, libraryID)
            VALUES (?, datetime('now'), datetime('now'), datetime('now'), ?, 1)
        """, (preprint_type_id if it.get("type") == "preprint" else journal_type_id, item_key))
        parent_item_id = c.lastrowid

        # 插入元数据字段
        def insert_val(fid, val):
            if not val or fid is None:
                return
            c.execute("SELECT valueID FROM itemDataValues WHERE value = ?", (str(val),))
            r_val = c.fetchone()
            if r_val:
                vid = r_val[0]
            else:
                c.execute("INSERT INTO itemDataValues (value) VALUES (?)", (str(val),))
                vid = c.lastrowid
            c.execute("INSERT INTO itemData (itemID, fieldID, valueID) VALUES (?, ?, ?)", (parent_item_id, fid, vid))

        insert_val(title_fid, title)
        insert_val(journal_fid, it.get("journal", ""))
        year = str(it.get("year") or "").strip()
        insert_val(date_fid, f"{year}-00-00 {year}" if re.fullmatch(r"\d{4}", year) else year)
        insert_val(doi_fid, it.get("doi", ""))
        insert_val(abstract_fid, it.get("abstract", ""))
        insert_val(volume_fid, it.get("volume", ""))
        insert_val(issue_fid, it.get("issue", ""))
        insert_val(pages_fid, it.get("pages", ""))
        insert_val(lang_fid, "zh-CN" if it.get("lang") == "zh" else "en")
        extra_bits = []
        if it.get("pmid"): extra_bits.append(f"PMID: {it['pmid']}")
        if it.get("pmcid"): extra_bits.append(f"PMCID: {it['pmcid']}")
        vst = (it.get("verification") or {}).get("status")
        if vst: extra_bits.append(f"verification: {vst}")
        insert_val(extra_fid, "\n".join(extra_bits))

        # 关联到集合
        c.execute("INSERT OR IGNORE INTO collectionItems (collectionID, itemID) VALUES (?, ?)", (collection_id, parent_item_id))

        # 写入作者 creators 与 itemCreators
        raw_authors = it.get("authors") or []
        if isinstance(raw_authors, str):
            raw_authors = [a.strip() for a in raw_authors.split(",") if a.strip()]

        for order_idx, author_name in enumerate(raw_authors):
            last_name, first_name, field_mode = parse_author_name(author_name)
            c.execute("SELECT creatorID FROM creators WHERE lastName = ? AND firstName = ? AND fieldMode = ?", (last_name, first_name, field_mode))
            r_creator = c.fetchone()
            if r_creator:
                creator_id = r_creator[0]
            else:
                c.execute("INSERT INTO creators (firstName, lastName, fieldMode) VALUES (?, ?, ?)", (first_name, last_name, field_mode))
                creator_id = c.lastrowid

            c.execute("""
                INSERT OR IGNORE INTO itemCreators (itemID, creatorID, creatorTypeID, orderIndex)
                VALUES (?, ?, ?, ?)
            """, (parent_item_id, creator_id, author_type_id, order_idx))

        # 复制物理 PDF 实体并建立 attachment 关联
        if source_pdf_path and os.path.exists(source_pdf_path) and attach_key:
            target_storage_dir = os.path.join(zotero_storage_dir, attach_key)
            os.makedirs(target_storage_dir, exist_ok=True)
            target_storage_pdf = os.path.join(target_storage_dir, pdf_filename)
            shutil.copyfile(source_pdf_path, target_storage_pdf)

            # 插入附件条目 (linkMode=0 imported_file)
            c.execute("""
                INSERT INTO items (itemTypeID, dateAdded, dateModified, clientDateModified, key, libraryID)
                VALUES (?, datetime('now'), datetime('now'), datetime('now'), ?, 1)
            """, (attach_type_id, attach_key))
            attach_item_id = c.lastrowid

            c.execute("""
                INSERT INTO itemAttachments (itemID, parentItemID, linkMode, contentType, path)
                VALUES (?, ?, 0, 'application/pdf', ?)
            """, (attach_item_id, parent_item_id, f"storage:{pdf_filename}"))

            print(f"  📎 [Zotero 物理挂载] [{item_key}] {title[:32]}... ({len(raw_authors)} 位作者已写入)")

        synced_items.append({"key": item_key, "title": title, "parent_item_id": parent_item_id})

    return synced_items
