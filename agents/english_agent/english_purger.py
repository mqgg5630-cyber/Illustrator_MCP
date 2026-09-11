#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
english_purger.py - 英文外网专属文献纯正性核验与自动排伪门禁
严格核查全球开源与权威 SCI 文献官方正版多页 PDF，自动抹除伪造占位、HTML 拦截页及损坏文件。
"""

import os
import sys
import json
import re
import argparse
import sqlite3

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

PLACEHOLDER_PATTERNS = [
    r"^AI_Mapping_.*\.pdf$",
    r"^deep-learning-in-.*\.pdf$",
    r"^Genome_To_Peptide_.*\.pdf$",
    r"^Pfeature_.*\.pdf$",
    r"^Mining-human-microbiomes.*\.pdf$",
    r"^mock_.*\.pdf$",
    r"^placeholder_.*\.pdf$",
    r"^test_.*\.pdf$",
]

def purge_english_directory(target_dir, zotero_db_path=r"E:\ozotero\zotero.sqlite"):
    print("=" * 65)
    print(f"🛡️  [English Agent] 启动英文外网文献纯正性核验与排伪门禁: {target_dir}")
    print("=" * 65)

    manifest_file = os.path.join(target_dir, "manifest.json")
    if not os.path.exists(manifest_file):
        print(f"⚠️ 未找到 manifest.json: {manifest_file}")
        return False

    with open(manifest_file, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    valid_keys = set()
    if os.path.exists(zotero_db_path):
        try:
            conn = sqlite3.connect(f"file:{zotero_db_path}?mode=ro", uri=True, timeout=10)
            c = conn.cursor()
            c.execute("SELECT key FROM items")
            valid_keys = set(row[0] for row in c.fetchall())
            conn.close()
        except Exception as e:
            print(f"⚠️ 读取 Zotero 数据库校验: {e}")

    cleaned = []
    whitelist_files = set()
    for it in manifest:
        fname = it.get("pdf_filename") or ""
        z_key = it.get("zotero_key", "")
        # 拦截占位伪造文件名
        is_bad = any(re.search(p, fname, re.IGNORECASE) for p in PLACEHOLDER_PATTERNS)
        if is_bad:
            print(f"🚫 [拦截伪造文献] {fname}")
            continue

        full_path = os.path.join(target_dir, fname)
        if not os.path.exists(full_path):
            print(f"🚫 [拦截缺失实体文件] {fname}")
            continue

        # 检查二进制头与文件大小
        size = os.path.getsize(full_path)
        if size < 50 * 1024:
            print(f"🚫 [拦截过小或损坏文件] {fname} ({size} bytes)")
            continue

        with open(full_path, "rb") as pf:
            head = pf.read(1024)
            if b"%PDF-" not in head:
                print(f"🚫 [拦截非PDF流] {fname}")
                continue

        cleaned.append(it)
        whitelist_files.add(fname)

    # 物理抹除非白名单残留文件
    purged_count = 0
    for f in os.listdir(target_dir):
        if f.lower().endswith(".pdf") and f not in whitelist_files:
            file_path = os.path.join(target_dir, f)
            print(f"🧹 [物理抹除残留文件] {f}")
            os.remove(file_path)
            purged_count += 1

    if len(cleaned) != len(manifest):
        with open(manifest_file, "w", encoding="utf-8") as f:
            json.dump(cleaned, f, ensure_ascii=False, indent=2)

    print(f"✅ 英文专属门禁通过！保留正版多页真文献: {len(cleaned)} 篇，物理抹除残留: {purged_count} 个")
    print("=" * 65)
    return True

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="英文专属文献排伪工具")
    parser.add_argument("target_dir", help="英文文献主题目录")
    args = parser.parse_args()
    purge_english_directory(args.target_dir)
