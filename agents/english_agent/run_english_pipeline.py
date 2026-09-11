#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_english_pipeline.py - 英文专属智能体 (English Academic Agent) 一键全自动工作流流水线
1. 自动检索与合法下载外网多页真实英文 PDF 并挂载 Zotero (english_downloader.py)
2. 纯正性审查与目录排伪净化 (english_purger.py)
3. 编译生成符合 SCI 顶刊规范的英文学术综述 DOCX 活体文档 (english_review_builder.py)
"""

import os
import sys
import subprocess
import argparse

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

def run_english_pipeline(query="antimicrobial peptides deep learning metagenomics", count=5, topic_dir=None):
    if not topic_dir:
        topic_dir = r"E:\0mcp-agv\ARTA_Agent_Output\AMP_Metagenomics_English_Review"
    os.makedirs(topic_dir, exist_ok=True)

    print("=" * 70)
    print("🚀 启动英文专属智能体 (English Academic Agent) 全流程编译流水线")
    print(f"📂 目标主题目录: {topic_dir}")
    print(f"🎯 检索关键词: {query} | 篇数: {count}")
    print("=" * 70)

    # Step 1: 检索下载真实 PDF 并挂载 Zotero
    print("\n👉 [Step 1/3] 执行外网多源真实多页英文 PDF 检索、下载与 Zotero 物理挂载...")
    downloader_script = os.path.join(CURRENT_DIR, "english_downloader.py")
    subprocess.run([sys.executable, downloader_script, "--query", query, "--count", str(count), "--out-dir", topic_dir, "--sync-zotero"], check=True)

    # Step 2: 排伪与目录净化
    print("\n👉 [Step 2/3] 执行英文真实文献排伪审查与目录物理净化...")
    purger_script = os.path.join(CURRENT_DIR, "english_purger.py")
    subprocess.run([sys.executable, purger_script, topic_dir], check=True)

    # Step 3: 编译生成 SCI 顶刊学术综述 DOCX
    print("\n👉 [Step 3/3] 编译生成 SCI 顶刊学术综述 DOCX (Zotero OpenXML Live)...")
    builder_script = os.path.join(CURRENT_DIR, "english_review_builder.py")
    subprocess.run([sys.executable, builder_script, "--theme-dir", topic_dir], check=True)

    print("\n" + "=" * 70)
    print("🎉 [English Academic Agent] 流水线执行完毕！外网真文献已入库，SCI 活体综述 DOCX 已交付！")
    print("=" * 70)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="英文专属智能体流水线")
    parser.add_argument("--query", default="antimicrobial peptides deep learning metagenomics", help="检索关键词")
    parser.add_argument("--count", type=int, default=5, help="检索篇数")
    parser.add_argument("--topic-dir", default=r"E:\0mcp-agv\ARTA_Agent_Output\AMP_Metagenomics_English_Review", help="落盘目录")
    args = parser.parse_args()
    run_english_pipeline(args.query, args.count, args.topic_dir)
