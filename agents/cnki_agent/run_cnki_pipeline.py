#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_cnki_pipeline.py - 知网专属智能体 (CNKI Academic Agent) 一键全自动工作流流水线
1. 校验与净化指定知网主题目录 (cnki_purger.py)
2. 确保 Zotero 物理挂载与零 C 盘关联 (cnki_downloader.py)
3. 编译生成高规格鲁东大学标准学术学位论文 DOCX 活体文档 (cnki_thesis_builder.py)
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

def run_cnki_pipeline(topic_dir):
    print("=" * 70)
    print("🚀 启动知网专属智能体 (CNKI Academic Agent) 全流程编译流水线")
    print(f"📂 目标主题目录: {topic_dir}")
    print("=" * 70)

    # Step 1: 排伪与物理净化
    print("\n👉 [Step 1/3] 执行纯知网真文献排伪审查与目录净化...")
    purger_script = os.path.join(CURRENT_DIR, "cnki_purger.py")
    subprocess.run([sys.executable, purger_script, topic_dir], check=True)

    # Step 2: 确认 Zotero 挂载
    print("\n👉 [Step 2/3] 核查 Zotero 零C盘物理附件挂载...")
    downloader_script = os.path.join(CURRENT_DIR, "cnki_downloader.py")
    subprocess.run([sys.executable, downloader_script, topic_dir], check=True)

    # Step 3: 编译标准学位论文 DOCX
    print("\n👉 [Step 3/3] 编译生成标准学位论文 / 长篇综述 DOCX (Zotero OpenXML Live)...")
    builder_script = os.path.join(CURRENT_DIR, "cnki_thesis_builder.py")
    subprocess.run([sys.executable, builder_script], check=True)

    print("\n" + "=" * 70)
    print("🎉 [CNKI Academic Agent] 流水线执行完毕！全部真文献已闭环，DOCX 活体论文已交付！")
    print("=" * 70)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="知网专属智能体流水线")
    parser.add_argument("--topic-dir", default=r"E:\0mcp-agv\ARTA_Agent_Output\阿尔兹海默症肠道宏基因组抗菌肽差异", help="主题目录")
    args = parser.parse_args()
    run_cnki_pipeline(args.topic_dir)
