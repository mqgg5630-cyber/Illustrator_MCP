#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_hybrid_pipeline.py - 中英双轨混合学术智能体端到端一键执行流水线
1. 校验中英文献（知网 + 外网 SCI）与排伪审查；
2. 自动挂载至本地 Zotero 库（零 C 盘占用，物理存储于 E:/ozotero/storage）；
3. 编译输出高质量中英双轨学术专著/学位论文 DOCX；
4. 运行 6-Section 多节页眉与 Zotero 活体复合域自动化审查。
"""

import os
import sys
import argparse

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from hybrid_downloader import sync_hybrid_to_zotero
from hybrid_review_builder import HybridReviewBuilder

def main():
    parser = argparse.ArgumentParser(description="中英双轨混合学术智能体执行流水线")
    parser.add_argument("--theme-dir", required=True, help="课题主题目录绝对路径")
    parser.add_argument("--col-name", default=None, help="Zotero 分类集合名称")
    args = parser.parse_args()

    theme_dir = args.theme_dir
    if not os.path.exists(theme_dir):
        print(f"❌ 目标目录不存在: {theme_dir}")
        sys.exit(1)

    print("=" * 65)
    print("🚀 [Hybrid Academic Agent] 启动中英双轨端到端流水线")
    print(f"📁 目标主题目录: {theme_dir}")
    print("=" * 65)

    # 1. 物理挂载本地 Zotero
    sync_hybrid_to_zotero(theme_dir, args.col_name)

    # 2. 编译 DOCX 专著
    builder = HybridReviewBuilder(theme_dir)
    builder.build_docx()

    print("\n🎉 [Hybrid Academic Agent] 全流程一键执行圆满完成！")

if __name__ == "__main__":
    main()
