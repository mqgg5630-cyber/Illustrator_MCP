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

WORKSPACE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if WORKSPACE_ROOT not in sys.path:
    sys.path.insert(0, WORKSPACE_ROOT)

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from hybrid_downloader import sync_hybrid_to_zotero
from hybrid_review_builder import HybridReviewBuilder, write_outline_skeleton

def main():
    parser = argparse.ArgumentParser(description="中英双轨混合学术智能体执行流水线")
    parser.add_argument("--theme-dir", required=True, help="课题主题目录绝对路径")
    parser.add_argument("--col-name", default=None, help="Zotero 分类集合名称")
    parser.add_argument("--skip-zotero", action="store_true", help="跳过 Zotero 挂载，仅编译 DOCX")
    parser.add_argument("--prepare", action="store_true", help="只执行 enrich → verify → digest → skeleton，供 Antigravity 撰写 outline.json；不编译")
    parser.add_argument("--no-enrich", action="store_true", help="跳过 enrich/verify/digest（manifest 已处理过时使用）")
    args = parser.parse_args()

    theme_dir = args.theme_dir
    if not os.path.exists(theme_dir):
        print(f"❌ 目标目录不存在: {theme_dir}")
        sys.exit(1)

    print("=" * 65)
    print("🚀 [Hybrid Academic Agent] 启动中英双轨端到端流水线")
    print(f"📁 目标主题目录: {theme_dir}")
    print("=" * 65)

    if not args.no_enrich:
        from agents.common.enrich import enrich_manifest
        from agents.common.verify import verify_manifest
        from agents.common.digest import write_digest
        print("🧩 [1/4] Enrich: 多源元数据聚合 (Crossref / OpenAlex / Europe PMC / JATS / S2)")
        enrich_manifest(theme_dir)
        print("🔎 [2/4] Verify: 外部权威反向核验，移出 unverified / retracted")
        _, vsum = verify_manifest(theme_dir, purge=True)
        print("📚 [3/4] Digest: 生成 digest/INDEX.md 与单篇卡片")
        write_digest(theme_dir)
        if vsum["verified"] + vsum["suspicious"] == 0:
            print("❌ 没有任何文献通过权威核验，终止。")
            sys.exit(3)

    outline = os.path.join(theme_dir, "outline.json")
    if args.prepare:
        skel = write_outline_skeleton(theme_dir, force=True)
        print(f"📝 [4/4] 已生成 {skel}\n   → 请 Antigravity 按 review-writing skill 阅读 digest/INDEX.md 撰写 outline.json，再运行本脚本（可加 --no-enrich）。")
        sys.exit(0)
    if not os.path.exists(outline):
        skel = write_outline_skeleton(theme_dir)
        print(f"❌ 缺少 {outline}")
        print(f"   已生成骨架 {skel}，请让 Antigravity 通读 manifest 文献后填写正文并另存为 outline.json 再运行。")
        sys.exit(2)

    # 1. 物理挂载本地 Zotero
    if not args.skip_zotero:
        sync_hybrid_to_zotero(theme_dir, args.col_name)

    # 2. 编译 DOCX 专著
    builder = HybridReviewBuilder(theme_dir)
    builder.build_docx()

    print("\n🎉 [Hybrid Academic Agent] 全流程一键执行圆满完成！")

if __name__ == "__main__":
    main()
