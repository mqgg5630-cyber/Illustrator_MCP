# -*- coding: utf-8 -*-
"""
agents.common.audit_gate
文献纯正性审查与物理排伪门禁
严格核验 PDF 二进制真实性（%PDF- 头、多页 >= 2 页、大小 >= 50KB），
自动抹除黑名单占位伪文献并净化 manifest.json、References.bib 与 References.ris。
"""

import os
import sys
import json
import re
import sqlite3

from .zotero_sync import export_bibtex_and_ris

PLACEHOLDER_PATTERNS = [
    r"^AI_Mapping_.*\.pdf$",
    r"^deep-learning-in-.*\.pdf$",
    r"^Genome_To_Peptide_.*\.pdf$",
    r"^Pfeature_.*\.pdf$",
    r"^Mining-human-microbiomes.*\.pdf$",
    r"^mock_.*\.pdf$",
    r"^placeholder_.*\.pdf$",
    r"^synthetic_.*\.pdf$",
    r"^test_.*\.pdf$",
]


def is_valid_pdf_file(filepath):
    """验证是否为真实合规的 PDF 文件 (非空、含 %PDF- 头、大小 >= 50KB、真实多页)"""
    if not os.path.exists(filepath):
        return False, "文件不存在"
    size = os.path.getsize(filepath)
    if size < 50 * 1024:
        return False, f"文件过小 ({size // 1024} KB < 50 KB 阈值，疑似拦截页或空壳)"

    try:
        with open(filepath, "rb") as f:
            header = f.read(1024)
            if b"%PDF-" not in header:
                return False, "非有效 PDF 二进制流 (缺少 %PDF- 头)"

            f.seek(0)
            content = f.read()
            page_matches = len(re.findall(rb"/Type\s*/Page\b", content))
            if page_matches == 1 and size < 80 * 1024:
                return False, f"疑似单页拦截/占位伪造文件 ({size // 1024} KB, 1页)"

        return True, f"有效 PDF (约 {max(1, page_matches)} 页, {size // 1024} KB)"
    except Exception as e:
        return False, f"读取校验失败: {e}"


def is_blacklisted_filename(filename):
    """检查是否匹配占位符/伪造命名黑名单"""
    for pat in PLACEHOLDER_PATTERNS:
        if re.search(pat, filename, re.IGNORECASE):
            return True
    return False


class LiteratureAuditor:
    """文献审查与排伪门禁类"""

    def __init__(self, target_dir, zotero_db_path=r"E:\ozotero\zotero.sqlite"):
        self.target_dir = target_dir
        self.zotero_db_path = zotero_db_path

    def audit_and_purge(self):
        """
        执行目录净化与排伪，返回 (clean_count, evicted_count)
        """
        print("=" * 65)
        print(f"🛡️  [LiteratureAuditor] 启动文献纯正性核验与自动排伪门禁: {self.target_dir}")
        print("=" * 65)

        if not os.path.exists(self.target_dir):
            print(f"❌ 目录不存在: {self.target_dir}")
            return 0, 0

        manifest_file = os.path.join(self.target_dir, "manifest.json")
        manifest_data = []
        if os.path.exists(manifest_file):
            try:
                with open(manifest_file, "r", encoding="utf-8") as f:
                    manifest_data = json.load(f)
            except Exception as e:
                print(f"⚠️ 读取 manifest.json 失败: {e}")

        # 扫描目录下的所有 PDF
        all_pdfs = [f for f in os.listdir(self.target_dir) if f.lower().endswith(".pdf")]
        evicted_files = set()

        for pdf_name in all_pdfs:
            pdf_path = os.path.join(self.target_dir, pdf_name)

            # 1. 检查黑名单
            if is_blacklisted_filename(pdf_name):
                print(f"🚫 [黑名单命中] 标记伪造/占位文献: {pdf_name}")
                evicted_files.add(pdf_name)
                continue

            # 2. 检查 PDF 真实性
            valid, reason = is_valid_pdf_file(pdf_path)
            if not valid:
                print(f"❌ [文件校验失败] {pdf_name}: {reason}")
                evicted_files.add(pdf_name)
                continue

        # 物理抹除无效文件
        for f_evicted in evicted_files:
            p = os.path.join(self.target_dir, f_evicted)
            try:
                os.remove(p)
                print(f"🗑️ [物理清理] 成功抹除无效文件: {f_evicted}")
            except Exception as e:
                print(f"⚠️ 抹除文件失败: {p}, {e}")

        # 净化 manifest.json
        clean_manifest = []
        for it in manifest_data:
            pdf_fn = it.get("pdf_filename") or os.path.basename(it.get("local_pdf", ""))
            if pdf_fn in evicted_files:
                continue
            # 确保物理文件依然存在
            if pdf_fn and not os.path.exists(os.path.join(self.target_dir, pdf_fn)):
                continue
            clean_manifest.append(it)

        with open(manifest_file, "w", encoding="utf-8") as f:
            json.dump(clean_manifest, f, ensure_ascii=False, indent=2)

        # 重新导出 bib / ris
        if clean_manifest:
            export_bibtex_and_ris(clean_manifest, self.target_dir)

        clean_count = len(clean_manifest)
        evicted_count = len(evicted_files)
        print(f"📊 [排伪门禁] 审核总结: {clean_count} 篇纯正有效, {evicted_count} 篇已排除净化。")
        return clean_count, evicted_count


def sanitize_theme_directory(target_dir, zotero_db_path=r"E:\ozotero\zotero.sqlite"):
    """向后兼容函数接口"""
    auditor = LiteratureAuditor(target_dir, zotero_db_path)
    clean_cnt, _ = auditor.audit_and_purge()
    return clean_cnt > 0
