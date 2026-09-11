#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
purge_unverified.py - 把「未通过权威核验」的文献从主题目录、Zotero SQLite 与 E:\\ozotero\\storage 中彻底清除。

清除来源（取并集）：
  1. <theme_dir>/manifest.rejected.json           —— verify_manifest(purge=True) 移出的条目
  2. manifest.json 中 verification.status ∈ {unverified, retracted}
  3. --keys K1,K2  手工指定 zotero_key（例如 init_collagen_manifest.py 里那 5 条虚构记录）
  4. --recheck     先联网重新跑 verify（不改写 manifest），再按结果清理

默认 dry-run 只打印；加 --apply 才真正删除。Zotero 端删除的是 items / itemData / itemCreators /
collectionItems / itemAttachments / 子附件条目 及 storage/<attach_key>/ 目录；主题目录下删除对应 PDF。

用法:
  python scripts/purge_unverified.py E:\\0mcp-agv\\ARTA_Agent_Output\\Hub_xxx
  python scripts/purge_unverified.py <theme_dir> --recheck --apply
  python scripts/purge_unverified.py <theme_dir> --keys 7WUA8PLB,254AZYKD --apply
"""
import os
import sys
import json
import shutil
import sqlite3
import argparse

WORKSPACE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if WORKSPACE_ROOT not in sys.path:
    sys.path.insert(0, WORKSPACE_ROOT)

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from agents.common.zotero_sync import ZOTERO_SQLITE, ZOTERO_STORAGE  # noqa: E402

BAD = {"unverified", "retracted"}


def _load(path):
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except Exception:
            return []


def collect_targets(theme_dir, keys=None, recheck=False, include_suspicious=False):
    manifest = _load(os.path.join(theme_dir, "manifest.json"))
    rejected = _load(os.path.join(theme_dir, "manifest.rejected.json"))
    targets = {}

    for it in rejected:
        if it.get("zotero_key"):
            targets[it["zotero_key"]] = it

    if recheck:
        from agents.common.verify import verify_record
        for it in manifest:
            v = verify_record(it)
            it["verification"] = {"status": v["status"], "reasons": v["reasons"]}
            print(f"  {v['status']:<11} {it.get('zotero_key')}  {(it.get('title') or '')[:60]}  {'; '.join(v['reasons'][:2])}")

    bad = BAD | ({"suspicious"} if include_suspicious else set())
    for it in manifest:
        st = (it.get("verification") or {}).get("status")
        if st in bad and it.get("zotero_key"):
            targets[it["zotero_key"]] = it

    for k in (keys or []):
        k = k.strip()
        if not k:
            continue
        hit = next((it for it in manifest + rejected if it.get("zotero_key") == k), None)
        targets[k] = hit or {"zotero_key": k, "title": "(manual key)"}

    return manifest, targets


def purge_zotero(conn, item_key, storage_dir, apply):
    c = conn.cursor()
    c.execute("SELECT itemID FROM items WHERE key = ?", (item_key,))
    row = c.fetchone()
    if not row:
        print(f"    · Zotero 中无 key={item_key}，跳过数据库删除")
        return 0
    parent_id = row[0]
    c.execute("SELECT itemID FROM itemAttachments WHERE parentItemID = ?", (parent_id,))
    child_ids = [r[0] for r in c.fetchall()]
    all_ids = [parent_id] + child_ids
    # 附件 storage 目录 = 附件条目 key
    child_keys = []
    for cid in child_ids:
        c.execute("SELECT key FROM items WHERE itemID = ?", (cid,))
        r = c.fetchone()
        if r:
            child_keys.append(r[0])
    print(f"    · Zotero itemID={parent_id} 附件 {child_ids} storage {child_keys}")
    if not apply:
        return len(all_ids)
    q = ",".join("?" * len(all_ids))
    for tbl in ("itemData", "itemCreators", "collectionItems", "itemNotes", "itemTags", "deletedItems"):
        try:
            c.execute(f"DELETE FROM {tbl} WHERE itemID IN ({q})", all_ids)
        except sqlite3.OperationalError:
            pass
    c.execute(f"DELETE FROM itemAttachments WHERE itemID IN ({q})", all_ids)
    c.execute(f"DELETE FROM items WHERE itemID IN ({q})", all_ids)
    for ck in child_keys:
        d = os.path.join(storage_dir, ck)
        if os.path.isdir(d):
            shutil.rmtree(d, ignore_errors=True)
            print(f"    · 已删除 storage 目录 {d}")
    return len(all_ids)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("theme_dir")
    ap.add_argument("--keys", default="", help="逗号分隔的 zotero_key，强制清除")
    ap.add_argument("--recheck", action="store_true", help="先联网重新核验")
    ap.add_argument("--include-suspicious", action="store_true", help="连 suspicious 一起清除（默认只清 unverified/retracted）")
    ap.add_argument("--apply", action="store_true", help="真正执行删除（默认 dry-run）")
    ap.add_argument("--zotero-db", default=ZOTERO_SQLITE)
    ap.add_argument("--zotero-storage", default=ZOTERO_STORAGE)
    args = ap.parse_args()

    theme_dir = os.path.abspath(args.theme_dir)
    manifest, targets = collect_targets(theme_dir, [k for k in args.keys.split(",")], args.recheck, args.include_suspicious)
    if not targets:
        print("✅ 没有需要清除的条目。")
        return
    print(f"\n{'🗑️  将清除' if args.apply else '🔍 [dry-run] 待清除'} {len(targets)} 条：")
    for k, it in targets.items():
        print(f"  - {k}  {(it.get('title') or '')[:70]}  [{(it.get('verification') or {}).get('status', 'manual')}]")

    conn = None
    if os.path.exists(args.zotero_db):
        conn = sqlite3.connect(args.zotero_db, timeout=15)
    else:
        print(f"⚠️ 找不到 Zotero 数据库 {args.zotero_db}，仅处理主题目录")

    for k, it in targets.items():
        print(f"\n▶ {k}")
        pdf = it.get("pdf_filename") or os.path.basename(it.get("local_pdf") or "")
        if pdf:
            p = os.path.join(theme_dir, pdf)
            if os.path.exists(p):
                print(f"    · 主题目录 PDF {p}")
                if args.apply:
                    os.remove(p)
        if conn:
            try:
                purge_zotero(conn, k, args.zotero_storage, args.apply)
            except sqlite3.OperationalError as e:
                print(f"    ⚠️ Zotero 数据库被占用或结构异常: {e}（请先关闭 Zotero 客户端）")

    if args.apply:
        if conn:
            conn.commit()
        keep = [it for it in manifest if it.get("zotero_key") not in targets]
        if len(keep) != len(manifest):
            with open(os.path.join(theme_dir, "manifest.json"), "w", encoding="utf-8") as f:
                json.dump(keep, f, ensure_ascii=False, indent=2)
        rej_path = os.path.join(theme_dir, "manifest.rejected.json")
        if os.path.exists(rej_path):
            os.rename(rej_path, rej_path + ".purged")
        print(f"\n✅ 清除完成：manifest 剩余 {len(keep)} 条。请重新运行 write_digest / --prepare 刷新 digest。")
    else:
        print("\nℹ️ dry-run 结束，确认无误后加 --apply 执行。")
    if conn:
        conn.close()


if __name__ == "__main__":
    main()
