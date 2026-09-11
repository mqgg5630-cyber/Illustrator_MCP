# -*- coding: utf-8 -*-
"""
agents.common.digest
把 enrich + verify 之后的 manifest 转成 Antigravity 可直接阅读的写作素材：

  <theme>/digest/INDEX.md            全部文献一览（每篇一行）+ 主题聚类 + 写作约束
  <theme>/digest/<zotero_key>.md     单篇卡片（≤ ~2 KB）：规范元数据 / 摘要 / 结果与结论摘录 / 可引用数值句
  <theme>/digest/literature_matrix.md 文献 × 维度 矩阵（可直接转三线表）

只使用 API 文本，不读 PDF。
"""

import json
import os
import re
from collections import Counter, defaultdict

MAX_CARD_CHARS = 2400


def _fmt_authors(rec, n=3):
    a = rec.get("authors") or []
    if isinstance(a, str):
        a = [x.strip() for x in a.split(",")]
    if not a:
        return "—"
    zh = rec.get("lang") == "zh"
    head = ("，" if zh else ", ").join(a[:n])
    if len(a) > n:
        head += "，等" if zh else ", et al."
    return head


def _cite_label(rec):
    a = rec.get("authors") or []
    first = a[0] if a else "Anon"
    if isinstance(first, str) and not re.fullmatch(r"[\u4e00-\u9fff]+", first):
        first = first.split(",")[0].split()[-1] if first else "Anon"
    y = rec.get("year", "n.d.")
    return f"{first} {y}"


def _status_mark(rec):
    s = (rec.get("verification") or {}).get("status", "unverified")
    return {"verified": "✅", "suspicious": "⚠️", "unverified": "❌", "retracted": "🛑"}.get(s, "❓")


def _clip(s, n):
    s = re.sub(r"\s+", " ", s or "").strip()
    return s if len(s) <= n else s[: n - 1].rstrip() + "…"


def _build_card_with(rec, abs_n, exc_n, facts_n):
    v = rec.get("verification") or {}
    lines = [
        "---",
        f"zotero_key: {rec.get('zotero_key', '')}",
        f"doi: {rec.get('doi', '')}",
        f"type: {rec.get('type', 'unknown')}",
        f"lang: {rec.get('lang', 'en')}",
        f"verification: {v.get('status', 'unverified')}",
        f"cited_by: {rec.get('cited_by_count', 0)}",
        "---",
        f"# {rec.get('title', '')}",
        "",
        f"**{_fmt_authors(rec, 5)}**. *{rec.get('journal', '')}*, {rec.get('year', '')}"
        + (f", {rec.get('volume')}" if rec.get("volume") else "")
        + (f"({rec.get('issue')})" if rec.get("issue") else "")
        + (f": {rec.get('pages')}" if rec.get("pages") else "")
        + ".",
        "",
    ]
    if rec.get("tldr"):
        lines += ["## TL;DR", _clip(rec["tldr"], 300), ""]
    kw = (rec.get("mesh") or [])[:8] or (rec.get("concepts") or [])[:8]
    if kw:
        lines += ["## Topics", "; ".join(kw), ""]
    if rec.get("abstract"):
        lines += [f"## Abstract ({rec.get('abstract_source', '')})", _clip(rec["abstract"], abs_n), ""]
    if rec.get("results_excerpt"):
        lines += ["## Results (JATS excerpt)", _clip(rec["results_excerpt"], exc_n), ""]
    if rec.get("conclusion_excerpt"):
        lines += ["## Conclusion (JATS excerpt)", _clip(rec["conclusion_excerpt"], exc_n), ""]
    ks = rec.get("key_sentences") or []
    if ks:
        lines += ["## Quotable facts（含数值，引注数值时只能用这里的句子）"]
        lines += [f"- {_clip(s, 260)}" for s in ks[:facts_n]]
        lines += [""]
    if v.get("status") == "suspicious":
        lines += ["> ⚠️ 核验存疑: " + "; ".join(v.get("reasons") or []), ""]
    return "\n".join(lines)


def build_card(rec):
    """渐进式裁剪：先缩摘要，再缩 Results/Conclusion，最后减少 quotable facts，直到 ≤ MAX_CARD_CHARS。"""
    budgets = [(1200, 700, 10), (700, 500, 8), (500, 350, 6), (350, 250, 5), (250, 200, 4)]
    text = ""
    for abs_n, exc_n, facts_n in budgets:
        text = _build_card_with(rec, abs_n, exc_n, facts_n)
        if len(text) <= MAX_CARD_CHARS:
            return text
    return text


def cluster_topics(manifest, top_n=6):
    """用 mesh/concepts 做粗聚类，给 Antigravity 一个章节结构起点。"""
    freq = Counter()
    per_doc = {}
    for rec in manifest:
        terms = set((rec.get("mesh") or []) + (rec.get("concepts") or []) + (rec.get("keywords") or []))
        terms = {t for t in terms if t and t.lower() not in {"humans", "animals", "biology", "chemistry", "medicine", "male", "female"}}
        per_doc[rec.get("zotero_key")] = terms
        freq.update(terms)
    clusters = []
    for term, n in freq.most_common(top_n * 3):
        if n < 2:
            continue
        members = [k for k, ts in per_doc.items() if term in ts]
        clusters.append((term, members))
        if len(clusters) >= top_n:
            break
    return clusters


def build_index(manifest, theme_name):
    verified = [r for r in manifest if (r.get("verification") or {}).get("status") == "verified"]
    susp = [r for r in manifest if (r.get("verification") or {}).get("status") == "suspicious"]
    reviews = [r for r in manifest if r.get("type") == "review"]
    lines = [
        f"# Literature Digest — {theme_name}",
        "",
        f"共 {len(manifest)} 篇 · ✅ verified {len(verified)} · ⚠️ suspicious {len(susp)} · review {len(reviews)} · 年份 "
        + (f"{min(int(str(r.get('year'))[:4]) for r in manifest if r.get('year'))}–{max(int(str(r.get('year'))[:4]) for r in manifest if r.get('year'))}" if any(r.get("year") for r in manifest) else "—"),
        "",
        "## 写作约束（review-writing skill 强制）",
        "1. 只能引用本文件列出的 `zotero_key`；`❌ unverified` / `🛑 retracted` 条目不得引用。",
        "2. 正文中出现的任何数值，必须能在对应 `<key>.md` 的 **Quotable facts** 或 Abstract/Results/Conclusion 中逐字找到。",
        "3. `type=review` 的文献只能用于研究现状/背景段，不得作为原研结论的唯一支撑。",
        "4. 每段综合 ≥ 2 篇文献；禁止按文献顺序逐篇复述。",
        "5. 无法从素材推出的方法细节（参数、样本量等）一律不写。",
        "",
        "## 文献一览",
        "",
        "| # | key | 状态 | 年份 | 类型 | 被引 | 标题 | 一句话结论 |",
        "|---|-----|------|------|------|------|------|------------|",
    ]
    for i, r in enumerate(sorted(manifest, key=lambda x: (str(x.get("year", "")), x.get("title", ""))), 1):
        lines.append(
            f"| {i} | `{r.get('zotero_key', '')}` | {_status_mark(r)} | {r.get('year', '')} | {r.get('type', '')} | {r.get('cited_by_count', 0)} "
            f"| {_clip(r.get('title', ''), 80)} | {_clip(r.get('tldr', ''), 140)} |"
        )
    clusters = cluster_topics(manifest)
    if clusters:
        lines += ["", "## 主题聚类（章节结构起点）", ""]
        for term, members in clusters:
            lines.append(f"- **{term}** ({len(members)}): " + ", ".join(f"`{m}`" for m in members))
    lines += ["", "## 单篇卡片", ""]
    lines += [f"- [`{r.get('zotero_key', '')}`](./{r.get('zotero_key', '')}.md) — {_cite_label(r)}" for r in manifest]
    return "\n".join(lines) + "\n"


def build_matrix(manifest):
    lines = [
        "# Literature Matrix",
        "",
        "> 直接可转为论文三线表的维度；仅含素材中可核实的字段，不虚构指标列。",
        "",
        "| key | 第一作者/年 | 类型 | 研究对象/主题 (mesh/concepts) | 主要结论 (TL;DR) | 关键数值 (首条 quotable fact) |",
        "|-----|------------|------|------------------------------|------------------|------------------------------|",
    ]
    for r in manifest:
        topics = "; ".join(((r.get("mesh") or []) + (r.get("concepts") or []))[:4])
        fact = (r.get("key_sentences") or [""])[0]
        lines.append(f"| `{r.get('zotero_key', '')}` | {_cite_label(r)} | {r.get('type', '')} | {_clip(topics, 90)} | {_clip(r.get('tldr', ''), 120)} | {_clip(fact, 140)} |")
    return "\n".join(lines) + "\n"


def write_digest(theme_dir, include_unverified=False):
    mf = os.path.join(theme_dir, "manifest.json")
    with open(mf, "r", encoding="utf-8") as f:
        manifest = json.load(f)
    if not include_unverified:
        manifest = [r for r in manifest if (r.get("verification") or {}).get("status") in ("verified", "suspicious")]
    out = os.path.join(theme_dir, "digest")
    os.makedirs(out, exist_ok=True)
    for old in os.listdir(out):
        if old.endswith(".md"):
            os.remove(os.path.join(out, old))
    for r in manifest:
        with open(os.path.join(out, f"{r.get('zotero_key', 'NOKEY')}.md"), "w", encoding="utf-8") as f:
            f.write(build_card(r))
    with open(os.path.join(out, "INDEX.md"), "w", encoding="utf-8") as f:
        f.write(build_index(manifest, os.path.basename(theme_dir)))
    with open(os.path.join(out, "literature_matrix.md"), "w", encoding="utf-8") as f:
        f.write(build_matrix(manifest))
    total = sum(os.path.getsize(os.path.join(out, x)) for x in os.listdir(out))
    print(f"📚 digest 写入 {out}: {len(manifest)} 张卡片 + INDEX.md + literature_matrix.md，共 {total // 1024} KB")
    return out


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("usage: python -m agents.common.digest <theme_dir> [--include-unverified]")
        sys.exit(1)
    write_digest(sys.argv[1], include_unverified="--include-unverified" in sys.argv)
