---
name: review-writing
description: 基于 digest/（元数据+摘要+结论段，不读 PDF）为已核验文献撰写综述 outline.json，供 hybrid_review_builder 编译为六段式 DOCX。触发词：写综述、写 outline、填 outline.json、review-writing。
---

# review-writing — 元数据线路综述写作规范

## 何时使用
主题目录里已经有 `digest/INDEX.md`（由 `run_hybrid_pipeline.py --prepare` 或 GUI 检索任务生成），
且存在 `outline.skeleton.json`。你的产出是同目录下的 `outline.json`。

## 唯一事实来源（硬约束）
1. **只读 `digest/`**，禁止打开 PDF、禁止凭记忆补充任何文献内容。
2. 只能引用 `digest/INDEX.md` 表格里 `verification ∈ {verified, suspicious}` 的 `zotero_key`。
3. 正文中出现的**每一个数值**（温度、时间、能量、Tm、RMSD、样本数、百分比……）必须能在对应
   `digest/<key>.md` 的 *Quotable facts / Abstract / Results / Conclusion* 里逐字找到，包括单位与约等号。
   找不到 → 改写为定性表述（"显著缩短"、"约一个数量级"），或不写。
4. 方法细节（力场、软件、步长、对接程序）只有出现在该篇卡片里才能提及。
5. `type = review` 的文献只能出现在「研究背景 / 研究现状」类段落，不能作为具体结论的证据。
6. 不得推断作者未说的因果；对比两篇文献时用 "A 报道…，而 B 观察到…" 的并列句式。

## 工作流
1. 读 `digest/INDEX.md`：看约束、文献总表、MeSH/概念聚类 → 决定 3–5 个主题簇。
2. 读 `digest/literature_matrix.md`：确定各篇的年份/类型/被引/TL;DR，安排叙述顺序（同簇内按年）。
3. 逐簇打开 `digest/<key>.md`，把要用的 Quotable facts 原句抄到草稿里，标注 key。
4. 按 `outline.skeleton.json` 的结构填写：
   - `meta`：题名、作者、单位、关键词（3–8 个，取自 Topics 高频词）。
   - `abstract`：≤300 字，不出现引注。
   - `sections[]`：每节 `heading` + `paragraphs[]`；每段 `{ "text": "...", "cites": ["KEY1","KEY2"] }`。
     - 每段 **≥2 篇**不同文献支撑（背景段允许 1 篇）；段内必须有综合句，不许 "A 做了…。B 做了…。" 罗列。
     - 每节末尾一段做小结 + 指出研究空白（空白本身不加引注）。
   - `tables[]`（可选）：列只能来自 digest 字段（年份、期刊、类型、方法关键词、核心数值原句）。
   - `conclusion`：综合全文，不引入新数值。
5. 删除 `_instructions` 与 `_available_references`，另存为 `outline.json`。

## 提交前自检（逐条打钩写进 outline.json 的 `_selfcheck`）
- [ ] 所有 `cites` 的 key 都在 INDEX.md 的 verified/suspicious 列表中。
- [ ] 用正则把正文里的数字全部抓出来，每个都能 `grep` 到对应 `digest/<key>.md`。
- [ ] 没有任何段落只引 review 型文献却给出具体结论。
- [ ] 每个 verified 文献至少被引用 1 次（否则说明该篇不该入库，回报用户）。
- [ ] 中文文献仍用 GB/T 7714 作者—年份格式由 builder 处理，你只填 key。

## 编译
```
python agents\hybrid_agent\run_hybrid_pipeline.py <theme_dir> --no-enrich
```
builder 会校验 cites 合法性并生成 DOCX；若报 key 不存在，回到 INDEX.md 核对。
