# 中英双轨混合学术智能体 (Hybrid Academic Agent)

## 1. 定位与架构原则
中英双轨混合学术智能体 (`hybrid-academic-agent`) 专用于中国知网 (CNKI) 与国际 SCI 顶刊 (PubMed/PMC/Europe PMC/PLOS) 真实文献的联合检索、双语清洗、统一挂载与高规格学位论文/学术专著 DOCX 活体编译。

- **独立性原则**：完全独立于 `cnki_agent` 与 `english_agent`，拥有独立的管线入口、元数据清洗层与融合引用解析器。
- **双语文献等权融合**：支持按用户需求自由配比（如知网 5 篇 + 英文 5 篇）。
- **双语引用合规性**：自动适配 GB/T 7714-2015 中英文双语引用标准（中文文献标注“等/卷期”，英文文献标注“et al./Vol”）。
- **Zotero 零 C 盘挂载**：PDF 实体物理存储于 `E:\ozotero\storage\<KEY>\<filename>.pdf`，写入本地 `zotero.sqlite`。
- **Word OpenXML 6-Section 铁律**：强制 6 大分节拓扑结构，彻底免疫“页眉全是目录”的全局渗透问题。

## 2. 核心文件清单
- `hybrid_downloader.py`：中英文双源文献元数据归一化、PDF 排伪审查与 Zotero 本地物理挂载。
- `hybrid_review_builder.py`：**`outline.json` 驱动**的通用编译器——克隆学位论文模板、渲染大纲章节/三线表、按首次引用顺序自动编号并注入双语 Zotero 活体复合域、6-Section 页眉净化、编译后自检。不含任何课题正文。
  - `--skeleton`：依据 `manifest.json` 生成 `outline.skeleton.json` 供 Antigravity 填写。
  - 环境变量 `HUB_THESIS_TEMPLATE` / `--template` 可覆盖模板路径；`HUB_SCRATCH_DIR` 覆盖临时目录。
- `run_hybrid_pipeline.py`：端到端入口（`--theme-dir` 必填，`--col-name`、`--skip-zotero` 可选）；缺少 `outline.json` 时退出码 2 并生成骨架。
- `examples/outline.collagen_example.json`：胶原蛋白 MD/对接综述的完整黄金示例（5 章 10 节 22 处引注 1 张表）。

## 3. 课题目录约定
```
<主题目录>/
├── manifest.json            # 文献元数据（zotero_key, lang, title, authors, journal, year, volume, issue, pages, doi, pdf_filename, abstract）
├── *.pdf                    # 真实多页 PDF
├── outline.skeleton.json    # --skeleton 生成，供撰写参考
├── outline.json             # Antigravity 撰写的正文大纲（编译输入）
└── <主题名>_学术专著论文.docx  # 编译输出
```

## 4. outline.json 结构速览
```json
{
  "meta": {"title_zh_line1": "...", "title_zh_line2": "...", "title_en": "...", "author_zh": "...", "...": "..."},
  "abstract_zh": ["段1", "段2"], "keywords_zh": "A；B；C",
  "abstract_en": ["para1", "para2"], "keywords_en": "A; B; C",
  "chapters": [
    {"title": "绪论", "sections": [
      {"title": "研究背景", "blocks": [
        {"type": "paragraph", "segments": [
          {"text": "……结论句", "cite": ["7WUA8PLB"]},
          {"text": "。无引注的过渡句。"}
        ]},
        {"type": "table", "caption": "关键指标对比", "headers": ["体系", "方法", "指标"], "rows": [["...", "...", "..."]]}
      ]}
    ]}
  ],
  "acknowledgement": ["段1", "段2"]
}
```
`cite` 可为多个 key 或 `"__ALL__"`；编号与参考文献顺序由编译器决定。
