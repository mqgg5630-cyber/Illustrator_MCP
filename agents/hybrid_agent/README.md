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
- `hybrid_review_builder.py`：高校权威学位论文标准模板克隆、中英文长篇综述编纂、双语活体复合域注入与 6-Section 页眉净化。
- `run_hybrid_pipeline.py`：端到端全流程一键命令行启动入口。
