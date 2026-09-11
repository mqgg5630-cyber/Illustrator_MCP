---
name: english-academic-agent
description: "英文专属智能体与外网学术文献工程技能：外网多源学术文献（PubMed/PMC, Europe PMC, OpenAlex, arXiv, bioRxiv, Unpaywall, Crossref）检索与正版多页英文 PDF 全文下载、零C盘 Zotero 实体附件物理挂载、排伪审查以及 SCI 顶刊综述/英文学位论文 DOCX 活体编译一键完成。"
---

# English Academic Agent — 外网英文学术文献检索、PDF下载、Zotero挂载与综述编译全流程

## 适用场景与触发词
- **触发词**：“英文综述”、“下载外网文献”、“english academic agent”、“外网pdf下载”、“sci综述”、“海外文献导入zotero”
- **核心定位**：专用于国际权威学术数据库（PubMed / PMC, Europe PMC, OpenAlex, arXiv, bioRxiv, Crossref, Unpaywall）的真实文献检索、正版多页英文 PDF 下载、Zotero 物理附件挂载与 SCI 顶刊综述 DOCX 活体编译。

---

## 核心工作流

### Step 1: 外网多源文献检索与真实 PDF 获取
- **数据源 API 矩阵**：
  - **Europe PMC & PMC**：调用 REST API `https://www.ebi.ac.uk/europepmc/webservices/rest/search` 获取 `isOpenAccess=Y` 与 `pmcid`，直连官方 PDF 端点：`https://europepmc.org/backend/ptpmcrender.fcgi?accid=<PMCID>&blobtype=pdf`；
  - **OpenAlex**：调用 `https://api.openalex.org/works?search=<query>`，提取 `open_access.oa_url` 与 `best_oa_location.pdf_url`；
  - **bioRxiv / medRxiv / arXiv**：通过 DOI 或 eprint ID 直连预印本官方二进制 PDF 数据流；
  - **Unpaywall**：通过 `https://api.unpaywall.org/v2/<DOI>?email=research@academic.org` 提取官方出版机构下发的 Legal OA PDF。
- **存储隔离**：文献集中存放于 `E:\0mcp-agv\ARTA_Agent_Output\<主题名称>\`。

### Step 2: 纯正性审查与自动排伪门禁 (Zero-Fake Gate)
- 运行 `python E:\0mcp-agv\scripts\auto_exclude_invalid_literature.py "<主题目录>"`；
- 检查文件前 1024 字节是否包含 `%PDF-` 二进制头；
- 验证有效页数 $\ge 2$ 页，自动剔除下载不完整、HTML 403 拦截页或错误文件；
- 自动清理与白名单不符的残留文件，确保物理文件 100% 真实有效。

### Step 3: 零 C 盘 Zotero 本地数据库物理挂载
- 运行 `english_pdf_downloader.py --sync-zotero`：
  - 将真实多页英文 PDF 写入 `E:\ozotero\storage\<KEY>\<filename>.pdf`；
  - 写入 `E:\ozotero\zotero.sqlite` 的 `items`、`itemData` 与 `itemAttachments` 表（`itemTypeID=4` 期刊条目，`itemTypeID=3` 且 `linkMode=0` 物理附件）；
  - 自动创建或关联主题集合（Collection）；
  - 导出同步更新的 `manifest.json`、`References.bib` 与 `References.ris`。

### Step 4: SCI 顶刊学术综述与学位论文 DOCX 活体编译
- **学术综述长篇架构**：
  - 撰写包含沙漏型引言、分类学分类对比表、核心量化数据（ACC/MCC/AUC/MIC/结合自由能）、构效关系机理讨论的万字级高密度长篇学术综述；
  - 表格采用标准科技三线表（顶底线 1.5pt，栏目线 0.75pt）；
  - 剥离各级标题手工序号（纯文本写入），100% 免疫双重标题。
- **OpenXML 活体复合域四大铁律注入**：
  - **正文活体引注域铁律（驱动基石）**：正文必须流式注入由 `w:fldChar` 包裹的 `ADDIN ZOTERO_ITEM CSL_CITATION` 复合域，必须绑定本地 `zotero.sqlite` 真实的 8 位 Key（`uris: ["http://zotero.org/users/local/.../items/<KEY>"]`），绝对严禁直接写入纯文本 `[1]`！因为 Zotero 文末参考文献是由正文引注动态驱动的，缺失引注域会导致 Zotero 判定引用数为 0。
  - **文末参考文献大容器包裹**：文末参考文献必须整体包裹在 `ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY` 复合域内（英文文档精准锚定在 `References` 至 `Acknowledgements` 之间）。
  - **属性切片与 ZIP 唯一性硬门禁**：在 `docProps/custom.xml` 中以 $\le 250$ 字符分段切片写入 `ZOTERO_PREF_1 ... ZOTERO_PREF_n`，并在 `[Content_Types].xml` 与 `_rels/.rels` 完成注册。重打包 ZIP 时强制跳过原模板旧文件（`if item.filename == "docProps/custom.xml": continue`），确保 ZIP 容器唯一，Word 打开零修复弹窗、点击 Zotero Refresh 即时联动！
  - **6-Section 拓扑与页眉隔离铁律**：必须构建完整 6 分节（封面、摘要、Abstract、Contents、正文 1~5 章、References/致谢）。正文末尾段落绑定 `header5.xml` 并重置页码为 1；References 样式必须为 `Heading 1 Unnumbered` 并绑定 `header8.xml`，回退文本净化为英文，彻底根除“页眉全是目录”或中文回退。


---

## 快速执行命令

```powershell
# 1. 批量检索并下载外网真实多页英文 PDF 并挂载 Zotero
python E:\0mcp-agv\scripts\english_pdf_downloader.py --query "antimicrobial peptides deep learning metagenomics" --count 20 --out-dir "E:\0mcp-agv\ARTA_Agent_Output\AMP_Metagenomics_English" --sync-zotero

# 2. 执行排伪与目录净化门禁
python E:\0mcp-agv\scripts\auto_exclude_invalid_literature.py "E:\0mcp-agv\ARTA_Agent_Output\AMP_Metagenomics_English"

# 3. 编译生成带 Zotero 活体双轨引用的 SCI 学术综述 DOCX
python E:\0mcp-agv\scripts\build_english_academic_review.py --theme-dir "E:\0mcp-agv\ARTA_Agent_Output\AMP_Metagenomics_English"
```
