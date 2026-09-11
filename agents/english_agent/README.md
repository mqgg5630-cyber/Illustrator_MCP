# 英文专属智能体工程 (English Academic Agent)

## 目录结构
```text
E:\0mcp-agv\agents\english_agent\
├── english_downloader.py         # 外网（arXiv/OpenAlex/PMC）真实英文 PDF 下载与 Zotero 挂载
├── english_purger.py             # 英文真实文献排伪审查与目录物理净化门禁
├── english_review_builder.py     # SCI 顶刊学术综述 DOCX 活体编译引擎
├── run_english_pipeline.py       # 英文一键端到端全流程运行入口
└── README.md                     # 本说明文档
```

## 一键运行命令
```powershell
python E:\0mcp-agv\agents\english_agent\run_english_pipeline.py --query "antimicrobial peptides deep learning metagenomics" --count 5 --topic-dir "E:\0mcp-agv\ARTA_Agent_Output\AMP_Metagenomics_English_Review"
```

## 交付规范与 Zotero 活体关联四大铁律
1. **真文献门禁 (Zero-Fake Gate)**：自动适配本地网络代理（127.0.0.1:10808），直连 arXiv / OpenAlex / PMC 获取官方正版多页英文 PDF，严格校验 `%PDF-` 二进制头与大小（$\ge 50\text{ KB}$）。
2. **Zotero 零 C 盘 (Zero-C-Drive)**：实体存入 `E:\ozotero\storage\<KEY>\<filename>.pdf`，写入本地 `zotero.sqlite` 生成 `journalArticle` 条目与物理附件。
3. **Word 活体引注四大铁律 (OpenXML Live Protocol)**：
   - **正文活体引注域（驱动基石）**：正文必须流式注入由 `w:fldChar` 包裹的 `ADDIN ZOTERO_ITEM CSL_CITATION` 复合域，必须绑定本地 `zotero.sqlite` 真实的 8 位 Key，绝对严禁直接写入纯文本 `[1]`！
   - **文末参考文献大容器包裹**：文末参考文献必须整体包裹在 `ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY` 复合域内（英文文档精准锚定在 `References` 至 `Acknowledgements` 之间）。
   - **属性切片与 ZIP 唯一性硬门禁**：在 `docProps/custom.xml` 中以 $\le 250$ 字符分段切片写入 `ZOTERO_PREF_1 ... ZOTERO_PREF_n`，并在 `[Content_Types].xml` 与 `_rels/.rels` 完成注册。重打包 ZIP 时强制跳过原模板旧文件（`if item.filename == "docProps/custom.xml": continue`），确保 ZIP 容器唯一，彻底根除 Word 修复弹窗！
   - **顶刊综述架构**：包含沙漏型引言、分类学分类对比表、核心量化数据（ACC/MCC/AUC/MIC/结合自由能）、构效关系讨论与标准科技三线表。

