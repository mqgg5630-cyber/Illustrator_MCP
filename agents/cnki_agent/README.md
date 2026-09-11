# 知网专属智能体工程 (CNKI Academic Agent)

## 目录结构
```text
E:\0mcp-agv\agents\cnki_agent\
├── cnki_downloader.py         # 知网真实多页 PDF 下载与 Zotero 物理挂载引擎
├── cnki_purger.py             # 知网真实文献排伪审查与目录物理净化门禁
├── cnki_thesis_builder.py     # 高校学位论文与学术长篇综述 DOCX 活体编译引擎
├── run_cnki_pipeline.py       # 知网一键端到端全流程运行入口
└── README.md                  # 本说明文档
```

## 一键运行命令
```powershell
python E:\0mcp-agv\agents\cnki_agent\run_cnki_pipeline.py --topic-dir "E:\0mcp-agv\ARTA_Agent_Output\阿尔兹海默症肠道宏基因组抗菌肽差异"
```

## 交付规范与 Zotero 活体关联四大铁律
1. **真文献门禁 (Zero-Fake Gate)**：仅收录知网官方 `%PDF-1.x` 纯正多页（10~30页）全文，自动抹除占位符或伪造文件。
2. **Zotero 零 C 盘 (Zero-C-Drive)**：实体存入 `E:\ozotero\storage\<KEY>\<filename>.pdf`，写入本地 `zotero.sqlite`。
3. **Word 活体引注四大铁律 (OpenXML Live Protocol)**：
   - **正文活体引注域（驱动基石）**：正文必须流式注入由 `w:fldChar` 包裹的 `ADDIN ZOTERO_ITEM CSL_CITATION` 复合域，必须绑定本地 `zotero.sqlite` 真实的 8 位 Key，绝对严禁直接写入纯文本 `[1]`！
   - **文末参考文献大容器包裹**：文末参考文献必须整体包裹在 `ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY` 复合域内（中文文档锚定在 `参 考 文 献` 至 `致  谢` 之间）。
   - **属性切片与 ZIP 唯一性硬门禁**：在 `docProps/custom.xml` 中以 $\le 250$ 字符分段切片写入 `ZOTERO_PREF_1 ... ZOTERO_PREF_n`，并在 `[Content_Types].xml` 与 `_rels/.rels` 完成注册。重打包 ZIP 时强制跳过原模板旧文件（`if item.filename == "docProps/custom.xml": continue`），确保 ZIP 容器唯一，彻底根除 Word 修复弹窗！
   - **标题多级编号剥离**：纯文本写入章节标题，自动剥离显式序号，100% 免疫双重标题。

