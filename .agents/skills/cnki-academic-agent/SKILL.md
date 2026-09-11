---
name: cnki-academic-agent
description: "知网专属智能体与全自动文献工程技能：全自动将中国知网 (CNKI) 原生真实多页学术文献检索、结构化元数据提取、官方原生 PDF 全文下载、零C盘 Zotero 物理附件挂载、排伪审查以及高规格高校学位论文/中文学术综述 DOCX 活体编译一键完成。"
---

# CNKI Academic Agent — 中国知网专属文献与论文写作工作流

## 适用场景与触发词
- **触发词**：“知网论文”、“知网文献检索”、“cnki agent”、“知网综述”、“知网导出zotero”
- **核心定位**：专注于中国知网 (CNKI) 原生真实多页学术文献的全生命周期处理，杜绝任何假文献。

## 核心工作流

### Step 1: 真实知网官方 PDF 检索与直连下载
- 采用 Edge CDP 鉴权会话（提取活动 Cookies）+ `Referer: https://kns.cnki.net/...` 握手头向 `docdown.cnki.net` 发起请求；
- 获取知网官方正版 `%PDF-1.x` 纯正多页（$\ge 2$ 页，通常 10~30 页）二进制全文；
- 存储于 `E:\0mcp-agv\ARTA_Agent_Output\<主题>\`。

### Step 2: 纯正性核验与自动排伪门禁
- 运行 `python E:\0mcp-agv\scripts\auto_exclude_invalid_literature.py "<主题目录>"`；
- 剔除伪造占位命名、损坏文件、单页拦截页以及非白名单文件，物理净化目录。

### Step 3: Zotero 本地数据库物理挂载 (零 C 盘)
- 写入 `E:\ozotero\zotero.sqlite`，将 PDF 实体无损存入 `E:\ozotero\storage\<KEY>\<filename>.pdf`；
- 建立 `itemTypeID=3`、`linkMode=0` 附件关联，客户端展示 `📎 附件`。
- 导出 `manifest.json`、`References.bib`、`References.ris`。

### Step 4: 顶刊学术综述与学位论文 OpenXML 活体编译
- 以鲁东大学标准模板为基准，剥离各级标题手工序号（100% 免疫双重标题）；
- **正文活体引注域铁律（驱动基石）**：正文每一个引注必须使用 OpenXML 复合域（`w:fldChar` 的 `begin`、`ADDIN ZOTERO_ITEM CSL_CITATION` 指令文本含真实 8 位 Zotero Key、`separate`、高亮预排版上标如 $^{[1]}$、`end`）完整包裹。**绝对严禁直接写入纯文本 `[1]`**！因为 Zotero 文末参考文献是由正文引注动态驱动的，缺失引注域会导致 Zotero 判定引用数为 0。
- **文末参考文献大容器包裹**：文末参考文献必须整体包裹在 `ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY` 复合域内（中文文档精准锚定在 `参 考 文 献` 至 `致  谢` 之间）。
- **属性切片与 ZIP 唯一性硬门禁**：在 `docProps/custom.xml` 中以 $\le 250$ 字符分段切片写入 `ZOTERO_PREF_1 ... ZOTERO_PREF_n`，并在 `[Content_Types].xml` 与 `_rels/.rels` 完成注册。重打包 ZIP 时强制跳过原模板旧文件（`if item.filename == "docProps/custom.xml": continue`），确保 ZIP 容器唯一，Word 打开零修复弹窗、点击 Zotero Refresh 即时联动！
- **6-Section 拓扑与页眉隔离铁律**：必须构建完整 6 分节（封面、摘要、Abstract、目录、正文 1~5 章、参考文献/致谢）。正文末尾段落绑定 `header5.xml` 并重置页码为 1；参考文献样式必须为 `Heading 1 Unnumbered` 并绑定 `header8.xml`，彻底根除“页眉全是目录”的全局渗透。

