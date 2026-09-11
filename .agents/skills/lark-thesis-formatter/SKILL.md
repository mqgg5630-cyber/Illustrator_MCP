---
name: lark-thesis-formatter
description: 全自动将学术论文、文献综述与科研文档排版为高校学士/硕士毕业论文（学位论文）标准格式规范。支持鲁东大学学术硕士学位论文权威模板克隆、校徽图徽完整保留、6行2列信息表填充、标准科技三线表、科学机制插图与 Zotero 双轨活体切片联动。
---

# 高校毕业论文（学位论文）全自动标准排版体系 Golden Standard

本技能提供从原始学术草稿到高校硕士/学士学位论文的工业级排版流水线，以《鲁东大学学术学位论文_Zotero活动引用版_new.docx》为唯一基准模板。

---

## 核心架构与操作指南

### 1. 唯一权威基准模板
* **路径**：`E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx`
* **原则**：直接通过 `docx.Document(TEMPLATE_PATH)` 深度克隆，严禁从空白文档从头组装，确保学校校徽/图徽、分节符、页眉页脚 100% 完整保留。

### 2. 标准封面与信息表填充
* **顶部元数据 (P00)**：分类号、单位代码、密级、学号。
* **论文题目 (P04)**：黑体 18pt 粗体居中。
* **信息表 (Table 0)**：6行2列无边框表格，宋体 14pt，左列普通右列加粗居中。

### 3. 正文清空与正文流内嵌三线表
* **正文清空**：通过 DOM 节点从目录后（P60 之后）彻底清空旧段落、旧表格与旧示例引文，根治 Zotero Refresh 重复编号。
* **章节标题多级自动编号剥离（杜绝双重标题）**：
  - 模板中 `Heading 1`、`Heading 2`、`Heading 3` 在底层 `styles.xml` 与 `numbering.xml` 中已绑定多级列表编号（`abstractNumId=6`），Word 会在渲染时动态前缀 `第%1章　`、`%1.%2　`、`%1.%2.%3　`。
  - 向 Word 写入正文标题时，必须通过正则自动剥离手工序号（例如传入 `"第1章 绪论"` 必须转为 `"绪论"`，传入 `"1.1 研究背景"` 必须转为 `"研究背景"`）。绝对严禁在段落文本中重复包含序号，否则将在正文、奇数页眉 STYLEREF 与大纲中产生严重的双重标题（如 `第1章 第1章 绪论`）。
  - 目录（`TOC 1/2/3`）样式无自动编号，必须保留显式编号（如 `第1章 绪论\t1`），确保用户在 Word 中按 F9 刷新目录时与静态目录完美对齐。
* **三线表规范**：顶底线 1.5pt 粗黑线，栏目线 0.75pt 细线，内外无竖线，表题黑体居中置于表上方。
* **插图规范**：高清科研矢量/位图，宽度 5.4~5.6 英寸居中排版，图题宋体居中置于图下方。

### 4. 封面、扉页与前置分区零残留规范 (Zero Placeholder Rule)
* **封面 (Cover, P04/Table 0)**：填写真实论文题目、作者、导师、专业与日期，绝对杜绝残留模板原作者姓名（如“温少华”）、原学院与旧题目。
* **中文扉页 (Chinese Flyleaf, P07/P09/P11)**：顶部保留 `鲁东大学硕士学位论文`，居中再次呈现论文题目，下方填写真实作者与导师信息（封面与扉页各出现一次题目符合高校装订国家标准）。
* **英文扉页 (English Flyleaf, P13..P15)**：填写真实英文论文题目、Candidate、Supervisor 与专业学院信息。
* **独创性声明页 (P17..P25)**：填写作者签名与导师签名日期，保留保密选项框。

### 5. Zotero 活体切片协议与四大铁律 (Zotero Live Integration Protocol)
* **真文献白名单校验与排伪 (Zero-Fake Citation Policy)**：正文引用与参考文献列表强制要求 100% 绑定本地 Zotero 数据库真实 8 位 Key，严禁使用任何 `AD_AMP_x` 或 `AI_Mapping_...` 伪造假条目。排版编译前自动执行 `auto_exclude_invalid_literature.py` 门禁，确保引用的全部是真实多页官方出版文献。
* **铁律 1：正文引注域驱动基石（严禁死文本）**：
  - 正文每个引注必须使用 OpenXML 复合域（`w:fldChar` 的 `begin`、`ADDIN ZOTERO_ITEM CSL_CITATION` 指令文本含真实 8 位 Zotero Key、`separate`、高亮预排版上标如 $^{[1]}$、`end`）完整包裹。
  - **绝对严禁直接写入纯文本 `[1]`**！Word 中文末参考文献（Bibliography）是由正文引注动态驱动生成的，若正文缺失复合域，Zotero 判定全文为“零引用”，点击 `Refresh` 将直接失效或清空文末列表。
* **铁律 2：文末参考文献大容器包裹（双语自适应）**：
  - 参考文献列表首尾段落必须包裹 `ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY`。
  - 中文文档锚定在 `参 考 文 献` 至 `致  谢` 之间；英文文档锚定在 `References` 至 `Acknowledgements` 之间。
* **铁律 3：属性注入与 ZIP 唯一性硬门禁（根除修复弹窗）**：
  - 在 `docProps/custom.xml` 中以每段 $\le 250$ 字符切片写入 `ZOTERO_PREF_1 ... ZOTERO_PREF_n`，并在 `[Content_Types].xml` 与 `_rels/.rels` 完成注册。
  - **重打包 ZIP 时强制跳过原模板中的同名文件（`if item.filename == "docProps/custom.xml": continue`）**，确保 ZIP 容器内条目绝对唯一，彻底根除 Word 打开时的“内容损坏/需要修复”弹窗。
* **铁律 4：零弹窗即时联动验证**：
  - WPS / Word 打开时直接显示预排版引用与文末列表，在 Word 中点击 Zotero 插件“Refresh”时自动即时刷新联动，绝对不出现“执行本操作前您需要插入引注”弹窗。

### 6. Word OpenXML 6-Section 拓扑与页眉隔离铁律 (Strict 6-Section Header Rule)
* **页眉全局渗透病因**：在清理模板残留正文段落（P61 之后）时，若误删 `<w:body>` 尾部分节符，会导致正文与参考文献全部坍缩并入 Section 3（目录节），造成全篇所有页眉强制显示为“目录”。
* **强制 6 大标准分节拓扑结构**：
  - **Section 0 (封面/扉页/声明页)**：无页眉，无页脚（`Header=[]`, `Footer=[]`）。
  - **Section 1 (中文摘要)**：独立页眉“摘要”，罗马数字页码 `I` 起算（`w:fmt="upperRoman" w:start="1"`）。
  - **Section 2 (英文 Abstract)**：独立页眉“Abstract”，罗马数字页码连续。
  - **Section 3 (目录 Contents)**：独立页眉“Contents”或“目录”，罗马数字页码连续。
  - **Section 4 (正文章节 Chapters 1~5)**：必须在正文最后一章末尾段落显式注入分节符（`s4_xml`），绑定 `header5.xml`（`STYLEREF "标题 1" \n` + `STYLEREF "标题 1"` 动态章节页眉），页脚重置为从 **`1`** 开始的阿拉伯数字（`w:fmt="decimal" w:start="1"`）。
  - **Section 5 (参考文献 References & 致谢)**：必须在 `<w:body>` 根节点末尾显式追加分节符（`s5_xml`），绑定 `header8.xml`（`STYLEREF "Heading 1 Unnumbered"` 独立页眉），阿拉伯页码自然连续。
* **样式严格绑定与回退文本净化**：
  - 参考文献与致谢一级标题样式必须显式指定为 `Heading 1 Unnumbered`，绝不允许使用 `Heading 1`。
  - `header4.xml`、`header5.xml`、`header8.xml` 内的 `<w:t>` 缓存回退文本必须与当前文档语种和标题完全一致，杜绝 Word 未刷新域时的中文/错误残留。


## 6. 权威成功案例与黄金标杆 (Golden Reference Case)
* **唯一权威基准模板**：`E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx`（严禁删除或改动）
* **标杆案例 1（食源性鲜味肽）**：`E:\0mcp-agv\ARTA_Agent_Output\基于机器学习的食源性鲜味肽高通量筛选与呈味机制解析_顶刊综述学位论文.docx`
* **标杆案例 2（深度学习抗菌肽）**：`E:\0mcp-agv\ARTA_Agent_Output\深度学习筛选抗菌肽\基于深度学习的抗菌肽高通量识别与智能设计研究_顶刊综述学位论文.docx`
* **知网真实多页 PDF 全文库**：落盘于 `E:\0mcp-agv\ARTA_Agent_Output\<主题名称>\` 与 `E:\ozotero\storage\`（零 C 盘占用）
* **全流程生成引擎**：`E:\0mcp-agv\agents\academic_thesis_agent\arta_agent.py`


