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

### 5. Zotero 活体切片协议
* **正文引用 (CSL Citation)**：正文每个引注使用 OpenXML 复合域（`w:fldChar` begin/separate/end 与 `ADDIN ZOTERO_ITEM CSL_CITATION`）包裹，预排版上标（如 $^{[1-3]}$，Times New Roman 10.5pt，克莱因蓝 `#002FA7`）。
* **参考文献域 (CSL Bibliography)**：首尾段落包裹 `ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY`。
* **首选项注册 (ZOTERO_PREF)**：在 `docProps/custom.xml` 中以 $\le 255$ 字符切片注入 GB/T 7714-2015 配置，并在 `[Content_Types].xml` 与 `_rels/.rels` 登记。
* **零弹窗保证**：WPS / Word 打开时直接显示预排版引用，点击 Zotero 插件“Refresh”时自动刷新编号，绝对不出现“执行本操作前您需要插入引注”弹窗。

---

## 6. 权威成功案例与黄金标杆 (Golden Reference Case)
* **唯一权威基准模板**：`E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx`（严禁删除或改动）
* **标杆案例 1（食源性鲜味肽）**：`E:\0mcp-agv\ARTA_Agent_Output\基于机器学习的食源性鲜味肽高通量筛选与呈味机制解析_顶刊综述学位论文.docx`
* **标杆案例 2（深度学习抗菌肽）**：`E:\0mcp-agv\ARTA_Agent_Output\深度学习筛选抗菌肽\基于深度学习的抗菌肽高通量识别与智能设计研究_顶刊综述学位论文.docx`
* **知网真实多页 PDF 全文库**：落盘于 `E:\0mcp-agv\ARTA_Agent_Output\<主题名称>\` 与 `E:\ozotero\storage\`（零 C 盘占用）
* **全流程生成引擎**：`E:\0mcp-agv\agents\academic_thesis_agent\arta_agent.py`


