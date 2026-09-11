---
name: hybrid-academic-agent
description: "中英双轨混合学术智能体与文献工程技能：知网 (CNKI) 中文核心与外网 SCI 顶刊 (PubMed/PMC/Europe PMC/PLOS) 真实文献联合检索、多页正版 PDF 批量下载、零 C 盘 Zotero 物理挂载、双语排伪审查、GB/T 7714 双语活体引用以及 6-Section 拓扑高规格学术专著/学位论文 DOCX 活体编译一键完成。"
---

# Hybrid Academic Agent — 中英双轨混合学术智能体工作流体系

## 1. 核心定位与适用场景
- **触发词**：“中英文综述”、“中英混合文献”、“双语学术专著”、“知网和英文混合”、“hybrid academic agent”、“中英双轨论文”
- **定位**：完全独立于 `cnki-academic-agent` 与 `english-academic-agent` 的第三套专属智能体，专用于中国知网 (CNKI) 与国际 SCI 顶刊的联合挖掘、等权双语归一化、双语活体复合域注入与 6-Section 完美多节学位论文/学术专著编译。

---

## 2. 核心工作流规范

### Step 1: 双轨真实多页 PDF 采集与存储隔离
- **国内通道**：知网 Edge CDP 鉴权会话直接获取知网官方正版出版多页 `%PDF-1.x` 数据流。
- **国际通道**：PubMed / PMC / PLOS / Europe PMC / OpenAlex 官方合法 Open Access 端点直连出版级完整多页 PDF。
- **存储隔离**：各课题文献集中落盘于 `E:\0mcp-agv\ARTA_Agent_Output\<主题名称>\`。

### Step 2: 纯正性核验与自动排伪门禁 (Zero-Fake Gate)
- 运行 `python E:\0mcp-agv\scripts\auto_exclude_invalid_literature.py "<主题目录>"`；
- 逐一检验前 1024 字节 `%PDF-` 标识符及多页（$\ge 2$ 页）二进制有效性；
- 物理抹除损坏、单页拦截页及与 `manifest.json` 不符的非白名单文件。

### Step 3: Zotero 零 C 盘本地数据库物理挂载
- 运行 `hybrid_downloader.py`：
  - 将中英文 PDF 物理落盘至 `E:\ozotero\storage\<KEY>\<filename>.pdf`；
  - 写入 `E:\ozotero\zotero.sqlite` 生成 `itemTypeID=4` 条目，并建立 `linkMode=0` 物理附件关联；
  - 自动在 Zotero 中新建独立的主题分类集合（Collection）；
  - 导出同步更新的双语 `manifest.json`、`References.bib` 与 `References.ris`。

### Step 4: 学术专著/学位论文 Word OpenXML 活体编译
- **唯一权威基准模板**：克隆 `E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx`。
- **强制 6-Section 多节拓扑架构（根除“页眉全是目录”）**：
  1. **Section 0 (封面/扉页/声明页)**：无页眉，无页脚。
  2. **Section 1 (中文摘要)**：独立页眉“摘要”，罗马数字页码 `I` 起算。
  3. **Section 2 (英文 Abstract)**：独立页眉“Abstract”，罗马数字页码连续。
  4. **Section 3 (目录 Contents)**：独立页眉“目  录”，罗马数字页码连续。
  5. **Section 4 (正文章节 Chapters 1~5)**：正文最后一章末尾显式注入 `s4_xml`（引用 `header5.xml`），绑定 `STYLEREF "标题 1" \n` 与 `STYLEREF "标题 1"` 动态章节页眉，页脚重置为从 **`1`** 开始的阿拉伯数字。
  6. **Section 5 (参考文献 References & 致谢)**：在 `<w:body>` 末尾显式追加 `s5_xml`（引用 `header8.xml`），绑定 `STYLEREF "Heading 1 Unnumbered"` 独立页眉，阿拉伯页码自然连续。
- **标题多级编号剥离（双重标题免疫）**：正文 `Heading 1/2/3` 标题通过正则自动剥离手动序号，防止在大纲与页眉产生双重编号。
- **双语活体复合域注入四大铁律**：
  - **正文活体引注域铁律**：正文流式注入 `ADDIN ZOTERO_ITEM CSL_CITATION` 复合域，中文文献作者标注“等”，英文文献作者标注“et al.”，绑定本地真实 8 位 Zotero Key。
  - **参考文献大容器包裹**：文末参考文献列表完整包裹于 `ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY` 复合域内。
  - **属性切片与 ZIP 唯一性硬门禁**：在 `docProps/custom.xml` 中以 $\le 250$ 字符分段切片写入 `ZOTERO_PREF_1 ... ZOTERO_PREF_n`，重打包 ZIP 时跳过同名旧文件，保证容器内条目绝对唯一，Word 打开零修复弹窗、Zotero 点击 Refresh 即时联动！
  - **页眉回退文本净化**：底层缓存文本清洗为对应章节标题，杜绝 Word 未刷新域时的中文/错误残留。

---

## 3. 权威成功案例与黄金标杆 (Golden Reference Case)
* **中英双轨成功标杆**：`E:\0mcp-agv\ARTA_Agent_Output\Collagen_Stability_MD_Docking_Hybrid_Review\Collagen_Stability_MD_Docking_Hybrid_Review_学术专著论文.docx`
* **知网与外网 10 篇真实多页 PDF 库**：落盘于 `E:\0mcp-agv\ARTA_Agent_Output\Collagen_Stability_MD_Docking_Hybrid_Review\` 与 `E:\ozotero\storage\`
* **独立工程管线**：`E:\0mcp-agv\agents\hybrid_agent\`
