# AGENTS.md - Workspace Behavioral Guidelines & Standards

## 1. 核心工作区规范 (Core Workspace Rules)
- **代码与脚本规范**：所有生成的 Python 脚本必须具备完整的错误处理与自检机制。
- **PPT 与图形设计标准**：
  - 强制 16:9 宽屏架构。
  - 必须严格遵循 `cyber-ppt` 与 `scientific-figure-shapes` 的排版阶梯（字号 $\ge 16\text{pt}$，充足呼吸感，杜绝拥挤重叠）。
  - 采用 Nature/Cell 顶刊级优雅配色（米白/淡雅冷灰底 + 饱和度适中的功能色）。

## 2. 幻灯片生成工具与风格路由准则 (Strict Tool Routing & Separation)
- **严禁工具与视觉体系混淆**，严格根据用户指定的工具/风格执行：
  - **`cyber-ppt` (PPTX)**：Nature/McKinsey 顶刊咨询级汇报，含纯矢量自绘 SVG 图标、象牙白底 `#F6F5F0`、带彩色顶栏的立柱卡片与金牌 Takeaway。
  - **`guizang-ppt-skill` (HTML/PPTX)**：单文件横向翻页网页 PPT 与双轨 PPTX 编译体系（含 WebGL 网格/ASCII 呼吸背景、键盘翻页、`P` 演讲者模式与 22 种锁定版式；字号严格 $\ge 18\text{pt}$，支持纯矢量可编辑 DrawingML 与 Playwright 高清像素级双轨导出）。
  - **`swiss-pptx` (PPTX)**：瑞士国际主义原生 16:9 PowerPoint，含克莱因蓝 IKB `#002FA7` 单一高亮色、1px 发丝线、2x2 数据 KPI 塔与对仗双栏。
  - **`banana-slides` (PPTX/AI Native)**：AI 原生演示文稿生成引擎，含香蕉暖金 `#F5A623` 与深板岩蓝 `#1E293B`、圆角卡片、多维计算漏斗与全流程 CLI。
  - **`ppt-master` (PPTX/DrawingML Native)**：AI 驱动的纯矢量 SVG 到 DrawingML 原生可编辑 PowerPoint 编译体系，支持 `academic-research` 学术研究规范、100% 独立内置质量自检门禁 (`svg_quality_checker.py`) 与无损编译。
  - **`dashi-ppt` (PPTX/HTML)**：基于 React 布局系统与 Headless Chromium 的高质量多主题幻灯片引擎，支持 8 套专业主题包（含 `theme07` 冷白调研风）、多变体版式与原生 100% 可编辑 PPTX 导出。

## 3. 强制自检闭环规范 (Mandatory Visual Inspection & Self-Check Loop)
- 生成任何 PPTX、HTML 或可视化图像后，必须执行自动导出与图像自检。
- 保证无文本截断、无坐标重叠、无溢出，重要内容严格 $\ge 20\text{pt}$。

## 4. 学位论文排版与标准模板规范 (Thesis Formatting & Template Standard)
- **基准模板唯一性**：高校毕业论文/学位论文排版必须以 `E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx` 为唯一权威基准模板。
- **标准封面与版面架构**：
  - 封面：顶部元数据（分类号/单位代码/密级/学号，宋体 12pt）+ `学术硕士学位论文`（黑体 30pt 粗体居中）+ 论文题目（黑体 18pt 粗体居中）+ 6行2列居中无边框信息表格（宋体 14pt，左列普通右列加粗）。
  - 扉页：中文扉页（二号宋体粗体 `鲁东大学硕士学位论文` + 题目 + 三号作者信息）与英文扉页（二号 Times New Roman 粗体 + 题目 + 三号英文作者信息）。
  - 声明页：`学位论文原创性声明和使用授权说明`（三号黑体粗体居中）+ 独创性声明 + 版权授权书（含保密选项勾选与双签名）。
  - 正文：A4 纸张，边距 上3.0cm/下2.5cm/左3.0cm/右3.0cm，页眉2.0cm/页脚2.2cm，模式一多级编号（第1章、1.1、1.1.1），标准三线表，公式编号居右对齐，Zotero 活体引用双轨兼容。

## 5. 知网真实多页 PDF 全文、Zotero 实体挂载与活体引用编译黄金标准 (Verified Gold Standard)
- **知网原生多页真实 PDF 下载规范**：
  - 严禁使用本地 1 页/占位伪造 PDF，严禁下载 CAJ 格式。
  - 必须通过文献详情页端点获取知网官方正版出版的完整 PDF（10~30页，含完整机制图、数据表、分子量、活性指标与参考文献）。
  - 采用 Edge CDP 鉴权会话（提取活动 Cookies）+ `Referer: https://kns.cnki.net/...` 握手头向 `docdown.cnki.net` 直连请求，获取官方 `%PDF-1.6` 纯正二进制数据流。
- **Zotero 实体附件深度挂载与零 C 盘规范 (Zero-C-Drive Rule)**：
  - 物理路径严格落盘于 E 盘：PDF 存储于 `E:\ozotero\storage\<KEY>\<filename>.pdf`，本地临时下载于 `E:\Users\文少\Downloads\`。绝对禁止向 C 盘写入任何数据。
  - 在 Zotero 数据库中建立 `itemTypeID=3`、`linkMode=0` (imported_file) 的物理附件关联，确保客户端直接显示真实可双击阅读的 `📎 附件`。
- **Word OpenXML 活体双轨编译规范（彻底根除“需要插入引注”弹窗与断联死文本）**：
  - **正文引注域铁律（驱动基石）**：正文每一个引注必须使用 OpenXML 复合域（`w:fldChar` 的 `begin`、`ADDIN ZOTERO_ITEM CSL_CITATION` 指令文本含真实 8 位 Zotero Key、`separate`、高亮预排版上标如 $^{[1]}$、`end`）完整包裹。**绝对严禁直接写入纯文本 `[1]`**！因为 Zotero 的文末参考文献是由正文引注动态驱动生成的，若正文缺失复合域，Zotero 判定全文为“零引用”，点击 `Refresh` 将直接失效或清空文末列表。
  - **参考文献域语言自适应包裹**：必须使用 `ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY` 将整个参考文献列表包裹于复合域内。正则表达式与切片锚点必须根据文档语种精准适配：中文文档锚定 `参 考 文 献` 至 `致  谢` 之间；英文文档锚定 `References` 至 `Acknowledgements` 之间。
  - **属性注入与 ZIP 唯一性硬门禁**：必须在 `docProps/custom.xml` 中以每段 $\le 250$ 字符切片写入 `ZOTERO_PREF_1 ... ZOTERO_PREF_n`，并在 `[Content_Types].xml` 与 `_rels/.rels` 完成注册。**重打包 ZIP 时强制跳过原模板中的同名文件（`if item.filename == "docProps/custom.xml": continue`）**，确保 ZIP 容器内条目绝对唯一，彻底根除 Word 打开时的“内容损坏/需要修复”弹窗。
  - **真实 8 位 Key 镜像绑定**：正文与文末引用的所有文献 ID，必须 100% 对应本地 `E:\ozotero\zotero.sqlite` 中实际注册的 8 位随机 Key（如 `TLM24G3U`），URI 格式强制为 `http://zotero.org/users/local/user/items/<KEY>`，严禁使用任何 `AD_AMP_x` 等临时字符串。
- **唯一基准模板与成功案例标杆**：
  - 模板唯一权威基准：`E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx`。
  - 中文成功标杆：`E:\0mcp-agv\ARTA_Agent_Output\阿尔兹海默症肠道宏基因组抗菌肽差异\基于深度学习预测阿尔兹海默症肠道宏基因组分阶段病人的抗菌肽差异_顶刊综述学位论文.docx`。
  - 英文成功标杆：`E:\0mcp-agv\ARTA_Agent_Output\AMP_Metagenomics_English_Review\AMP_Metagenomics_English_Review_SCI_Review_Monograph.docx`。
  - 中英混合双轨版式与 6-Section 拓扑格式验证用例 (Layout Demo, 含合成版式测试数据，非真实知网抓取标杆)：`E:\0mcp-agv\ARTA_Agent_Output\Collagen_Stability_MD_Docking_Hybrid_Review\Collagen_Stability_MD_Docking_Hybrid_Review_学术专著论文.docx`。


## 6. 顶刊长篇学术综述深度撰写规范 (Deep Review Monograph Standard)
- **拒绝表面化摘要罗列**：面对数十页、上百页真实知网/SCI PDF 全文，绝对禁止仅生成几个概括性段落。必须输出万字以上、具备专著分量的模式一多级编号深度学术论述。
- **全要素深度数据萃取**：正文必须详实引述文献中的具体量化数据（序列电荷数、疏水矩、APD3/DRAMP 收录量、SVM/CNN/GAT 的 ACC/MCC 真实测试值、微量肉汤稀释法测定的 MIC 真实值、结合自由能 $\Delta G$）。
- **模式一深度结构化展开**：每个章节细化至三级标题（如 1.1、1.1.1、1.1.2、1.2、1.2.1...），对方法演进、互作机理、技术瓶颈展开严密的批判性推演与交叉对比。

## 7. 实验型科技论文标准架构规范 (Experimental Research Article Standard)
- **严格遵循 IMRAD 实验论文体系**：
  - **引言 (Introduction)**：明确科学假说（Hypothesis）、未解难题与创新切入点。
  - **实验材料与方法 (Materials and Methods)**：严谨可重复，包含标准菌株（ATCC 编号）、仪器型号（HPLC、MALDI-TOF-MS、SEM）、化学合成多肽纯度（$\ge 95\%$）、分子动力学模拟体系（POPC/POPG 膜、100 ns GROMACS、MM-PBSA）、体外抑菌 MIC/MBC 测定、溶血实验 HC50、场发射扫描电镜样品制备与统计学处理（$N \ge 3$、$P < 0.05$）。
  - **实验结果与分析 (Results)**：数据与图表双驱动，包含计算模型消融对比表、宏基因组挖掘特征雷达图、体外抑菌 MIC 三线表、溶血毒性选择性指数（TI = HC50 / MIC）与扫描电镜膜穿孔形貌。
  - **机理解析与讨论 (Discussion)**：构效关系深入归因（阳离子残基与疏水核心配比、跨膜孔道动力学演变、耐药阻遏优势）。
  - **结论 (Conclusions)**：提炼 3~4 条定量结论并指明应用前景。

## 8. Light-Skills 写作学术严谨性与论证闭环体系 (Light Academic Rigor & Claim-Evidence System)
- **Claim-Evidence Binding (论点-证据严格绑定)**：所有核心论断（Claim）必须强制绑定具象化的量化实验数据或文献实测数据作为证据支撑（如真实数据集样本量、特征维度、MIC 实测值、ACC/MCC/AUC 评测指标、结合自由能、HC50 溶血毒性等），杜绝任何脱离数据的空泛修辞。
- **Argument Contract (论证逻辑闭环契约)**：引言与各核心章节必须严格遵循“科学痛点 (Pain) $\to$ 现有文献瓶颈 (Gap) $\to$ 创新理论/计算洞察 (Insight) $\to$ 核心贡献 (Contribution) $\to$ 决定性实验证据 (Decisive Evidence)”的严密逻辑链。
- **Contribution Consistency (三处绝对一致性原则)**：摘要 (Abstract)、引言贡献段 (Introduction Contributions) 与结论 (Conclusions) 中声明的科研贡献、核心方法论及定量指标数值必须 100% 镜像对齐，杜绝前后矛盾或数值漂移。
- **Anti-Overclaim (严防过度宣称与边界限定)**：学术措辞强度必须严格匹配证据等级（未见显著差异者必须严谨陈述为“未见显著差异”，计算预测与体外抑菌、溶血实测必须清晰界定验证边界），严禁无凭据的绝对化定性词汇。
- **跨材料术语与指标一致性 (Cross-Material Consistency)**：全篇专业术语（如 AMPs、ESM-2、BiLSTM、MIC、HC50、TI）、指标命名、数学符号及缩写首注必须高度统一。

## 9. 标题多级编号剥离与双重标题免疫规范 (Heading Number Stripping & Double-Title Immunity)
- **Word 原生多级编号叠加机理**：基准模板（鲁东大学官方标准模板）中的 `Heading 1`（标题 1）、`Heading 2`（标题 2）、`Heading 3`（标题 3）在底层 `styles.xml` 与 `numbering.xml` 中已深度绑定自动多级列表编号（`abstractNumId=6`），Word/WPS 在排版渲染时会自动在段落前插入 `第%1章　`、`%1.%2　`、`%1.%2.%3　`。
- **正文写入纯文本铁律**：向 Word 写入正文章节标题时，必须通过正则自动剥离手工章节序号（如将 `"第1章 绪论"` 剥离为 `"绪论"`，将 `"1.1 研究背景"` 剥离为 `"研究背景"`）。绝对严禁在段落文本中重复包含序号，否则将导致正文、奇数页眉动态 STYLEREF 域与 Word 左侧大纲导航栏全部沦为严重的“双重标题”（如 `第1章 第1章 绪论`、`1.1 1.1 xxx`）。
- **目录（TOC）双轨契约**：目录样式（`TOC 1/2/3`）本身无自动编号，必须保留显式编号文本（如 `第1章 绪论\t1`）；当用户在 Word 中按 F9 刷新目录时，Word 从剥离后的 `Heading 1` 提取出的正好是单一标准的 `第1章 绪论`，实现静态与动态 100% 镜像对齐。
- **封面与扉页双重标题规范界定**：第 1 页封面（外封）与第 2 页中文扉页（内封）各自独立保留论文题目，是教育部与学校学位办标准装订架构的法定设计，不得遗漏或误判。

## 10. 真实知网文献批量采集与主题目录隔离规范 (Multi-Theme CNKI Batch Ingest & Isolation Standard)
- **主题专属目录隔离**：每个研究主题必须在 `E:\0mcp-agv\ARTA_Agent_Output\<主题名称>\` 下建立独立目录，默认检索并精准下载至多 20 篇官方真实出版的知网/SCI 多页 PDF。
- **零 C 盘与物理挂载闭环**：所有 PDF 实体无损落盘至 `E:\ozotero\storage\<KEY>\<filename>.pdf`，并自动向本地 Zotero 数据库建立关联条目，输出对应的 `.bib` 与 `.ris` 文献库文件，支持跨章节长篇深度数据挖掘。

## 11. 真实文献纯正性核验与伪造/错误文献自动排除门禁 (Automated Fake Literature Detection & Eviction Gate)
- **真文献白名单绝对准则 (Authentic Literature Whitelist)**：
  - 进入知识库与最终论文的文献必须具备中国知网 (CNKI) 或 SCI 权威数据库官方正版出版元数据，且必须拥有真实的 `%PDF-1.x` 多页（$\ge 2$ 页）二进制全文。
  - **严禁虚构、伪造、推测或合成任何参考文献**，严禁使用非真实存在的文章名（如通用模式生成的 `AI_Mapping_...`、`deep-learning-in-...` 等占位符文件）。
- **自动化检测与排伪流程 (Automated Detection & Eviction)**：
  - 在执行任何综述撰写、数据萃取、Zotero 导入或 `manifest.json` 导出前，必须前置执行纯正性审查脚本（`auto_exclude_invalid_literature.py`）。
  - **自动判定并排除四类错误文献**：
    1. **伪造/占位文献**：文件名包含常见占位词、缺少真实出版年卷期或期刊号者；
    2. **伪 PDF 文件**：非 `%PDF-` 二进制头、0 字节或由本地死循环/文本伪造的 1 页虚假 PDF；
    3. **非目标主题残留文献**：历史遗留、跨主题交叉污染或未收录在本次 `manifest.json` 中的多余文件；
    4. **虚假 Zotero Key**：未在本地 `zotero.sqlite` 注册或使用临时字符串（如 `AD_AMP_x`）的条目。
  - **物理磁盘自动净化 (Directory Sanitization)**：
    - 一旦检出伪造、错误或未受纳文献，系统必须立即将其从目标目录物理抹除，确保目标工作目录与 `manifest.json`、`References.bib`、`References.ris` 保持 100% 镜像洁净。

## 12. 双轨专属学术智能体架构与路由准则 (Dual-Track Academic Agent Architecture)
为确保国内学术与国际学术两套文献技术栈互不污染，工作区正式确立并部署双轨专属智能体：

### 1. 知网专属智能体 (CNKI Academic Agent, `cnki-academic-agent`)
- **定位与适用**：中国知网 (CNKI) 中文学术文献检索、硕博学位论文、中文核心期刊（北大核心、CSCD）挖掘。
- **采集通道**：Edge CDP 鉴权会话（提取活动 Cookies）+ `Referer: https://kns.cnki.net/...` 握手头直连 `docdown.cnki.net` 获取官方正版 `%PDF-1.x` 纯正多页数据流。
- **知识库挂载**：零 C 盘落盘至 `E:\ozotero\storage\<KEY>\<filename>.pdf`，写入本地 `zotero.sqlite` 生成实体条目与 `📎 附件`。
- **论文输出**：高校标准学位论文模式（鲁东大学官方模板克隆、校徽图徽保留、纯文本剥离标题序号、正文复合域 `ADDIN ZOTERO_ITEM` + 参考文献 `ADDIN ZOTERO_BIBL` + `docProps/custom.xml` 注册 GB/T 7714 活体双轨复合域引用）。

### 2. 英文专属智能体 (English Academic Agent, `english-academic-agent`)
- **定位与适用**：国际开源与顶刊学术文献（PubMed Central / Europe PMC, OpenAlex, arXiv, bioRxiv, Crossref, Unpaywall）检索与正版英文多页 PDF 采集。
- **采集通道**：通过本地网络代理环境自动适配直连，通过官方 Open Access 接口与 Legal OA URL 获取出版级完整 PDF 全文。
- **排伪门禁**：自动调用 `auto_exclude_invalid_literature.py` 过滤 0 字节、HTML 403 拦截页及占位文献。
- **知识库挂载**：零 C 盘落盘至 `E:\ozotero\storage\<KEY>\<filename>.pdf`，建立 `itemTypeID=4` (journalArticle) 与 `linkMode=0` 物理关联。
### 3. 中英混合专属智能体 (Hybrid Academic Agent, `hybrid-academic-agent`)
- **定位与适用**：中英双轨真实学术文献联合检索与交叉挖掘，面向国家重大攻关、跨国对比综述及硕博混合文献学位论文。
- **独立性原则**：完全独立于 `cnki-academic-agent` 与 `english-academic-agent`，拥有独立的管线入口、元数据清洗层与融合引用解析器。
- **采集通道**：同时支持知网 CDP 鉴权下载与国际 PubMed/PMC/OpenAlex OA 直连，按中英文分别归一化元数据字段。
- **引用双轨融合**：在 CSL 引用引擎中支持 GB/T 7714 中英双语混编规范（中文文献标注等/卷期，英文文献标注 et al./Vol），自动注入复合域 `ADDIN ZOTERO_ITEM`。
- **知识库挂载**：零 C 盘落盘至 `E:\ozotero\storage\<KEY>\<filename>.pdf`，写入统一的 SQLite 库。

## 13. Word OpenXML 6-Section 分节架构与页眉隔离铁律 (Strict 6-Section & Header Isolation Standard)
- **页眉全局渗透病因**：在清洗模板残留正文段落时，若误删 `<w:body>` 尾部分节符，会导致正文与参考文献全部坍缩并入 Section 3（目录节），造成全篇所有页眉强制显示为“目录”。
- **强制 6 大标准分节拓扑结构**：
  1. **Section 0 (封面/扉页/声明页)**：无页眉，无页脚（`Header=[]`, `Footer=[]`）。
  2. **Section 1 (中文摘要)**：独立页眉“摘要”，罗马数字页码 `I` 起算（`w:fmt="upperRoman" w:start="1"`）。
  3. **Section 2 (英文 Abstract)**：独立页眉“Abstract”，罗马数字页码连续。
  4. **Section 3 (目录 Contents)**：独立页眉“Contents”或“目录”，罗马数字页码连续。
  5. **Section 4 (正文章节 Chapters 1~5)**：必须在正文最后一章末尾段落显式注入分节符（`s4_xml`），绑定 `header5.xml`（`STYLEREF "标题 1" \n` + `STYLEREF "标题 1"` 动态章节页眉），页脚重置为从 **`1`** 开始的阿拉伯数字（`w:fmt="decimal" w:start="1"`）。
  6. **Section 5 (参考文献 References & 致谢)**：必须在 `<w:body>` 根节点末尾显式追加分节符（`s5_xml`），绑定 `header8.xml`（`STYLEREF "Heading 1 Unnumbered"` 独立页眉），阿拉伯页码自然连续。
- **样式严格绑定契约**：
  - 正文各章一级标题必须为 `Heading 1`（标题 1），且必须通过正则剥离显式数字编号（免疫双重标题）。
  - 参考文献与致谢一级标题样式必须显式指定为 `Heading 1 Unnumbered`，绝不允许使用 `Heading 1`，避免污染章节编号与页眉动态域。
- **页眉回退文本净化 (Fallback Text Sanitization)**：
  - `header4.xml`、`header5.xml`、`header8.xml` 内的 `<w:t>` 缓存文本必须与当前文档语种和标题完全一致，杜绝 Word 未刷新域时的中文/错误残留。
