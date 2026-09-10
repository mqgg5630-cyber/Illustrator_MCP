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
- **Word OpenXML 活体双轨编译规范（彻底根除“需要插入引注”弹窗）**：
  - 正文引用：必须使用 OpenXML 复合域（`w:fldChar` 的 `begin`、`ADDIN ZOTERO_ITEM CSL_CITATION` 指令文本、`separate`、高亮预排版上标、`end`）完整包裹。严禁直接写入死文本。
  - 参考文献：必须使用 `ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY` 将整个参考文献列表包裹于复合域内。
  - 属性注入：必须在 `docProps/custom.xml` 中以每段 $\le 255$ 字符切片写入 `ZOTERO_PREF_1 ... ZOTERO_PREF_n`，并在 `[Content_Types].xml` 与 `_rels/.rels` 完成注册，实现 Word/WPS 打开即认、点击 Zotero Refresh 即时联动。
- **唯一基准模板与成功案例标杆**：
  - 模板唯一权威基准：`E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx`。
  - 终极成功标杆案例：`E:\0mcp-agv\ARTA_Agent_Output\基于机器学习的食源性鲜味肽高通量筛选与呈味机制解析_顶刊综述学位论文.docx`。

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

