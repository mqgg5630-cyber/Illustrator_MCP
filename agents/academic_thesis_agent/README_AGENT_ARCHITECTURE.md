# Academic-Review-Thesis-Agent (ARTA) 架构设计白皮书与工程规范

## 一、 系统定位与核心价值
**ARTA (Academic-Review-Thesis-Agent)** 是一套面向科研人员、高校师生及学术机构的**端到端全自动科研综述生成、学位论文排版与答辩演示文稿生成智能体**。
系统实现了从原始文献捕获、知识图谱与综述生成、Zotero 动态活体无感入库，到高校学士/硕士毕业论文（以鲁东大学等为标杆）法定全套格式化排版，以及多风格 PPT 演示文稿全自动生成的完整闭环。

---

## 二、 整体分层架构图 (Architecture Diagram)

```mermaid
flowchart TD
    subgraph S1 [1. 文献感知与抽取层 Ingestion]
        A1[知网 CNKI] --> A0[文献元数据解析器]
        A2[SCI / CrossRef / DOI] --> A0
        A3[本地 RIS / BibTeX / PDF] --> A0
    end

    subgraph S2 [2. Zotero 本地活体通信层 Connector]
        A0 --> B1[ZoteroLocalConnector 23119 端口]
        B1 -->|saveItems 批量入库| B2[(Zotero 本地数据库)]
        B2 -->|提取唯一 Item Key| B3[Key-URI 活体映射表]
    end

    subgraph S3 [3. 智能综述合成与大纲撰写层 Synthesis]
        B3 --> C1[学术综述大纲引擎]
        C1 --> C2[段落撰写与核心论点提炼]
        C1 --> C3[三线对比表与构效关系模型]
    end

    subgraph S4 [4. 双轨 Word 活体编译层 Compiler]
        C2 & C3 --> D1[DualTrackWordCompiler]
        D1 -->|1. 渲染层| D2[零 Refresh 预排版呈现]
        D1 -->|2. 协议层| D3[ADDIN ZOTERO_ITEM 复杂域]
        D1 -->|3. 首选项| D4[docProps/custom.xml 切片]
        D2 & D3 & D4 --> D5[学位论文完整底本 .docx]
    end

    subgraph S5 [5. 高校毕业论文标准排版层 Formatter]
        E0[权威基准模板: E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx] --> E1[LarkThesisFormatterAdapter / TemplateCloningEngine]
        D5 --> E1
        E1 -->|学校图徽完整保留 / A4页边距 / 模式一多级编号 / 章节分节 / 页眉页脚 / 三线表 / 机制插图| E2[🎓 最终毕业论文交付文档 .docx]
        E1 -->|同步导出| E3[📚 配套 RIS & BibTeX 文献库]
    end

    subgraph S6 [6. 全量多仓库 PPT 演示文稿生成层 PPT Router]
        C1 --> F1[PPTRouter 统一调度引擎]
        F1 -->|ppt-master| F2[PPT-Master 纯矢量 DrawingML 可编辑 PPTX / 案例模板库]
        F1 -->|cyber-ppt| F3[CyberPPT Nature/McKinsey 象牙白立柱卡片 PPTX]
        F1 -->|swiss-pptx| F4[Swiss-PPTX 瑞士国际主义克莱因蓝 PPTX]
        F1 -->|guizang-ppt| F5[Guizang-PPT 归藏 WebGL 呼吸网格横向翻页 PPT / PPTX]
        F1 -->|banana-slides| F6[Banana-Slides AI Native 香蕉暖金 PPTX]
        F1 -->|dashi-ppt| F7[Dashi-PPT 多主题响应式冷白调研风 HTML / PPTX]
        F2 & F3 & F4 & F5 & F6 & F7 --> F8[📊 学术答辩 / 汇报演示文稿 (字号严格 >= 18pt)]
    end
```

---

## 三、 七大核心模块设计与职责划分

| 模块名称 | 对应文件 / 类 | 核心职责 | 扩展优化点 |
| :--- | :--- | :--- | :--- |
| **1. 数据契约层** | `PaperItem`, `ThesisStudentInfo` | 统一定义学术论文元数据、作者、期刊与学生学籍信息 | 支持增加基金项目号、答辩委员会名单等字段 |
| **2. Zotero 通信层** | `ZoteroLocalConnector` | 与 Zotero 本地 23119 端口通信，静默批量入库并打标，提取真实 Item Key | 可增加 Web API 密钥模式以支持云端建文件夹 |
| **3. 综述语义层** | `AcademicThesisAgent._build_*` | 提炼研究背景、特征工程、算法对比、受体机制与结论展望 | 可接入大模型 Prompt 进行自动化段落深度扩写 |
| **4. 活体编译层** | `DualTrackWordCompiler` | 构建 Zotero CSL 复杂域与 `custom.xml`，实现“零 Refresh 直出 + 活体兼容” | 支持在 Nature、IEEE、GB/T 7714 间任意切换 |
| **5. 毕业论文排版层** | `LarkThesisFormatterAdapter` | 驱动 `Lark-Formatter` 引擎（Python 3.10 / `lark` 环境）执行高校学位论文标准排版 | 可扩展接入清华、浙大、中科院等其他高校模板 |
| **6. PPT 演示文稿引擎** | `PPTRouter` / `PPTMasterEngine` / `CyberPPTEngine` / `SwissPPTXEngine` / `GuizangPPTEngine` / `BananaSlidesEngine` / `DashiPPTEngine` | 全面接入工作区内 6 大 PPT 仓库与案例（含 `ppt-master` 的 `umami-peptide-ml` 等全部模版），支持纯矢量 DrawingML、WebGL 翻页与原生 PPTX 导出 | **严格执行排版阶梯：正文/要点 $\ge 18\text{pt}$，重点 $\ge 20\text{pt}$，大标题 $\ge 28\text{pt}$**，杜绝拥挤重叠 |
| **7. 主控调度层** | `AcademicThesisAgent` | 编排 6 阶段流水线，提供错误自检与统一成果交付 | 可封装为 Web UI 或 FastMCP 微服务 |

---

## 四、 快速使用与二次开发指南

### 1. 命令行快速运行 (CLI)
```bash
# 激活 lark 环境并运行 Agent (一键生成学位论文 + Zotero 库 + Dashi-PPT 汇报)
E:\spider\Library\envs\lark\python.exe e:\0mcp-agv\agents\academic_thesis_agent\arta_agent.py
```

### 2. 代码调用示例 (Python API)
```python
from agents.academic_thesis_agent.arta_agent import AcademicThesisAgent, PaperItem, AuthorInfo, ThesisStudentInfo

# 1. 实例化 Agent
agent = AcademicThesisAgent(workspace_dir="e:/my_thesis_project")

# 2. 定义学生学籍与论文信息
student = ThesisStudentInfo(
    school_name="鲁东大学",
    school_code="10451",
    student_name="文  少",
    degree_field="食品科学与工程",
    degree_type="硕士学位论文"
)

# 3. 传入文献列表、章节与 PPT 配置，一键生成全套学术产物
results = agent.execute_pipeline(
    topic="基于机器学习的食源性鲜味肽高通量筛选与呈味机制解析",
    papers=my_papers,
    student_info=student,
    chinese_abstract="...",
    english_abstract="...",
    chapters=my_chapters,
    generate_ppt=True,
    ppt_engine="dashi-ppt",
    ppt_theme="theme07"  # 自由切换: theme07 (冷白调研风), theme02 (科技紫绿), theme09 (深蓝杂志)
)
print("毕业论文已生成:", results["thesis_docx"])
print("答辩 PPT 已生成:", results["ppt_deck"])
```

---

## 五、 本次 Dashi-PPT 鲜味肽答辩幻灯片页面大纲与亮点

1. **Slide 01 (封面页 - Hero 架构)**: 《基于机器学习的食源性鲜味肽高通量筛选与呈味机制解析》鲁东大学硕士学位论文开题与研究成果答辩
2. **Slide 02 (研究背景与核心痛点)**: 天然减盐需求 (30%+) / 传统筛选瓶颈 (6-12 个月) / AI 虚拟筛选突破 (100x+ 高通量扫描)
3. **Slide 03 (多维特征工程与表征构建)**: 序列离散组成 AAC/DPC (420维) / 伪氨基酸组分 PseAAC / 深度语义嵌入 ProtBERT (1024维)
4. **Slide 04 (多分类器与深度神经网络性能对比)**: iUmami-SCM (86.5% Acc 可解释性) / SVM/RF (89.2% Acc 抗过拟合) / DeepUmami (93.4% Acc 端到端)
5. **Slide 05 (鲜味受体 T1R1/T1R3 互作机制解析)**: GPCR 二聚体靶点 / 4 个关键残基结合位点 (Arg151, Arg277, Ser172, His71) / 结合自由能评估 (-8.5 kcal/mol)
6. **Slide 06 (结论与未来展望)**: 负样本基准库扩充 / 鲜味阈值定量回归预测 / 自动化微流控湿实验闭环
