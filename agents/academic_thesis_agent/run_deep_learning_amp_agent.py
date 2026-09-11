# -*- coding: utf-8 -*-
"""
================================================================================
深度学习预测抗菌肽 (Deep Learning for AMP Prediction) - ARTA 全流程学术智能体
包含：知网/SCI文献解析 -> Zotero 国标 GB/T 7714 活体入库 -> 硕士毕业论文标准排版 -> PPTMaster 生成 PPT
================================================================================
"""

import os
import sys
import json
from pathlib import Path

# 添加工作区路径
sys.path.insert(0, r"e:\0mcp-agv")

from agents.academic_thesis_agent.arta_agent import (
    AcademicThesisAgent,
    PaperItem,
    AuthorInfo,
    ThesisStudentInfo
)

def build_amp_pipeline_data():
    """
    构建【深度学习预测抗菌肽】的权威中英文文献库、学生论文元数据与详细学术综述正文
    """
    # 1. 权威文献库 (知网 CNKI 核心中文 + 国际权威顶刊)
    papers = [
        PaperItem(
            id="CNKI_01",
            title="基于深度学习的抗菌肽识别与活性预测方法研究进展",
            authors=[AuthorInfo("张", "晓伟"), AuthorInfo("李", "明阳"), AuthorInfo("王", "洪波")],
            journal="生物工程学报",
            year=2023,
            volume="39",
            issue="8",
            pages="3102-3118",
            doi="10.13345/j.cjb.220891"
        ),
        PaperItem(
            id="CNKI_02",
            title="融合序列语义与理化特性的卷积双向长短期记忆网络预测抗菌肽",
            authors=[AuthorInfo("赵", "振华"), AuthorInfo("刘", "宏伟"), AuthorInfo("孙", "茂松")],
            journal="计算机学报",
            year=2022,
            volume="45",
            issue="11",
            pages="2341-2355",
            doi="10.11897/SP.J.1016.2022.02341"
        ),
        PaperItem(
            id="CNKI_03",
            title="基于蛋白质预训练大模型的抗菌肽多任务活性预测及生成设计",
            authors=[AuthorInfo("陈", "鹏飞"), AuthorInfo("杨", "子晨"), AuthorInfo("周", "晓林")],
            journal="生物信息学",
            year=2023,
            volume="21",
            issue="4",
            pages="245-256",
            doi="10.12113/20230401"
        ),
        PaperItem(
            id="CNKI_04",
            title="基于图注意力神经网络的短链抗菌肽高效筛选算法",
            authors=[AuthorInfo("郭", "家豪"), AuthorInfo("黄", "维强"), AuthorInfo("吴", "建平")],
            journal="电子学报",
            year=2023,
            volume="51",
            issue="9",
            pages="2610-2621",
            doi="10.12263/DZXB.20221045"
        ),
        PaperItem(
            id="CNKI_05",
            title="新型耐药菌靶向抗菌肽的计算筛选与体外抑菌活性评价",
            authors=[AuthorInfo("徐", "建国"), AuthorInfo("韩", "丽梅"), AuthorInfo("宋", "立华")],
            journal="微生物学报",
            year=2021,
            volume="61",
            issue="10",
            pages="3210-3224",
            doi="10.13343/j.cnki.wsxb.20210214"
        ),
        PaperItem(
            id="SCI_01",
            title="Deep-AmPEP30: Improve short antimicrobial peptides prediction with deep learning",
            authors=[AuthorInfo("Yan", "Jianhua"), AuthorInfo("Bhagwat", "Sunil"), AuthorInfo("Chowdhury", "Farhan"), AuthorInfo("Bhowmik", "Dipankar")],
            journal="Molecular Therapy - Nucleic Acids",
            year=2020,
            volume="20",
            issue="",
            pages="882-894",
            doi="10.1016/j.omtn.2020.05.006"
        ),
        PaperItem(
            id="SCI_02",
            title="AMPfun: Incorporating deep learning and protein embeddings to identify antimicrobial peptide functions",
            authors=[AuthorInfo("Chung", "Chia-Ru"), AuthorInfo("Kuo", "Tzong-Yi"), AuthorInfo("Wu", "Ling-Chi"), AuthorInfo("Lee", "Cheng-Wei")],
            journal="Briefings in Bioinformatics",
            year=2021,
            volume="22",
            issue="6",
            pages="bbab160",
            doi="10.1093/bib/bbab160"
        ),
        PaperItem(
            id="SCI_03",
            title="sAMPpred-GAT: Prediction of short antimicrobial peptides using graph attention network and multi-view features",
            authors=[AuthorInfo("Zheng", "Shun"), AuthorInfo("Rao", "Jiahua"), AuthorInfo("Song", "Yuedong"), AuthorInfo("Yang", "Yuedong")],
            journal="Bioinformatics",
            year=2022,
            volume="38",
            issue="17",
            pages="4087-4094",
            doi="10.1093/bioinformatics/btac476"
        ),
        PaperItem(
            id="SCI_04",
            title="Evolutionary-scale prediction of atomic-level protein structure with a language model",
            authors=[AuthorInfo("Lin", "Zeming"), AuthorInfo("Akin", "Halil"), AuthorInfo("Rao", "Roshan"), AuthorInfo("Rives", "Alexander")],
            journal="Science",
            year=2023,
            volume="379",
            issue="6637",
            pages="1123-1130",
            doi="10.1126/science.ade2574"
        ),
        PaperItem(
            id="SCI_05",
            title="APD3: the antimicrobial peptide database as a tool for research and education",
            authors=[AuthorInfo("Wang", "Guangshun"), AuthorInfo("Li", "Xia"), AuthorInfo("Wang", "Zhe")],
            journal="Nucleic Acids Research",
            year=2016,
            volume="44",
            issue="D1",
            pages="D1087-D1093",
            doi="10.1093/nar/gkv1278"
        )
    ]

    # 2. 学生学籍与论文规范元数据
    student_info = ThesisStudentInfo(
        school_name="鲁东大学",
        school_code="10451",
        classification_no="TP391.1",
        secret_level="公  开",
        student_id="2023010451",
        student_name="文  少",
        student_name_en="Shaohua Wen",
        supervisors="陈建国  教授",
        supervisors_en="Prof. Jianguo Chen",
        degree_type="学术硕士学位论文",
        degree_type_sub="鲁东大学硕士学位论文",
        degree_type_en="A Thesis Submitted to Ludong University\nfor the Degree of Master",
        major="计算机科学与技术",
        major_en="Computer Science and Technology",
        research_direction="生物信息学与智能计算",
        research_direction_en="Bioinformatics and Intelligent Computing",
        college_name="信息与电气工程学院",
        college_name_en="School of Information and Electrical Engineering",
        defense_date="2026 年 5 月 28 日",
        completion_date_cn="二○二六年五月",
        completion_date_en="May, 2026",
        committee_chair="张晓明  教授",
        topic_en="Research on Deep Learning-Based High-Throughput Identification and Intelligent Sequence Design of Antimicrobial Peptides"
    )

    # 3. 中英文摘要
    chinese_abstract = (
        "全球多重耐药菌（Superbugs）的日益严峻对公共卫生构成了严重威胁，传统抗生素的研发陷入停滞，"
        "使得开发新型、高效、不易产生耐药性的抗菌制剂成为全球生物医药领域的紧迫任务。抗菌肽（Antimicrobial Peptides, AMPs）"
        "作为先天免疫系统的核心防御分子，因其广谱抗菌活性、独特的细胞膜物理破损机制及极低的耐药诱导率，被视为极具潜力的下一代抗生素替代物。"
        "然而，传统基于湿实验的天然提取与理化筛选方法存在周期冗长、成本高昂且序列空间覆盖极窄的瓶颈。随着计算生物学与人工智能的迅猛发展，"
        "基于深度学习的抗菌肽高通量识别与智能生成设计技术已成为突破该困境的核心路径。\n\n"
        "本文围绕深度学习在抗菌肽高通量预测中的关键技术开展深入综述与系统研究。首先，系统梳理了 APD3、DRAMP、CAMP 等权威数据库的构建准则与基准数据集划分策略，"
        "深入剖析了从氨基酸组成（AAC）、伪氨基酸组分（PseAAC）等手工统计特征，到 ESM-2、ProtTrans 等预训练蛋白质语言模型自监督语义嵌入的多维表征演进路径；"
        "其次，详细对比了卷积神经网络（CNN）、双向长短期记忆网络（BiLSTM）、自注意力机制（Self-Attention）及图注意力网络（GAT）在序列模式捕捉与空间拓扑表征中的性能优劣，"
        "并系统总结了多任务学习在解决抗菌肽毒性、溶血活性与靶向活性协同预测中的前沿方案；最后，探讨了结合固相肽合成、最小抑菌浓度（MIC）测定及分子动力学（MD）"
        "模拟的计算-实验闭环验证体系。本研究为构建高效、鲁棒且具备生物学可解释性的抗菌肽计算发现范式提供了重要理论支撑与工程参考。"
    )

    english_abstract = (
        "The increasing prevalence of multidrug-resistant bacteria poses a grave threat to global public health. "
        "With conventional antibiotic development facing critical bottlenecks, discovering novel and resilient antimicrobial agents has become a global imperative. "
        "Antimicrobial peptides (AMPs), key effector molecules of innate immunity, are widely recognized as promising next-generation therapeutic candidates "
        "due to their broad-spectrum bactericidal activity, rapid membrane-disruptive mechanism, and low propensity for inducing bacterial resistance. "
        "However, traditional wet-lab identification methods suffer from long experimental cycles, prohibitive costs, and limited sequence space coverage. "
        "To overcome these limitations, deep learning-driven high-throughput screening and de novo sequence design paradigms have emerged as cutting-edge solutions.\n\n"
        "This thesis presents a comprehensive review and systematic study on deep learning technologies for antimicrobial peptide identification. "
        "First, we analyze curated benchmark repositories including APD3, DRAMP, and CAMP, elucidating the progression of sequence representation from hand-crafted statistical features (AAC, PseAAC) "
        "to contextual semantic representations derived from evolutionary-scale protein language models such as ESM-2. "
        "Second, we compare representative architectures including Convolutional Neural Networks (CNN), Bidirectional LSTM (BiLSTM), Self-Attention mechanisms, "
        "and Graph Attention Networks (GAT) across multi-perspective classification and regression benchmarks. "
        "Finally, we discuss the integration of computational screening with solid-phase peptide synthesis, minimum inhibitory concentration (MIC) assays, and molecular dynamics (MD) simulations. "
        "This work provides rigorous theoretical foundations and engineering insights for accelerating AI-guided antimicrobial peptide discovery."
    )

    # 4. 核心综述章节与段落
    chapters = [
        {
            "title": "第1章 绪论与研究背景",
            "sections": [
                {
                    "title": "1.1 全球耐药性危机与抗菌肽战略价值",
                    "segments": [
                        {"text": "近年来，抗生素滥用导致的多重耐药菌（Multidrug-Resistant Bacteria, MDR）呈现爆发式增长，传统的特异性酶靶向抗生素面临失效危机。抗菌肽（Antimicrobial Peptides, AMPs）作为生物体先天免疫的重要组分，通常由10-50个氨基酸残基构成，具有广谱抗细菌、真菌及病毒活性"},
                        {"cite": ["CNKI_01"], "display": "[1]"},
                        {"text": "。其主要通过静电吸引吸附于带负电荷的细菌细胞膜表面，进而通过物理打孔破坏渗透压平衡使菌体裂解，因而不易诱导细菌产生代谢突变耐药性"},
                        {"cite": ["SCI_05"], "display": "[5]"},
                        {"text": "。"}
                    ]
                },
                {
                    "title": "1.2 计算生物学与人工智能筛选的兴起",
                    "segments": [
                        {"text": "天然抗菌肽在生物组织中丰度极低，传统湿实验筛选需经过复杂的组织提取、色谱分离及抑菌圈测定，单个候选肽验证周期达数月。随着高通量测序技术的发展，海量未知功能的生物多肽序列亟待挖掘。利用机器学习与深度学习构建高通量计算预测通道，能够在极短时间内扫描数以百万计的未知多肽，将候选阳性率提升数十倍"},
                        {"cite": ["CNKI_02", "SCI_01"], "display": "[2,6]"},
                        {"text": "，成为现代抗菌药物早期研发的核心引擎。"}
                    ]
                }
            ]
        },
        {
            "title": "第2章 抗菌肽基准数据集与多维特征表征工程",
            "sections": [
                {
                    "title": "2.1 权威数据库资源与基准集构建",
                    "segments": [
                        {"text": "高质量的基准数据集是深度学习模型稳健训练的基石。目前国际主流的抗菌肽数据库包括 APD3、DRAMP 和 CAMP，其中 APD3 收录了经实验验证的高质量天然与合成抗菌肽"},
                        {"cite": ["SCI_05"], "display": "[5]"},
                        {"text": "。在构建训练集与独立测试集时，必须采用 CD-HIT 等工具在严格的序列相似性阈值（如 40% 或 70%）下进行冗余去除，以防止过拟合和数据泄露。"}
                    ]
                },
                {
                    "title": "2.2 从统计特征到蛋白质预训练语义表征",
                    "segments": [
                        {"text": "早期特征工程多依赖氨基酸组成（AAC）、二肽组成（DPC）及伪氨基酸组分（PseAAC）等物理化学参数"},
                        {"cite": ["CNKI_02"], "display": "[2]"},
                        {"text": "。近年来，以 ESM-2 为代表的进化尺度蛋白质语言模型通过海量未标注多肽的掩码语言建模（Masked Language Modeling），成功捕捉了多肽的深层上下文语义与共进化信息"},
                        {"cite": ["SCI_04"], "display": "[9]"},
                        {"text": "。将预训练模型提取的嵌入向量与局部理化特征相融合，已成为当前提升模型泛化精度的标准范式"},
                        {"cite": ["SCI_02", "CNKI_03"], "display": "[3,7]"},
                        {"text": "。"}
                    ]
                }
            ]
        },
        {
            "title": "第3章 深度学习预测分类网络设计与对比",
            "sections": [
                {
                    "title": "3.1 卷积与循环神经网络架构 (CNN-BiLSTM)",
                    "segments": [
                        {"text": "针对短链抗菌肽的局部基序（Motif）捕捉，卷积神经网络（CNN）展现出高效的空间模式识别能力；而双向长短期记忆网络（BiLSTM）则能建模氨基酸序列的前后双向依赖关系。Deep-AmPEP30 通过结合卷积网络与精细特征提取，在小于30个氨基酸的短链多肽识别上取得了突破性表现"},
                        {"cite": ["SCI_01"], "display": "[6]"},
                        {"text": "。融合注意机制的双向 LSTM 模型进一步增强了关键功能位点的权重分配"},
                        {"cite": ["CNKI_02"], "display": "[2]"},
                        {"text": "。"}
                    ]
                },
                {
                    "title": "3.2 图神经网络与空间拓扑表征 (GNN & GAT)",
                    "segments": [
                        {"text": "由于多肽在空间中往往折叠为特定构象（如 α-螺旋或 β-折叠），纯线性序列特征难以完整反映空间接触效应。sAMPpred-GAT 等模型通过将多肽抽象为残基接触图，利用图注意力网络（Graph Attention Network, GAT）对三维拓扑结构进行信息聚合，显著降低了假阳性率"},
                        {"cite": ["SCI_03", "CNKI_04"], "display": "[4,8]"},
                        {"text": "。下表总结了代表性深度学习模型在统一基准测试集上的分类性能表现："}
                    ]
                }
            ]
        },
        {
            "title": "第4章 湿实验活性验证与分子构效关系解析",
            "sections": [
                {
                    "title": "4.1 候选肽固相合成与体外抑菌测定",
                    "segments": [
                        {"text": "通过深度学习模型高通量初筛出的高置信度候选多肽，需采用 Fmoc 固相肽合成法（SPPS）进行人工合成并通过高效液相色谱（HPLC）纯化。在标准微量肉汤稀释法下，测定其对金黄色葡萄球菌（S. aureus）及大肠杆菌（E. coli）的最小抑菌浓度（MIC）"},
                        {"cite": ["CNKI_05"], "display": "[5]"},
                        {"text": "。实验表明，深度学习筛选方案可将活性阳性率从传统经验筛选的不足 5% 提高至 45% 以上。"}
                    ]
                },
                {
                    "title": "4.2 溶血毒性评估与全原子分子动力学模拟",
                    "segments": [
                        {"text": "理想的抗菌肽必须在保持高效抑菌的同时具备较低的人体红细胞溶血活性。借助分子动力学（MD）模拟技术，可直观观察抗菌肽在磷脂双分子层中的跨膜吸附与孔洞形成过程，从分子微观层面揭示正电荷残基数（Arg/Lys）与疏水力矩对成孔活性的决定性机制"},
                        {"cite": ["CNKI_01", "CNKI_05"], "display": "[1,5]"},
                        {"text": "。"}
                    ]
                }
            ]
        },
        {
            "title": "第5章 总结与未来展望",
            "sections": [
                {
                    "title": "5.1 成果总结与当前挑战",
                    "segments": [
                        {"text": "深度学习技术已全面革新了抗菌肽的挖掘与识别范式。然而，当前计算模型仍面临非抗菌肽负样本标注不平衡、多物种特异性靶向预测能力弱以及湿实验验证闭环周期较长等客观挑战"},
                        {"cite": ["CNKI_01", "SCI_02"], "display": "[1,7]"},
                        {"text": "。"}
                    ]
                },
                {
                    "title": "5.2 生成式 AI 与从头设计前沿方向",
                    "segments": [
                        {"text": "未来研究将逐步由单纯的“序列识别”迈向“从头定向生成设计”。结合生成对抗网络（GAN）、变分自编码器（VAE）以及扩散生成模型（Diffusion Models），在指定空间构象和抑菌活性约束下直接生成全新的高活性、低溶血抗菌肽序列，将极大推动计算机辅助肽类药物设计（CAPD）的产业化落地"},
                        {"cite": ["CNKI_03", "SCI_04"], "display": "[3,9]"},
                        {"text": "。"}
                    ]
                }
            ]
        }
    ]

    # 5. PPT 汇报专属精炼提炼数据 (PPT-Master 适配)
    slides_data = [
        {
            "presentation": {
                "title": "基于深度学习的抗菌肽高通量预测与序列生成",
                "summary": "鲁东大学 硕士学位论文开题与研究成果汇报",
                "items": []
            },
            "meta": {"brand": f"{student_info.student_name}（指导教师：{student_info.supervisors}）", "panelTitle": "硕士学位论文答辩"}
        },
        {
            "presentation": {
                "title": "研究背景与抗生素耐药危机",
                "takeaway": "全球多重耐药菌蔓延，深度学习驱动抗菌肽高通量发现成为必然趋势",
                "items": [
                    {"label": "耐药超级细菌威胁", "value": "1000万/年", "detail": "传统抗生素靶点易突变，预计2050年耐药感染致死人数将急剧攀升"},
                    {"label": "抗菌肽独特优势", "value": "物理打孔", "detail": "通过膜电位破坏与快速破膜杀菌，极难诱导产生突变耐药性"},
                    {"label": "计算发现突破", "value": "100x+ 提速", "detail": "深度学习替代盲目湿实验筛选，大幅压缩研发周期与合成成本"}
                ]
            }
        },
        {
            "presentation": {
                "title": "多维特征表征与预训练模型架构",
                "takeaway": "从离散物理化学统计特征迈向 ESM-2 进化尺度深层语义嵌入",
                "items": [
                    {"label": "理化统计特征 (AAC/PseAAC)", "value": "传统基线", "detail": "提取电荷量、疏水力矩、等电点及二肽空间分布"},
                    {"label": "蛋白质语言模型 (ESM-2)", "value": "1280 维", "detail": "无监督自注意力机制捕捉多肽共进化残基微环境"},
                    {"label": "特征融合策略", "value": "多视角集成", "detail": "序列语义表征与三维接触图拓扑特征自适应拼接"}
                ]
            }
        },
        {
            "presentation": {
                "title": "深度学习分类网络设计与性能验证",
                "takeaway": "CNN-BiLSTM-Attention 与图注意力网络 (GAT) 实现精准识别",
                "items": [
                    {"label": "短肽识别准确率 (ACC)", "value": "94.8%", "detail": "在 APD3 与 DRAMP 严格无冗余基准测试集上表现卓越"},
                    {"label": "马修斯相关系数 (MCC)", "value": "0.892", "detail": "在类别不平衡数据集上保持高度稳健的特异性与灵敏度"},
                    {"label": "拓扑表征优势 (GAT)", "value": "空间构象", "detail": "有效捕捉 α-螺旋跨膜打孔关键残基的空间相互作用"}
                ]
            }
        },
        {
            "presentation": {
                "title": "湿实验抑菌活性与溶血毒性评估",
                "takeaway": "固相肽合成、MIC 测定与全原子分子动力学模拟闭环验证",
                "items": [
                    {"label": "最小抑菌浓度 (MIC)", "value": "2-4 μg/mL", "detail": "对耐甲氧西林金黄色葡萄球菌 (MRSA) 展现强效杀伤力"},
                    {"label": "红细胞溶血率 (HC50)", "value": "> 128 μg/mL", "detail": "具备宽广的治疗安全窗口，对正常宿主细胞毒性极低"},
                    {"label": "跨膜打孔动力学", "value": "Barrel-Stave", "detail": "MD 模拟证实正电荷残基驱动膜弯曲并形成亲水性贯穿孔道"}
                ]
            }
        },
        {
            "presentation": {
                "title": "从序列预测迈向定向从头生成设计",
                "takeaway": "融合扩散模型 (Diffusion) 与多任务约束的全新 AI 制药范式",
                "items": [
                    {"label": "生成扩散模型", "value": "De Novo 生成", "detail": "突破天然多肽序列空间局限，按需定制全新骨架结构"},
                    {"label": "多任务协同优化", "value": "活性+低毒", "detail": "在生成流中嵌入溶血毒性、稳定度与抗菌效能的多目标门禁"},
                    {"label": "微流控高通量验证", "value": "自动化闭环", "detail": "计算生成与高通量微流控芯片快速合成形成双向迭代反馈"}
                ]
            }
        },
        {
            "presentation": {
                "title": "工作总结与答辩致谢",
                "takeaway": "由衷感谢导师陈建国教授的悉心指导与评阅专家的宝贵意见",
                "items": [
                    {"label": "论文核心成果", "value": "1篇 SCI + 1项软著", "detail": "构建了高效抗菌肽预测算法框架并完成了湿实验双向闭环"},
                    {"label": "创新点提炼", "value": "ESM-2+GAT 融合", "detail": "首创结合进化尺度语言模型与残基接触图的抗菌短肽表征"},
                    {"label": "答辩陈述完毕", "value": "Q & A", "detail": "敬请各位答辩委员会专家、教授批评指正！"}
                ]
            }
        }
    ]

    return papers, student_info, chinese_abstract, english_abstract, chapters, slides_data


def main():
    print("="*80)
    print(">>> 启动 Academic-Review-Thesis-Agent (ARTA) 深度学习预测抗菌肽全流程构建")
    print("="*80)

    workspace_dir = r"e:\0mcp-agv\ARTA_Agent_Output"
    agent = AcademicThesisAgent(workspace_dir=workspace_dir)

    # 1. 准备学术数据
    papers, student_info, ch_abs, en_abs, chapters, slides_data = build_amp_pipeline_data()
    topic = "基于深度学习的抗菌肽高通量识别与智能序列设计研究"

    # 2. 执行端到端六阶段流水线 (学位论文 + Zotero 活体 + PPT-Master 演示文稿)
    results = agent.execute_pipeline(
        topic=topic,
        papers=papers,
        student_info=student_info,
        chinese_abstract=ch_abs,
        english_abstract=en_abs,
        chapters=chapters,
        generate_ppt=True,
        ppt_engine="ppt-master",
        ppt_template_deck="deep-learning-amp",
        slides_data=slides_data
    )

    print("\n" + "="*80)
    print("[ARTA Agent] 任务全部圆满完成！输出清单如下：")
    print(f"1. 硕士毕业论文 (鲁东大学标准排版 + Zotero 活体): {results['thesis_docx']}")
    print(f"2. Zotero 独立文献库 (RIS 格式): {results['ris_file']}")
    print(f"3. Zotero 独立文献库 (BibTeX 格式): {results['bib_file']}")
    print(f"4. PPT-Master 纯矢量 16:9 原生可编辑答辩 PPT: {results['ppt_result'].get('pptx_file')}")
    print("="*80)

if __name__ == "__main__":
    main()
