# -*- coding: utf-8 -*-
"""
ARTA 命令行 CLI 控制台 (支持单步独立执行与一键全流程)
"""

import sys
import argparse
from arta_engine import ARTAEngine

def main():
    parser = argparse.ArgumentParser(description="Academic-Review-Thesis-Agent (ARTA) 模块化 CLI 执行器")
    parser.add_argument("--step", type=int, choices=[1, 2, 3, 4, 5, 6], help="执行单独某一个步骤 (1~6)")
    parser.add_argument("--all", action="store_true", help="一键执行 Step 1 ~ Step 6 端到端全流程")
    parser.add_argument("--topic", type=str, default="基于深度学习的抗菌肽高通量识别与智能序列设计研究", help="研究课题/论文题目")
    parser.add_argument("--engine", type=str, default="ppt-master", choices=["ppt-master", "cyber-ppt", "swiss-pptx", "dashi-ppt", "banana-slides", "guizang-ppt"], help="PPT 渲染引擎")
    
    args = parser.parse_args()
    engine = ARTAEngine()

    if args.all or len(sys.argv) == 1:
        engine.run_full_pipeline(topic=args.topic, engine=args.engine)
    elif args.step == 1:
        engine.step1_parse_literature()
    elif args.step == 2:
        engine.step2_zotero_import()
    elif args.step == 3:
        engine.step3_synthesize_review(topic=args.topic, mode="deterministic")
    elif args.step == 4:
        engine.step4_build_thesis_base()
    elif args.step == 5:
        engine.step5_format_thesis()
    elif args.step == 6:
        engine.step6_generate_ppt(engine=args.engine)

if __name__ == "__main__":
    main()
