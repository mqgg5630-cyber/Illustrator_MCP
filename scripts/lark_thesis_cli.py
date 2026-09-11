# -*- coding: utf-8 -*-
"""
Lark Thesis Formatter CLI - 毕业论文与学术文档全自动排版命令行工具
基准模板：E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx
基于 Lark-Formatter 核心排版引擎 (Python 3.10+ / lark 环境)
"""
import os
import sys
import argparse
from pathlib import Path

# 添加 Lark-Formatter 源码路径
LARK_ROOT = r"E:\0writing\Lark-Formatter"
if LARK_ROOT not in sys.path:
    sys.path.insert(0, LARK_ROOT)

DEFAULT_STANDARD_TEMPLATE = r"E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx"

try:
    from src.scene.manager import load_default_scene, load_preset, list_presets, load_scene
    from src.engine.pipeline import Pipeline
except ImportError as e:
    print(f"[Error] Failed to import Lark-Formatter modules from {LARK_ROOT}: {e}")
    sys.exit(1)

def format_document(input_path: str, output_path: str = None, preset_name: str = None, scene_path: str = None, template_path: str = None):
    """
    执行毕业论文 / 学位论文全自动规范排版
    """
    in_p = Path(input_path).resolve()
    if not in_p.exists():
        print(f"[Error] Input file not found: {in_p}")
        return False
    
    if output_path is None:
        out_p = in_p.parent / f"{in_p.stem}_thesis_formatted{in_p.suffix}"
    else:
        out_p = Path(output_path).resolve()
    
    out_p.parent.mkdir(parents=True, exist_ok=True)
    
    # 加载配置
    if scene_path and os.path.exists(scene_path):
        config = load_scene(scene_path)
        print(f"[Info] Loaded custom scene config: {scene_path}")
    elif preset_name and preset_name not in ("thesis", "ludong_master", "default"):
        try:
            config = load_preset(preset_name)
            print(f"[Info] Loaded preset: {preset_name}")
        except Exception:
            print(f"[Warning] Preset '{preset_name}' not found, falling back to default thesis scene.")
            config = load_default_scene()
    else:
        config = load_default_scene()
        ref_tmpl = template_path or DEFAULT_STANDARD_TEMPLATE
        print(f"[Info] Loaded standard thesis formatting scene (鲁东大学学术硕士学位论文标准规范)")
        print(f"[Info] Authoritative Template: {ref_tmpl}")

    print(f"[Running] Formatting: {in_p.name} -> {out_p.name} ...")
    pipeline = Pipeline(config=config)
    res = pipeline.run(doc_path=str(in_p))
    
    if res.success and res.doc:
        res.doc.save(str(out_p))
        print(f"[Success] Formatted thesis document saved to: {out_p}")
        return True
    elif res.doc:
        # 部分阶段成功也保存文档
        res.doc.save(str(out_p))
        print(f"[Warning] Formatting pipeline status [{res.status}], saved output to: {out_p}")
        return True
    else:
        print(f"[Failed] Formatting ended with status: {res.status}")
        return False

def main():
    parser = argparse.ArgumentParser(
        description="Lark-Formatter CLI: 自动化高校毕业论文与学术文档排版工具 (鲁东大学学术硕士标准定稿版)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
示例:
  lark-thesis.bat -i "input.docx" -o "formatted_thesis.docx"
  python lark_thesis_cli.py -i "{DEFAULT_STANDARD_TEMPLATE}" -o "鲁东大学学术硕士学位论文_定稿排版.docx"
        """
    )
    parser.add_argument("-i", "--input", required=True, help="输入的 Word 文档路径 (.docx)")
    parser.add_argument("-o", "--output", default=None, help="排版后输出的 Word 文档路径 (.docx)")
    parser.add_argument("-p", "--preset", default=None, help="预设模板名称 (如 thesis, ludong_master, standard)")
    parser.add_argument("-t", "--template", default=DEFAULT_STANDARD_TEMPLATE, help="基准模板文件路径")
    parser.add_argument("-c", "--config", default=None, help="自定义 JSON 场景配置文件路径")
    parser.add_argument("--list-presets", action="store_true", help="列出所有可用的预设排版方案")

    args = parser.parse_args()

    if args.list_presets:
        presets = list_presets()
        print("Available presets:", presets)
        return

    success = format_document(
        input_path=args.input,
        output_path=args.output,
        preset_name=args.preset,
        scene_path=args.config,
        template_path=args.template
    )
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
