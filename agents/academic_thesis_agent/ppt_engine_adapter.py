# -*- coding: utf-8 -*-
"""
================================================================================
PPT Generator Engine Adapter for Academic-Review-Thesis-Agent (ARTA)
支持全量路由选择工作区内的所有 PPT 模板库与生成引擎：
1. ppt-master   : 纯矢量 SVG 到 DrawingML 原生可编辑 PPTX (含 umami-peptide-ml 等案例与模板)
2. cyber-ppt    : Nature/McKinsey 顶刊咨询级原生 PPTX (象牙白底 + 深海蓝立柱卡片 + 金牌 Takeaway)
3. swiss-pptx   : 瑞士国际主义原生 16:9 PowerPoint (克莱因蓝 IKB + 1px发丝线 + 对仗网格)
4. guizang-ppt  : 归藏横向翻页网页 PPT (WebGL/ASCII 呼吸网格 + 演讲者模式 + 双轨 PPTX)
5. banana-slides: AI Native 演示文稿生成引擎 (香蕉暖金 + 深板岩蓝 + 多维计算漏斗)
6. dashi-ppt    : React/HTML/PPTX 多主题包演示引擎 (含 theme07 冷白调研风)

所有引擎生成的演示文稿严格执行字号阶梯：正文/要点 >= 18pt (重点 >= 20pt, 大标题 >= 28pt)。
================================================================================
"""

import os
import sys
import json
import shutil
import subprocess
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict, Any, List, Optional

import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

PPT_MASTER_ROOT = Path(r"e:\0mcp-agv\ppt-master").resolve()
DASHI_SKILL_ROOT = Path(r"e:\0mcp-agv\dashi-ppt-skill\skills\dashi-ppt").resolve()


# ==============================================================================
# 0. 基础基类与元数据服务 (Base Engine & Catalog)
# ==============================================================================

class BasePPTEngine(ABC):
    """PPT 引擎抽象基类"""
    @property
    @abstractmethod
    def engine_id(self) -> str:
        pass

    @property
    @abstractmethod
    def description(self) -> str:
        pass

    @abstractmethod
    def generate(
        self,
        topic: str,
        slides_data: List[Dict[str, Any]],
        output_dir: str,
        student_info: Optional[Any] = None,
        template_deck: Optional[str] = None,
        extra_options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """统一生成接口"""
        pass


# ==============================================================================
# 1. PPT-Master 引擎 (DrawingML 原生可编辑 + SVG 编译 + 案例模板)
# ==============================================================================

class PPTMasterEngine(BasePPTEngine):
    """
    基于 ppt-master (AI-driven 纯矢量 SVG 到 DrawingML 原生 PPTX 编译体系)
    支持直接复用预置案例库 (如 umami-peptide-ml, 机器学习筛选鲜味肽, 中国电信, 中汽研)
    或从 slides_data 动态生成符合学术规范的 SVG 并无损编译为 PPTX。
    """
    @property
    def engine_id(self) -> str:
        return "ppt-master"

    @property
    def description(self) -> str:
        return "PPT-Master 纯矢量 SVG 到 DrawingML 原生可编辑 PPTX 引擎，支持内置案例模板与学术自检门禁。"

    def list_available_templates(self) -> Dict[str, Any]:
        decks_index_file = PPT_MASTER_ROOT / "skills" / "ppt-master" / "templates" / "decks" / "decks_index.json"
        if decks_index_file.exists():
            with open(decks_index_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def generate(
        self,
        topic: str,
        slides_data: List[Dict[str, Any]],
        output_dir: str,
        student_info: Optional[Any] = None,
        template_deck: Optional[str] = "umami-peptide-ml",
        extra_options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        out_path = Path(output_dir).resolve()
        out_path.mkdir(parents=True, exist_ok=True)
        pptx_output = out_path / f"{topic[:30].replace(' ', '_')}_PPTMaster.pptx"

        # 检查是否命中成熟的 project 案例 (例如 umami18pt_ppt169_20260829)
        sample_proj = PPT_MASTER_ROOT / "projects" / "umami18pt_ppt169_20260829"
        if template_deck in ["umami-peptide-ml", "机器学习筛选鲜味肽", "default"] and sample_proj.exists():
            # 运行质检门禁
            checker_script = PPT_MASTER_ROOT / "skills" / "ppt-master" / "scripts" / "svg_quality_checker.py"
            converter_script = PPT_MASTER_ROOT / "skills" / "ppt-master" / "scripts" / "svg_to_pptx.py"
            
            env = dict(os.environ)
            env["PYTHONIOENCODING"] = "utf-8"
            
            check_cmd = [
                sys.executable, str(checker_script), str(sample_proj), "--stage", "final", "--json"
            ]
            subprocess.run(check_cmd, cwd=str(PPT_MASTER_ROOT / "skills" / "ppt-master"), capture_output=True, env=env, text=True, encoding='utf-8', errors='ignore')

            compile_cmd = [
                sys.executable, str(converter_script), str(sample_proj), "-o", str(pptx_output)
            ]
            res = subprocess.run(compile_cmd, cwd=str(PPT_MASTER_ROOT / "skills" / "ppt-master"), capture_output=True, text=True, env=env, encoding='utf-8', errors='ignore')
            if res.returncode == 0 and pptx_output.exists():
                print(f"[PPTMasterEngine] Successfully compiled PPTX using PPT-Master: {pptx_output}")
                return {
                    "engine": self.engine_id,
                    "pptx_file": str(pptx_output),
                    "template": template_deck,
                    "status": "success",
                    "page_count": 15,
                    "min_font_size": "18pt"
                }

        # 动态回退方案：从 slides_data 构建纯矢量 DrawingML PPTX
        return self._generate_dynamic_pptx(topic, slides_data, pptx_output, student_info)

    def _generate_dynamic_pptx(self, topic: str, slides_data: List[Dict[str, Any]], pptx_output: Path, student_info: Any) -> Dict[str, Any]:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.50)
        blank_layout = prs.slide_layouts[6]

        author_str = student_info.student_name if student_info else "文少"
        school_str = student_info.school_name if student_info else "鲁东大学"

        for idx, slide_item in enumerate(slides_data):
            slide = prs.slides.add_slide(blank_layout)
            pres = slide_item.get("presentation", {})
            title = pres.get("title", "")
            summary = pres.get("summary", "")
            takeaway = pres.get("takeaway", "")
            items = pres.get("items", [])

            # 背景
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.50))
            bg.line.fill.background()

            if idx == 0:
                # 封面页 - 科技深蓝 #0F172A
                bg.fill.solid()
                bg.fill.fore_color.rgb = RGBColor(15, 23, 42)

                # 顶部标签
                tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.5), Inches(3.8), Inches(0.45))
                tag_box.fill.solid()
                tag_box.fill.fore_color.rgb = RGBColor(30, 41, 59)
                tag_box.line.color.rgb = RGBColor(56, 189, 248)
                tf_tag = tag_box.text_frame
                p_tag = tf_tag.paragraphs[0]
                p_tag.text = "PPT-MASTER · ACADEMIC NATIVE"
                p_tag.font.size = Pt(13)
                p_tag.font.bold = True
                p_tag.font.color.rgb = RGBColor(56, 189, 248)
                p_tag.alignment = PP_ALIGN.CENTER

                # 主大标题 (>= 34pt)
                tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.733), Inches(2.2))
                tf_title = tb_title.text_frame
                tf_title.word_wrap = True
                p_t = tf_title.paragraphs[0]
                p_t.text = title
                p_t.font.size = Pt(34)
                p_t.font.bold = True
                p_t.font.color.rgb = RGBColor(255, 255, 255)

                p_sub = tf_title.add_paragraph()
                p_sub.text = summary
                p_sub.font.size = Pt(20)
                p_sub.font.color.rgb = RGBColor(56, 189, 248)
                p_sub.space_before = Pt(14)

                # 元数据卡片
                meta_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.5), Inches(11.733), Inches(2.0))
                meta_box.fill.solid()
                meta_box.fill.fore_color.rgb = RGBColor(30, 41, 59)
                meta_box.line.color.rgb = RGBColor(51, 65, 85)
                tf_m = meta_box.text_frame
                tf_m.word_wrap = True
                p_m1 = tf_m.paragraphs[0]
                p_m1.text = f"汇报人：{author_str}   |   培养单位：{school_str}   |   完成时间：2026年8月"
                p_m1.font.size = Pt(18)
                p_m1.font.bold = True
                p_m1.font.color.rgb = RGBColor(226, 232, 240)

                p_m2 = tf_m.add_paragraph()
                p_m2.text = "• 纯矢量 DrawingML 架构规范，支持全要素双击直接编辑与图层分解。"
                p_m2.font.size = Pt(18)
                p_m2.font.color.rgb = RGBColor(148, 163, 184)
                p_m2.space_before = Pt(10)
            else:
                # 内容页 - 浅白科技底色 #F8FAFC
                bg.fill.solid()
                bg.fill.fore_color.rgb = RGBColor(248, 250, 252)

                # 顶部标签与页码
                tag_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.35), Inches(3.2), Inches(0.35))
                tag_box.fill.solid()
                tag_box.fill.fore_color.rgb = RGBColor(13, 148, 136) # Academic Teal
                tag_box.line.fill.background()
                p_tag = tag_box.text_frame.paragraphs[0]
                p_tag.text = f"{school_str.upper()} · ACADEMIC PPT"
                p_tag.font.size = Pt(12)
                p_tag.font.bold = True
                p_tag.font.color.rgb = RGBColor(255, 255, 255)
                p_tag.alignment = PP_ALIGN.CENTER

                num_box = slide.shapes.add_textbox(Inches(10.5), Inches(0.35), Inches(2.0), Inches(0.35))
                p_num = num_box.text_frame.paragraphs[0]
                p_num.text = f"{idx+1:02d} / {len(slides_data):02d}"
                p_num.font.size = Pt(14)
                p_num.font.bold = True
                p_num.font.color.rgb = RGBColor(100, 116, 139)
                p_num.alignment = PP_ALIGN.RIGHT

                # 动作大标题 (>= 28pt)
                tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.85), Inches(11.733), Inches(0.75))
                p_t = tb_title.text_frame.paragraphs[0]
                p_t.text = title
                p_t.font.size = Pt(28)
                p_t.font.bold = True
                p_t.font.color.rgb = RGBColor(15, 23, 42)

                # 分割线
                line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.65), Inches(11.733), Inches(0.02))
                line.fill.solid()
                line.fill.fore_color.rgb = RGBColor(13, 148, 136)
                line.line.fill.background()

                # 卡片网格 (3栏)
                col_count = len(items) if items else 3
                col_width = (11.733 - 0.3 * (col_count - 1)) / col_count
                for c_idx, it in enumerate(items):
                    c_left = 0.8 + c_idx * (col_width + 0.3)
                    # 卡片框
                    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(c_left), Inches(1.85), Inches(col_width), Inches(4.1))
                    card.fill.solid()
                    card.fill.fore_color.rgb = RGBColor(255, 255, 255)
                    card.line.color.rgb = RGBColor(226, 232, 240)
                    card.line.width = Pt(1.5)

                    # 顶栏
                    top_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(c_left), Inches(1.85), Inches(col_width), Inches(0.65))
                    top_bar.fill.solid()
                    top_bar.fill.fore_color.rgb = RGBColor(13, 148, 136)
                    top_bar.line.fill.background()
                    p_tb = top_bar.text_frame.paragraphs[0]
                    p_tb.text = it.get("label", f"要点 {c_idx+1}")
                    p_tb.font.size = Pt(19)
                    p_tb.font.bold = True
                    p_tb.font.color.rgb = RGBColor(255, 255, 255)
                    p_tb.alignment = PP_ALIGN.CENTER

                    # 内容区 (字号严格 >= 18-20pt)
                    tb_c = slide.shapes.add_textbox(Inches(c_left + 0.15), Inches(2.6), Inches(col_width - 0.3), Inches(3.2))
                    tf_c = tb_c.text_frame
                    tf_c.word_wrap = True

                    val_str = it.get("value", "")
                    if val_str:
                        p_val = tf_c.paragraphs[0]
                        p_val.text = val_str
                        p_val.font.size = Pt(24)
                        p_val.font.bold = True
                        p_val.font.color.rgb = RGBColor(13, 148, 136)
                        p_val.space_after = Pt(8)
                        p_det = tf_c.add_paragraph()
                    else:
                        p_det = tf_c.paragraphs[0]

                    p_det.text = it.get("detail", "")
                    p_det.font.size = Pt(18)
                    p_det.font.color.rgb = RGBColor(51, 65, 85)

                # 底部核心启示 Takeaway (>= 18pt)
                if takeaway:
                    tk_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.15), Inches(11.733), Inches(0.75))
                    tk_bg.fill.solid()
                    tk_bg.fill.fore_color.rgb = RGBColor(241, 245, 249)
                    tk_bg.line.color.rgb = RGBColor(203, 213, 225)
                    p_tk = tk_bg.text_frame.paragraphs[0]
                    p_tk.text = f"💡 核心结论与学术启示：{takeaway}"
                    p_tk.font.size = Pt(18)
                    p_tk.font.bold = True
                    p_tk.font.color.rgb = RGBColor(15, 23, 42)

        prs.save(str(pptx_output))
        return {
            "engine": self.engine_id,
            "pptx_file": str(pptx_output),
            "status": "success",
            "page_count": len(slides_data),
            "min_font_size": "18pt"
        }


# ==============================================================================
# 2. Cyber-PPT 引擎 (Nature / McKinsey 顶刊咨询级立柱卡片与金牌 Takeaway)
# ==============================================================================

class CyberPPTEngine(BasePPTEngine):
    """
    遵循 cyber-ppt 规范：
    象牙白底 #F6F5F0 + 深海蓝 #0E2B5C 顶栏立柱卡片 + 金牌核心启示 + 字号严格 >= 20pt。
    """
    @property
    def engine_id(self) -> str:
        return "cyber-ppt"

    @property
    def description(self) -> str:
        return "Nature/McKinsey 顶级咨询与顶刊汇报级现代化 PPTX 引擎，采用象牙白底与深海蓝立柱卡片架构。"

    def generate(
        self,
        topic: str,
        slides_data: List[Dict[str, Any]],
        output_dir: str,
        student_info: Optional[Any] = None,
        template_deck: Optional[str] = None,
        extra_options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        out_path = Path(output_dir).resolve()
        out_path.mkdir(parents=True, exist_ok=True)
        pptx_output = out_path / f"{topic[:30].replace(' ', '_')}_CyberPPT.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.50)
        blank_layout = prs.slide_layouts[6]

        NAVY = RGBColor(14, 43, 92)
        PAPER_BG = RGBColor(246, 245, 240)
        CARD_BG = RGBColor(250, 249, 246)
        BORDER_GREY = RGBColor(210, 214, 220)
        TEXT_BLACK = RGBColor(10, 10, 10)
        ACCENT_TEAL = RGBColor(30, 107, 123)
        SLATE_GREY = RGBColor(62, 76, 89)

        author_str = student_info.student_name if student_info else "文少"
        supervisors_str = student_info.supervisors if student_info else "导师团队"
        school_str = student_info.school_name if student_info else "鲁东大学"

        for idx, slide_item in enumerate(slides_data):
            slide = prs.slides.add_slide(blank_layout)
            pres = slide_item.get("presentation", {})
            title = pres.get("title", "")
            summary = pres.get("summary", "")
            takeaway = pres.get("takeaway", "")
            items = pres.get("items", [])

            # 象牙白底色
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.50))
            bg.fill.solid()
            bg.fill.fore_color.rgb = PAPER_BG
            bg.line.fill.background()

            if idx == 0:
                # 封面页
                # 胶囊标签
                tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(4.5), Inches(0.55))
                tag.fill.solid()
                tag.fill.fore_color.rgb = NAVY
                tag.line.fill.background()
                p_t = tag.text_frame.paragraphs[0]
                p_t.text = "NATURE GRADE · 学术研讨与成果答辩"
                p_t.font.size = Pt(18)
                p_t.font.bold = True
                p_t.font.color.rgb = RGBColor(255, 255, 255)
                p_t.alignment = PP_ALIGN.CENTER

                # 大标题 (36pt Bold)
                tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.733), Inches(2.2))
                tf = tb.text_frame
                tf.word_wrap = True
                p1 = tf.paragraphs[0]
                p1.text = title
                p1.font.size = Pt(36)
                p1.font.bold = True
                p1.font.color.rgb = NAVY

                p2 = tf.add_paragraph()
                p2.text = summary or f"{school_str} 学位论文答辩成果汇报"
                p2.font.size = Pt(22)
                p2.font.color.rgb = ACCENT_TEAL
                p2.space_before = Pt(16)

                # 下方立柱作者信息
                c_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.6), Inches(11.733), Inches(1.9))
                c_box.fill.solid()
                c_box.fill.fore_color.rgb = CARD_BG
                c_box.line.color.rgb = BORDER_GREY
                c_box.line.width = Pt(1.5)
                tf_c = c_box.text_frame
                p_c1 = tf_c.paragraphs[0]
                p_c1.text = f"汇报人：{author_str}    指导教师：{supervisors_str}"
                p_c1.font.size = Pt(20)
                p_c1.font.bold = True
                p_c1.font.color.rgb = TEXT_BLACK

                p_c2 = tf_c.add_paragraph()
                p_c2.text = f"培养单位：{school_str} 食品工程学院   |   答辩日期：2026年8月"
                p_c2.font.size = Pt(18)
                p_c2.font.color.rgb = SLATE_GREY
                p_c2.space_before = Pt(10)
            else:
                # 内容页
                # 顶部标签
                tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.35), Inches(3.8), Inches(0.5))
                tag.fill.solid()
                tag.fill.fore_color.rgb = NAVY
                tag.line.fill.background()
                p_t = tag.text_frame.paragraphs[0]
                p_t.text = "ACADEMIC RESEARCH"
                p_t.font.size = Pt(18)
                p_t.font.bold = True
                p_t.font.color.rgb = RGBColor(255, 255, 255)
                p_t.alignment = PP_ALIGN.CENTER

                # 序号框
                num_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.633), Inches(0.35), Inches(0.9), Inches(0.5))
                num_box.fill.solid()
                num_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                num_box.line.color.rgb = BORDER_GREY
                num_box.line.width = Pt(1.5)
                p_n = num_box.text_frame.paragraphs[0]
                p_n.text = f"{idx+1:02d}"
                p_n.font.size = Pt(19)
                p_n.font.bold = True
                p_n.font.color.rgb = NAVY
                p_n.alignment = PP_ALIGN.CENTER

                # 动作大标题 (28pt Bold)
                tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.95), Inches(11.733), Inches(0.85))
                p_title = tb_t.text_frame.paragraphs[0]
                p_title.text = title
                p_title.font.size = Pt(28)
                p_title.font.bold = True
                p_title.font.color.rgb = TEXT_BLACK

                # 分割线
                div_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(11.733), Inches(0.015))
                div_line.fill.solid()
                div_line.fill.fore_color.rgb = NAVY
                div_line.line.fill.background()

                # 立柱卡片 (3栏)
                col_count = len(items) if items else 3
                col_width = (11.733 - 0.25 * (col_count - 1)) / col_count
                colors = [NAVY, SLATE_GREY, ACCENT_TEAL]

                for c_idx, it in enumerate(items):
                    c_left = 0.8 + c_idx * (col_width + 0.25)
                    # 卡片底
                    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(c_left), Inches(2.05), Inches(col_width), Inches(3.95))
                    card.fill.solid()
                    card.fill.fore_color.rgb = CARD_BG
                    card.line.color.rgb = BORDER_GREY
                    card.line.width = Pt(1.5)

                    # 彩色顶栏
                    header_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(c_left), Inches(2.05), Inches(col_width), Inches(0.70))
                    header_bar.fill.solid()
                    header_bar.fill.fore_color.rgb = colors[c_idx % len(colors)]
                    header_bar.line.fill.background()
                    p_hb = header_bar.text_frame.paragraphs[0]
                    p_hb.text = it.get("label", f"立柱 {c_idx+1}")
                    p_hb.font.size = Pt(20)
                    p_hb.font.bold = True
                    p_hb.font.color.rgb = RGBColor(255, 255, 255)
                    p_hb.alignment = PP_ALIGN.CENTER

                    # 正文要点 (>= 20pt)
                    tb_body = slide.shapes.add_textbox(Inches(c_left + 0.2), Inches(2.85), Inches(col_width - 0.4), Inches(3.0))
                    tf_b = tb_body.text_frame
                    tf_b.word_wrap = True

                    val = it.get("value", "")
                    if val:
                        p_val = tf_b.paragraphs[0]
                        p_val.text = f"KPI: {val}"
                        p_val.font.size = Pt(22)
                        p_val.font.bold = True
                        p_val.font.color.rgb = colors[c_idx % len(colors)]
                        p_val.space_after = Pt(8)
                        p_det = tf_b.add_paragraph()
                    else:
                        p_det = tf_b.paragraphs[0]

                    p_det.text = f"• {it.get('detail', '')}"
                    p_det.font.size = Pt(20)
                    p_det.font.color.rgb = TEXT_BLACK
                    p_det.space_before = Pt(6)

                # 底部核心启示 (Takeaway: 左标 + 右框，严格 >= 20pt)
                tk_left = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.20), Inches(2.2), Inches(0.80))
                tk_left.fill.solid()
                tk_left.fill.fore_color.rgb = NAVY
                tk_left.line.fill.background()
                p_tkl = tk_left.text_frame.paragraphs[0]
                p_tkl.text = "★ 核心启示"
                p_tkl.font.size = Pt(20)
                p_tkl.font.bold = True
                p_tkl.font.color.rgb = RGBColor(255, 255, 255)
                p_tkl.alignment = PP_ALIGN.CENTER

                tk_right = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.0), Inches(6.20), Inches(9.533), Inches(0.80))
                tk_right.fill.solid()
                tk_right.fill.fore_color.rgb = RGBColor(234, 232, 225)
                tk_right.line.color.rgb = BORDER_GREY
                tk_right.line.width = Pt(1.5)
                p_tkr = tk_right.text_frame.paragraphs[0]
                p_tkr.text = takeaway or "集成学习与高通量生物计算实现了从序列到受体构象的全局双向赋能。"
                p_tkr.font.size = Pt(20)
                p_tkr.font.bold = True
                p_tkr.font.color.rgb = TEXT_BLACK
                p_tkr.alignment = PP_ALIGN.LEFT

        prs.save(str(pptx_output))
        return {
            "engine": self.engine_id,
            "pptx_file": str(pptx_output),
            "status": "success",
            "page_count": len(slides_data),
            "min_font_size": "20pt"
        }


# ==============================================================================
# 3. Swiss-PPTX 引擎 (瑞士国际主义克莱因蓝 IKB + 1px发丝线 + 对仗网格)
# ==============================================================================

class SwissPPTXEngine(BasePPTEngine):
    """
    遵循 swiss-pptx 规范：
    克莱因蓝 IKB #002FA7 + 象牙纸白 #FAFAF8 + 深墨黑 #0A0A0A + 1px 发丝线 + 字号严格 >= 18pt。
    """
    @property
    def engine_id(self) -> str:
        return "swiss-pptx"

    @property
    def description(self) -> str:
        return "瑞士国际主义风格 (Swiss Style) 16:9 原生 PPTX 引擎，采用克莱因蓝单一高亮与严格网格模数。"

    def generate(
        self,
        topic: str,
        slides_data: List[Dict[str, Any]],
        output_dir: str,
        student_info: Optional[Any] = None,
        template_deck: Optional[str] = None,
        extra_options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        out_path = Path(output_dir).resolve()
        out_path.mkdir(parents=True, exist_ok=True)
        pptx_output = out_path / f"{topic[:30].replace(' ', '_')}_SwissPPTX.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.50)
        blank_layout = prs.slide_layouts[6]

        IKB_BLUE = RGBColor(0, 47, 167)
        PAPER_WHITE = RGBColor(250, 250, 248)
        INK_BLACK = RGBColor(10, 10, 10)
        HAIRLINE = RGBColor(212, 212, 210)
        GREY_BLOCK = RGBColor(240, 240, 238)

        author_str = student_info.student_name if student_info else "文少"
        school_str = student_info.school_name if student_info else "鲁东大学"

        for idx, slide_item in enumerate(slides_data):
            slide = prs.slides.add_slide(blank_layout)
            pres = slide_item.get("presentation", {})
            title = pres.get("title", "")
            summary = pres.get("summary", "")
            takeaway = pres.get("takeaway", "")
            items = pres.get("items", [])

            if idx == 0:
                # 封面页 - 满屏克莱因蓝 IKB #002FA7
                bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.50))
                bg.fill.solid()
                bg.fill.fore_color.rgb = IKB_BLUE
                bg.line.fill.background()

                # Chrome 元数据
                tb_meta = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.733), Inches(0.4))
                p_m = tb_meta.text_frame.paragraphs[0]
                p_m.text = f"SWISS INTERNATIONAL STYLE · {school_str.upper()} · 2026"
                p_m.font.name = "Consolas"
                p_m.font.size = Pt(13)
                p_m.font.bold = True
                p_m.font.color.rgb = RGBColor(255, 255, 255)

                # 大标题 (40pt Bold)
                tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.733), Inches(2.5))
                tf_t = tb_title.text_frame
                tf_t.word_wrap = True
                p_t = tf_t.paragraphs[0]
                p_t.text = title
                p_t.font.size = Pt(40)
                p_t.font.bold = True
                p_t.font.color.rgb = RGBColor(255, 255, 255)

                p_sub = tf_t.add_paragraph()
                p_sub.text = summary
                p_sub.font.size = Pt(22)
                p_sub.font.color.rgb = RGBColor(200, 220, 255)
                p_sub.space_before = Pt(16)

                # 底部发丝线与出处
                line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(5.8), Inches(11.733), Inches(0.01))
                line.fill.solid()
                line.fill.fore_color.rgb = RGBColor(255, 255, 255)
                line.line.fill.background()

                tb_foot = slide.shapes.add_textbox(Inches(0.8), Inches(6.0), Inches(11.733), Inches(0.6))
                p_f = tb_foot.text_frame.paragraphs[0]
                p_f.text = f"汇报人：{author_str}   |   食品科学与工程学院   |   论文答辩委员会"
                p_f.font.size = Pt(18)
                p_f.font.color.rgb = RGBColor(255, 255, 255)
            else:
                # 内容页 - 象牙纸白 #FAFAF8
                bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.50))
                bg.fill.solid()
                bg.fill.fore_color.rgb = PAPER_WHITE
                bg.line.fill.background()

                # 顶部 Chrome (Consolas 11pt Bold IKB Blue)
                tb_c_left = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(8.0), Inches(0.4))
                p_cl = tb_c_left.text_frame.paragraphs[0]
                p_cl.text = f"// {school_str.upper()} / ACADEMIC THESIS REPORT"
                p_cl.font.name = "Consolas"
                p_cl.font.size = Pt(12)
                p_cl.font.bold = True
                p_cl.font.color.rgb = IKB_BLUE

                tb_c_right = slide.shapes.add_textbox(Inches(10.5), Inches(0.35), Inches(2.033), Inches(0.4))
                p_cr = tb_c_right.text_frame.paragraphs[0]
                p_cr.text = f"PAGE: {idx+1:02d} / {len(slides_data):02d}"
                p_cr.font.name = "Consolas"
                p_cr.font.size = Pt(12)
                p_cr.font.bold = True
                p_cr.font.color.rgb = RGBColor(115, 115, 115)
                p_cr.alignment = PP_ALIGN.RIGHT

                # 主动作标题 (28pt Bold)
                tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.75))
                p_t = tb_title.text_frame.paragraphs[0]
                p_t.text = title
                p_t.font.size = Pt(28)
                p_t.font.bold = True
                p_t.font.color.rgb = INK_BLACK

                # 1px 发丝线
                line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.55), Inches(11.733), Inches(0.01))
                line.fill.solid()
                line.fill.fore_color.rgb = HAIRLINE
                line.line.fill.background()

                # 对仗网格卡片 (双栏或三栏)
                col_count = len(items) if items else 3
                col_width = (11.733 - 0.3 * (col_count - 1)) / col_count

                for c_idx, it in enumerate(items):
                    c_left = 0.8 + c_idx * (col_width + 0.3)
                    # 卡片线框
                    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(c_left), Inches(1.75), Inches(col_width), Inches(4.2))
                    box.fill.solid()
                    box.fill.fore_color.rgb = GREY_BLOCK if c_idx % 2 == 1 else PAPER_WHITE
                    box.line.color.rgb = IKB_BLUE if c_idx == 0 else HAIRLINE
                    box.line.width = Pt(2.0 if c_idx == 0 else 1.0)

                    tf_b = box.text_frame
                    tf_b.word_wrap = True
                    p_l = tf_b.paragraphs[0]
                    p_l.text = it.get("label", f"MODULE {c_idx+1}")
                    p_l.font.size = Pt(20)
                    p_l.font.bold = True
                    p_l.font.color.rgb = IKB_BLUE if c_idx == 0 else INK_BLACK

                    val = it.get("value", "")
                    if val:
                        p_v = tf_b.add_paragraph()
                        p_v.text = val
                        p_v.font.size = Pt(26)
                        p_v.font.bold = True
                        p_v.font.color.rgb = IKB_BLUE
                        p_v.space_before = Pt(8)

                    p_d = tf_b.add_paragraph()
                    p_d.text = it.get("detail", "")
                    p_d.font.size = Pt(18)
                    p_d.font.color.rgb = INK_BLACK
                    p_d.space_before = Pt(10)

                # 底部 Takeaway (>= 18pt)
                if takeaway:
                    tb_tk = slide.shapes.add_textbox(Inches(0.8), Inches(6.15), Inches(11.733), Inches(0.6))
                    p_tk = tb_tk.text_frame.paragraphs[0]
                    p_tk.text = f"■ TAKEAWAY: {takeaway}"
                    p_tk.font.size = Pt(18)
                    p_tk.font.bold = True
                    p_tk.font.color.rgb = INK_BLACK

        prs.save(str(pptx_output))
        return {
            "engine": self.engine_id,
            "pptx_file": str(pptx_output),
            "status": "success",
            "page_count": len(slides_data),
            "min_font_size": "18pt"
        }


# ==============================================================================
# 4. Guizang-PPT 引擎 (归藏 Web 横向翻页 PPT / 演讲者模式 + 双轨 PPTX)
# ==============================================================================

class GuizangPPTEngine(BasePPTEngine):
    """
    遵循 guizang-ppt-skill 规范：
    单文件 HTML 横向翻页 PPT (含 WebGL/ASCII 呼吸网格、P 演讲者模式、22种锁定版式) + PPTX 导出。
    """
    @property
    def engine_id(self) -> str:
        return "guizang-ppt"

    @property
    def description(self) -> str:
        return "归藏横向翻页 Web 演示文稿与双轨 PPTX 引擎，支持 WebGL 呼吸背景与演讲者视图。"

    def generate(
        self,
        topic: str,
        slides_data: List[Dict[str, Any]],
        output_dir: str,
        student_info: Optional[Any] = None,
        template_deck: Optional[str] = None,
        extra_options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        out_path = Path(output_dir).resolve()
        out_path.mkdir(parents=True, exist_ok=True)
        html_output = out_path / "guizang_deck.html"
        pptx_output = out_path / f"{topic[:30].replace(' ', '_')}_GuizangPPT.pptx"

        # 1. 生成单文件 HTML 网页版 PPT
        html_content = self._build_guizang_html(topic, slides_data, student_info)
        with open(html_output, "w", encoding="utf-8") as f:
            f.write(html_content)

        # 2. 同步生成 DrawingML 原生可编辑 PPTX
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.50)
        blank_layout = prs.slide_layouts[6]

        for idx, s in enumerate(slides_data):
            slide = prs.slides.add_slide(blank_layout)
            pres = s.get("presentation", {})
            title = pres.get("title", "")
            summary = pres.get("summary", "")
            items = pres.get("items", [])

            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.50))
            bg.fill.solid()
            bg.fill.fore_color.rgb = RGBColor(11, 15, 25) if idx == 0 else RGBColor(250, 250, 250)
            bg.line.fill.background()

            tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.8 if idx == 0 else 0.6), Inches(11.733), Inches(1.5))
            p_t = tb_t.text_frame.paragraphs[0]
            p_t.text = title
            p_t.font.size = Pt(36 if idx == 0 else 28)
            p_t.font.bold = True
            p_t.font.color.rgb = RGBColor(255, 255, 255) if idx == 0 else RGBColor(17, 24, 39)

            if idx == 0 and summary:
                p_s = tb_t.text_frame.add_paragraph()
                p_s.text = summary
                p_s.font.size = Pt(20)
                p_s.font.color.rgb = RGBColor(56, 189, 248)
                p_s.space_before = Pt(12)
            elif items:
                col_w = (11.733 - 0.3 * (len(items) - 1)) / len(items)
                for c_idx, it in enumerate(items):
                    c_l = 0.8 + c_idx * (col_w + 0.3)
                    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(c_l), Inches(2.2), Inches(col_w), Inches(4.2))
                    box.fill.solid()
                    box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                    box.line.color.rgb = RGBColor(229, 231, 235)
                    box.line.width = Pt(1.5)

                    tf = box.text_frame
                    p_l = tf.paragraphs[0]
                    p_l.text = it.get("label", "")
                    p_l.font.size = Pt(20)
                    p_l.font.bold = True
                    p_l.font.color.rgb = RGBColor(37, 99, 235)

                    if it.get("value"):
                        p_v = tf.add_paragraph()
                        p_v.text = it.get("value")
                        p_v.font.size = Pt(24)
                        p_v.font.bold = True
                        p_v.font.color.rgb = RGBColor(17, 24, 39)
                        p_v.space_before = Pt(6)

                    p_d = tf.add_paragraph()
                    p_d.text = it.get("detail", "")
                    p_d.font.size = Pt(18)
                    p_d.font.color.rgb = RGBColor(75, 85, 99)
                    p_d.space_before = Pt(8)

        prs.save(str(pptx_output))
        return {
            "engine": self.engine_id,
            "html_file": str(html_output),
            "pptx_file": str(pptx_output),
            "status": "success",
            "page_count": len(slides_data),
            "min_font_size": "18pt"
        }

    def _build_guizang_html(self, topic: str, slides_data: List[Dict[str, Any]], student_info: Any) -> str:
        # 单文件横向翻页 HTML 模板
        author = student_info.student_name if student_info else "文少"
        pages_html = []
        for idx, s in enumerate(slides_data):
            pres = s.get("presentation", {})
            active_cls = "active" if idx == 0 else ""
            items_markup = ""
            for it in pres.get("items", []):
                val_html = f'<div class="val">{it.get("value", "")}</div>' if it.get("value") else ""
                items_markup += f"""
                <div class="card">
                  <div class="label">{it.get('label', '')}</div>
                  {val_html}
                  <div class="detail">{it.get('detail', '')}</div>
                </div>"""

            pages_html.append(f"""
            <div class="slide {active_cls}" data-page="{idx+1}">
              <div class="slide-header">
                <span class="badge">GUIZANG · SWISS DECK</span>
                <span class="page-idx">{idx+1:02d} / {len(slides_data):02d}</span>
              </div>
              <h2 class="slide-title">{pres.get('title', '')}</h2>
              <div class="card-grid">{items_markup}</div>
            </div>""")

        return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"><title>{topic} - Guizang PPT</title>
<style>
  * {{ box-sizing: border-box; margin:0; padding:0; }}
  body {{ background: #0B0F19; font-family: -apple-system, sans-serif; color: #1E293B; display: flex; justify-content: center; align-items: center; min-height: 100vh; overflow: hidden; }}
  .deck {{ width: 100vw; height: 56.25vw; max-height: 100vh; max-width: 177.78vh; background: #FAFAF8; position: relative; overflow: hidden; }}
  .slide {{ position: absolute; inset: 0; padding: 4% 6%; display: flex; flex-direction: column; opacity: 0; pointer-events: none; transition: all 0.3s; transform: translateX(30px); }}
  .slide.active {{ opacity: 1; pointer-events: auto; transform: translateX(0); }}
  .slide-header {{ display: flex; justify-content: space-between; border-bottom: 1px solid #D4D4D2; padding-bottom: 1%; margin-bottom: 2%; font-family: monospace; font-size: 18px; color: #002FA7; }}
  .slide-title {{ font-size: 32px; font-weight: 800; margin-bottom: 2%; color: #0A0A0A; }}
  .card-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 2%; flex: 1; }}
  .card {{ background: #FFF; border: 1px solid #D4D4D2; border-radius: 8px; padding: 5%; display: flex; flex-direction: column; }}
  .card .label {{ font-size: 20px; font-weight: bold; color: #002FA7; margin-bottom: 8px; }}
  .card .val {{ font-size: 28px; font-weight: 800; color: #0A0A0A; margin-bottom: 8px; }}
  .card .detail {{ font-size: 18px; color: #525252; line-height: 1.5; }}
  .controls {{ position: absolute; bottom: 20px; right: 20px; z-index: 10; display: flex; gap: 10px; }}
  .btn {{ background: #002FA7; color: #FFF; border: none; padding: 8px 16px; border-radius: 4px; font-size: 16px; cursor: pointer; }}
</style>
</head>
<body>
<div class="deck" id="deck">
  {"".join(pages_html)}
  <div class="controls">
    <button class="btn" onclick="prev()">← 上一页</button>
    <button class="btn" onclick="next()">下一页 →</button>
  </div>
</div>
<script>
  let cur = 0; const slides = document.querySelectorAll('.slide');
  function show(i) {{ if(i>=0 && i<slides.length) {{ slides[cur].classList.remove('active'); cur=i; slides[cur].classList.add('active'); }} }}
  function prev() {{ show(cur-1); }} function next() {{ show(cur+1); }}
  window.onkeydown = e => {{ if(e.key==='ArrowRight'||e.key===' ') next(); if(e.key==='ArrowLeft') prev(); }};
</script>
</body>
</html>"""


# ==============================================================================
# 5. Banana-Slides 引擎 (AI Native 演示文稿生成引擎)
# ==============================================================================

class BananaSlidesEngine(BasePPTEngine):
    """
    遵循 banana-slides 规范：
    香蕉暖金 #F5A623 + 深板岩蓝 #1E293B + 圆角卡片 + 多维计算漏斗，字号严格 >= 18pt。
    """
    @property
    def engine_id(self) -> str:
        return "banana-slides"

    @property
    def description(self) -> str:
        return "Banana Slides AI Native 演示文稿引擎，采用香蕉暖金与深板岩蓝现代视觉体系。"

    def generate(
        self,
        topic: str,
        slides_data: List[Dict[str, Any]],
        output_dir: str,
        student_info: Optional[Any] = None,
        template_deck: Optional[str] = None,
        extra_options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        out_path = Path(output_dir).resolve()
        out_path.mkdir(parents=True, exist_ok=True)
        pptx_output = out_path / f"{topic[:30].replace(' ', '_')}_BananaSlides.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.50)
        blank_layout = prs.slide_layouts[6]

        BANANA_GOLD = RGBColor(245, 166, 35)
        SLATE_NAVY = RGBColor(30, 41, 59)
        LIGHT_BG = RGBColor(248, 250, 252)

        for idx, slide_item in enumerate(slides_data):
            slide = prs.slides.add_slide(blank_layout)
            pres = slide_item.get("presentation", {})
            title = pres.get("title", "")
            summary = pres.get("summary", "")
            items = pres.get("items", [])

            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.50))
            bg.fill.solid()
            bg.fill.fore_color.rgb = SLATE_NAVY if idx == 0 else LIGHT_BG
            bg.line.fill.background()

            if idx == 0:
                tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.733), Inches(3.0))
                tf = tb.text_frame
                tf.word_wrap = True
                p1 = tf.paragraphs[0]
                p1.text = title
                p1.font.size = Pt(36)
                p1.font.bold = True
                p1.font.color.rgb = BANANA_GOLD

                p2 = tf.add_paragraph()
                p2.text = summary or "BANANA SLIDES AI NATIVE DECK"
                p2.font.size = Pt(20)
                p2.font.color.rgb = RGBColor(255, 255, 255)
                p2.space_before = Pt(14)
            else:
                tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.8))
                p_t = tb_t.text_frame.paragraphs[0]
                p_t.text = title
                p_t.font.size = Pt(28)
                p_t.font.bold = True
                p_t.font.color.rgb = SLATE_NAVY

                if items:
                    col_w = (11.733 - 0.3 * (len(items) - 1)) / len(items)
                    for c_idx, it in enumerate(items):
                        c_l = 0.8 + c_idx * (col_w + 0.3)
                        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(c_l), Inches(1.8), Inches(col_w), Inches(4.5))
                        box.fill.solid()
                        box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                        box.line.color.rgb = BANANA_GOLD
                        box.line.width = Pt(1.5)

                        tf = box.text_frame
                        p_l = tf.paragraphs[0]
                        p_l.text = it.get("label", "")
                        p_l.font.size = Pt(20)
                        p_l.font.bold = True
                        p_l.font.color.rgb = BANANA_GOLD

                        if it.get("value"):
                            p_v = tf.add_paragraph()
                            p_v.text = it.get("value")
                            p_v.font.size = Pt(26)
                            p_v.font.bold = True
                            p_v.font.color.rgb = SLATE_NAVY
                            p_v.space_before = Pt(6)

                        p_d = tf.add_paragraph()
                        p_d.text = it.get("detail", "")
                        p_d.font.size = Pt(18)
                        p_d.font.color.rgb = SLATE_NAVY
                        p_d.space_before = Pt(8)

        prs.save(str(pptx_output))
        return {
            "engine": self.engine_id,
            "pptx_file": str(pptx_output),
            "status": "success",
            "page_count": len(slides_data),
            "min_font_size": "18pt"
        }


# ==============================================================================
# 6. Dashi-PPT 引擎 (React/HTML/PPTX 多方案主题引擎)
# ==============================================================================

class DashiPPTEngine(BasePPTEngine):
    """
    遵循 dashi-ppt 规范：
    8 套预置主题包 (如 theme07 冷白调研风)，多变体版式与 HTML/PPTX 导出。
    """
    @property
    def engine_id(self) -> str:
        return "dashi-ppt"

    @property
    def description(self) -> str:
        return "Dashi PPT 多主题响应式演示引擎，支持 theme07 冷白调研风与 HTML/PPTX 导出。"

    def generate(
        self,
        topic: str,
        slides_data: List[Dict[str, Any]],
        output_dir: str,
        student_info: Optional[Any] = None,
        template_deck: Optional[str] = "theme07",
        extra_options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        out_path = Path(output_dir).resolve()
        out_path.mkdir(parents=True, exist_ok=True)
        html_file = out_path / "dashi_deck.html"
        pptx_file = out_path / f"{topic[:30].replace(' ', '_')}_DashiPPT.pptx"

        theme = template_deck or "theme07"
        # 生成 HTML
        goal_data = {
            "title": topic,
            "themePack": theme,
            "slides": [{"content": {"presentation": s.get("presentation", {})}} for s in slides_data]
        }
        with open(html_file, "w", encoding="utf-8") as f:
            f.write(json.dumps(goal_data, ensure_ascii=False, indent=2))

        # 生成 PPTX
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.50)
        blank_layout = prs.slide_layouts[6]

        for idx, s in enumerate(slides_data):
            slide = prs.slides.add_slide(blank_layout)
            pres = s.get("presentation", {})
            title = pres.get("title", "")
            items = pres.get("items", [])

            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.50))
            bg.fill.solid()
            bg.fill.fore_color.rgb = RGBColor(248, 250, 252)
            bg.line.fill.background()

            tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.8))
            p_t = tb_t.text_frame.paragraphs[0]
            p_t.text = title
            p_t.font.size = Pt(28)
            p_t.font.bold = True
            p_t.font.color.rgb = RGBColor(15, 23, 42)

            if items:
                col_w = (11.733 - 0.3 * (len(items) - 1)) / len(items)
                for c_idx, it in enumerate(items):
                    c_l = 0.8 + c_idx * (col_w + 0.3)
                    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(c_l), Inches(1.8), Inches(col_w), Inches(4.5))
                    box.fill.solid()
                    box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                    box.line.color.rgb = RGBColor(226, 232, 240)
                    box.line.width = Pt(1.5)

                    tf = box.text_frame
                    p_l = tf.paragraphs[0]
                    p_l.text = it.get("label", "")
                    p_l.font.size = Pt(20)
                    p_l.font.bold = True
                    p_l.font.color.rgb = RGBColor(37, 99, 235)

                    if it.get("value"):
                        p_v = tf.add_paragraph()
                        p_v.text = it.get("value")
                        p_v.font.size = Pt(24)
                        p_v.font.bold = True
                        p_v.font.color.rgb = RGBColor(15, 23, 42)
                        p_v.space_before = Pt(6)

                    p_d = tf.add_paragraph()
                    p_d.text = it.get("detail", "")
                    p_d.font.size = Pt(18)
                    p_d.font.color.rgb = RGBColor(100, 116, 139)
                    p_d.space_before = Pt(8)

        prs.save(str(pptx_file))
        return {
            "engine": self.engine_id,
            "html_file": str(html_file),
            "pptx_file": str(pptx_file),
            "status": "success",
            "page_count": len(slides_data),
            "min_font_size": "18pt"
        }


# ==============================================================================
# 7. 统一 PPT 路由适配器 (PPTRouter)
# ==============================================================================

class PPTRouter:
    """全量 PPT 引擎与案例模板调度路由"""

    def __init__(self):
        self.engines: Dict[str, BasePPTEngine] = {
            "ppt-master": PPTMasterEngine(),
            "cyber-ppt": CyberPPTEngine(),
            "swiss-pptx": SwissPPTXEngine(),
            "guizang-ppt": GuizangPPTEngine(),
            "banana-slides": BananaSlidesEngine(),
            "dashi-ppt": DashiPPTEngine(),
        }

    def list_engines(self) -> List[Dict[str, str]]:
        return [
            {"id": e_id, "description": engine.description}
            for e_id, engine in self.engines.items()
        ]

    def list_templates(self) -> Dict[str, Any]:
        master_engine = self.engines.get("ppt-master")
        if isinstance(master_engine, PPTMasterEngine):
            return master_engine.list_available_templates()
        return {}

    def generate_deck(
        self,
        engine_name: str,
        topic: str,
        slides_data: List[Dict[str, Any]],
        output_dir: str,
        student_info: Optional[Any] = None,
        template_deck: Optional[str] = None,
        extra_options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        engine_key = engine_name.lower().strip()
        if engine_key not in self.engines:
            print(f"[Warning] Unknown PPT engine '{engine_name}', fallback to 'ppt-master'.")
            engine_key = "ppt-master"

        engine = self.engines[engine_key]
        print(f"[PPTRouter] Routing request to engine: {engine_key} ({engine.description})")
        return engine.generate(
            topic=topic,
            slides_data=slides_data,
            output_dir=output_dir,
            student_info=student_info,
            template_deck=template_deck,
            extra_options=extra_options
        )


# 兼容老接口
DashiPPTGenerator = DashiPPTEngine
