# -*- coding: utf-8 -*-
"""
Academic-Review-Thesis-Agent (ARTA) 现代化图形界面控制台 (GUI)
支持：
1. 一键端到端全自动运行 (One-Click Pipeline)
2. 6 大独立步骤单项执行 (Step-by-Step Modular Execution)
3. 100% 确定性去 AI 化算法与模板克隆体系
4. 实时日志控制台与物料一键打开
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

sys.path.insert(0, r"e:\0mcp-agv")
from agents.academic_thesis_agent.arta_engine import ARTAEngine, ThesisStudentInfo

class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, s):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, s)
        self.widget.see(tk.END)
        self.widget.configure(state='disabled')

    def flush(self):
        pass


class ARTAGUIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ARTA 学术综述·学位论文·答辩PPT 全自动工程工作台 (v2.5 去AI高可靠版)")
        self.root.geometry("1080x820")
        self.root.minsize(960, 720)
        self.root.configure(bg="#F4F5F7")

        self.engine = ARTAEngine()
        self._setup_styles()
        self._build_ui()

        # 重定向标准输出到 GUI 控制台
        sys.stdout = TextRedirector(self.log_text)

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        # 基础配色
        style.configure(".", background="#F4F5F7", foreground="#1E293B", font=("Microsoft YaHei UI", 9))
        style.configure("Header.TLabel", font=("Microsoft YaHei UI", 16, "bold"), foreground="#002FA7", background="#F4F5F7")
        style.configure("SubHeader.TLabel", font=("Microsoft YaHei UI", 10), foreground="#64748B", background="#F4F5F7")
        style.configure("Card.TLabelframe", background="#FFFFFF", relief="flat")
        style.configure("Card.TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"), foreground="#0F172A", background="#FFFFFF")
        
        # 按钮样式
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 11, "bold"), foreground="#FFFFFF", background="#002FA7", padding=8)
        style.map("Primary.TButton", background=[("active", "#002070")])

        style.configure("Step.TButton", font=("Microsoft YaHei UI", 9, "bold"), foreground="#0F172A", background="#E2E8F0", padding=5)
        style.map("Step.TButton", background=[("active", "#CBD5E1")])

        style.configure("Open.TButton", font=("Microsoft YaHei UI", 9), foreground="#002FA7", background="#EFF6FF", padding=4)
        style.map("Open.TButton", background=[("active", "#DBEAFE")])

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding="15 15 15 15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部 Header 区域
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(header_frame, text="🎓 ARTA 学位论文与答辩演示文稿生成引擎", style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(header_frame, text="支持权威模板原生克隆（校徽100%保留）· 正文内嵌三线表/机制插图 · Zotero 活体切片协议 · 100% 确定性去AI化", style="SubHeader.TLabel").pack(anchor=tk.W, pady=(2, 0))

        # 2. 中上部：参数配置卡片
        config_card = ttk.LabelFrame(main_frame, text=" 课题与排版元数据配置 ", style="Card.TLabelframe", padding="12 10 12 10")
        config_card.pack(fill=tk.X, pady=(0, 10))

        # 第一行：论文题目与 PPT 引擎
        r1 = ttk.Frame(config_card, style="Card.TLabelframe")
        r1.pack(fill=tk.X, pady=3)
        ttk.Label(r1, text="论文主标题：", font=("Microsoft YaHei UI", 9, "bold"), background="#FFFFFF").pack(side=tk.LEFT)
        self.topic_var = tk.StringVar(value="基于深度学习的抗菌肽高通量识别与智能序列设计研究")
        ttk.Entry(r1, textvariable=self.topic_var, font=("Microsoft YaHei UI", 9), width=58).pack(side=tk.LEFT, padx=(5, 20), fill=tk.X, expand=True)

        ttk.Label(r1, text="PPT 渲染引擎：", font=("Microsoft YaHei UI", 9, "bold"), background="#FFFFFF").pack(side=tk.LEFT)
        self.engine_var = tk.StringVar(value="ppt-master")
        engine_combo = ttk.Combobox(r1, textvariable=self.engine_var, values=["ppt-master", "cyber-ppt", "swiss-pptx", "dashi-ppt", "banana-slides"], state="readonly", width=14)
        engine_combo.pack(side=tk.LEFT, padx=(5, 0))

        # 第二行：学生信息配置
        r2 = ttk.Frame(config_card, style="Card.TLabelframe")
        r2.pack(fill=tk.X, pady=4)
        
        ttk.Label(r2, text="作者姓名：", background="#FFFFFF").pack(side=tk.LEFT)
        self.author_var = tk.StringVar(value="文  少")
        ttk.Entry(r2, textvariable=self.author_var, width=10).pack(side=tk.LEFT, padx=(2, 15))

        ttk.Label(r2, text="指导教师：", background="#FFFFFF").pack(side=tk.LEFT)
        self.advisor_var = tk.StringVar(value="刘新建  教授")
        ttk.Entry(r2, textvariable=self.advisor_var, width=14).pack(side=tk.LEFT, padx=(2, 15))

        ttk.Label(r2, text="学科专业：", background="#FFFFFF").pack(side=tk.LEFT)
        self.major_var = tk.StringVar(value="计算机科学与技术")
        ttk.Entry(r2, textvariable=self.major_var, width=16).pack(side=tk.LEFT, padx=(2, 15))

        ttk.Label(r2, text="答辩日期：", background="#FFFFFF").pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value="2026 年 5 月 28 日")
        ttk.Entry(r2, textvariable=self.date_var, width=16).pack(side=tk.LEFT, padx=(2, 0))

        # 3. 中部：操作卡片 (一键运行 + 6 大单步执行项)
        actions_card = ttk.LabelFrame(main_frame, text=" 执行控制中心 (One-Click Pipeline & Modular Execution) ", style="Card.TLabelframe", padding="12 10 12 10")
        actions_card.pack(fill=tk.X, pady=(0, 10))

        # 一键总执行按钮区
        btn_all_frame = ttk.Frame(actions_card, style="Card.TLabelframe")
        btn_all_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_run_all = ttk.Button(btn_all_frame, text="🚀 一键全自动执行端到端全流程 (Stage 1 ~ Stage 6)", style="Primary.TButton", command=self._on_run_all)
        self.btn_run_all.pack(fill=tk.X, ipady=4)

        # 6 大独立步骤按钮网格
        steps_grid = ttk.Frame(actions_card, style="Card.TLabelframe")
        steps_grid.pack(fill=tk.X)

        self.btn_s1 = ttk.Button(steps_grid, text="Step 1: 文献解析与清洗", style="Step.TButton", command=lambda: self._run_async(self.engine.step1_parse_literature))
        self.btn_s1.grid(row=0, column=0, padx=4, pady=3, sticky="ew")

        self.btn_s2 = ttk.Button(steps_grid, text="Step 2: Zotero 活体入库绑定", style="Step.TButton", command=lambda: self._run_async(self.engine.step2_zotero_import))
        self.btn_s2.grid(row=0, column=1, padx=4, pady=3, sticky="ew")

        self.btn_s3 = ttk.Button(steps_grid, text="Step 3: 学术综述章节合成", style="Step.TButton", command=lambda: self._run_async(lambda: self.engine.step3_synthesize_review(self.topic_var.get(), "deterministic")))
        self.btn_s3.grid(row=0, column=2, padx=4, pady=3, sticky="ew")

        self.btn_s4 = ttk.Button(steps_grid, text="Step 4: 论文底本生成(含三线表)", style="Step.TButton", command=lambda: self._run_async(lambda: self.engine.step4_build_thesis_base(self._get_current_student_info())))
        self.btn_s4.grid(row=1, column=0, padx=4, pady=3, sticky="ew")

        self.btn_s5 = ttk.Button(steps_grid, text="Step 5: 高校论文标准排版", style="Step.TButton", command=lambda: self._run_async(self.engine.step5_format_thesis))
        self.btn_s5.grid(row=1, column=1, padx=4, pady=3, sticky="ew")

        self.btn_s6 = ttk.Button(steps_grid, text="Step 6: 答辩 PPT 编译生成", style="Step.TButton", command=lambda: self._run_async(lambda: self.engine.step6_generate_ppt(self.engine_var.get())))
        self.btn_s6.grid(row=1, column=2, padx=4, pady=3, sticky="ew")

        for col in range(3):
            steps_grid.columnconfigure(col, weight=1)

        # 4. 下部：控制台日志卡片与文件快捷打开
        log_card = ttk.LabelFrame(main_frame, text=" 实时执行日志与产物导出 ", style="Card.TLabelframe", padding="10 8 10 8")
        log_card.pack(fill=tk.BOTH, expand=True)

        # 快捷打开按钮栏
        open_bar = ttk.Frame(log_card, style="Card.TLabelframe")
        open_bar.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(open_bar, text="📄 打开论文底本 (DOCX)", style="Open.TButton", command=self._open_base_docx).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(open_bar, text="🎓 打开最终排版论文 (DOCX)", style="Open.TButton", command=self._open_final_docx).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(open_bar, text="📊 打开答辩幻灯片 (PPTX)", style="Open.TButton", command=self._open_pptx).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(open_bar, text="📁 打开输出物料目录", style="Open.TButton", command=self._open_output_dir).pack(side=tk.RIGHT)

        # ScrolledText 实时日志
        self.log_text = scrolledtext.ScrolledText(log_card, wrap=tk.WORD, font=("Consolas", 9), bg="#0F172A", fg="#38BDF8", insertbackground="#38BDF8", relief="flat")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.insert(tk.END, "=== ARTA 学术综述与学位论文生成工作台就绪 ===\n点击上方【🚀 一键全自动执行】或按需点击单个步骤按钮开始。\n\n")
        self.log_text.configure(state='disabled')

    def _get_current_student_info(self) -> ThesisStudentInfo:
        return ThesisStudentInfo(
            category_no="TP18",
            school_code="10446",
            security_level="公开",
            student_id="2023020888",
            author_name=self.author_var.get(),
            author_en="Shao Wen",
            advisor=self.advisor_var.get(),
            advisor_en="Prof. Xinjian Liu",
            college="生命科学学院 / 计算机学院",
            major=self.major_var.get(),
            major_en="Computer Science and Technology",
            research_direction="生物信息学与智能计算",
            defense_date=self.date_var.get(),
            defense_chair="李  军  教授"
        )

    def _on_run_all(self):
        topic = self.topic_var.get().strip()
        engine_name = self.engine_var.get()
        student_info = self._get_current_student_info()

        def _worker():
            self._set_buttons_state(False)
            try:
                self.engine.student_info = student_info
                self.engine.run_full_pipeline(topic=topic, engine=engine_name)
                messagebox.showinfo("执行成功", "恭喜！ARTA 端到端学位论文与答辩 PPT 全流程已生成完毕！")
            except Exception as e:
                print(f"\n[ERROR] 执行出现异常: {e}")
                messagebox.showerror("执行错误", str(e))
            finally:
                self._set_buttons_state(True)

        threading.Thread(target=_worker, daemon=True).start()

    def _run_async(self, func):
        def _worker():
            self._set_buttons_state(False)
            try:
                func()
                print("\n[Step 完成] 步骤执行完毕。")
            except Exception as e:
                print(f"\n[ERROR] 步骤执行异常: {e}")
                messagebox.showerror("执行错误", str(e))
            finally:
                self._set_buttons_state(True)

        threading.Thread(target=_worker, daemon=True).start()

    def _set_buttons_state(self, enabled: bool):
        st = tk.NORMAL if enabled else tk.DISABLED
        self.btn_run_all.configure(state=st)
        self.btn_s1.configure(state=st)
        self.btn_s2.configure(state=st)
        self.btn_s3.configure(state=st)
        self.btn_s4.configure(state=st)
        self.btn_s5.configure(state=st)
        self.btn_s6.configure(state=st)

    def _open_base_docx(self):
        p = self.engine.base_docx_path or os.path.join(self.engine.output_dir, "鲁东大学_学术硕士学位论文_基于深度学习的抗菌肽高通量识别与智能序列设计研究_底本.docx")
        if os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showwarning("提示", f"文档尚未生成或不存在: {p}")

    def _open_final_docx(self):
        p = self.engine.final_docx_path or os.path.join(self.engine.output_dir, "鲁东大学_学术硕士学位论文_基于深度学习的抗菌肽高通量识别与智能序列设计研究_最终排版完成版.docx")
        if os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showwarning("提示", f"文档尚未生成或不存在: {p}")

    def _open_pptx(self):
        p = self.engine.ppt_path or os.path.join(self.engine.output_dir, "基于深度学习的抗菌肽高通量识别与智能序列设计研究_PPT_ppt-master", "基于深度学习的抗菌肽高通量识别与智能序列设计研究_PPTMaster.pptx")
        if os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showwarning("提示", f"PPT 尚未生成或不存在: {p}")

    def _open_output_dir(self):
        if os.path.exists(self.engine.output_dir):
            os.startfile(self.engine.output_dir)


if __name__ == "__main__":
    root = tk.Tk()
    app = ARTAGUIApp(root)
    root.mainloop()
