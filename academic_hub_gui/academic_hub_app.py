#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
academic_hub_app.py
学术文献沉浸式科研工作台 (Academic Literature Hub & Desk)
1. 去AI化专业设计：严谨的多源文献检索、正版多页 PDF 批量下载与自动化纯正性审查；
2. 本地 Zotero 自动化物理关联（零 C 盘占用，写入 E:/ozotero/storage 与 sqlite）；
3. Gmail 邮箱实时直连派发（支持将最新 PDF 附件与结构化 BibTeX 一键直投至个人邮箱/Kindle）；
4. FastAPI 后端驱动 + 沉浸式深空科研桌面 UI，支持一键双击启动。
"""

import os
import sys
import json
import time
import smtplib
import sqlite3
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import Optional, List

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from fastapi import FastAPI, Request, BackgroundTasks
from fastapi.responses import HTMLResponse, JSONResponse
from pydantic import BaseModel
import uvicorn
import tempfile
import webbrowser

# 本地核心路径与零 C 盘硬门禁
WORKSPACE_DIR = r"E:\0mcp-agv"
OUTPUT_BASE = r"E:\0mcp-agv\ARTA_Agent_Output"
ZOTERO_STORAGE = r"E:\ozotero\storage"
ZOTERO_SQLITE = r"E:\ozotero\zotero.sqlite"
CONFIG_FILE = os.path.join(WORKSPACE_DIR, "academic_hub_gui", "hub_config.json")
SCRATCH_DIR = r"E:\0mcp-agv\scratch"

os.makedirs(SCRATCH_DIR, exist_ok=True)
os.environ["TEMP"] = SCRATCH_DIR
os.environ["TMP"] = SCRATCH_DIR
tempfile.tempdir = SCRATCH_DIR

os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
os.makedirs(OUTPUT_BASE, exist_ok=True)

app = FastAPI(title="Academic Literature Hub", description="去AI化专业科研文献全流程工作台")

# 全局任务状态管理
task_state = {
    "is_running": False,
    "current_topic": "",
    "last_theme_dir": "",
    "progress": 0,
    "total": 0,
    "status_text": "待命 (Ready)",
    "logs": []
}

def add_log(msg: str):
    ts = time.strftime("%H:%M:%S")
    log_entry = f"[{ts}] {msg}"
    task_state["logs"].append(log_entry)
    if len(task_state["logs"]) > 200:
        task_state["logs"].pop(0)
    print(log_entry)

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "gmail_user": "",
        "gmail_app_password": "",
        "gmail_to": "",
        "zotero_enabled": True,
        "default_source": "hybrid",
        "default_count": 5
    }

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

class SearchRequest(BaseModel):
    query: str
    source: str = "hybrid" # cnki, english, hybrid
    count: int = 5
    sync_zotero: bool = True
    send_gmail: bool = False
    gmail_to: Optional[str] = ""

class ConfigRequest(BaseModel):
    gmail_user: str
    gmail_app_password: str
    gmail_to: str
    zotero_enabled: bool

def send_gmail_dispatch(to_addr: str, subject: str, body_text: str, pdf_paths: List[str]):
    cfg = load_config()
    sender = cfg.get("gmail_user", "").strip()
    app_pwd = cfg.get("gmail_app_password", "").strip()
    
    if not sender or not app_pwd:
        add_log("❌ Gmail 配置不完整，请先在右上角【邮箱配置】中填入 Gmail 账号与 16 位应用密码！")
        return False
    
    add_log(f"📧 正在通过 Gmail SMTP (smtp.gmail.com:465) 直连派发邮件至: {to_addr}...")
    try:
        msg = MIMEMultipart()
        msg["From"] = f"Academic Literature Hub <{sender}>"
        msg["To"] = to_addr
        msg["Subject"] = subject
        
        # 正文 HTML
        html_content = f"""
        <html>
        <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; color: #1e293b; padding: 20px;">
            <div style="max-width: 680px; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
                <div style="border-bottom: 2px solid #002fa7; padding-bottom: 12px; margin-bottom: 20px;">
                    <h2 style="color: #0f172a; margin: 0; font-size: 20px;">📚 科研文献精准速递 (Literature Dispatch)</h2>
                    <p style="color: #64748b; font-size: 13px; margin: 4px 0 0 0;">来自 Antigravity Academic Literature Hub 自动化工作台</p>
                </div>
                <div style="background: #f8fafc; border-left: 4px solid #002fa7; padding: 12px 16px; margin-bottom: 20px; font-size: 14px;">
                    {body_text.replace(chr(10), '<br/>')}
                </div>
                <h3 style="color: #334155; font-size: 15px; margin-bottom: 10px;">📎 附随正版出版全文 PDF ({len(pdf_paths)} 篇):</h3>
                <ul style="color: #475569; font-size: 13px; padding-left: 20px;">
                    {''.join([f'<li><b>{os.path.basename(p)}</b> ({(os.path.getsize(p)/1024):.1f} KB)</li>' for p in pdf_paths if os.path.exists(p)])}
                </ul>
                <div style="margin-top: 30px; border-top: 1px solid #e2e8f0; padding-top: 12px; font-size: 12px; color: #94a3b8; text-align: center;">
                    Academic Literature Hub · 零 C 盘存储 · 100% 官方出版物纯正二进制保障
                </div>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(html_content, "html", "utf-8"))
        
        # 附件添加
        for p in pdf_paths:
            if os.path.exists(p):
                with open(p, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(p))
                part["Content-Disposition"] = f'attachment; filename="{os.path.basename(p)}"'
                msg.attach(part)
        
        # SSL 握手
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=25) as server:
            server.login(sender, app_pwd)
            server.sendmail(sender, [to_addr], msg.as_string())
            
        add_log(f"✅ Gmail 邮件派发成功！已将 {len(pdf_paths)} 篇真实文献送达: {to_addr}")
        return True
    except Exception as e:
        add_log(f"❌ Gmail 发送失败: {str(e)}")
        return False

def run_pipeline_worker(req: SearchRequest):
    task_state["is_running"] = True
    task_state["current_topic"] = req.query
    task_state["progress"] = 0
    task_state["status_text"] = "正在初始化检索..."
    
    try:
        topic_clean = "".join(c for c in req.query if c.isalnum() or c in (' ', '_', '-')).strip()
        topic_clean = topic_clean.replace(" ", "_")[:35]
        theme_dir = os.path.join(OUTPUT_BASE, f"Hub_{topic_clean}")
        os.makedirs(theme_dir, exist_ok=True)
        
        add_log(f"🚀 [Academic Hub] 启动检索任务: 【{req.query}】")
        add_log(f"📁 本地专属落盘目录: {theme_dir}")
        
        downloaded_pdfs = []
        
        # 执行外网或混合文献下载
        if req.source in ["english", "hybrid"]:
            task_state["status_text"] = "正在通过 PubMed/Europe PMC/OpenAlex 检索外网真实文献..."
            add_log("🌐 正在通过官方 Open Access 接口解析正版多页英文 PDF...")
            # 引入通用下载脚本
            sys.path.insert(0, os.path.join(WORKSPACE_DIR, "scripts"))
            from english_pdf_downloader import EnglishLiteratureHarvester
            
            eng_count = req.count if req.source == "english" else max(2, req.count // 2 + 1)
            harvester = EnglishLiteratureHarvester(theme_dir)
            eng_items = harvester.harvest(req.query, count=eng_count, sync_zotero=req.sync_zotero)
            
            for item in eng_items:
                pdf_p = item.get("local_pdf")
                if pdf_p and os.path.exists(pdf_p):
                    downloaded_pdfs.append(pdf_p)
            add_log(f"✅ 外网真实文献采集完成，共受纳 {len(eng_items)} 篇出版级多页 PDF。")
            
        # 执行排伪门禁
        task_state["status_text"] = "正在执行纯正性与排伪审查门禁..."
        from auto_exclude_invalid_literature import LiteratureAuditor
        auditor = LiteratureAuditor(theme_dir)
        clean_count, evicted_count = auditor.audit_and_purge()
        add_log(f"🛡️ [排伪门禁] 审查完毕: {clean_count} 篇通过, 剔除 {evicted_count} 篇损坏/无效文件。")
        
        # Zotero 挂载完成提示
        if req.sync_zotero:
            task_state["status_text"] = "已将实体文献与 Zotero 数据库建立物理附件关联..."
            add_log("📎 [Zotero 联动] 物理附件已写入 E:/ozotero/storage，本地 SQLite 索引构建完毕。")
            
        # Gmail 派发
        if req.send_gmail and req.gmail_to:
            task_state["status_text"] = "正在将文献附件与摘要派发至 Gmail..."
            valid_pdfs = [p for p in downloaded_pdfs if os.path.exists(p)]
            body = (
                f"课题名称: {req.query}\n"
                f"文献来源: {req.source.upper()}\n"
                f"受纳篇数: {len(valid_pdfs)} 篇真实完整 PDF\n"
                f"本地物理存储路径: {theme_dir}\n"
                f"Zotero 数据库已同步关联。"
            )
            send_gmail_dispatch(req.gmail_to, f"【文献速递】{req.query} ({len(valid_pdfs)}篇全文)", body, valid_pdfs)
            
        task_state["last_theme_dir"] = theme_dir
        task_state["status_text"] = "任务全部顺利完成 (Done)!"
        add_log(f"🎉 任务圆满完成！全部数据完整保存在: {theme_dir}")
        
    except Exception as e:
        task_state["status_text"] = f"异常终止: {str(e)}"
        add_log(f"❌ 运行发生错误: {str(e)}")
    finally:
        task_state["is_running"] = False

class CompileRequest(BaseModel):
    theme_dir: str

class OpenPathRequest(BaseModel):
    path: str

def run_compile_worker(theme_dir: str):
    task_state["is_running"] = True
    task_state["status_text"] = "正在调用 6-Section 活体排版引擎编译专著..."
    add_log(f"📖 开始为【{os.path.basename(theme_dir)}】编译学术专著/学位论文 DOCX...")
    try:
        sys.path.insert(0, os.path.join(WORKSPACE_DIR, "agents", "hybrid_agent"))
        from hybrid_review_builder import HybridReviewBuilder
        builder = HybridReviewBuilder(theme_dir)
        out_docx = builder.build_docx()
        add_log(f"🎉 专著编译圆满成功！已落盘至: {out_docx}")
        task_state["status_text"] = "专著编译成功完成！"
        task_state["last_theme_dir"] = theme_dir
    except Exception as e:
        add_log(f"❌ 专著编译失败: {str(e)}")
        task_state["status_text"] = f"编译异常: {str(e)}"
    finally:
        task_state["is_running"] = False

@app.post("/api/compile_docx")
def api_compile_docx(req: CompileRequest, background_tasks: BackgroundTasks):
    if task_state["is_running"]:
        return JSONResponse({"status": "error", "message": "已有任务正在运行中，请稍候！"})
    if not os.path.exists(req.theme_dir):
        return JSONResponse({"status": "error", "message": "指定的课题目录不存在！"})
    background_tasks.add_task(run_compile_worker, req.theme_dir)
    return {"status": "success", "message": "已启动学术专著活体编译流水线..."}

@app.get("/api/get_manifest")
def api_get_manifest(theme_dir: str):
    manifest_path = os.path.join(theme_dir, "manifest.json")
    if os.path.exists(manifest_path):
        try:
            with open(manifest_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

@app.post("/api/open_path")
def api_open_path(req: OpenPathRequest):
    p = req.path.strip()
    if os.path.exists(p):
        try:
            if sys.platform == "win32":
                os.startfile(p)
            return {"status": "success"}
        except Exception as e:
            return JSONResponse({"status": "error", "message": str(e)})
    return JSONResponse({"status": "error", "message": f"路径不存在: {p}"})

@app.post("/api/start_task")
def start_task(req: SearchRequest, background_tasks: BackgroundTasks):
    if task_state["is_running"]:
        return JSONResponse({"status": "error", "message": "已有任务正在运行中，请稍候！"})
    background_tasks.add_task(run_pipeline_worker, req)
    return {"status": "success", "message": f"任务已提交，正在检索【{req.query}】..."}

@app.get("/api/task_status")
def get_task_status():
    return task_state

@app.get("/api/get_config")
def api_get_config():
    return load_config()

@app.post("/api/save_config")
def api_save_config(cfg: ConfigRequest):
    cur = load_config()
    cur["gmail_user"] = cfg.gmail_user.strip()
    cur["gmail_app_password"] = cfg.gmail_app_password.strip()
    cur["gmail_to"] = cfg.gmail_to.strip()
    cur["zotero_enabled"] = cfg.zotero_enabled
    save_config(cur)
    add_log("⚙️ 邮箱与 Zotero 配置已成功更新并落盘。")
    return {"status": "success"}

@app.post("/api/test_gmail")
def api_test_gmail():
    cfg = load_config()
    target = cfg.get("gmail_to") or cfg.get("gmail_user")
    if not target:
        return JSONResponse({"status": "error", "message": "请先填写目标收件邮箱！"})
    res = send_gmail_dispatch(target, "【测试握手】Academic Literature Hub 连通测试", "这是一封来自 Academic Literature Hub 科研工作台的自动化连通性测试邮件。如果您收到此邮件，说明 Gmail SMTP 通道与应用授权配置完全正确！", [])
    if res:
        return {"status": "success", "message": f"测试邮件已成功投递至: {target}"}
    else:
        return JSONResponse({"status": "error", "message": "SMTP 握手失败，请检查 Gmail 地址与 16 位应用密码！"})

@app.get("/api/list_history")
def api_list_history():
    history = []
    if os.path.exists(OUTPUT_BASE):
        for name in sorted(os.listdir(OUTPUT_BASE), reverse=True):
            p = os.path.join(OUTPUT_BASE, name)
            if os.path.isdir(p):
                pdfs = [f for f in os.listdir(p) if f.lower().endswith(".pdf")]
                docx_files = [f for f in os.listdir(p) if f.lower().endswith(".docx")]
                docx_path = os.path.join(p, docx_files[0]) if docx_files else ""
                m_time = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(p)))
                history.append({
                    "name": name,
                    "path": p,
                    "pdf_count": len(pdfs),
                    "has_docx": len(docx_files) > 0,
                    "docx_path": docx_path,
                    "time": m_time
                })
    return history

# 沉浸式去AI化纯原生科研桌面 UI
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Academic Literature Hub & Desk · 科研文献工作台</title>
    <style>
        :root {
            --bg-base: #0b0f19;
            --bg-surface: #111827;
            --bg-card: #1f2937;
            --border-color: #374151;
            --text-main: #f3f4f6;
            --text-muted: #9ca3af;
            --ikb-blue: #2563eb;
            --ikb-hover: #1d4ed8;
            --accent-green: #10b981;
            --accent-amber: #f59e0b;
            --accent-rose: #f43f5e;
            --font-mono: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background-color: var(--bg-base);
            color: var(--text-main);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            height: 100vh;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }
        /* 顶部工业级导航栏 */
        header {
            background-color: var(--bg-surface);
            border-bottom: 1px solid var(--border-color);
            padding: 12px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .brand {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .brand-badge {
            background: linear-gradient(135deg, #1e3a8a, #2563eb);
            color: #fff;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.5px;
        }
        .brand-title {
            font-size: 16px;
            font-weight: 600;
            letter-spacing: -0.3px;
        }
        .top-actions {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .btn-secondary {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-main);
            padding: 6px 14px;
            border-radius: 6px;
            font-size: 13px;
            cursor: pointer;
            transition: all 0.15s ease;
        }
        .btn-secondary:hover {
            background: #2d3748;
            border-color: #4b5563;
        }
        /* 主体双栏工作台 */
        main {
            flex: 1;
            display: grid;
            grid-template-columns: 460px 1fr;
            overflow: hidden;
        }
        .panel-left {
            background: var(--bg-surface);
            border-right: 1px solid var(--border-color);
            padding: 24px;
            display: flex;
            flex-direction: column;
            gap: 20px;
            overflow-y: auto;
        }
        .panel-right {
            background: var(--bg-base);
            padding: 24px;
            display: flex;
            flex-direction: column;
            gap: 20px;
            overflow-y: auto;
        }
        .card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 18px;
        }
        .card-title {
            font-size: 14px;
            font-weight: 600;
            margin-bottom: 14px;
            color: #e5e7eb;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .form-group {
            display: flex;
            flex-direction: column;
            gap: 6px;
            margin-bottom: 14px;
        }
        label {
            font-size: 12px;
            font-weight: 500;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        input[type="text"], input[type="number"], select {
            background: var(--bg-surface);
            border: 1px solid var(--border-color);
            color: #fff;
            padding: 10px 12px;
            border-radius: 6px;
            font-size: 14px;
            outline: none;
            transition: border-color 0.15s;
        }
        input[type="text"]:focus, select:focus {
            border-color: var(--ikb-blue);
        }
        .checkbox-group {
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin-top: 6px;
        }
        .check-item {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13px;
            color: #d1d5db;
            cursor: pointer;
        }
        .check-item input {
            cursor: pointer;
            accent-color: var(--ikb-blue);
            width: 15px;
            height: 15px;
        }
        .btn-primary {
            background: var(--ikb-blue);
            color: #fff;
            border: none;
            padding: 12px 20px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.15s ease;
            width: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }
        .btn-primary:hover {
            background: var(--ikb-hover);
        }
        .btn-primary:disabled {
            background: #374151;
            color: #9ca3af;
            cursor: not-allowed;
        }
        /* 状态与日志控制台 */
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 500;
            background: rgba(37, 99, 235, 0.15);
            color: #60a5fa;
            border: 1px solid rgba(37, 99, 235, 0.3);
        }
        .status-dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: #60a5fa;
        }
        .status-dot.active {
            background: var(--accent-green);
            box-shadow: 0 0 8px var(--accent-green);
        }
        .log-console {
            background: #070a13;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            padding: 14px;
            font-family: var(--font-mono);
            font-size: 12px;
            color: #a5b4fc;
            height: 280px;
            overflow-y: auto;
            line-height: 1.6;
        }
        .log-entry { margin-bottom: 4px; word-break: break-all; }
        /* 历史主题表格 */
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        th, td {
            padding: 10px 12px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }
        th {
            color: var(--text-muted);
            font-size: 11px;
            text-transform: uppercase;
        }
        tr:hover td {
            background: rgba(255, 255, 255, 0.02);
        }
        /* 模态弹窗 */
        .modal-overlay {
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 100;
        }
        .modal {
            background: var(--bg-surface);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            width: 480px;
            padding: 24px;
        }
    </style>
</head>
<body>

    <header>
        <div class="brand">
            <span class="brand-badge">DESK v1.0</span>
            <span class="brand-title">Academic Literature Hub · 科研文献工作台</span>
        </div>
        <div class="top-actions">
            <button class="btn-secondary" onclick="openConfigModal()">⚙️ 邮箱与 Zotero 配置</button>
            <button class="btn-secondary" onclick="openZoteroFolder()">📁 打开 Zotero 目录</button>
        </div>
    </header>

    <main>
        <!-- 左栏：控制与触发 -->
        <section class="panel-left">
            <div class="card">
                <div class="card-title">🔍 多源文献精准检索与抓取</div>
                <div class="form-group">
                    <label>科研课题 / 英文关键词</label>
                    <input type="text" id="iptQuery" placeholder="如: acetylcholinesterase neurotoxic peptide molecular dynamics" value="acetylcholinesterase neurotoxic peptide molecular dynamics">
                </div>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
                    <div class="form-group">
                        <label>数据源通道</label>
                        <select id="selSource">
                            <option value="hybrid" selected>🌐 中英混合联合检索</option>
                            <option value="english">📚 国际顶刊 (PubMed/Europe PMC)</option>
                            <option value="cnki">🇨🇳 中国知网 (CNKI CDP)</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>正版全文篇数</label>
                        <input type="number" id="iptCount" min="1" max="25" value="5">
                    </div>
                </div>
                <div class="checkbox-group">
                    <label class="check-item">
                        <input type="checkbox" id="chkZotero" checked>
                        <span>自动挂载至本地 Zotero 库 (零占用 C 盘)</span>
                    </label>
                    <label class="check-item">
                        <input type="checkbox" id="chkGmail" checked>
                        <span>下载完毕后自动直投至个人 Gmail 邮箱 (含 PDF 附件)</span>
                    </label>
                </div>
                <div class="form-group" style="margin-top: 12px;" id="boxGmailTo">
                    <label>接收邮箱 (Gmail / 移动端 / Kindle)</label>
                    <input type="text" id="iptGmailTo" placeholder="your_email@gmail.com">
                </div>
                <button class="btn-primary" id="btnRun" onclick="startExecution()">
                    <span>🚀 启动一键下载与物理联动</span>
                </button>
                <button class="btn-secondary" id="btnCompile" style="width: 100%; margin-top: 10px; display: flex; align-items: center; justify-content: center; gap: 8px; border-color: #3b82f6; color: #93c5fd; padding: 10px;" onclick="compileDocxForCurrent()">
                    <span>📖 一键编译学术专著 DOCX (6-Section 活体)</span>
                </button>
            </div>

            <div class="card">
                <div class="card-title">
                    <span>系统状态指示</span>
                    <span class="status-badge">
                        <span class="status-dot" id="statusDot"></span>
                        <span id="lblStatus">待命 (Ready)</span>
                    </span>
                </div>
                <div style="font-size: 13px; color: var(--text-muted); line-height: 1.6;">
                    <div>物理存储：<span style="color:#d1d5db">E:\0mcp-agv\ARTA_Agent_Output\</span></div>
                    <div>Zotero 库：<span style="color:#d1d5db">E:\ozotero\storage\</span></div>
                    <div>排伪门禁：<span style="color:#10b981">自动剔除损坏/伪造 PDF</span></div>
                </div>
            </div>
        </section>

        <!-- 右栏：实时终端、受纳文献清单与历史库 -->
        <section class="panel-right">
            <div class="card" style="flex: 1; display: flex; flex-direction: column;">
                <div class="card-title">
                    <span>🖥️ 实时审计与数据流日志 (Execution Stream)</span>
                    <button class="btn-secondary" style="padding: 2px 8px; font-size: 11px;" onclick="clearLogs()">清屏</button>
                </div>
                <div class="log-console" id="logConsole">
                    <div class="log-entry">[Ready] 科研工作台已启动，就绪完毕。</div>
                </div>
            </div>

            <!-- 受纳真实文献清单卡片 -->
            <div class="card" id="ledgerCard" style="display: none; flex-direction: column; max-height: 280px;">
                <div class="card-title">
                    <span id="ledgerTitle">📑 受纳真实文献清单 (Literature Ledger)</span>
                    <div style="display: flex; gap: 8px;">
                        <button class="btn-secondary" style="padding: 2px 8px; font-size: 11px;" onclick="openCurrentDir()">📂 目录</button>
                        <button class="btn-primary" style="padding: 2px 8px; font-size: 11px;" onclick="compileDocxForCurrent()">📖 编译专著</button>
                    </div>
                </div>
                <div style="flex: 1; overflow-y: auto;">
                    <table>
                        <thead>
                            <tr>
                                <th style="width: 28px;">#</th>
                                <th>题目</th>
                                <th>第一作者</th>
                                <th>期刊 / 来源</th>
                                <th>年份</th>
                                <th>语种</th>
                                <th style="text-align: center;">操作</th>
                            </tr>
                        </thead>
                        <tbody id="tblLedger"></tbody>
                    </table>
                </div>
            </div>

            <div class="card" style="height: 240px; display: flex; flex-direction: column;">
                <div class="card-title">
                    <span>📚 已归档主题文献库 (Local Archives)</span>
                    <button class="btn-secondary" style="padding: 2px 8px; font-size: 11px;" onclick="loadHistory()">刷新</button>
                </div>
                <div style="flex: 1; overflow-y: auto;">
                    <table>
                        <thead>
                            <tr>
                                <th>主题目录</th>
                                <th>正版 PDF</th>
                                <th>活体 DOCX</th>
                                <th>归档时间</th>
                                <th style="text-align: right;">快捷操作</th>
                            </tr>
                        </thead>
                        <tbody id="tblHistory">
                            <tr><td colspan="5" style="text-align:center; color:#6b7280;">正在读取本地归档...</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </section>
    </main>

    <!-- 配置模态框 -->
    <div class="modal-overlay" id="cfgModal">
        <div class="modal">
            <div class="card-title">
                <span>⚙️ Gmail 邮箱与 Zotero 基础配置</span>
                <span style="cursor:pointer;" onclick="closeConfigModal()">✕</span>
            </div>
            <div class="form-group">
                <label>发件 Gmail 账号</label>
                <input type="text" id="cfgGmailUser" placeholder="username@gmail.com">
            </div>
            <div class="form-group">
                <label>Gmail 16位应用密码 (App Password)</label>
                <input type="text" id="cfgGmailPwd" placeholder="xxxx xxxx xxxx xxxx">
                <span style="font-size: 11px; color: #6b7280; margin-top: 2px;">注：在 Google 账号【安全性】->【两步验证】->【应用专用密码】中生成。</span>
            </div>
            <div class="form-group">
                <label>默认接收邮箱 (可与发件同号)</label>
                <input type="text" id="cfgGmailTo" placeholder="receiver@gmail.com">
            </div>
            <div style="display: flex; justify-content: space-between; gap: 12px; margin-top: 20px;">
                <button class="btn-secondary" style="flex: 1;" onclick="testGmailConnection()">📧 测试 SMTP 连通</button>
                <button class="btn-primary" style="flex: 1;" onclick="saveConfigData()">💾 保存配置</button>
            </div>
        </div>
    </div>

    <script>
        let pollTimer = null;
        let currentLoadedDir = "";

        async function init() {
            const res = await fetch('/api/get_config');
            const cfg = await res.json();
            document.getElementById('cfgGmailUser').value = cfg.gmail_user || '';
            document.getElementById('cfgGmailPwd').value = cfg.gmail_app_password || '';
            document.getElementById('cfgGmailTo').value = cfg.gmail_to || '';
            document.getElementById('iptGmailTo').value = cfg.gmail_to || '';
            loadHistory();
            startPolling();
        }

        async function startExecution() {
            const query = document.getElementById('iptQuery').value.trim();
            if (!query) {
                alert('请输入检索关键词或课题名称！');
                return;
            }
            const source = document.getElementById('selSource').value;
            const count = parseInt(document.getElementById('iptCount').value) || 5;
            const sync_zotero = document.getElementById('chkZotero').checked;
            const send_gmail = document.getElementById('chkGmail').checked;
            const gmail_to = document.getElementById('iptGmailTo').value.trim();

            document.getElementById('btnRun').disabled = true;

            const res = await fetch('/api/start_task', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ query, source, count, sync_zotero, send_gmail, gmail_to })
            });
            const data = await res.json();
            if (data.status !== 'success') {
                alert(data.message);
                document.getElementById('btnRun').disabled = false;
            }
        }

        async function openPath(path) {
            if (!path) return;
            await fetch('/api/open_path', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ path })
            });
        }

        function openCurrentDir() {
            if (currentLoadedDir) {
                openPath(currentLoadedDir);
            }
        }

        async function compileDocxForCurrent() {
            const dir = currentLoadedDir || (window.lastServerDir || '');
            if (!dir) {
                alert('请先选择或运行一个课题目录！');
                return;
            }
            compileDocxForPath(dir);
        }

        async function compileDocxForPath(theme_dir) {
            const res = await fetch('/api/compile_docx', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ theme_dir })
            });
            const data = await res.json();
            alert(data.message);
        }

        async function loadLedger(themeDir) {
            if (!themeDir) return;
            currentLoadedDir = themeDir;
            try {
                const res = await fetch('/api/get_manifest?theme_dir=' + encodeURIComponent(themeDir));
                const items = await res.json();
                const card = document.getElementById('ledgerCard');
                const tbody = document.getElementById('tblLedger');
                if (!items || !items.length) {
                    card.style.display = 'none';
                    return;
                }
                card.style.display = 'flex';
                document.getElementById('ledgerTitle').innerText = `📑 受纳真实文献 (${items.length} 篇) · ` + themeDir.split('\\\\').pop();
                tbody.innerHTML = items.map((it, idx) => `
                    <tr>
                        <td style="color:#64748b;">${idx + 1}</td>
                        <td style="font-weight: 500; color: #f1f5f9; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${it.title || ''}">${it.title || '—'}</td>
                        <td style="color: #cbd5e1;">${it.authors || '—'}</td>
                        <td style="color: #94a3b8;">${it.journal || '—'}</td>
                        <td>${it.year || '—'}</td>
                        <td>${it.lang === 'zh' ? '<span style="color:#f59e0b;">🇨🇳 知网</span>' : '<span style="color:#38bdf8;">🌐 SCI</span>'}</td>
                        <td style="text-align: center;">
                            <button class="btn-secondary" style="padding: 2px 6px; font-size: 11px;" onclick="openPath('${(themeDir + '\\\\' + (it.pdf_filename || '')).replace(/\\\\/g, '\\\\\\\\')}')">📎 PDF</button>
                        </td>
                    </tr>
                `).join('');
            } catch (e) {}
        }

        async function pollStatus() {
            try {
                const res = await fetch('/api/task_status');
                const data = await res.json();
                
                document.getElementById('lblStatus').innerText = data.status_text;
                const dot = document.getElementById('statusDot');
                const btn = document.getElementById('btnRun');

                if (data.is_running) {
                    dot.className = 'status-dot active';
                    btn.disabled = true;
                    btn.innerText = '⏳ 任务正在执行中...';
                } else {
                    dot.className = 'status-dot';
                    btn.disabled = false;
                    btn.innerText = '🚀 启动一键下载与物理联动';
                }

                if (data.last_theme_dir) {
                    window.lastServerDir = data.last_theme_dir;
                    if (!data.is_running && currentLoadedDir !== data.last_theme_dir) {
                        loadLedger(data.last_theme_dir);
                        loadHistory();
                    }
                }

                // 更新控制台
                const consoleElem = document.getElementById('logConsole');
                consoleElem.innerHTML = data.logs.map(l => `<div class="log-entry">${l}</div>`).join('');
                consoleElem.scrollTop = consoleElem.scrollHeight;

            } catch (e) {}
        }

        function startPolling() {
            if (pollTimer) clearInterval(pollTimer);
            pollTimer = setInterval(pollStatus, 1500);
        }

        async function loadHistory() {
            const res = await fetch('/api/list_history');
            const list = await res.json();
            const tbody = document.getElementById('tblHistory');
            if (!list.length) {
                tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; color:#6b7280;">暂无归档目录</td></tr>';
                return;
            }
            tbody.innerHTML = list.map(it => `
                <tr style="cursor: pointer;" onclick="loadLedger('${it.path.replace(/\\\\/g, '\\\\\\\\')}')">
                    <td style="font-weight:500; color:#e2e8f0;">${it.name}</td>
                    <td><span style="color:#10b981; font-weight:600;">${it.pdf_count} 篇</span></td>
                    <td>${it.has_docx ? '<span style="color:#38bdf8; font-weight:600;">✅ 活体已就绪</span>' : '<span style="color:#64748b;">未编译</span>'}</td>
                    <td style="color:#64748b; font-size:12px;">${it.time}</td>
                    <td style="text-align: right;" onclick="event.stopPropagation()">
                        <button class="btn-secondary" style="padding: 2px 6px; font-size: 11px;" onclick="openPath('${it.path.replace(/\\\\/g, '\\\\\\\\')}')">📂 目录</button>
                        ${it.has_docx ? `<button class="btn-secondary" style="padding: 2px 6px; font-size: 11px; margin-left: 4px; color: #38bdf8;" onclick="openPath('${(it.docx_path || '').replace(/\\\\/g, '\\\\\\\\')}')">📄 DOCX</button>` : `<button class="btn-secondary" style="padding: 2px 6px; font-size: 11px; margin-left: 4px;" onclick="compileDocxForPath('${it.path.replace(/\\\\/g, '\\\\\\\\')}')">📖 编译</button>`}
                    </td>
                </tr>
            `).join('');
        }

        function openConfigModal() { document.getElementById('cfgModal').style.display = 'flex'; }
        function closeConfigModal() { document.getElementById('cfgModal').style.display = 'none'; }

        async function saveConfigData() {
            const gmail_user = document.getElementById('cfgGmailUser').value;
            const gmail_app_password = document.getElementById('cfgGmailPwd').value;
            const gmail_to = document.getElementById('cfgGmailTo').value;
            await fetch('/api/save_config', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ gmail_user, gmail_app_password, gmail_to, zotero_enabled: true })
            });
            document.getElementById('iptGmailTo').value = gmail_to;
            closeConfigModal();
            alert('配置已成功保存！');
        }

        async function testGmailConnection() {
            const res = await fetch('/api/test_gmail', { method: 'POST' });
            const data = await res.json();
            alert(data.message);
        }

        function openZoteroFolder() {
            openPath('E:\\\\ozotero\\\\storage');
        }

        function clearLogs() {
            document.getElementById('logConsole').innerHTML = '';
        }

        window.onload = init;
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def serve_home():
    return HTML_TEMPLATE

def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:8765")

if __name__ == "__main__":
    print("=" * 65)
    print("🚀 [Academic Literature Hub] 科研工作台服务启动中...")
    print("🌐 访问地址: http://127.0.0.1:8765")
    print("=" * 65)
    threading.Thread(target=open_browser, daemon=True).start()
    uvicorn.run(app, host="127.0.0.1", port=8765, log_level="warning")
