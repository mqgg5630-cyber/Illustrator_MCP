# Academic Literature Hub · 去AI化科研文献工作台使用指南

> **定位**：拒绝廉价 AI 对话气泡与 Prompt 泛滥，打造专业、严肃、沉浸的工业级科研文献采集与分流工作台。支持中国知网 (CNKI) 与国际顶刊 (PubMed/PMC/Europe PMC/OpenAlex) 真实多页 PDF 批量下载、本地 Zotero 数据库自动挂载（零 C 盘占用），并支持一键通过 Gmail 邮箱安全通道直投至个人邮箱、手机端或 Kindle。

---

## 快速开启方式 (How to Launch)

工作台支持以下两种便捷启动方式：

### 方式一：Windows 桌面 / 资源管理器一键双击启动（推荐）
直接在 Windows 文件资源管理器中找到并双击运行批处理脚本：
👉 **[Start_Academic_Hub.bat](./Start_Academic_Hub.bat)**

- 双击后，系统会自动启动后台轻量级服务，并在默认浏览器中自动打开工作台交互界面：`http://127.0.0.1:8765`。

### 方式二：终端命令行启动
打开 PowerShell 或命令提示符，执行以下指令：
```powershell
cd E:\0mcp-agv
python E:\0mcp-agv\academic_hub_gui\academic_hub_app.py
```
终端启动后将输出访问地址，并在 1 秒后自动弹出浏览器窗口。

---

## 核心功能与操作流程 (Workflow)

### 1. 首次配置：Gmail 直投通道与 Zotero 联动
点击界面右上角的 **【⚙️ 邮箱与 Zotero 配置】** 按钮：
1. **发件 Gmail 账号**：填写您的个人 Gmail 邮箱地址（如 `yourname@gmail.com`）。
2. **Gmail 16 位应用密码 (App Password)**：
   - 登录 Google 账号，进入 **【安全性】 (Security)**；
   - 确保已启用【两步验证】；
   - 在搜索框搜索“应用专用密码”（App Passwords），名称输入 `AcademicHub`，系统将生成一组 16 位字母密码（如 `abcd efgh ijkl mnop`）；
   - 将该 16 位密码粘贴到输入框中。
3. **默认接收邮箱**：填写接收文献速递的目标邮箱（可填本人的 Gmail，或发送给团队成员、平板阅读邮箱）。
4. **测试握手**：点击【📧 测试 SMTP 连通】，确认成功收到握手测试邮件后，点击【💾 保存配置】。

---

### 2. 多源检索与批量下载操作
在左侧控制面板中配置下载参数：
1. **课题名称 / 关键词 (Query)**：
   - 输入中文主题（如：`食源性鲜味肽 呈味机理 分子动力学`）；
   - 或输入英文检索词（如：`acetylcholinesterase neurotoxic peptide molecular dynamics`）。
2. **数据源通道选择 (Source)**：
   - **🌐 中英混合联合检索 (Hybrid)**：同时兼顾国内知网中文核心与国际 SCI 顶刊；
   - **📚 国际顶刊 (PubMed/Europe PMC/OpenAlex)**：抓取国际权威 Open Access 官方多页出版级 PDF；
   - **🇨🇳 中国知网 (CNKI CDP)**：调用 CDP 鉴权会话直接抓取知网官方多页 PDF。
3. **篇数设定**：默认设置为 5 篇，可按需指定 1~25 篇。
4. **联动开关**：
   - `[√] 自动挂载至本地 Zotero 库`：文献下载后自动存入 `E:\ozotero\storage\`，并在本地 SQLite 中建立带 `📎 附件` 的条目。
   - `[√] 下载完毕后自动直投至 Gmail 邮箱`：自动打包全部正版 PDF 作为附件，并整理结构化清单投递至您的邮箱。
5. **点击执行**：点击 **【🚀 启动一键下载与物理联动】**。

---

### 3. 实时审计与历史归档 (Audit & Archives)
- **实时控制台 (Execution Stream)**：右上方黑色控制台实时流式输出检索状态、下载进度、`%PDF-` 二进制头审查与排伪门禁过滤信息。
- **历史归档列表 (Local Archives)**：右下方自动列出 `E:\0mcp-agv\ARTA_Agent_Output\` 下已归档的所有课题、受纳 PDF 篇数与 DOCX 综述状态，方便随时调用。

---

## 本地物理落盘与零 C 盘保障 (Storage Path)

本工作台遵循最高学术存储安全规范，所有数据与文件**严格落盘于 E 盘**，绝对杜绝占用系统盘：
* **工作台源码与配置文件**：`E:\0mcp-agv\academic_hub_gui\`
* **文献与课题落盘目录**：`E:\0mcp-agv\ARTA_Agent_Output\Hub_<课题名>\`
* **Zotero 附件物理存储库**：`E:\ozotero\storage\<KEY>\<filename>.pdf`
* **Zotero 本地数据库**：`E:\ozotero\zotero.sqlite`

---

## 常见问题与排错 (FAQ)

**Q1：Gmail 发送时提示认证失败 (Authentication Required)？**
- 请勿使用 Google 账号的日常登录密码，必须使用 Google 官方生成的 **16 位应用专用密码 (App Password)**。
- 检查网络环境是否能正常访问 `smtp.gmail.com:465`。

**Q2：Zotero 客户端打开后看不到新文献？**
- 工作台直接向本地 `E:\ozotero\zotero.sqlite` 与物理附件目录写入数据；若 Zotero 正在运行，可在 Zotero 中按 `Ctrl + R` 或切换下集合刷新显示。

**Q3：下载的 PDF 为什么有的被自动排除了？**
- 工作台内置强制纯正性审查门禁（`auto_exclude_invalid_literature.py`），任何非正版二进制、损坏文件、单页 HTML 拦截页均会被自动剔除，保障知识库 100% 纯正。
