# 在 Google Antigravity（反重力）中安装并使用 Illustrator MCP

本文档说明如何把本仓库的 MCP 服务器接入 **Google Antigravity**（IDE / 2.0 桌面版 / CLI 均可）。
英文版总览见仓库根目录 `README.md`。

---

## 0. 前置条件

| 依赖 | 版本 | 说明 |
|---|---|---|
| Google Antigravity | 任意近期版本 | IDE、2.0 桌面版或 CLI |
| Python | 3.10+ | 建议使用 conda / venv 环境 |
| Adobe Illustrator | 25.0（CC 2021）及以上 | 需要 CEP 面板与服务器通信 |
| Node.js | LTS（可选） | 仅用于构建 CEP 面板（`cep-extension`） |

> **Windows 用户**：Antigravity 启动 MCP 子进程时**不会继承你 PowerShell 里的 PATH**，
> 而且工作目录也不是你的项目目录。所以配置里必须写**绝对路径**——本仓库的安装脚本会自动帮你写好。

---

## 1. 一键安装（推荐）

```powershell
# 1) 克隆代码（示例目录 E:\0mcp-agv）
git clone -b arena/01a0328a-illustrator-mcp https://github.com/mqgg5630-cyber/Illustrator_MCP.git E:\0mcp-agv
cd E:\0mcp-agv

# 2) 安装 Python 包（在当前 conda/venv 环境里）
python -m pip install -e .
#   可选：布尔路径运算 path_boolean 需要 pyclipper
python -m pip install -e ".[geometry]"

# 3) 写入 Antigravity 配置（全局 + 工作区）
python scripts\install_antigravity.py
```

Windows 上也可以直接双击 / 运行 `install-antigravity.bat`，它会自动完成第 2、3 步。

脚本会写入：

| 作用域 | 文件 | 说明 |
|---|---|---|
| 全局（2.0 / IDE / CLI / SDK 共用） | `~/.gemini/config/mcp_config.json` | 所有工作区都能用 |
| 全局（旧版 IDE 1.x） | `~/.gemini/antigravity/mcp_config.json` | 仅当该目录已存在时才写 |
| 全局（旧版 CLI） | `~/.gemini/antigravity-cli/mcp_config.json` | 仅当该目录已存在时才写 |
| 工作区（当前默认） | `<项目>/.agents/mcp_config.json` | 只在该项目里生效 |
| 工作区（旧写法） | `<项目>/.agent/mcp_config.json` | 仅当该目录已存在时才写 |

生成的条目形如（全部为绝对路径）：

```json
{
  "mcpServers": {
    "illustrator": {
      "command": "C:\\Users\\you\\miniconda3\\python.exe",
      "args": ["-m", "illustrator_mcp.server"],
      "env": {
        "PYTHONUNBUFFERED": "1",
        "PYTHONIOENCODING": "utf-8",
        "WS_PORT": "8081",
        "TIMEOUT": "30"
      },
      "cwd": "E:\\0mcp-agv"
    }
  }
}
```

写入前会自动备份已有配置（`mcp_config.json.bak-<时间戳>`），并且**只增改 `illustrator` 这一项**，
不会动你其他 MCP 服务器。

### 常用参数

```powershell
python scripts\install_antigravity.py --scope global      # 只写全局配置
python scripts\install_antigravity.py --scope workspace   # 只写当前项目 .agents\mcp_config.json
python scripts\install_antigravity.py --port 8090         # 换 WebSocket 端口
python scripts\install_antigravity.py --python C:\path\to\python.exe   # 指定解释器
python scripts\install_antigravity.py --print-config      # 只打印 JSON，自己粘贴
python scripts\install_antigravity.py --verify            # 检查配置是否正确
python scripts\install_antigravity.py --dry-run --json    # 预演，不写文件
python scripts\install_antigravity.py --uninstall         # 移除本服务器条目
```

---

## 2. 安装 Illustrator 侧的 CEP 面板

MCP 服务器与 Illustrator 之间通过 WebSocket（默认 `8081`）通信，Illustrator 里必须装好面板：

```powershell
cd cep-extension
npm install
npm run build          # 生成 dist/
cd ..
.\install-cep.bat      # 需要"以管理员身份运行"
```

然后：

1. 启动 Adobe Illustrator；
2. 菜单 **窗口 → 扩展功能 → MCP Control**；
3. 在面板里点 **Connect**。

macOS 用 `./install-cep.sh`。

---

## 3. 验证

### 3.1 命令行自检

```powershell
python -m illustrator_mcp.doctor                 # 逐项检查
python -m illustrator_mcp.doctor --handshake     # 额外做一次真实的 MCP stdio 握手
python -m illustrator_mcp.doctor --json          # 机器可读输出
```

`--handshake` 会真正启动服务器并完成 `initialize` + `tools/list`，
Antigravity 连接失败时用它最能定位问题。退出码 0 = 全部通过（允许有 warn）。

### 3.2 在 Antigravity 里确认

- **Antigravity IDE**：Agent 侧边栏右上角 **…** → **MCP Servers**，确认 `illustrator` 已连接；
  点 **Manage MCP Servers → View raw config** 可以直接编辑配置文件。
- **Antigravity 2.0**：左下角 **Settings → Customizations → Installed MCP Servers**。
- **Antigravity CLI**：在对话框输入 `/mcp`，查看状态环与连接日志。

连上之后，工具会以 `illustrator_*` 前缀出现（共 12 个），例如：

```
illustrator_execute_script   illustrator_execute_task   illustrator_document
illustrator_export_document  illustrator_history        illustrator_place_file
illustrator_set_reference    illustrator_get_document   illustrator_query_items
illustrator_preflight_check  illustrator_path_boolean   illustrator_path_import_svg
```

试着对 Agent 说：「新建一个 A4 文档，画一个红色矩形」。

### 3.3 权限

Antigravity 默认对未配置的 MCP 工具使用 **Ask** 模式（每次执行都要你点确认）。
要免确认，可以在权限策略里放行：

```
mcp(illustrator/*)
```

或精确到单个工具，例如 `mcp(illustrator/illustrator_get_document)`。

---

## 4. 手动配置（不想跑脚本时）

把下面内容合并进 `~/.gemini/config/mcp_config.json`
（Windows: `%USERPROFILE%\.gemini\config\mcp_config.json`），
**把 `command` 换成你环境里 `python.exe` 的绝对路径**
（PowerShell 里执行 `(Get-Command python).Source` 可以拿到）：

```json
{
  "mcpServers": {
    "illustrator": {
      "command": "C:\\Users\\you\\miniconda3\\python.exe",
      "args": ["-m", "illustrator_mcp.server"],
      "env": {
        "PYTHONUNBUFFERED": "1",
        "PYTHONIOENCODING": "utf-8",
        "WS_PORT": "8081",
        "TIMEOUT": "30"
      },
      "cwd": "E:\\0mcp-agv"
    }
  }
}
```

支持的字段：`command` / `args` / `env` / `cwd`（stdio）、`serverUrl` / `headers` /
`authProviderType` / `oauth`（远程）、`disabled`、`disabledTools`。
**不要**写 `type`、`url`、`transport` 这类字段：官方属性表里没有它们，社区反馈会导致配置校验失败
（传输方式由 `command` / `serverUrl` 自动推断）。`--verify` 会帮你检查这一点。

改完 **重启 Antigravity**。

---

## 5. 常见问题

| 现象 | 原因 | 处理 |
|---|---|---|
| 服务器一直 disconnected / `failed to initialize: EOF` | `command` 写的是 `python` 而不是绝对路径，或该解释器没装本包 | 重跑 `python scripts\install_antigravity.py`，或用 `--python` 指定 |
| `No module named illustrator_mcp` | 配置里的解释器不是 `pip install -e .` 用的那个 | 用 `--python` 指向正确的 `python.exe`，然后重跑安装脚本 |
| 面板显示但工具为 0 | 服务器启动即崩溃 | 手动跑 `python -m illustrator_mcp.server` 看报错；再跑 `python -m illustrator_mcp.doctor` |
| `ModuleNotFoundError: No module named 'mcp.server.fastmcp'` | 装了 mcp 2.x，但代码是旧版 | 拉取本分支（已内置 1.x/2.x 兼容层 `illustrator_mcp/compat.py`），或 `pip install "mcp>=1.9.0,<3.0.0"` |
| 工具能列出但执行报未连接 | CEP 面板没连上 | Illustrator 里打开面板点 **Connect**；确认 `WS_PORT` 与配置一致 |
| `port 8081 already has a listener` | 有旧进程占用 | 结束旧进程，或用 `--port 8090` 重新安装（面板端也要改） |
| 改了 `.env` 不生效 | 旧版本只从工作目录读 `.env`，而 Antigravity 的工作目录不是项目目录 | 本分支已修复：`.env` 会同时从项目根目录读取；也可直接把变量写进配置的 `env` |
| 配置文件被判定非法 | 里面写了注释或 `type` 字段 | 用 `--print-config` 重新生成，安装脚本会把坏文件备份成 `mcp_config.json.invalid-<时间戳>.bak` |

日志：服务器所有日志走 **stderr**（stdout 专用于 JSON-RPC），
Antigravity 的 MCP 面板里能看到；命令行手动运行即可直接看到。

---

## 6. 卸载

```powershell
python scripts\install_antigravity.py --uninstall     # 移除配置条目
python -m pip uninstall illustrator-mcp               # 可选：卸载 Python 包
```

CEP 面板在 `%APPDATA%\Adobe\CEP\extensions\com.illustrator.mcp.panel` 手动删除即可。
