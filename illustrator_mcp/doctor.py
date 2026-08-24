"""
Self-check for an Illustrator MCP install ("does this actually work here?").

Run it after installing, and again whenever Antigravity shows the server as
disconnected::

    python -m illustrator_mcp.doctor            # human readable
    python -m illustrator_mcp.doctor --json     # machine readable
    python -m illustrator_mcp.doctor --handshake  # also run a real MCP stdio handshake

Exit code is 0 when every check passed (warnings allowed), 1 otherwise.

The module imports nothing but the standard library at module scope so it can
diagnose an environment whose dependencies are missing or broken.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import socket
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from illustrator_mcp.antigravity import (
    DEFAULT_SERVER_NAME,
    SERVER_MODULE,
    Finding,
    project_root,
    verify,
)

__all__ = ["run_checks", "handshake", "main", "Finding"]

CEP_EXTENSION_ID = "com.illustrator.mcp.panel"
REQUIRED_MODULES = (
    ("mcp", "pip install 'mcp>=1.9.0,<3.0.0'"),
    ("pydantic", "pip install 'pydantic>=2.0.0'"),
    ("pydantic_settings", "pip install 'pydantic-settings>=2.0.0'"),
    ("websockets", "pip install 'websockets>=12.0'"),
    ("PIL", "pip install 'Pillow>=10.0.0'"),
)
OPTIONAL_MODULES = (
    ("pyclipper", "pip install -e '.[geometry]' — enables path_boolean"),
)
HANDSHAKE_TIMEOUT = 60


# ── individual checks ────────────────────────────────────────────────────


def _check_python() -> List[Finding]:
    ok = sys.version_info >= (3, 10)
    return [
        Finding(
            "ok" if ok else "error",
            "python",
            f"{'.'.join(map(str, sys.version_info[:3]))} at {sys.executable}",
            "" if ok else "Python 3.10+ is required",
        )
    ]


def _check_dependencies() -> List[Finding]:
    findings: List[Finding] = []
    import importlib

    for module, hint in REQUIRED_MODULES:
        try:
            importlib.import_module(module)
            findings.append(Finding("ok", f"dependency.{module}", "installed"))
        except Exception as exc:  # noqa: BLE001 - report any import failure
            findings.append(
                Finding("error", f"dependency.{module}", f"missing ({exc})", hint)
            )

    for module, hint in OPTIONAL_MODULES:
        try:
            importlib.import_module(module)
            findings.append(Finding("ok", f"optional.{module}", "installed"))
        except Exception:  # noqa: BLE001 - optional by design
            findings.append(Finding("warn", f"optional.{module}", "not installed", hint))

    try:
        from illustrator_mcp.compat import describe

        findings.append(Finding("ok", "mcp.sdk", describe()))
    except Exception as exc:  # noqa: BLE001
        findings.append(
            Finding(
                "error",
                "mcp.sdk",
                f"cannot resolve the MCP server class ({exc})",
                "pip install 'mcp>=1.9.0,<3.0.0'",
            )
        )
    return findings


def _check_package(root: Path) -> List[Finding]:
    findings: List[Finding] = []
    try:
        import illustrator_mcp

        version = getattr(illustrator_mcp, "__version__", "unknown")
        location = Path(getattr(illustrator_mcp, "__file__", "?")).resolve()
        findings.append(
            Finding("ok", "package", f"illustrator_mcp {version} at {location.parent}")
        )
    except Exception as exc:  # noqa: BLE001
        findings.append(
            Finding(
                "error",
                "package",
                f"cannot import illustrator_mcp ({exc})",
                f"cd {root} && pip install -e .",
            )
        )
    return findings


def _cep_extension_dirs() -> List[Path]:
    if sys.platform == "win32":
        base = os.environ.get("APPDATA")
        if not base:
            return []
        return [Path(base) / "Adobe" / "CEP" / "extensions" / CEP_EXTENSION_ID]
    if sys.platform == "darwin":
        return [
            Path.home()
            / "Library"
            / "Application Support"
            / "Adobe"
            / "CEP"
            / "extensions"
            / CEP_EXTENSION_ID
        ]
    return [Path.home() / ".config" / "Adobe" / "CEP" / "extensions" / CEP_EXTENSION_ID]


def _check_cep_panel(root: Path) -> List[Finding]:
    findings: List[Finding] = []

    source = root / "cep-extension"
    dist = source / "dist"
    if not source.is_dir():
        findings.append(
            Finding(
                "error",
                "cep.source",
                f"cep-extension/ not found in {root}",
                "run the doctor from inside the Illustrator_MCP checkout",
            )
        )
    elif not dist.is_dir():
        findings.append(
            Finding(
                "warn",
                "cep.build",
                "cep-extension/dist is missing — the panel has not been built",
                "cd cep-extension && npm install && npm run build",
            )
        )
    else:
        findings.append(Finding("ok", "cep.build", f"panel built at {dist}"))

    installed = [p for p in _cep_extension_dirs() if p.exists()]
    if installed:
        findings.append(Finding("ok", "cep.installed", f"panel installed at {installed[0]}"))
    else:
        script = "install-cep.bat" if sys.platform == "win32" else "install-cep.sh"
        findings.append(
            Finding(
                "warn",
                "cep.installed",
                f"CEP panel {CEP_EXTENSION_ID} is not installed",
                f"run {script} (as Administrator on Windows), then open "
                "Illustrator → Window → Extensions → MCP Control → Connect",
            )
        )

    if not shutil.which("npm"):
        findings.append(
            Finding(
                "warn",
                "cep.toolchain",
                "npm not found on PATH — cannot rebuild the CEP panel",
                "install Node.js LTS if you need to rebuild the panel",
            )
        )
    return findings


def _check_port(port: int) -> List[Finding]:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        in_use = sock.connect_ex(("127.0.0.1", port)) == 0
    if in_use:
        return [
            Finding(
                "warn",
                "bridge.port",
                f"port {port} already has a listener",
                "fine if Illustrator's panel is connected to a running server; "
                "otherwise another process is holding the port — set WS_PORT to a free one",
            )
        ]
    return [Finding("ok", "bridge.port", f"port {port} is free for the CEP bridge")]


def _check_stdio_purity(root: Path) -> List[Finding]:
    """Importing the server must print nothing on stdout (JSON-RPC owns it)."""
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    try:
        proc = subprocess.run(
            [sys.executable, "-c", f"import {SERVER_MODULE}"],
            cwd=str(root),
            env=env,
            capture_output=True,
            timeout=HANDSHAKE_TIMEOUT,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return [Finding("error", "stdio.purity", f"could not run the import check ({exc})")]

    stdout = proc.stdout or b""
    if proc.returncode != 0:
        tail = (proc.stderr or b"").decode("utf-8", "replace").strip().splitlines()
        detail = tail[-1] if tail else f"exit code {proc.returncode}"
        return [
            Finding(
                "error",
                "stdio.purity",
                f"importing the server failed: {detail}",
                "run: pip install -e .",
            )
        ]
    if stdout.strip():
        return [
            Finding(
                "error",
                "stdio.purity",
                f"import wrote {len(stdout)} bytes to stdout: {stdout[:120]!r}",
                "stdout carries MCP JSON-RPC — any stray print breaks Antigravity",
            )
        ]
    return [Finding("ok", "stdio.purity", "import is silent on stdout (stderr-only logging)")]


# ── live MCP stdio handshake ─────────────────────────────────────────────


def handshake(root: Path, port: Optional[int] = None) -> List[Finding]:
    """Spawn the server and speak MCP over stdio, exactly like Antigravity."""
    env = dict(
        os.environ,
        PYTHONUNBUFFERED="1",
        PYTHONIOENCODING="utf-8",
    )
    if port:
        env["WS_PORT"] = str(port)

    proc = subprocess.Popen(
        [sys.executable, "-m", SERVER_MODULE],
        cwd=str(root),
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
    )
    assert proc.stdin and proc.stdout

    def _send(payload: Dict[str, Any]) -> None:
        proc.stdin.write(json.dumps(payload).encode("utf-8") + b"\n")  # type: ignore[union-attr]
        proc.stdin.flush()  # type: ignore[union-attr]

    findings: List[Finding] = []
    try:
        _send(
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2025-06-18",
                    "capabilities": {},
                    "clientInfo": {"name": "illustrator-mcp-doctor", "version": "1.0"},
                },
            }
        )
        line = proc.stdout.readline()
        if not line:
            stderr = (proc.stderr.read() or b"").decode("utf-8", "replace")[-400:] if proc.stderr else ""
            return [
                Finding(
                    "error",
                    "handshake",
                    f"server closed stdout without answering initialize {stderr}",
                    "run the server manually to see the traceback: "
                    f"{sys.executable} -m {SERVER_MODULE}",
                )
            ]
        response = json.loads(line)
        server_name = (response.get("result") or {}).get("serverInfo", {}).get("name", "?")
        protocol = (response.get("result") or {}).get("protocolVersion", "?")
        findings.append(
            Finding("ok", "handshake.initialize", f"server '{server_name}' answered ({protocol})")
        )

        _send({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})
        _send({"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
        tools_line = proc.stdout.readline()
        tools = json.loads(tools_line).get("result", {}).get("tools", [])
        names = sorted(t.get("name", "?") for t in tools)
        if len(names) < 12:
            findings.append(
                Finding(
                    "error",
                    "handshake.tools",
                    f"only {len(names)} tools advertised: {names}",
                    "expected the 12-tool inventory from illustrator_mcp.tools",
                )
            )
        else:
            findings.append(
                Finding("ok", "handshake.tools", f"{len(names)} tools advertised")
            )
    except (OSError, ValueError, subprocess.TimeoutExpired) as exc:
        findings.append(
            Finding("error", "handshake", f"handshake failed: {exc}", "check the server logs")
        )
    finally:
        try:
            if proc.stdin:
                proc.stdin.close()
        except OSError:  # pragma: no cover
            pass
        try:
            proc.wait(timeout=20)
        except subprocess.TimeoutExpired:  # pragma: no cover
            proc.kill()
            proc.wait(timeout=10)

    if proc.returncode not in (0, None):
        findings.append(
            Finding(
                "warn",
                "handshake.shutdown",
                f"server exited with code {proc.returncode}",
                "Antigravity kills the process on disconnect, so a clean 0 is expected",
            )
        )
    else:
        findings.append(Finding("ok", "handshake.shutdown", f"clean shutdown (exit {proc.returncode})"))
    return findings


# ── orchestration ────────────────────────────────────────────────────────


def run_checks(
    root: Optional[Path] = None,
    server_name: str = DEFAULT_SERVER_NAME,
    scope: str = "all",
    with_handshake: bool = False,
    port: Optional[int] = None,
) -> List[Finding]:
    """Run every check and return the findings in a stable order."""
    root = Path(root).resolve() if root else project_root()

    findings: List[Finding] = []
    findings += _check_python()
    findings += _check_dependencies()
    findings += _check_package(root)
    findings += _check_cep_panel(root)

    try:
        from illustrator_mcp.config import config

        findings += _check_port(int(config.ws_port))
        ws_port: Optional[int] = port or int(config.ws_port)
    except Exception as exc:  # noqa: BLE001 - config import can fail on broken deps
        findings.append(Finding("error", "config", f"cannot load config ({exc})"))
        ws_port = port or 8081

    findings += _check_stdio_purity(root)
    findings += verify(name=server_name, scope=scope, root=root)
    if with_handshake:
        findings += handshake(root, port=ws_port)
    return findings


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Console entry point. Returns a process exit code."""
    parser = argparse.ArgumentParser(
        prog="illustrator-mcp-doctor",
        description="Verify that the Illustrator MCP server is installed and usable.",
    )
    parser.add_argument("--project", default=None, help="project root (default: this checkout)")
    parser.add_argument("--name", default=DEFAULT_SERVER_NAME, help="Antigravity server name")
    parser.add_argument(
        "--scope",
        choices=["all", "global", "workspace"],
        default="all",
        help="which Antigravity configs to audit",
    )
    parser.add_argument(
        "--handshake",
        action="store_true",
        help="also spawn the server and perform a real MCP stdio handshake",
    )
    parser.add_argument("--port", type=int, default=None, help="WS_PORT override for --handshake")
    parser.add_argument("--json", action="store_true", help="machine-readable output")
    args = parser.parse_args(argv)

    findings = run_checks(
        root=Path(args.project) if args.project else None,
        server_name=args.name,
        scope=args.scope,
        with_handshake=args.handshake,
        port=args.port,
    )

    if args.json:
        print(
            json.dumps(
                {
                    "ok": not any(f.level == "error" for f in findings),
                    "findings": [f.to_dict() for f in findings],
                },
                indent=2,
                ensure_ascii=False,
            )
        )
    else:
        print(f"Illustrator MCP doctor — {project_root(args.project)}")
        print("-" * 70)
        for finding in findings:
            print(finding.render())
        errors = sum(1 for f in findings if f.level == "error")
        warnings = sum(1 for f in findings if f.level == "warn")
        print("-" * 70)
        print(f"{len(findings) - errors - warnings} ok · {warnings} warn · {errors} fail")

    return 1 if any(f.level == "error" for f in findings) else 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
