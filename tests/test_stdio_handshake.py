"""
End-to-end stdio handshake tests.

Antigravity (反重力) talks to this server the same way Claude Desktop does:
it spawns ``python -m illustrator_mcp.server`` and exchanges newline-delimited
JSON-RPC over stdin/stdout. These tests reproduce that exchange with raw pipes
(no MCP client SDK involved) so they work on both mcp 1.x and mcp 2.x, and
assert the two invariants a stdio host depends on:

* stdout carries **only** JSON-RPC frames — logging goes to stderr
* ``tools/list`` advertises the full 12-tool inventory
"""

import json
import os
import subprocess
import sys
import threading
from pathlib import Path

import pytest

from illustrator_mcp.tools import EXPECTED_TOOL_NAMES

REPO_ROOT = Path(__file__).resolve().parent.parent
# Deliberately not 8081: a developer may have a live bridge on the default port.
TEST_PORT = 8199
READ_TIMEOUT = 90


def _spawn():
    env = dict(
        os.environ,
        PYTHONUNBUFFERED="1",
        PYTHONIOENCODING="utf-8",
        WS_PORT=str(TEST_PORT),
    )
    proc = subprocess.Popen(
        [sys.executable, "-m", "illustrator_mcp.server"],
        cwd=str(REPO_ROOT),
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
    )
    stderr_lines: list = []

    def _drain():
        assert proc.stderr is not None
        for raw in proc.stderr:
            stderr_lines.append(raw.decode("utf-8", "replace"))

    thread = threading.Thread(target=_drain, daemon=True)
    thread.start()
    return proc, stderr_lines, thread


def _send(proc, payload) -> None:
    assert proc.stdin is not None
    proc.stdin.write(json.dumps(payload).encode("utf-8") + b"\n")
    proc.stdin.flush()


def _read_json(proc, timeout: float = READ_TIMEOUT) -> dict:
    """Read one stdout line as JSON, with a timeout so CI cannot hang."""
    box: dict = {}

    def _reader():
        assert proc.stdout is not None
        box["raw"] = proc.stdout.readline()

    thread = threading.Thread(target=_reader, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        raise TimeoutError(f"server produced no stdout line within {timeout}s")
    raw = box.get("raw") or b""
    if not raw.strip():
        raise EOFError("server closed stdout without sending a response")
    return json.loads(raw.decode("utf-8"))


@pytest.fixture(scope="module")
def session():
    """One full MCP session: initialize → tools/list → EOF shutdown."""
    proc, stderr_lines, drain_thread = _spawn()
    result: dict = {"stderr_lines": stderr_lines, "stdout_lines": 0}
    try:
        _send(
            proc,
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2025-06-18",
                    "capabilities": {},
                    "clientInfo": {"name": "pytest-stdio", "version": "1.0"},
                },
            },
        )
        init = _read_json(proc)
        result["stdout_lines"] += 1
        result["initialize"] = init

        _send(proc, {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})
        _send(proc, {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
        tools = _read_json(proc)
        result["stdout_lines"] += 1
        result["tools"] = tools
    finally:
        if proc.stdin:
            proc.stdin.close()
        try:
            result["returncode"] = proc.wait(timeout=30)
        except subprocess.TimeoutExpired:  # pragma: no cover - hang is a failure
            proc.kill()
            result["returncode"] = proc.wait(timeout=15)
            result["timed_out"] = True
        drain_thread.join(timeout=10)
        result["stderr"] = "".join(stderr_lines)
    return result


def test_initialize_handshake(session):
    payload = session["initialize"]
    assert payload.get("jsonrpc") == "2.0"
    assert payload.get("id") == 1
    server_info = payload["result"]["serverInfo"]
    assert server_info["name"] == "illustrator_mcp"
    assert payload["result"]["protocolVersion"]


def test_tools_list_advertises_full_inventory(session):
    tools = session["tools"]["result"]["tools"]
    names = {tool["name"] for tool in tools}
    assert names == EXPECTED_TOOL_NAMES
    # Every tool must carry a description — Antigravity shows it in the picker.
    assert all(tool.get("description") for tool in tools)


def test_stdout_is_pure_json_rpc(session):
    # Two requests were answered; anything extra on stdout would be a banner
    # or a stray print, which corrupts the JSON-RPC stream for the host.
    assert session["stdout_lines"] == 2


def test_logs_go_to_stderr(session):
    assert "LIFESPAN STARTUP" in session["stderr"]
    assert "tools registered" in session["stderr"]


def test_clean_shutdown_on_stdin_close(session):
    assert not session.get("timed_out"), "server ignored EOF on stdin"
    assert session["returncode"] == 0, session["stderr"][-800:]


def test_importing_the_server_is_silent_on_stdout():
    """A stray print at import time would break every stdio host."""
    proc = subprocess.run(
        [sys.executable, "-c", "import illustrator_mcp.server"],
        cwd=str(REPO_ROOT),
        capture_output=True,
        timeout=READ_TIMEOUT,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")[-800:]
    assert proc.stdout == b"", f"stdout must stay empty, got: {proc.stdout[:200]!r}"
