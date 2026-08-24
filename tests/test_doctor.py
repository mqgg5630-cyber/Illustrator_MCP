"""
Tests for the ``illustrator-mcp-doctor`` self-check.

The doctor is the thing a user runs when Antigravity shows the server as
disconnected, so its checks and exit codes are covered here against the real
checkout.
"""

import json
from pathlib import Path

from illustrator_mcp import doctor

REPO_ROOT = Path(__file__).resolve().parent.parent
HANDSHAKE_PORT = 8299  # avoid the default bridge port


def _levels(findings):
    return {f.level for f in findings}


def test_run_checks_covers_the_main_subsystems():
    findings = doctor.run_checks(root=REPO_ROOT, scope="workspace")
    keys = {f.key for f in findings}

    assert "python" in keys
    assert "mcp.sdk" in keys
    assert "stdio.purity" in keys
    assert "bridge.port" in keys
    assert any(k.startswith("dependency.") for k in keys)
    assert any(k.startswith("antigravity.") for k in keys)


def test_python_and_stdio_checks_pass_on_this_environment():
    findings = doctor.run_checks(root=REPO_ROOT, scope="workspace")
    by_key = {f.key: f for f in findings}

    assert by_key["python"].level == "ok"
    assert by_key["stdio.purity"].level == "ok", by_key["stdio.purity"].message
    assert by_key["mcp.sdk"].level == "ok"
    assert by_key["mcp.sdk"].message.startswith("mcp ")
    assert by_key["package"].level == "ok"


def test_missing_panel_is_a_warning_not_an_error():
    findings = doctor.run_checks(root=REPO_ROOT, scope="workspace")
    by_key = {f.key: f for f in findings}
    # The CEP panel is Windows/macOS-only tooling; in CI it is simply absent.
    assert by_key["cep.installed"].level == "warn"
    assert "install-cep" in by_key["cep.installed"].hint


def test_main_json_output_and_exit_code(capsys):
    code = doctor.main(["--json", "--scope", "workspace", "--project", str(REPO_ROOT)])
    data = json.loads(capsys.readouterr().out)

    assert set(data) == {"ok", "findings"}
    assert data["ok"] is (code == 0)
    assert any(f["key"] == "python" for f in data["findings"])
    assert all(f["level"] in {"ok", "warn", "error"} for f in data["findings"])


def test_main_human_output_summarises(capsys):
    doctor.main(["--scope", "workspace", "--project", str(REPO_ROOT)])
    out = capsys.readouterr().out
    assert "Illustrator MCP doctor" in out
    assert "ok ·" in out and "fail" in out


def test_handshake_against_a_real_server():
    findings = doctor.handshake(REPO_ROOT, port=HANDSHAKE_PORT)
    by_key = {f.key: f for f in findings}

    assert by_key["handshake.initialize"].level == "ok", by_key["handshake.initialize"].message
    assert "illustrator_mcp" in by_key["handshake.initialize"].message
    assert by_key["handshake.tools"].level == "ok", by_key["handshake.tools"].message
    assert by_key["handshake.shutdown"].level == "ok", by_key["handshake.shutdown"].message
