"""
Config resolution tests.

MCP hosts differ in the working directory they launch the server with —
Google Antigravity starts stdio servers from its own program folder — so the
project-root ``.env`` must be honoured no matter where the process runs from.
"""

import os
from pathlib import Path

import pytest

from illustrator_mcp.config import ENV_FILES, PROJECT_ROOT, Config


def test_env_files_cover_project_root_and_cwd():
    assert PROJECT_ROOT == Path(__file__).resolve().parent.parent
    assert PROJECT_ROOT / ".env" in ENV_FILES
    assert ".env" in ENV_FILES
    # The CWD copy comes last so it wins over the project-root copy.
    assert list(ENV_FILES).index(PROJECT_ROOT / ".env") < list(ENV_FILES).index(".env")


@pytest.fixture
def project_dotenv():
    """Temporarily provide a project-root .env, skipping if one already exists."""
    path = PROJECT_ROOT / ".env"
    if path.exists():
        pytest.skip(f"real {path} present — not overwriting it")
    path.write_text("WS_PORT=8123\nTIMEOUT=45\n", encoding="utf-8")
    try:
        yield path
    finally:
        path.unlink(missing_ok=True)


def test_project_root_env_is_read_from_a_foreign_cwd(tmp_path, monkeypatch, project_dotenv):
    """Exactly the Antigravity situation: CWD is not the project root."""
    monkeypatch.chdir(tmp_path)
    for key in ("WS_PORT", "TIMEOUT"):
        monkeypatch.delenv(key, raising=False)

    config = Config()

    assert config.ws_port == 8123
    assert config.timeout == 45.0


def test_process_environment_still_wins(tmp_path, monkeypatch, project_dotenv):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("WS_PORT", "8199")

    assert Config().ws_port == 8199
