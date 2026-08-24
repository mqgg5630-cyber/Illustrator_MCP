#!/usr/bin/env python3
"""
Register this Illustrator MCP server with Google Antigravity (反重力).

Works before or after ``pip install -e .`` — the Antigravity helpers are
stdlib-only, so the repo root is simply added to ``sys.path`` when the package
is not installed yet.

Examples
--------
    python scripts/install_antigravity.py                     # write global + workspace configs
    python scripts/install_antigravity.py --scope workspace   # only <project>/.agents/mcp_config.json
    python scripts/install_antigravity.py --port 8090 --json
    python scripts/install_antigravity.py --verify
    python scripts/install_antigravity.py --uninstall
"""

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from illustrator_mcp.antigravity import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
