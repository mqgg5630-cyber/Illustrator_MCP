"""
Google Antigravity (反重力) integration: config discovery, install, verify.

Antigravity does **not** read ``claude_desktop_config.json`` and it does not
inherit your shell ``PATH``. It reads a ``mcp_config.json`` with a top-level
``mcpServers`` object from one of these locations:

=================================  =============================================
Scope                              Path
=================================  =============================================
Global (2.0 / IDE / CLI / SDK)     ``~/.gemini/config/mcp_config.json``
Global (Antigravity IDE 1.x)       ``~/.gemini/antigravity/mcp_config.json``
Global (Antigravity CLI, legacy)   ``~/.gemini/antigravity-cli/mcp_config.json``
Workspace (current default)        ``<project>/.agents/mcp_config.json``
Workspace (legacy spelling)        ``<project>/.agent/mcp_config.json``
=================================  =============================================

Two Antigravity behaviours drive everything this module writes:

1. **No inherited PATH.** A bare ``"command": "python"`` or
   ``"command": "illustrator-mcp"`` often fails to resolve when the GUI spawns
   the process, so entries always use the *absolute* interpreter path.
2. **CWD is not your project.** Stdio servers are launched from the
   Antigravity program folder, so entries carry an absolute ``cwd`` and the
   module is started with ``-m illustrator_mcp.server``.

The module is deliberately stdlib-only so it can run *before* the package is
installed (``python scripts/install_antigravity.py``).
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple, Union

__all__ = [
    "CONFIG_FILENAME",
    "DEFAULT_SERVER_NAME",
    "SERVER_MODULE",
    "Finding",
    "ConfigTarget",
    "TargetResult",
    "InstallReport",
    "project_root",
    "global_targets",
    "workspace_targets",
    "resolve_targets",
    "build_entry",
    "load_config",
    "write_config",
    "install",
    "uninstall",
    "verify",
    "main",
]

CONFIG_FILENAME = "mcp_config.json"
DEFAULT_SERVER_NAME = "illustrator"
SERVER_MODULE = "illustrator_mcp.server"

#: Global config locations, newest first. ``legacy`` paths are only written
#: when they already exist (an older Antigravity install created them) or when
#: the caller passes ``include_all_legacy``.
GLOBAL_LOCATIONS: Tuple[Tuple[str, Tuple[str, ...], bool], ...] = (
    ("config", (".gemini", "config"), False),
    ("antigravity", (".gemini", "antigravity"), True),
    ("antigravity-cli", (".gemini", "antigravity-cli"), True),
)

#: Workspace config directories, current default first.
WORKSPACE_DIRS: Tuple[Tuple[str, bool], ...] = ((".agents", False), (".agent", True))


# ── diagnostics plumbing (shared with illustrator_mcp.doctor) ────────────


@dataclass
class Finding:
    """A single check result. ``level`` is one of ok / warn / error."""

    level: str
    key: str
    message: str
    hint: str = ""

    def to_dict(self) -> Dict[str, str]:
        return {"level": self.level, "key": self.key, "message": self.message, "hint": self.hint}

    def render(self) -> str:
        badge = {"ok": "[ OK ]", "warn": "[WARN]", "error": "[FAIL]"}.get(self.level, "[ ?? ]")
        line = f"{badge} {self.key}: {self.message}"
        if self.hint:
            line += f"\n         → {self.hint}"
        return line


# ── targets ──────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class ConfigTarget:
    """One ``mcp_config.json`` location Antigravity may read."""

    scope: str  # "global" | "workspace"
    label: str
    path: Path
    legacy: bool = False

    @property
    def exists(self) -> bool:
        return self.path.is_file()

    @property
    def parent_exists(self) -> bool:
        return self.path.parent.is_dir()

    def to_dict(self) -> Dict[str, Any]:
        return {
            "scope": self.scope,
            "label": self.label,
            "path": str(self.path),
            "legacy": self.legacy,
            "exists": self.exists,
        }

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        return f"{self.label} ({self.path})"


def project_root(start: Optional[Union[str, Path]] = None) -> Path:
    """Locate the Illustrator_MCP checkout containing this package."""
    here = Path(start).resolve() if start else Path(__file__).resolve().parent.parent
    if here.is_file():
        here = here.parent
    for candidate in [here, *here.parents]:
        pyproject = candidate / "pyproject.toml"
        if pyproject.is_file():
            try:
                if "illustrator-mcp" in pyproject.read_text(encoding="utf-8"):
                    return candidate
            except OSError:  # pragma: no cover - unreadable file
                continue
    # Fall back to the parent of the package directory (editable installs).
    return Path(__file__).resolve().parent.parent


def global_targets(
    home_dir: Optional[Union[str, Path]] = None,
    include_all_legacy: bool = False,
) -> List[ConfigTarget]:
    """Global ``mcp_config.json`` locations for the current user."""
    home = Path(home_dir).expanduser() if home_dir else Path.home()
    targets: List[ConfigTarget] = []
    for label, parts, legacy in GLOBAL_LOCATIONS:
        path = home.joinpath(*parts) / CONFIG_FILENAME
        if legacy and not include_all_legacy and not path.parent.is_dir():
            continue
        targets.append(
            ConfigTarget(scope="global", label=f"global/{label}", path=path, legacy=legacy)
        )
    return targets


def workspace_targets(
    root: Optional[Union[str, Path]] = None,
    include_all_legacy: bool = False,
) -> List[ConfigTarget]:
    """Workspace-level ``mcp_config.json`` locations inside a project."""
    base = Path(root).resolve() if root else project_root()
    targets: List[ConfigTarget] = []
    for dirname, legacy in WORKSPACE_DIRS:
        directory = base / dirname
        path = directory / CONFIG_FILENAME
        if legacy and not include_all_legacy and not directory.is_dir():
            continue
        targets.append(
            ConfigTarget(scope="workspace", label=f"workspace/{dirname}", path=path, legacy=legacy)
        )
    return targets


def resolve_targets(
    scope: str = "all",
    root: Optional[Union[str, Path]] = None,
    home_dir: Optional[Union[str, Path]] = None,
    include_all_legacy: bool = False,
) -> List[ConfigTarget]:
    """Expand a ``--scope`` value into concrete config targets."""
    scope = (scope or "all").lower()
    if scope not in {"all", "global", "workspace"}:
        raise ValueError(f"unknown scope {scope!r} (expected all|global|workspace)")
    targets: List[ConfigTarget] = []
    if scope in {"all", "global"}:
        targets += global_targets(home_dir=home_dir, include_all_legacy=include_all_legacy)
    if scope in {"all", "workspace"}:
        targets += workspace_targets(root=root, include_all_legacy=include_all_legacy)
    return targets


# ── entry construction ───────────────────────────────────────────────────


def build_entry(
    python_exe: Optional[Union[str, Path]] = None,
    project_dir: Optional[Union[str, Path]] = None,
    ws_port: Optional[int] = None,
    timeout: Optional[Union[int, float]] = None,
    env: Optional[Dict[str, str]] = None,
    module: str = SERVER_MODULE,
    include_cwd: bool = True,
) -> Dict[str, Any]:
    """Build the ``mcpServers.<name>`` object Antigravity should spawn.

    The command is always an absolute interpreter path because Antigravity's
    spawn environment does not inherit the user shell ``PATH``.
    """
    interpreter = Path(python_exe).resolve() if python_exe else Path(sys.executable).resolve()
    cwd = Path(project_dir).resolve() if project_dir else project_root()

    entry_env: Dict[str, str] = {
        # Keep stdio framing intact and make stderr logs UTF-8 on Windows,
        # where the default console codec would mangle the ✓/✗ markers.
        "PYTHONUNBUFFERED": "1",
        "PYTHONIOENCODING": "utf-8",
    }
    if ws_port is not None:
        entry_env["WS_PORT"] = str(int(ws_port))
    if timeout is not None:
        entry_env["TIMEOUT"] = str(timeout)
    if env:
        entry_env.update({str(k): str(v) for k, v in env.items()})

    entry: Dict[str, Any] = {
        "command": str(interpreter),
        "args": ["-m", module],
        "env": entry_env,
    }
    if include_cwd:
        entry["cwd"] = str(cwd)
    return entry


# ── config file IO ───────────────────────────────────────────────────────


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d%H%M%S")


def load_config(path: Union[str, Path]) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """Read a ``mcp_config.json``.

    Returns ``(data, None)`` on success, ``(None, error)`` when the file is
    missing or unparseable. A missing file is not an error condition for the
    installer — it returns ``({}, None)``.
    """
    path = Path(path)
    if not path.is_file():
        return {}, None
    try:
        raw = path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        return None, f"cannot read {path}: {exc}"
    if not raw.strip():
        return {}, None
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        return None, f"{path} is not valid JSON ({exc.msg} at line {exc.lineno} column {exc.colno})"
    if not isinstance(data, dict):
        return None, f"{path} must contain a JSON object at the top level"
    return data, None


def write_config(path: Union[str, Path], data: Dict[str, Any]) -> Path:
    """Atomically write JSON config, creating parent directories."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(data, indent=2, ensure_ascii=False) + "\n"
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(payload, encoding="utf-8")
    os.replace(tmp, path)
    return path


# ── install / uninstall ──────────────────────────────────────────────────


@dataclass
class TargetResult:
    """What happened to one config target."""

    target: ConfigTarget
    action: str  # created | updated | unchanged | removed | absent | planned
    backup: Optional[str] = None
    warnings: List[str] = field(default_factory=list)
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "target": self.target.to_dict(),
            "action": self.action,
            "backup": self.backup,
            "warnings": self.warnings,
            "error": self.error,
        }


@dataclass
class InstallReport:
    """Aggregate result of an install/uninstall run."""

    results: List[TargetResult] = field(default_factory=list)

    @property
    def errors(self) -> List[str]:
        return [r.error for r in self.results if r.error]

    @property
    def warnings(self) -> List[str]:
        return [w for r in self.results for w in r.warnings]

    @property
    def ok(self) -> bool:
        return not self.errors

    def to_dict(self) -> Dict[str, Any]:
        return {
            "ok": self.ok,
            "results": [r.to_dict() for r in self.results],
            "warnings": self.warnings,
            "errors": self.errors,
        }


def _mutate(
    targets: Sequence[ConfigTarget],
    name: str,
    mutate_fn,
    dry_run: bool = False,
) -> InstallReport:
    """Apply ``mutate_fn(data, name) -> action`` to every target config."""
    report = InstallReport()
    for target in targets:
        warnings: List[str] = []
        data, err = load_config(target.path)
        if err is not None:
            if data is None and target.exists:
                # Never destroy a config we cannot parse: park it aside.
                backup = target.path.with_name(
                    f"{CONFIG_FILENAME}.invalid-{_timestamp()}.bak"
                )
                if not dry_run:
                    try:
                        shutil.copy2(target.path, backup)
                    except OSError as exc:  # pragma: no cover - FS failure
                        report.results.append(
                            TargetResult(target, "unchanged", None, warnings, f"{err} ({exc})")
                        )
                        continue
                warnings.append(f"existing config was not valid JSON ({err}); kept a copy at {backup}")
                data = {}
            else:
                report.results.append(TargetResult(target, "unchanged", None, warnings, err))
                continue

        assert data is not None
        servers = data.get("mcpServers")
        if servers is None:
            servers = {}
            data["mcpServers"] = servers
        elif not isinstance(servers, dict):
            warnings.append("'mcpServers' was not an object; replaced it")
            servers = {}
            data["mcpServers"] = servers

        action = mutate_fn(servers, name)
        if action == "unchanged":
            report.results.append(TargetResult(target, action, None, warnings))
            continue
        if action == "absent":
            report.results.append(TargetResult(target, action, None, warnings))
            continue

        backup_path: Optional[str] = None
        if dry_run:
            report.results.append(TargetResult(target, f"planned:{action}", None, warnings))
            continue

        if target.exists:
            backup = target.path.with_name(f"{CONFIG_FILENAME}.bak-{_timestamp()}")
            try:
                shutil.copy2(target.path, backup)
                backup_path = str(backup)
            except OSError as exc:  # pragma: no cover - FS failure
                warnings.append(f"could not back up existing config: {exc}")

        try:
            write_config(target.path, data)
        except OSError as exc:
            report.results.append(
                TargetResult(target, "unchanged", backup_path, warnings, f"write failed: {exc}")
            )
            continue

        report.results.append(TargetResult(target, action, backup_path, warnings))
    return report


def install(
    targets: Sequence[ConfigTarget],
    name: str = DEFAULT_SERVER_NAME,
    entry: Optional[Dict[str, Any]] = None,
    dry_run: bool = False,
) -> InstallReport:
    """Write (or refresh) the ``illustrator`` entry in every target config."""
    entry = entry if entry is not None else build_entry()

    # Replace the whole entry so stale keys (a relative ``command``, a
    # rejected ``type`` field, an old port) cannot survive an upgrade.
    def _set(servers: Dict[str, Any], key: str) -> str:
        if servers.get(key) == entry:
            return "unchanged"
        action = "created" if key not in servers else "updated"
        servers[key] = entry
        return action

    return _mutate(targets, name, _set, dry_run=dry_run)


def uninstall(
    targets: Sequence[ConfigTarget],
    name: str = DEFAULT_SERVER_NAME,
    dry_run: bool = False,
) -> InstallReport:
    """Remove only the ``illustrator`` entry, leaving other servers alone."""

    def _drop(servers: Dict[str, Any], key: str) -> str:
        if key not in servers:
            return "absent"
        servers.pop(key)
        return "removed"

    return _mutate(targets, name, _drop, dry_run=dry_run)


# ── verification ─────────────────────────────────────────────────────────


def verify(
    name: str = DEFAULT_SERVER_NAME,
    targets: Optional[Sequence[ConfigTarget]] = None,
    scope: str = "all",
    root: Optional[Union[str, Path]] = None,
    home_dir: Optional[Union[str, Path]] = None,
) -> List[Finding]:
    """Audit Antigravity configs without touching them."""
    if targets is None:
        targets = resolve_targets(scope=scope, root=root, home_dir=home_dir)

    findings: List[Finding] = []
    if not targets:
        findings.append(
            Finding(
                "warn",
                "antigravity.targets",
                "no mcp_config.json target resolved for this scope",
                "pass --scope all, or re-run the installer",
            )
        )
        return findings

    found_any = False
    for target in targets:
        data, err = load_config(target.path)
        if err is not None:
            findings.append(
                Finding(
                    "error",
                    f"antigravity.{target.label}",
                    f"{target.path} is unusable: {err}",
                    "fix the JSON or delete the file and re-run the installer",
                )
            )
            continue
        if not target.exists:
            findings.append(
                Finding(
                    "warn" if found_any else "error",
                    f"antigravity.{target.label}",
                    f"{target.path} does not exist",
                    "run: python scripts/install_antigravity.py --scope all",
                )
            )
            continue

        servers = (data or {}).get("mcpServers") or {}
        entry = servers.get(name) if isinstance(servers, dict) else None
        if not isinstance(entry, dict):
            findings.append(
                Finding(
                    "warn" if found_any else "error",
                    f"antigravity.{target.label}",
                    f"no '{name}' server in {target.path}",
                    "run: python scripts/install_antigravity.py --scope all",
                )
            )
            continue

        found_any = True
        findings.extend(_verify_entry(target, name, entry))

    if not found_any:
        findings.append(
            Finding(
                "error",
                "antigravity.server",
                f"'{name}' is not configured for Antigravity anywhere",
                "run: python scripts/install_antigravity.py --scope all",
            )
        )
    return findings


def _verify_entry(target: ConfigTarget, name: str, entry: Dict[str, Any]) -> List[Finding]:
    findings: List[Finding] = []
    key = f"antigravity.{target.label}.{name}"

    if "serverUrl" not in entry and not entry.get("command"):
        findings.append(
            Finding(
                "error",
                key,
                "entry has neither 'command' (stdio) nor 'serverUrl' (remote)",
                "re-run the installer to regenerate the entry",
            )
        )
        return findings

    for rejected in ("type", "url", "transport"):
        if rejected in entry:
            findings.append(
                Finding(
                    "error",
                    key,
                    f"entry contains unsupported key '{rejected}' — Antigravity rejects it",
                    "remove the key; transport is inferred from command/serverUrl",
                )
            )

    command = entry.get("command")
    if command:
        if not Path(command).is_absolute():
            findings.append(
                Finding(
                    "error",
                    key,
                    f"command '{command}' is not an absolute path",
                    "Antigravity does not inherit your shell PATH; re-run the installer",
                )
            )
        elif not Path(command).exists():
            findings.append(
                Finding(
                    "error",
                    key,
                    f"command '{command}' does not exist",
                    "re-run the installer from the Python environment that has the package",
                )
            )
        else:
            findings.append(Finding("ok", key, f"command OK: {command}"))

    args = entry.get("args") or []
    if command and "-m" not in args and SERVER_MODULE not in " ".join(map(str, args)):
        findings.append(
            Finding(
                "warn",
                key,
                f"args {args} do not start the server module",
                f"expected ['-m', '{SERVER_MODULE}']",
            )
        )

    cwd = entry.get("cwd")
    if cwd and not Path(str(cwd)).is_dir():
        findings.append(
            Finding("warn", key, f"cwd '{cwd}' does not exist", "re-run the installer")
        )
    return findings


# ── CLI ──────────────────────────────────────────────────────────────────


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="install_antigravity",
        description=(
            "Register the Illustrator MCP server with Google Antigravity "
            "(反重力) by writing an absolute-path stdio entry into "
            "mcp_config.json."
        ),
    )
    parser.add_argument(
        "--scope",
        choices=["all", "global", "workspace"],
        default="all",
        help="which mcp_config.json files to write (default: all)",
    )
    parser.add_argument("--name", default=DEFAULT_SERVER_NAME, help="server name in mcpServers")
    parser.add_argument(
        "--project",
        default=None,
        help="project root used for 'cwd' (default: this checkout)",
    )
    parser.add_argument(
        "--python",
        default=None,
        help="interpreter that has illustrator_mcp installed (default: this one)",
    )
    parser.add_argument("--port", type=int, default=None, help="WS_PORT for the CEP bridge")
    parser.add_argument("--timeout", type=float, default=None, help="TIMEOUT in seconds")
    parser.add_argument(
        "--env",
        action="append",
        default=[],
        metavar="KEY=VALUE",
        help="extra environment variable (repeatable)",
    )
    parser.add_argument(
        "--all-legacy",
        action="store_true",
        help="also write legacy config locations (~/.gemini/antigravity*, .agent/)",
    )
    parser.add_argument("--no-cwd", action="store_true", help="omit the 'cwd' key")
    parser.add_argument("--uninstall", action="store_true", help="remove the server entry")
    parser.add_argument("--verify", action="store_true", help="audit configs and exit")
    parser.add_argument("--print-config", action="store_true", help="print the JSON snippet only")
    parser.add_argument("--dry-run", action="store_true", help="report changes without writing")
    parser.add_argument("--json", action="store_true", help="machine-readable output")
    return parser


def _parse_env_pairs(pairs: Iterable[str]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for pair in pairs:
        if "=" not in pair:
            raise SystemExit(f"--env expects KEY=VALUE, got {pair!r}")
        key, value = pair.split("=", 1)
        out[key.strip()] = value
    return out


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Console entry point. Returns a process exit code."""
    args = _parser().parse_args(argv)
    root = Path(args.project).resolve() if args.project else project_root()
    entry = build_entry(
        python_exe=args.python,
        project_dir=root,
        ws_port=args.port,
        timeout=args.timeout,
        env=_parse_env_pairs(args.env),
        include_cwd=not args.no_cwd,
    )
    targets = resolve_targets(
        scope=args.scope, root=root, include_all_legacy=args.all_legacy
    )

    if args.print_config:
        print(json.dumps({"mcpServers": {args.name: entry}}, indent=2, ensure_ascii=False))
        return 0

    if args.verify:
        findings = verify(name=args.name, targets=targets)
        if args.json:
            print(json.dumps([f.to_dict() for f in findings], indent=2, ensure_ascii=False))
        else:
            for finding in findings:
                print(finding.render())
        return 1 if any(f.level == "error" for f in findings) else 0

    if args.uninstall:
        report = uninstall(targets, name=args.name, dry_run=args.dry_run)
    else:
        report = install(targets, name=args.name, entry=entry, dry_run=args.dry_run)

    if args.json:
        print(json.dumps(report.to_dict(), indent=2, ensure_ascii=False))
    else:
        verb = "Removed" if args.uninstall else "Wrote"
        print(f"Illustrator MCP → Google Antigravity ({args.scope} scope)")
        for result in report.results:
            marker = {"unchanged": "=", "absent": "-", "removed": "x"}.get(
                result.action, "+" if not result.action.startswith("planned") else "~"
            )
            print(f"  [{marker}] {result.target.label}: {result.action}  {result.target.path}")
            for warning in result.warnings:
                print(f"      ! {warning}")
            if result.error:
                print(f"      x {result.error}")
        if not args.uninstall and report.ok:
            print()
            print(f"{verb} mcpServers.{args.name} → {entry['command']} {' '.join(entry['args'])}")
            print("Next: restart Antigravity, then open Agent panel → … → MCP Servers to confirm")
            print(f"      '{args.name}' is connected (or run /mcp in Antigravity CLI).")

    return 0 if report.ok else 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
