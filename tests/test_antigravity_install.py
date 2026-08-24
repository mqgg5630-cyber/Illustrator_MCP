"""
Tests for the Google Antigravity (反重力) installer/verifier.

Antigravity reads ``mcp_config.json`` from ``~/.gemini/config/`` (global) or
``<workspace>/.agents/`` (project) and requires absolute commands because it
does not inherit the user shell PATH. These tests pin that behaviour using a
temporary home directory, so nothing touches the developer's real config.
"""

import json
import sys
from pathlib import Path

import pytest

from illustrator_mcp import antigravity as ag


# ── helpers ──────────────────────────────────────────────────────────────


def _global_path(home: Path) -> Path:
    return home / ".gemini" / "config" / "mcp_config.json"


def _write(path: Path, data) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return path


@pytest.fixture
def entry(tmp_path):
    """A deterministic entry pointing at a real interpreter and project dir."""
    return ag.build_entry(
        python_exe=sys.executable,
        project_dir=tmp_path / "proj",
        ws_port=8090,
        timeout=15,
    )


# ── entry construction ───────────────────────────────────────────────────


class TestBuildEntry:
    def test_command_is_absolute(self, entry):
        assert Path(entry["command"]).is_absolute()
        assert Path(entry["command"]) == Path(sys.executable).resolve()

    def test_args_start_the_server_module(self, entry):
        assert entry["args"] == ["-m", "illustrator_mcp.server"]

    def test_env_values_are_strings(self, entry):
        assert entry["env"]["WS_PORT"] == "8090"
        assert entry["env"]["TIMEOUT"] == "15"
        assert entry["env"]["PYTHONIOENCODING"] == "utf-8"
        assert all(isinstance(v, str) for v in entry["env"].values())

    def test_cwd_is_absolute_project_dir(self, tmp_path, entry):
        assert Path(entry["cwd"]).is_absolute()
        assert Path(entry["cwd"]) == (tmp_path / "proj").resolve()

    def test_no_transport_field_antigravity_rejects(self, entry):
        # Antigravity rejects a 'type' key and infers transport from command.
        assert "type" not in entry
        assert "transport" not in entry
        assert "url" not in entry

    def test_extra_env_merged(self, tmp_path):
        entry = ag.build_entry(project_dir=tmp_path, env={"ILLUSTRATOR_MCP_DEBUG": "1"})
        assert entry["env"]["ILLUSTRATOR_MCP_DEBUG"] == "1"

    def test_cwd_can_be_omitted(self, tmp_path):
        assert "cwd" not in ag.build_entry(project_dir=tmp_path, include_cwd=False)


# ── target resolution ────────────────────────────────────────────────────


class TestTargets:
    def test_global_default_is_gemini_config(self, tmp_path):
        targets = ag.global_targets(home_dir=tmp_path)
        assert [t.label for t in targets] == ["global/config"]
        assert targets[0].path == _global_path(tmp_path)

    def test_legacy_global_dirs_skipped_until_they_exist(self, tmp_path):
        assert len(ag.global_targets(home_dir=tmp_path)) == 1
        (tmp_path / ".gemini" / "antigravity").mkdir(parents=True)
        labels = [t.label for t in ag.global_targets(home_dir=tmp_path)]
        assert "global/antigravity" in labels

    def test_all_legacy_flag_writes_every_location(self, tmp_path):
        labels = [t.label for t in ag.global_targets(home_dir=tmp_path, include_all_legacy=True)]
        assert labels == ["global/config", "global/antigravity", "global/antigravity-cli"]

    def test_workspace_default_is_dot_agents(self, tmp_path):
        targets = ag.workspace_targets(root=tmp_path)
        assert [t.label for t in targets] == ["workspace/.agents"]

    def test_workspace_legacy_dir_included_when_present(self, tmp_path):
        (tmp_path / ".agent").mkdir()
        labels = [t.label for t in ag.workspace_targets(root=tmp_path)]
        assert labels == ["workspace/.agents", "workspace/.agent"]

    def test_scope_all_combines_global_and_workspace(self, tmp_path):
        targets = ag.resolve_targets(scope="all", root=tmp_path / "proj", home_dir=tmp_path)
        assert {t.scope for t in targets} == {"global", "workspace"}

    def test_unknown_scope_rejected(self, tmp_path):
        with pytest.raises(ValueError):
            ag.resolve_targets(scope="everywhere", root=tmp_path, home_dir=tmp_path)


# ── install ──────────────────────────────────────────────────────────────


class TestInstall:
    def test_creates_global_config(self, tmp_path, entry):
        targets = ag.global_targets(home_dir=tmp_path)
        report = ag.install(targets, entry=entry)

        assert report.ok
        config = json.loads(_global_path(tmp_path).read_text(encoding="utf-8"))
        assert config["mcpServers"]["illustrator"] == entry

    def test_preserves_other_servers(self, tmp_path, entry):
        other = {"command": "npx", "args": ["-y", "some-other-mcp"]}
        _write(_global_path(tmp_path), {"mcpServers": {"other": other}})

        ag.install(ag.global_targets(home_dir=tmp_path), entry=entry)

        config = json.loads(_global_path(tmp_path).read_text(encoding="utf-8"))
        assert config["mcpServers"]["other"] == other
        assert config["mcpServers"]["illustrator"] == entry

    def test_replaces_stale_entry(self, tmp_path, entry):
        # A hand-written entry with a relative command and a rejected 'type'.
        _write(
            _global_path(tmp_path),
            {"mcpServers": {"illustrator": {"command": "python", "type": "stdio"}}},
        )

        report = ag.install(ag.global_targets(home_dir=tmp_path), entry=entry)

        assert report.results[0].action == "updated"
        config = json.loads(_global_path(tmp_path).read_text(encoding="utf-8"))
        assert config["mcpServers"]["illustrator"] == entry

    def test_idempotent_second_run_writes_nothing(self, tmp_path, entry):
        targets = ag.global_targets(home_dir=tmp_path)
        ag.install(targets, entry=entry)
        before = _global_path(tmp_path).read_text(encoding="utf-8")

        report = ag.install(targets, entry=entry)

        assert report.results[0].action == "unchanged"
        assert _global_path(tmp_path).read_text(encoding="utf-8") == before
        assert not list(_global_path(tmp_path).parent.glob("*.bak-*"))

    def test_backup_created_when_content_changes(self, tmp_path, entry):
        _write(_global_path(tmp_path), {"mcpServers": {"other": {"command": "npx"}}})
        ag.install(ag.global_targets(home_dir=tmp_path), entry=entry)
        assert list(_global_path(tmp_path).parent.glob("mcp_config.json.bak-*"))

    def test_invalid_existing_config_is_parked_not_destroyed(self, tmp_path, entry):
        broken = _global_path(tmp_path)
        broken.parent.mkdir(parents=True)
        broken.write_text("{ this is not json", encoding="utf-8")

        report = ag.install(ag.global_targets(home_dir=tmp_path), entry=entry)

        assert report.ok
        assert report.warnings and "not valid JSON" in report.warnings[0]
        parked = list(broken.parent.glob("mcp_config.json.invalid-*.bak"))
        assert parked and parked[0].read_text(encoding="utf-8") == "{ this is not json"
        config = json.loads(broken.read_text(encoding="utf-8"))
        assert config["mcpServers"]["illustrator"] == entry

    def test_dry_run_leaves_disk_untouched(self, tmp_path, entry):
        report = ag.install(ag.global_targets(home_dir=tmp_path), entry=entry, dry_run=True)
        assert report.results[0].action == "planned:created"
        assert not _global_path(tmp_path).exists()

    def test_workspace_scope_writes_dot_agents(self, tmp_path, entry):
        root = tmp_path / "proj"
        report = ag.install(ag.workspace_targets(root=root), entry=entry)
        assert report.ok
        written = root / ".agents" / "mcp_config.json"
        assert written.is_file()
        assert json.loads(written.read_text(encoding="utf-8"))["mcpServers"]["illustrator"] == entry


# ── uninstall ────────────────────────────────────────────────────────────


class TestUninstall:
    def test_removes_only_our_entry(self, tmp_path, entry):
        other = {"command": "npx", "args": ["-y", "some-other-mcp"]}
        _write(_global_path(tmp_path), {"mcpServers": {"illustrator": entry, "other": other}})

        report = ag.uninstall(ag.global_targets(home_dir=tmp_path))

        assert report.results[0].action == "removed"
        config = json.loads(_global_path(tmp_path).read_text(encoding="utf-8"))
        assert "illustrator" not in config["mcpServers"]
        assert config["mcpServers"]["other"] == other

    def test_absent_entry_is_not_an_error(self, tmp_path):
        report = ag.uninstall(ag.global_targets(home_dir=tmp_path))
        assert report.ok
        assert report.results[0].action == "absent"


# ── verification ─────────────────────────────────────────────────────────


class TestVerify:
    def test_reports_unconfigured_install(self, tmp_path):
        findings = ag.verify(targets=ag.global_targets(home_dir=tmp_path))
        assert any(f.level == "error" for f in findings)
        assert any("not configured" in f.message for f in findings)

    def test_clean_after_install(self, tmp_path, entry):
        targets = ag.global_targets(home_dir=tmp_path)
        ag.install(targets, entry=entry)
        findings = ag.verify(targets=targets)
        assert not [f for f in findings if f.level == "error"], [f.render() for f in findings]
        assert any(f.key.startswith("antigravity.global/config.illustrator") for f in findings)

    def test_flags_relative_command(self, tmp_path):
        _write(
            _global_path(tmp_path),
            {"mcpServers": {"illustrator": {"command": "python", "args": ["-m", ag.SERVER_MODULE]}}},
        )
        findings = ag.verify(targets=ag.global_targets(home_dir=tmp_path))
        assert any("not an absolute path" in f.message for f in findings)

    def test_flags_rejected_type_key(self, tmp_path, entry):
        bad = dict(entry, type="stdio")
        _write(_global_path(tmp_path), {"mcpServers": {"illustrator": bad}})
        findings = ag.verify(targets=ag.global_targets(home_dir=tmp_path))
        assert any("unsupported key 'type'" in f.message for f in findings)

    def test_flags_unparseable_config(self, tmp_path):
        broken = _global_path(tmp_path)
        broken.parent.mkdir(parents=True)
        broken.write_text("nope", encoding="utf-8")
        findings = ag.verify(targets=ag.global_targets(home_dir=tmp_path))
        assert any(f.level == "error" and "unusable" in f.message for f in findings)


# ── CLI ──────────────────────────────────────────────────────────────────


class TestCli:
    def test_print_config_emits_mergeable_snippet(self, capsys, tmp_path):
        code = ag.main(
            [
                "--print-config",
                "--scope",
                "workspace",
                "--project",
                str(tmp_path),
                "--port",
                "8099",
            ]
        )
        assert code == 0
        data = json.loads(capsys.readouterr().out)
        assert data["mcpServers"]["illustrator"]["env"]["WS_PORT"] == "8099"
        assert Path(data["mcpServers"]["illustrator"]["command"]).is_absolute()

    def test_install_then_verify_via_cli(self, capsys, tmp_path):
        project = tmp_path / "proj"
        project.mkdir()
        argv = ["--scope", "workspace", "--project", str(project)]

        assert ag.main(argv) == 0
        assert (project / ".agents" / "mcp_config.json").is_file()

        assert ag.main([*argv, "--verify"]) == 0

        assert ag.main([*argv, "--uninstall"]) == 0
        config = json.loads((project / ".agents" / "mcp_config.json").read_text(encoding="utf-8"))
        assert config["mcpServers"] == {}

    def test_json_output_shape(self, capsys, tmp_path):
        assert ag.main(["--scope", "workspace", "--project", str(tmp_path), "--json"]) == 0
        data = json.loads(capsys.readouterr().out)
        assert data["ok"] is True
        assert data["results"][0]["action"] == "created"
        assert data["results"][0]["target"]["scope"] == "workspace"

    def test_env_pair_validation(self, tmp_path):
        with pytest.raises(SystemExit):
            ag.main(["--scope", "workspace", "--project", str(tmp_path), "--env", "BADPAIR"])
