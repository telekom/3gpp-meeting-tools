"""Unit tests for continuous GitHub master self-updater and UI dialog."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
for _p in (str(ROOT_DIR), str(SRC_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from PyQt5.QtWidgets import QApplication

from core.ui.update_dialog import AppUpdateDialog
from core.updater.app_updater import (
    CommitInfo,
    LocalInstallInfo,
    UpdateCheckResult,
    _safe_extract_zip,
    evaluate_update_status,
    inspect_local_install,
    parse_commits_atom,
    should_auto_check,
    stage_zip_update,
)
from core.updater.apply_worker import (
    APP_VERSION_FILENAME,
    apply_staged_tree,
    is_protected_relative_path,
    is_seed_only_relative_path,
)

SAMPLE_ATOM = b"""<?xml version="1.0" encoding="UTF-8"?>
<feed xmlns="http://www.w3.org/2005/Atom">
  <title>Recent Commits to 3gpp-meeting-tools:master</title>
  <entry>
    <id>tag:github.com,2008:Grit::Commit/aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa</id>
    <link type="text/html" rel="alternate" href="https://github.com/telekom/3gpp-meeting-tools/commit/aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa" />
    <title>feat(updater): add continuous master self-updater</title>
    <updated>2026-11-14T12:00:00Z</updated>
    <author><name>Alice</name></author>
    <content type="html">&lt;pre style='white-space:pre-wrap'&gt;feat(updater): add continuous master self-updater

- Add Atom feed watcher
- Preserve local DBs and configs&lt;/pre&gt;</content>
  </entry>
  <entry>
    <id>tag:github.com,2008:Grit::Commit/bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb</id>
    <link type="text/html" rel="alternate" href="https://github.com/telekom/3gpp-meeting-tools/commit/bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb" />
    <title>fix(ui): polish status bar badge</title>
    <updated>2026-11-13T10:00:00Z</updated>
    <author><name /><email>bob@example.com</email></author>
    <content type="html">&lt;pre&gt;fix(ui): polish status bar badge&lt;/pre&gt;</content>
  </entry>
  <entry>
    <id>tag:github.com,2008:Grit::Commit/cccccccccccccccccccccccccccccccccccccccc</id>
    <link type="text/html" rel="alternate" href="https://github.com/telekom/3gpp-meeting-tools/commit/cccccccccccccccccccccccccccccccccccccccc" />
    <title>chore: initial baseline</title>
    <updated>2026-11-10T08:00:00Z</updated>
    <author><name>Carol</name></author>
    <content type="html">&lt;pre&gt;chore: initial baseline&lt;/pre&gt;</content>
  </entry>
</feed>
"""


class TestAppUpdater(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_parse_commits_atom_extracts_multiline_message_and_author(self) -> None:
        commits = parse_commits_atom(SAMPLE_ATOM)
        self.assertEqual(len(commits), 3)

        c0 = commits[0]
        self.assertEqual(c0.sha, "a" * 40)
        self.assertEqual(c0.short_sha, "aaaaaaa")
        self.assertEqual(c0.title, "feat(updater): add continuous master self-updater")
        self.assertIn("Preserve local DBs and configs", c0.message)
        self.assertEqual(c0.author, "Alice")
        self.assertGreater(c0.updated_ts, 0.0)
        self.assertTrue(c0.url.endswith("a" * 40))

        # Falls back to email when <author><name/> is empty
        self.assertEqual(commits[1].author, "bob@example.com")

    def test_parse_commits_atom_rejects_malformed_xml(self) -> None:
        with self.assertRaises(ValueError):
            parse_commits_atom(b"<not-closed>")

    def test_inspect_local_install_modes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "3GPP Tools"
            (root / "src").mkdir(parents=True)
            main_py = root / "src" / "main_tools.py"
            main_py.write_text("# main\n", encoding="utf-8")

            # 1. Manual ZIP without .git or .app_version.json
            info_manual = inspect_local_install(root)
            self.assertEqual(info_manual.install_mode, "zip_manual")
            self.assertEqual(info_manual.installed_sha, "")
            self.assertGreater(info_manual.estimated_ts, 0.0)

            # 2. Tracked ZIP via .app_version.json
            (root / APP_VERSION_FILENAME).write_text(
                json.dumps(
                    {
                        "installed_sha": "b" * 40,
                        "short_sha": "bbbbbbb",
                        "managed_files": ["src/main_tools.py"],
                    }
                ),
                encoding="utf-8",
            )
            info_tracked = inspect_local_install(root)
            self.assertEqual(info_tracked.install_mode, "zip_tracked")
            self.assertEqual(info_tracked.installed_sha, "b" * 40)
            self.assertEqual(info_tracked.managed_files, ["src/main_tools.py"])

            # 3. Git checkout detection via .git/HEAD
            git_dir = Path(tmp) / ".git"
            (git_dir / "refs" / "heads").mkdir(parents=True)
            (git_dir / "HEAD").write_text("ref: refs/heads/master\n", encoding="utf-8")
            (git_dir / "refs" / "heads" / "master").write_text("c" * 40 + "\n", encoding="utf-8")
            info_git = inspect_local_install(root)
            self.assertEqual(info_git.install_mode, "git")
            self.assertEqual(info_git.installed_sha, "c" * 40)

    def test_evaluate_update_status_tracked_and_manual(self) -> None:
        commits = parse_commits_atom(SAMPLE_ATOM)

        # Up-to-date tracked
        res_ok = evaluate_update_status(
            LocalInstallInfo("zip_tracked", Path("."), installed_sha="a" * 40, short_sha="aaaaaaa"),
            commits,
        )
        self.assertEqual(res_ok.status, "up_to_date")
        self.assertEqual(res_ok.pending_commits, [])

        # Behind by 2 commits
        res_behind = evaluate_update_status(
            LocalInstallInfo("zip_tracked", Path("."), installed_sha="c" * 40, short_sha="ccccccc"),
            commits,
        )
        self.assertEqual(res_behind.status, "update_available")
        self.assertEqual([c.short_sha for c in res_behind.pending_commits], ["aaaaaaa", "bbbbbbb"])
        self.assertFalse(res_behind.history_truncated)

        # Older than all 3 commits in feed -> history_truncated=True
        res_very_old = evaluate_update_status(
            LocalInstallInfo("zip_tracked", Path("."), installed_sha="d" * 40, short_sha="ddddddd"),
            commits,
        )
        self.assertEqual(res_very_old.status, "update_available")
        self.assertEqual(len(res_very_old.pending_commits), 3)
        self.assertTrue(res_very_old.history_truncated)

        # Manual ZIP timestamp comparison
        res_manual_behind = evaluate_update_status(
            LocalInstallInfo("zip_manual", Path("."), estimated_ts=commits[1].updated_ts),
            commits,
        )
        self.assertEqual(res_manual_behind.status, "update_available")
        self.assertEqual([c.short_sha for c in res_manual_behind.pending_commits], ["aaaaaaa"])

        res_manual_latest = evaluate_update_status(
            LocalInstallInfo("zip_manual", Path("."), estimated_ts=commits[0].updated_ts),
            commits,
        )
        self.assertEqual(res_manual_latest.status, "up_to_date")

        # Developer git repo ahead of origin/master (latest remote commit already in HEAD)
        ahead_local = LocalInstallInfo(
            "git", Path("."), installed_sha="f" * 40, short_sha="fffffff", git_root=Path(".")
        )
        with patch("core.updater.app_updater._git_is_ancestor", return_value=True):
            res_ahead = evaluate_update_status(ahead_local, commits)
            self.assertEqual(res_ahead.status, "up_to_date")
            self.assertEqual(res_ahead.pending_commits, [])

    def test_should_auto_check_throttling(self) -> None:
        now = 1_700_000_000.0
        self.assertFalse(should_auto_check({"auto_check_enabled": False, "last_check_ts": 0.0}, now))
        self.assertFalse(
            should_auto_check(
                {"auto_check_enabled": True, "check_interval_hours": 12, "last_check_ts": now - 3600},
                now,
            )
        )
        self.assertTrue(
            should_auto_check(
                {"auto_check_enabled": True, "check_interval_hours": 12, "last_check_ts": now - 13 * 3600},
                now,
            )
        )

    def test_protected_and_seed_paths(self) -> None:
        self.assertTrue(is_protected_relative_path("src/config/network_config.json"))
        self.assertTrue(is_protected_relative_path("src/config/plantuml_presets_user.json"))
        self.assertTrue(is_protected_relative_path("src/assets/plantuml.jar"))
        self.assertTrue(is_protected_relative_path("src/db/user_tdocs.db"))
        self.assertTrue(is_protected_relative_path("output/diagram.vsdx"))
        self.assertTrue(is_seed_only_relative_path("src/db/3gpp_data.db"))
        self.assertTrue(is_seed_only_relative_path("src/config/useful_links.json"))
        self.assertFalse(is_protected_relative_path("src/main_tools.py"))
        self.assertFalse(is_protected_relative_path("src/config/plantuml_presets.json"))

    def test_apply_staged_tree_preserves_user_data_and_cleans_obsolete_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            app_root = tmp_path / "install" / "3GPP Tools"
            staged = tmp_path / "staged" / "3GPP Tools"
            app_data = tmp_path / "appdata"

            # Existing installation with user DB, configs, plantuml.jar, output, and obsolete .py
            (app_root / "src" / "db").mkdir(parents=True)
            (app_root / "src" / "config").mkdir(parents=True)
            (app_root / "src" / "assets").mkdir(parents=True)
            (app_root / "output").mkdir(parents=True)
            (app_root / "src" / "main_tools.py").write_text("OLD_MAIN", encoding="utf-8")
            (app_root / "src" / "obsolete_module.py").write_text("REMOVE_ME", encoding="utf-8")
            (app_root / "src" / "db" / "3gpp_data.db").write_bytes(b"USER_SYNCED_DB")
            (app_root / "src" / "config" / "network_config.json").write_text('{"proxy":"corp"}', encoding="utf-8")
            (app_root / "src" / "config" / "useful_links.json").write_text('{"custom":true}', encoding="utf-8")
            (app_root / "src" / "assets" / "plantuml.jar").write_bytes(b"JAR_BYTES")
            (app_root / "output" / "chart.vsdx").write_bytes(b"USER_OUTPUT")

            # Staged update tree from GitHub ZIP
            (staged / "src" / "db").mkdir(parents=True)
            (staged / "src" / "config").mkdir(parents=True)
            (staged / "src" / "main_tools.py").write_text("NEW_MAIN", encoding="utf-8")
            (staged / "src" / "new_feature.py").write_text("NEW_FEATURE", encoding="utf-8")
            (staged / "src" / "config" / "plantuml_presets.json").write_text('{"v":2}', encoding="utf-8")
            (staged / "src" / "config" / "useful_links.json").write_text('{"default":true}', encoding="utf-8")
            (staged / "src" / "db" / "3gpp_data.db").write_bytes(b"EMPTY_SEED_DB")

            summary = apply_staged_tree(
                staged_app_dir=staged,
                app_root=app_root,
                commit_meta={
                    "sha": "a" * 40,
                    "short_sha": "aaaaaaa",
                    "title": "feat: update",
                    "updated_iso": "2026-11-14T12:00:00Z",
                },
                prev_managed_files=["src/main_tools.py", "src/obsolete_module.py"],
                app_data_root=app_data,
                log_fn=lambda _msg: None,
            )

            self.assertEqual(summary["removed"], 1)
            self.assertEqual((app_root / "src" / "main_tools.py").read_text(encoding="utf-8"), "NEW_MAIN")
            self.assertEqual((app_root / "src" / "new_feature.py").read_text(encoding="utf-8"), "NEW_FEATURE")
            self.assertFalse((app_root / "src" / "obsolete_module.py").exists())

            # Existing user files preserved verbatim
            self.assertEqual((app_root / "src" / "db" / "3gpp_data.db").read_bytes(), b"USER_SYNCED_DB")
            self.assertEqual((app_root / "src" / "config" / "useful_links.json").read_text(encoding="utf-8"), '{"custom":true}')
            self.assertEqual((app_root / "src" / "config" / "network_config.json").read_text(encoding="utf-8"), '{"proxy":"corp"}')
            self.assertEqual((app_root / "src" / "assets" / "plantuml.jar").read_bytes(), b"JAR_BYTES")
            self.assertEqual((app_root / "output" / "chart.vsdx").read_bytes(), b"USER_OUTPUT")

            # Seed files DO get copied if missing on a fresh install
            fresh_root = tmp_path / "fresh"
            (fresh_root / "src").mkdir(parents=True)
            apply_staged_tree(
                staged_app_dir=staged,
                app_root=fresh_root,
                commit_meta={"sha": "a" * 40, "short_sha": "aaaaaaa", "title": "fresh"},
                log_fn=lambda _msg: None,
            )
            self.assertEqual((fresh_root / "src" / "db" / "3gpp_data.db").read_bytes(), b"EMPTY_SEED_DB")
            self.assertEqual((fresh_root / "src" / "config" / "useful_links.json").read_text(encoding="utf-8"), '{"default":true}')

    def test_zipslip_protection_and_stage_zip_update(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            bad_zip = tmp_path / "evil.zip"
            with zipfile.ZipFile(bad_zip, "w") as zf:
                zf.writestr("../escape.txt", "bad")
            with self.assertRaises(ValueError):
                _safe_extract_zip(bad_zip, tmp_path / "out")

            # Valid GitHub archive layout staging
            app_root = tmp_path / "install" / "3GPP Tools"
            (app_root / "src").mkdir(parents=True)
            (app_root / "requirements.txt").write_text("PyQt5>=5.15.0\n", encoding="utf-8")

            good_zip = tmp_path / "good.zip"
            with zipfile.ZipFile(good_zip, "w") as zf:
                zf.writestr("3gpp-meeting-tools-aaa/3GPP Tools/src/main_tools.py", "print('ok')\n")
                zf.writestr("3gpp-meeting-tools-aaa/3GPP Tools/requirements.txt", "PyQt5>=5.15.0\nrequests>=2.25.1\n")

            def _fake_download(url, dest_path, **_kwargs):
                Path(dest_path).write_bytes(good_zip.read_bytes())
                return good_zip.stat().st_size

            commit = parse_commits_atom(SAMPLE_ATOM)[0]
            local_info = LocalInstallInfo("zip_tracked", app_root=app_root, installed_sha="b" * 40)
            with (
                patch("core.updater.app_updater.get_temp_root", return_value=tmp_path / "temp"),
                patch("core.updater.app_updater.NetworkSession.download_file", side_effect=_fake_download),
            ):
                staged = stage_zip_update(commit, local_info=local_info)
                self.assertTrue((staged.staged_app_dir / "src" / "main_tools.py").is_file())
                self.assertTrue(staged.requirements_changed)

    def test_update_dialog_renders_states(self) -> None:
        commits = parse_commits_atom(SAMPLE_ATOM)
        local = LocalInstallInfo("zip_tracked", Path("."), installed_sha="c" * 40, short_sha="ccccccc")
        res = evaluate_update_status(local, commits)

        dlg = AppUpdateDialog(initial_result=res)
        self.assertIn("2 new changes", dlg.status_title_lbl.text())
        self.assertTrue(dlg.apply_btn.isEnabled())
        self.assertIn("aaaaaaa", dlg.changes_browser.toHtml())
        self.assertIn("Preserve local DBs and configs", dlg.changes_browser.toHtml())
        self.assertFalse(hasattr(dlg, "auto_check_cb"))
        dlg.close()

        # Dirty git workspace disables overwrite
        dirty_local = LocalInstallInfo("git", Path("."), installed_sha="c" * 40, short_sha="ccccccc", is_git_dirty=True)
        res_dirty = evaluate_update_status(dirty_local, commits)
        dlg2 = AppUpdateDialog(initial_result=res_dirty)
        self.assertFalse(dlg2.apply_btn.isEnabled())
        self.assertFalse(dlg2.notice_lbl.isHidden())
        dlg2.close()

    def test_status_bar_update_pill(self) -> None:
        from core.ui.system_status_bar import SystemStatusBar

        bar = SystemStatusBar()
        try:
            self.assertTrue(bar.update_indicator.isHidden())
            fired = []
            bar.update_check_requested.connect(lambda: fired.append(True))
            bar.set_update_available(3, "abc1234", "fix(ui): sample")
            self.assertFalse(bar.update_indicator.isHidden())
            self.assertIn("3", bar.update_indicator.text())
            bar.update_indicator.click()
            self.assertEqual(fired, [True])
            bar.clear_update_available()
            self.assertTrue(bar.update_indicator.isHidden())
        finally:
            bar.stop_monitors()
            bar.mcp_timer.stop()


if __name__ == "__main__":
    unittest.main()
