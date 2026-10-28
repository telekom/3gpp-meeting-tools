# --- File: tests/test_nas_reload.py ---
"""Unit tests verifying specification reload / re-index functionality in the Protocols tab."""
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtWidgets import QApplication, QAction, QMessageBox, QWidget

from modules.nas.ui.nas_components import NASVersionTreeWidget
from modules.nas.ui.nas_ingest_controller import (
    build_spec_reload_task,
    is_ingestion_running,
    reload_checked_versions,
    reload_single_version,
    reload_spec_group,
    start_batch_ingestion,
)
from modules.specifications.core.database import SpecsDatabase

app = QApplication.instance() or QApplication(sys.argv)


class TestNASReload(unittest.TestCase):
    """Test suite for specification reload and re-indexing in the Protocols tab."""

    def setUp(self):
        self.specs_db_path = Path("src/db/3gpp_data.db")
        self.specs_db = SpecsDatabase(self.specs_db_path) if self.specs_db_path.exists() else None

    def test_build_spec_reload_task_with_specs_db(self):
        """Verifies build_spec_reload_task retrieves official filename and url from specs_db."""
        if not self.specs_db:
            self.skipTest("3gpp_data.db not found for integration test")

        mock_tab = MagicMock()
        mock_tab.specs_db = self.specs_db
        mock_tab.cache_dir = Path("cache/specs")

        task = build_spec_reload_task(mock_tab, "TS 24.501", "v18.4.0", prompt_on_missing=False)
        self.assertIsNotNone(task)
        self.assertEqual(task["spec_number"], "24.501")
        self.assertEqual(task["version"], "18.4.0")
        self.assertEqual(task["filename"], "24501-i40.zip")
        self.assertIn("https://www.3gpp.org", task["file_url"])

    def test_build_spec_reload_task_missing_silent(self):
        """Verifies build_spec_reload_task returns None without prompt when prompt_on_missing=False."""
        mock_tab = MagicMock()
        mock_tab.specs_db = None
        mock_tab.cache_dir = Path("non_existent_cache_dir_12345")

        task = build_spec_reload_task(mock_tab, "99.999", "1.0.0", prompt_on_missing=False)
        self.assertIsNone(task)

    def test_version_tree_signals_defined(self):
        """Verifies NASVersionTreeWidget defines reload signals with correct signatures."""
        tree = NASVersionTreeWidget()
        self.assertTrue(hasattr(tree, "reload_version_requested"))
        self.assertTrue(hasattr(tree, "reload_spec_requested"))

    def test_version_context_menu_has_reload_action(self):
        """Verifies right-clicking a version item exposes the '🔄 Reload TS ...' QAction."""
        tree = NASVersionTreeWidget()
        tree.resize(400, 300)

        versions_data = [
            {"id": 1, "spec_number": "24.501", "version": "18.4.0"},
        ]
        tree.populate(versions_data)

        # Locate version item (child 0 under top-level spec)
        spec_item = tree.topLevelItem(1)
        self.assertIsNotNone(spec_item)
        version_item = spec_item.child(0)
        self.assertIsNotNone(version_item)

        tree.setCurrentItem(version_item)
        pos = tree.visualItemRect(version_item).center()

        created_actions = []

        def mock_exec_(point):
            # Capture actions from active menu
            for child in tree.findChildren(QAction):
                created_actions.append(child.text())
            return None

        with patch("PyQt5.QtWidgets.QMenu.exec_", side_effect=mock_exec_):
            tree._on_context_menu(pos)

        reload_actions = [txt for txt in created_actions if "Reload TS 24.501 v18.4.0" in txt]
        self.assertTrue(len(reload_actions) > 0, f"Expected Reload action in context menu, got: {created_actions}")

    def test_spec_context_menu_has_reload_action(self):
        """Verifies right-clicking a spec item exposes the '🔄 Reload all versions...' QAction."""
        tree = NASVersionTreeWidget()
        tree.resize(400, 300)

        versions_data = [
            {"id": 1, "spec_number": "24.501", "version": "18.4.0"},
            {"id": 2, "spec_number": "24.501", "version": "17.7.0"},
        ]
        tree.populate(versions_data)

        spec_item = tree.topLevelItem(1)
        self.assertIsNotNone(spec_item)

        tree.setCurrentItem(spec_item)
        pos = tree.visualItemRect(spec_item).center()

        created_actions = []

        def mock_exec_(point):
            for child in tree.findChildren(QAction):
                created_actions.append(child.text())
            return None

        with patch("PyQt5.QtWidgets.QMenu.exec_", side_effect=mock_exec_):
            tree._on_context_menu(pos)

        reload_actions = [txt for txt in created_actions if "Reload all versions of TS 24.501" in txt]
        self.assertTrue(len(reload_actions) > 0, f"Expected Reload all action, got: {created_actions}")

    def test_reload_single_version_dispatches_batch_ingestion(self):
        """Verifies reload_single_version prompts confirmation, builds task, and starts ingestion."""
        mock_tab = MagicMock()
        mock_tab.thread = None

        with patch("modules.nas.ui.nas_ingest_controller.QMessageBox.question", return_value=QMessageBox.Yes), \
             patch("modules.nas.ui.nas_ingest_controller.build_spec_reload_task", return_value={"spec_number": "24.501", "version": "18.4.0"}), \
             patch("modules.nas.ui.nas_ingest_controller.start_batch_ingestion") as mock_start:

            reload_single_version(mock_tab, "24.501", "18.4.0")

            mock_tab._save_config.assert_called_once()
            mock_start.assert_called_once_with(mock_tab, [{"spec_number": "24.501", "version": "18.4.0"}])

    def test_reload_single_version_aborted_on_rejection(self):
        """Verifies reload_single_version does not proceed if user rejects confirmation."""
        mock_tab = MagicMock()
        mock_tab.thread = None

        with patch("modules.nas.ui.nas_ingest_controller.QMessageBox.question", return_value=QMessageBox.No), \
             patch("modules.nas.ui.nas_ingest_controller.start_batch_ingestion") as mock_start:

            reload_single_version(mock_tab, "24.501", "18.4.0")
            mock_start.assert_not_called()

    def test_reload_spec_group_dispatches_all_children(self):
        """Verifies reload_spec_group builds tasks for all child versions."""
        mock_tab = MagicMock()
        mock_tab.thread = None
        mock_tab.version_tree.get_all_versions_for_spec.return_value = [
            {"spec_number": "24.501", "version": "18.4.0"},
            {"spec_number": "24.501", "version": "17.7.0"},
        ]

        def fake_task(tab, s, v, prompt_on_missing=False):
            return {"spec_number": s, "version": v}

        with patch("modules.nas.ui.nas_ingest_controller.QMessageBox.question", return_value=QMessageBox.Yes), \
             patch("modules.nas.ui.nas_ingest_controller.build_spec_reload_task", side_effect=fake_task), \
             patch("modules.nas.ui.nas_ingest_controller.start_batch_ingestion") as mock_start:

            reload_spec_group(mock_tab, "24.501", 2)

            mock_tab._save_config.assert_called_once()
            self.assertEqual(mock_start.call_count, 1)
            tasks_arg = mock_start.call_args[0][1]
            self.assertEqual(len(tasks_arg), 2)
            self.assertEqual(tasks_arg[0]["version"], "18.4.0")
            self.assertEqual(tasks_arg[1]["version"], "17.7.0")

    def test_reload_checked_versions_dispatches_checked_only(self):
        """Verifies reload_checked_versions gathers and processes checked versions."""
        mock_tab = MagicMock()
        mock_tab.thread = None
        mock_tab.version_tree.get_checked_versions_info.return_value = [
            {"spec_number": "24.501", "version": "18.4.0"},
            {"spec_number": "38.331", "version": "17.5.0"},
        ]

        def fake_task(tab, s, v, prompt_on_missing=False):
            return {"spec_number": s, "version": v}

        with patch("modules.nas.ui.nas_ingest_controller.QMessageBox.question", return_value=QMessageBox.Yes), \
             patch("modules.nas.ui.nas_ingest_controller.build_spec_reload_task", side_effect=fake_task), \
             patch("modules.nas.ui.nas_ingest_controller.start_batch_ingestion") as mock_start:

            reload_checked_versions(mock_tab)

            mock_tab._save_config.assert_called_once()
            tasks_arg = mock_start.call_args[0][1]
            self.assertEqual(len(tasks_arg), 2)

    def test_ingestion_in_progress_guard(self):
        """Verifies start_batch_ingestion and reload actions prevent concurrent ingestion threads."""
        mock_tab = MagicMock()
        mock_thread = MagicMock()
        mock_thread.isRunning.return_value = True
        mock_tab.thread = mock_thread

        with patch("modules.nas.ui.nas_ingest_controller.QMessageBox.warning") as mock_warn:
            reload_single_version(mock_tab, "24.501", "18.4.0")
            mock_warn.assert_called_once()
            self.assertIn("Ingestion In Progress", mock_warn.call_args[0][1])

    def test_is_ingestion_running_handles_real_qwidget_without_crashing(self):
        """Verifies is_ingestion_running correctly handles QWidget instances where tab.thread is a bound method."""
        widget = QWidget()
        # QWidget.thread is a bound method returning QThread affinity, not an ingestion thread
        self.assertTrue(callable(widget.thread))
        self.assertFalse(is_ingestion_running(widget))

    def test_reload_single_version_with_real_qwidget_no_attribute_error(self):
        """Verifies reload_single_version does not crash with AttributeError on uninitialized QWidget."""
        widget = QWidget()
        with patch("modules.nas.ui.nas_ingest_controller.QMessageBox.question", return_value=QMessageBox.No):
            # Should not raise AttributeError: 'builtin_function_or_method' object has no attribute 'isRunning'
            reload_single_version(widget, "24.501", "18.4.0")


if __name__ == "__main__":
    unittest.main()
