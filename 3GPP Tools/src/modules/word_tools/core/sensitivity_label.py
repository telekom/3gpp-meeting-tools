# --- File: modules/word_tools/core/sensitivity_label.py ---
import traceback
import uuid
import platform
import logging
from enum import Enum
from datetime import datetime, timezone
from typing import Any

from modules.word_tools.core.word_config import WordConfig

logger = logging.getLogger(__name__)


def _log(message: str, level: int = logging.INFO, ui_logger=None) -> None:
    """Log through Python logging; optionally support legacy signal/logger callers."""
    logger.log(level, message)


class MsoAssignmentMethod(Enum):
    NOT_SET = -1  # The assignment method value is not set.
    STANDARD = 0  # The label is applied by default.
    PRIVILEGED = 1  # The label was manually selected.
    AUTO = 2  # The label is applied automatically.

    @staticmethod
    def from_int(value_int: int):
        try:
            return MsoAssignmentMethod(value_int)
        except Exception as e:
            logger.warning(f'Value does not exist in Enum: {e}')
            return None


def set_sensitivity_label(document: Any, ui_logger=None):
    """
    Applies a corporate sensitivity label to a Word COM Document object to bypass manual save popups.
    """
    if platform.system() != 'Windows':
        return document

    config = WordConfig.load()

    # ---> NEW: The Gatekeeper check
    if not config.get("enable_sensitivity_labels", False):
        _log("   ➔ Sensitivity labels disabled in config. Skipping...", logging.DEBUG, ui_logger)
        return document

    label_id = config.get("sensitivity_level_label_id")
    label_name = config.get("sensitivity_level_label_name")

    if not label_id or not label_name:
        _log("⚠️ Sensitivity labels enabled, but ID or Name is missing in config!", logging.WARNING, ui_logger)
        return document

    did_something = False
    try:
        # Check if the document actually supports sensitivity labels
        if getattr(document, 'SensitivityLabel', None) is None:
            return document

        sensitivity_label = document.SensitivityLabel.GetLabel()
        assignment_method = MsoAssignmentMethod.from_int(sensitivity_label.AssignmentMethod)

        if assignment_method == MsoAssignmentMethod.NOT_SET:
            _log(f"   ➔ Applying corporate Sensitivity Label: {label_name}...", logging.INFO, ui_logger)

            did_something = True
            new_sl = document.SensitivityLabel.CreateLabelInfo()
            new_sl.ActionId = str(uuid.uuid1())
            new_sl.AssignmentMethod = 1
            new_sl.ContentBits = 0
            new_sl.IsEnabled = True
            new_sl.Justification = ''
            new_sl.LabelId = label_id
            new_sl.LabelName = label_name
            # Dynamically generate current UTC time
            new_sl.SetDate = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
            new_sl.SiteId = str(uuid.uuid1())

            document.SensitivityLabel.SetLabel(new_sl, new_sl)

            if config.get("save_document_after_setting_sensitivity_label"):
                document.Save()
        else:
            _log(f"   ➔ Sensitivity label already set. Skipping...", logging.DEBUG, ui_logger)

    except Exception as e:
        _log(f"⚠️ Could not set sensitivity label (Feature may be disabled): {e}", logging.WARNING, ui_logger)
        if did_something:
            logger.exception("Sensitivity label operation failed after modifying the label.")

    return document