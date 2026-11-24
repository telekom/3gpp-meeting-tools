"""Single-document Word COM sensitivity label application."""

from __future__ import annotations

import logging
import platform
import uuid
from datetime import datetime, timezone
from typing import Any, Optional

from modules.word_tools.core.word_config import WordConfig
from .inspector import MsoAssignmentMethod, _log

logger = logging.getLogger(__name__)


def set_sensitivity_label(
    document: Any,
    label_id: Optional[str] = None,
    label_name: Optional[str] = None,
    ui_logger=None,
    return_changed: bool = False,
    force_overwrite: bool = False,
):
    """
    Applies a corporate sensitivity label to a Word COM Document object to bypass manual save popups.

    Args:
        document: Microsoft Word COM Document instance.
        label_id: Optional GUID of the label. If None, loaded from active WordConfig.
        label_name: Optional display name of the label. If None, loaded from active WordConfig.
        ui_logger: Optional UI logging target.
        return_changed: If True, returns boolean indicating whether the label was modified.
        force_overwrite: If True, overwrites an existing label even if one was already set.
    """
    if platform.system() != "Windows":
        return False if return_changed else document

    config = WordConfig.load()

    # The Gatekeeper check
    if not config.get("enable_sensitivity_labels", False):
        _log("   ➔ Sensitivity labels disabled in config. Skipping...", logging.DEBUG, ui_logger)
        return False if return_changed else document

    # Resolve active label if not explicitly provided
    active_label = WordConfig.get_active_sensitivity_label()
    if not label_id or not label_name:
        label_id = label_id or active_label.get("id") or config.get("sensitivity_level_label_id")
        label_name = label_name or active_label.get("name") or config.get("sensitivity_level_label_name")

    if not label_id or not label_name:
        _log(
            "   ➔ No sensitivity label configured or detected in user policy. Skipping...",
            logging.DEBUG,
            ui_logger,
        )
        return False if return_changed else document

    # Resolve TenantId / SiteId for Purview compliance
    site_id = ""
    for lbl in config.get("sensitivity_labels", []):
        if isinstance(lbl, dict) and (lbl.get("id") or "").lower() == str(label_id).lower():
            site_id = lbl.get("tenant_id", "")
            break
    site_id = (
        site_id
        or active_label.get("tenant_id")
        or config.get("sensitivity_tenant_id")
        or str(uuid.uuid1())
    )

    did_something = False
    try:
        # Check if the document actually supports sensitivity labels
        if getattr(document, "SensitivityLabel", None) is None:
            return False if return_changed else document

        sensitivity_label = document.SensitivityLabel.GetLabel()
        assignment_method = MsoAssignmentMethod.from_int(sensitivity_label.AssignmentMethod)

        needs_setting = assignment_method == MsoAssignmentMethod.NOT_SET
        if not needs_setting and force_overwrite:
            current_id = getattr(sensitivity_label, "LabelId", "")
            if current_id != label_id:
                needs_setting = True

        if needs_setting:
            _log(
                f"   ➔ Applying corporate Sensitivity Label: {label_name} ({label_id})...",
                logging.INFO,
                ui_logger,
            )

            new_sl = document.SensitivityLabel.CreateLabelInfo()
            new_sl.ActionId = str(uuid.uuid1())
            new_sl.AssignmentMethod = 1  # Privileged
            new_sl.ContentBits = 0
            new_sl.IsEnabled = True
            new_sl.Justification = ""
            new_sl.LabelId = label_id
            new_sl.LabelName = label_name
            # Dynamically generate current UTC time
            new_sl.SetDate = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            new_sl.SiteId = site_id

            document.SensitivityLabel.SetLabel(new_sl, new_sl)
            did_something = True

            if config.get("save_document_after_setting_sensitivity_label"):
                document.Save()
        else:
            _log(
                f"   ➔ Sensitivity label already set ({getattr(sensitivity_label, 'LabelName', 'Classified')}). Skipping...",
                logging.DEBUG,
                ui_logger,
            )

    except Exception as e:
        _log(
            f"⚠️ Could not set sensitivity label (Feature may be disabled): {e}",
            logging.WARNING,
            ui_logger,
        )
        if did_something:
            logger.exception("Sensitivity label operation failed after modifying the label.")

    return did_something if return_changed else document
