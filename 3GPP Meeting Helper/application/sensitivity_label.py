import traceback

import uuid
from enum import Enum
import platform

from config.word import WordConfig


class MsoAssignmentMethod(Enum):
    NOT_SET = -1  # The assignment method value is not set.
    STANDARD = 0  # The label is applied by default.
    PRIVILEGED = 1  # The label was manually selected.
    AUTO = 2  # The label is applied automatically.

    @staticmethod
    def from_int(value_int: int):
        try:
            assignment_method = MsoAssignmentMethod(value_int)
            return assignment_method
        except Exception as e:
            print(f'Value does not exist in Enum: {e}')
            return None


def set_sensitivity_label(document):
    if platform.system() != 'Windows':
        return None

    # Only do something if we are to do something
    if WordConfig.sensitivity_level_label_name is None or WordConfig.sensitivity_level_label_id is None:
        return document

    # Added code in case sensitivity labels are required. Meeting documents are public
    did_something = False
    try:
        sensitivity_label = document.SensitivityLabel.GetLabel()
        assignment_method = MsoAssignmentMethod.from_int(sensitivity_label.AssignmentMethod)
        print(f'Sensitivity level assignment method: {assignment_method}')
        # See https://docs.microsoft.com/en-us/office/vba/api/overview/library-reference/labelinfo-members-office
        if assignment_method == MsoAssignmentMethod.NOT_SET:
            # Need to set default sensitivity label. This part may vary depending on your organization
            print(f'Setting sensitivity label. Current assignment method: {sensitivity_label.AssignmentMethod}')
            did_something = True
            new_sl = document.SensitivityLabel.CreateLabelInfo()
            new_sl.ActionId = str(uuid.uuid1())
            new_sl.AssignmentMethod = 1
            new_sl.ContentBits = 0
            new_sl.IsEnabled = True
            new_sl.Justification = ''
            new_sl.LabelId = WordConfig.sensitivity_level_label_id
            new_sl.LabelName = WordConfig.sensitivity_level_label_name
            new_sl.SetDate = '2022-08-18T09:04:30Z'
            new_sl.SiteId = str(uuid.uuid1())
            document.SensitivityLabel.SetLabel(new_sl, new_sl)
            print(f'Set SensitivityLabel to {new_sl.LabelName}')

            if WordConfig.save_document_after_setting_sensitivity_label:
                document.Save()
                print(f'Saved document after setting sensitivity label')
        else:
            print('Not setting sensitivity label (already set)')
    except Exception as e:
        print(f'Could not get sensitivity label info. Probably this feature is not used by your installation: {e}')
        if did_something:
            traceback.print_exc()

    return document
