import traceback

import uuid


def set_sensitivity_label(document):
    # Added code in case sensitivity labels are required. Meeting documents are public
    did_something = False
    try:
        sensitivity_label = document.SensitivityLabel.GetLabel()
        # See https://docs.microsoft.com/en-us/office/vba/api/overview/library-reference/labelinfo-members-office
        if sensitivity_label.AssignmentMethod == -1:
            # Need to set default sensitivity label. This part may vary depending on your organization
            print('Setting sensitivity label')
            did_something = True
            new_sl = document.SensitivityLabel.CreateLabelInfo()
            new_sl.ActionId = str(uuid.uuid1())
            new_sl.AssignmentMethod = 1
            new_sl.ContentBits = 0
            new_sl.IsEnabled = True
            new_sl.Justification = ''
            new_sl.LabelId = '55339bf0-f345-473a-9ec8-6ca7c8197055'
            new_sl.LabelName = 'OFFEN'
            new_sl.SetDate = '2022-08-18T09:04:30Z'
            new_sl.SiteId = str(uuid.uuid1())
            document.SensitivityLabel.SetLabel(new_sl, new_sl)
            print('Set SensitivityLabel to {0}'.format(new_sl.LabelName))
        else:
            print('Not setting sensitivity label (already set)')
    except:
        print('Could not get sensitivity label info. Probably this feature is not used by your installation')
        if did_something:
            traceback.print_exc()

    return document
