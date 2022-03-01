import win32com.client

# Global Word instance
word = None

def get_word():
    global word
    if word is not None:
        return word

    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except:
        try:
            word = win32com.client.Dispatch("Word.Application")
        except:
            word = None
    if word is not None:
        try:
            word.Visible = True
        except:
            print('Could not set property "Visible" from Word to "True"')
        try:
            word.DisplayAlerts = False
        except:
            print('Could not set property "DisplayAlerts" from Word to "False"')
    return word


def open_word_document(filename='', set_as_active_document=True):
    if (filename is None) or (filename == ''):
        doc = get_word().Documents.Add()
    else:
        doc = get_word().Documents.Open(filename)
    if set_as_active_document:
        get_word().Activate()
        doc.Activate()
    return doc


