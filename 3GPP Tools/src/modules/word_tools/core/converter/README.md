# `modules.word_tools.core.converter`

Modular Word document format converter package decomposed from the former `word_converter.py` monolith:

- [`com_utils.py`](com_utils.py): Word `WD_FORMAT_*` constants, NTFS attribute/Zone.Identifier cleanup, RPC disconnect detection, `faulthandler` shield, unattended Word COM configuration, and safe COM teardown.
- [`doc_to_docx.py`](doc_to_docx.py): Synchronous `.doc` → `.docx` conversion (Word COM + LibreOffice fallback) and isolated Word PDF export helper.
- [`worker.py`](worker.py): `WordConverterThread` (`QThread`) supporting Word COM export, legacy EMF/WMF PDF media repair, LibreOffice fallback, and OpenXML Markdown export.
