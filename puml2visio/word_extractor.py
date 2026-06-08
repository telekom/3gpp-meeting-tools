import zipfile
import io
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal


class WordExtractorThread(QThread):
    ui_log_msg = pyqtSignal(str)

    def __init__(self, docx_path: str):
        super().__init__()
        self.docx_path = Path(docx_path)

    def run(self):
        self.ui_log_msg.emit(f"\n📄 Analyzing Word Document: {self.docx_path.name}...")
        output_dir = self.docx_path.parent
        extracted_files = []

        try:
            with zipfile.ZipFile(self.docx_path, 'r') as z:
                direct_vsdx = [f for f in z.namelist() if f.endswith('.vsdx')]
                for f in direct_vsdx:
                    data = z.read(f)
                    out_name = output_dir / f"{self.docx_path.stem}_{Path(f).name}"
                    with open(out_name, 'wb') as out:
                        out.write(data)
                    extracted_files.append(out_name)
                    self.ui_log_msg.emit(f"✅ Extracted native Visio object: {out_name.name}")

                bins = [f for f in z.namelist() if f.startswith('word/embeddings/') and f.endswith('.bin')]
                for i, emb in enumerate(bins):
                    data = z.read(emb)
                    start_idx = data.find(b'PK\x03\x04')
                    if start_idx != -1:
                        vsdx_data = data[start_idx:]
                        try:
                            with zipfile.ZipFile(io.BytesIO(vsdx_data)) as test_z:
                                if 'visio/document.xml' in test_z.namelist() or '[Content_Types].xml' in test_z.namelist():
                                    out_name = output_dir / f"{self.docx_path.stem}_embedded_{i + 1}.vsdx"
                                    counter = 1
                                    while out_name.exists():
                                        out_name = output_dir / f"{self.docx_path.stem}_embedded_{i + 1}_{counter}.vsdx"
                                        counter += 1
                                    with open(out_name, 'wb') as out:
                                        out.write(vsdx_data)
                                    extracted_files.append(out_name)
                                    self.ui_log_msg.emit(f"✅ Extracted OLE Visio object: {out_name.name}")
                        except zipfile.BadZipFile:
                            pass

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Error reading Word file: {e}")

        if not extracted_files:
            self.ui_log_msg.emit("⚠️ No embedded Visio files found in this document.")
        else:
            self.ui_log_msg.emit(
                f"🎉 Successfully extracted {len(extracted_files)} Visio file(s) to the document's folder!")