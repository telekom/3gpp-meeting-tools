import os
import shutil
import logging
from pathlib import Path
from docx import Document

# --- CORRECTED IMPORTS ---
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
# -------------------------

from PyQt5.QtCore import QThread, pyqtSignal


class DocxSplitter:
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)

    def iter_block_items(self, parent):
        """Yields every paragraph and table sequentially to maintain XML order."""
        for child in parent.element.body:
            if isinstance(child, CT_P):
                yield child
            elif isinstance(child, CT_Tbl):
                yield child

    def _get_heading_level(self, text: str):
        """Determines the heading level based on the numbering scheme (e.g., '6.1.4')."""
        parts = text.split()[0].split('.')
        clean_parts = [p for p in parts if p.strip()]
        return len(clean_parts)

    def split(self, target_clause_prefix: str, split_depth: int, output_dir: str):
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        doc = Document(self.file_path)
        blocks = list(self.iter_block_items(doc))
        toc = []

        for i, block in enumerate(blocks):
            if isinstance(block, CT_P):
                from docx.text.paragraph import Paragraph
                para = Paragraph(block, doc)

                if para.style.name.startswith('Heading'):
                    text = para.text.strip()
                    if text.startswith(target_clause_prefix):
                        level = self._get_heading_level(text)

                        if level == split_depth:
                            safe_name = "".join(c for c in text if c.isalnum() or c in (' ', '.', '-')).strip()
                            toc.append({
                                'title': safe_name,
                                'start_idx': i,
                                'end_idx': len(blocks)
                            })

        if not toc:
            raise ValueError(
                f"Could not find any clauses matching prefix '{target_clause_prefix}' at depth {split_depth}.")

        for i in range(len(toc) - 1):
            toc[i]['end_idx'] = toc[i + 1]['start_idx']

        generated_files = []
        for section in toc:
            out_file = Path(output_dir) / f"{section['title']}.docx"
            shutil.copy(self.file_path, out_file)

            sub_doc = Document(out_file)
            sub_blocks = list(self.iter_block_items(sub_doc))

            # Delete everything BEFORE the start index
            for block in sub_blocks[:section['start_idx']]:
                block.getparent().remove(block)

            # Delete everything AFTER the end index
            for block in sub_blocks[section['end_idx']:]:
                block.getparent().remove(block)

            sub_doc.save(out_file)
            generated_files.append(out_file)

        return generated_files


class DocxSplitterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()

    def __init__(self, file_path: str, prefix: str, depth: int):
        super().__init__()
        self.file_path = file_path
        self.prefix = prefix
        self.depth = depth

    def run(self):
        try:
            # Force the output to a subfolder named after the document
            file_path = Path(self.file_path)
            output_dir = file_path.parent / f"{file_path.stem}_split"

            self.ui_log_msg.emit(f"⏳ Splitting document into subfolder: {output_dir.name}", 20)

            splitter = DocxSplitter(self.file_path)
            generated_files = splitter.split(self.prefix, self.depth, str(output_dir))

            self.ui_log_msg.emit(f"✅ Created {len(generated_files)} files in: {output_dir}", 20)
        except Exception as e:
            self.ui_log_msg.emit(f"❌ Splitter Error: {str(e)}", 40)
        finally:
            self.finished.emit()