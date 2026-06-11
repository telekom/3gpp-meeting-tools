import os
import shutil
import logging
import re
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
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

    def _prune_unused_media(self, doc):
        """
        Garbage Collector: Scans the surviving XML for active Relationship IDs (rId).
        Deletes any hidden media/OLE relationships that are no longer used in the clause.
        """
        xml_str = doc.element.body.xml
        # Find every relationship ID currently active in the text body
        used_rids = set(re.findall(r'r:id="([^"]+)"', xml_str))

        rels = doc.part.rels
        for rId in list(rels.keys()):
            if rId not in used_rids:
                rel = rels[rId]
                # Aggressively unlink heavy media types. (We leave styles/fonts alone).
                if "image" in rel.reltype or "oleObject" in rel.reltype or "package" in rel.reltype:
                    del rels[rId]

    def _process_section(self, section, output_dir, progress_callback):
        """The isolated task run by the parallel ThreadPool."""
        out_file = Path(output_dir) / f"{section['title']}.docx"
        shutil.copy(self.file_path, out_file)

        sub_doc = Document(out_file)
        sub_blocks = list(self.iter_block_items(sub_doc))

        # Delete everything BEFORE the target clause
        for block in sub_blocks[:section['start_idx']]:
            block.getparent().remove(block)

        # Delete everything AFTER the target clause
        for block in sub_blocks[section['end_idx']:]:
            block.getparent().remove(block)

        # Run the Garbage Collector to remove orphaned Visio/Image bloat!
        self._prune_unused_media(sub_doc)

        sub_doc.save(out_file)

        if progress_callback:
            progress_callback(section['title'])

        return out_file

    def split(self, target_clause_prefix: str, split_depth: int, output_dir: str, progress_callback=None):
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        doc = Document(self.file_path)
        blocks = list(self.iter_block_items(doc))
        toc = []

        # Build the Map
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

        # Determine Boundaries
        for i in range(len(toc) - 1):
            toc[i]['end_idx'] = toc[i + 1]['start_idx']

        generated_files = []

        # Calculate a safe thread limit to prevent RAM spikes (Max 3 threads)
        safe_threads = min(3, os.cpu_count() or 1)

        # Run the Subtractive Slicing in a THROTTLED Parallel Pool!
        with ThreadPoolExecutor(max_workers=safe_threads) as executor:
            futures = []
            for section in toc:
                futures.append(executor.submit(self._process_section, section, output_dir, progress_callback))

            for future in as_completed(futures):
                generated_files.append(future.result())

        return generated_files


class DocxSplitterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()

    def __init__(self, file_path: str, prefix: str, depth: int):
        super().__init__()
        self.file_path = file_path
        self.prefix = prefix
        self.depth = depth
        self.completed_count = 0

    def _on_section_complete(self, section_title: str):
        """Thread-safe callback triggered every time a parallel worker finishes a file."""
        self.completed_count += 1
        self.ui_log_msg.emit(f"   ↳ [{self.completed_count}] Extracted: {section_title}", logging.INFO)

    def run(self):
        try:
            self.ui_log_msg.emit(
                f"⏳ Initiating parallel document slicing (Prefix: '{self.prefix}', Depth: {self.depth})...",
                logging.INFO)
            output_dir = Path(self.file_path).parent / f"{Path(self.file_path).stem}_split"

            splitter = DocxSplitter(self.file_path)

            # Pass our callback directly into the engine
            generated_files = splitter.split(self.prefix, self.depth, str(output_dir), self._on_section_complete)

            self.ui_log_msg.emit(
                f"✅ Successfully split document into {len(generated_files)} optimized files at:\n{output_dir}",
                logging.INFO)
        except Exception as e:
            self.ui_log_msg.emit(f"❌ Splitter Error: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()