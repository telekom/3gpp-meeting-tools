"""Heading models, deterministic source sorting, and python-docx heading scanner."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import List, Sequence

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph


@dataclass(frozen=True)
class HeadingInfo:
    id: str
    source_path: str
    source_index: int
    block_index: int
    end_index: int
    level: int
    text: str


def natural_key(path: str):
    name = Path(path).name.lower()
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", name)]


def order_sources(paths: Sequence[str], mode: str = "filename") -> List[str]:
    clean = [str(Path(p)) for p in paths if p]
    return sorted(clean, key=natural_key) if mode == "filename" else clean


class WordHeadingScanner:
    @staticmethod
    def _blocks(doc):
        return [x for x in doc.element.body if isinstance(x, (CT_P, CT_Tbl))]

    @staticmethod
    def _heading_level(para: Paragraph) -> int:
        style_name = ""
        try:
            style_name = para.style.name or ""
        except Exception:
            pass
        match = re.match(r"^Heading\s+(\d+)$", style_name, re.I)
        if match:
            return int(match.group(1))
        try:
            outline = para._p.pPr.outlineLvl
            if outline is not None:
                return int(outline.val) + 1
        except Exception:
            pass
        return 0

    def scan(
        self,
        paths: Sequence[str],
        order_mode: str = "filename",
        progress_callback=None,
    ) -> List[HeadingInfo]:
        ordered = order_sources(paths, order_mode)
        result: List[HeadingInfo] = []
        total_files = len(ordered)
        for source_index, path in enumerate(ordered):
            if progress_callback:
                progress_callback(source_index, total_files, 0, 1, path)
            doc = Document(path)
            blocks = self._blocks(doc)
            total_blocks = max(1, len(blocks))
            raw = []
            for block_index, block in enumerate(blocks):
                if progress_callback and (
                    block_index % 100 == 0 or block_index == total_blocks - 1
                ):
                    progress_callback(source_index, total_files, block_index + 1, total_blocks, path)
                if not isinstance(block, CT_P):
                    continue
                para = Paragraph(block, doc)
                text = para.text.strip()
                if not text:
                    continue
                level = self._heading_level(para)
                if level:
                    raw.append((block_index, level, text))
            for i, (block_index, level, text) in enumerate(raw):
                end_index = len(blocks)
                for next_block, next_level, _ in raw[i + 1 :]:
                    if next_level <= level:
                        end_index = next_block
                        break
                result.append(
                    HeadingInfo(
                        id=f"{source_index}:{block_index}",
                        source_path=path,
                        source_index=source_index,
                        block_index=block_index,
                        end_index=end_index,
                        level=level,
                        text=text,
                    )
                )
        return result
