from core.queue_manager import register_task
from modules.word_tools.core.word_converter import WordConverterThread
from modules.word_tools.core.word_extractor import WordExtractorThread
from modules.word_tools.core.docx_splitter import DocxSplitterThread
from modules.word_tools.core.word_comparator import WordComparatorThread


def register_word_plugin():
    """Announces all Word-related tools to the QueueManager."""
    register_task(
        target_format="extract_visio",
        display_name="EXTRACT OLE",
        thread_factory=lambda f, p, ctx: WordExtractorThread(str(f))
    )
    register_task(
        target_format="split_docx",
        display_name="SPLIT CLAUSES",
        thread_factory=lambda f, p, ctx: DocxSplitterThread(str(f), p.get('prefix'), p.get('depth'))
    )
    register_task(
        target_format="compare_docx",
        display_name="COMPARE DOCS",
        thread_factory=lambda f, p, ctx: WordComparatorThread(p.get('doc_a'), p.get('doc_b'), p.get('keep_open'))
    )
    register_task(
        target_format="word_convert",
        display_name="Format Conversion",
        thread_factory=lambda file_path, params, ctx: WordConverterThread(
            str(file_path),
            params.get("fmt", "pdf"),
            engine=params.get("engine", "auto"),
            revisions=params.get("revisions", "markup"),
            extract_media=params.get("extract_media", True),
            output_dir=params.get("output_dir"),
            include_header=params.get("include_header", True),
        )
    )
    register_task(
        target_format="word_to_markdown",
        display_name="WORD → MARKDOWN",
        thread_factory=lambda file_path, params, ctx: WordConverterThread(
            str(file_path),
            "md",
            engine=params.get("engine", "auto"),
            revisions=params.get("revisions", "markup"),
            extract_media=params.get("extract_media", True),
            output_dir=params.get("output_dir"),
            include_header=params.get("include_header", True),
        )
    )