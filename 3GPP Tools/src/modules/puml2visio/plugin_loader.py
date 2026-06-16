from core.queue_manager import register_task


def register_plugin():
    """
    Called once at startup. Announces all available modules and their
    background threads to the core QueueManager.
    """
    # 1. Import your specific modules here
    from modules.puml2visio.core.visio_converter import ConverterThread, SvgConverterThread
    from modules.puml2visio.core.powerpoint_converter import PptxConverterThread
    from modules.puml2visio.core.word_extractor import WordExtractorThread
    from modules.puml2visio.core.docx_splitter import DocxSplitterThread
    from modules.puml2visio.core.word_comparator import WordComparatorThread
    from modules.puml2visio.core.ascii_converter import AsciiConverterThread

    # 2. Register PlantUML / Visio tasks (These need the jar_path from context)
    register_task(
        target_format="vsdx",
        display_name="To .VSDX",
        thread_factory=lambda f, p, ctx: ConverterThread(f, ctx.get('jar_path'))
    )
    register_task(
        target_format="svg",
        display_name="To .SVG",
        thread_factory=lambda f, p, ctx: SvgConverterThread(f, ctx.get('jar_path'))
    )
    register_task(
        target_format="pptx",
        display_name="To .PPTX",
        thread_factory=lambda f, p, ctx: PptxConverterThread(f, ctx.get('jar_path'))
    )

    # 3. Register Word tasks (These use the params dictionary)
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
        target_format="ascii",
        display_name="To .TXT",
        thread_factory=lambda f, p, ctx: AsciiConverterThread(f, ctx.get('jar_path'))
    )