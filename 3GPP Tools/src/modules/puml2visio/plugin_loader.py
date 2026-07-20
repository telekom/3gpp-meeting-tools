from core.queue_manager import register_task
from modules.puml2visio.core.pptx2visio_converter import PptxToVisioConverterThread


def register_puml2visio_plugin():
    """Announces PlantUML and Visio module tools to the QueueManager."""
    from modules.puml2visio.core.visio_converter import ConverterThread, SvgConverterThread
    from modules.puml2visio.core.powerpoint_converter import PptxConverterThread
    from modules.puml2visio.core.ascii_converter import AsciiConverterThread

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
    register_task(
        target_format="ascii",
        display_name="To .TXT",
        thread_factory=lambda f, p, ctx: AsciiConverterThread(f, ctx.get('jar_path'))
    )
    register_task(
        target_format="pptx_to_visio",
        display_name="PowerPoint to Visio",
        # Notice we don't need the jar_path for this specific thread, so we only pass the file path 'f'
        thread_factory=lambda f, p, ctx: PptxToVisioConverterThread(f)
    )