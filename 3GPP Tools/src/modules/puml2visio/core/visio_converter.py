import logging
import re
from pathlib import Path

import pythoncom
import win32com.client
import gc
from PyQt5.QtCore import QThread, pyqtSignal

from modules.puml2visio.utils.utils import strip_watermark, generate_cleaned_svg
from modules.puml2visio.config.paths import PLANTUML_WATERMARK


class VisioReaderThread(QThread):
    text_extracted = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, vsdx_path):
        super().__init__()
        self.vsdx_path = vsdx_path

    def run(self):
        pythoncom.CoInitialize()
        visio = None
        try:
            visio = win32com.client.DispatchEx("Visio.Application")
            visio.Visible = False
            visio.AlertResponse = 7

            doc = visio.Documents.OpenEx(str(Path(self.vsdx_path).resolve()), 2)
            source_code = ""
            for i in range(1, doc.Pages.Count + 1):
                page = doc.Pages(i)
                if page.Name == "PlantUML Source":
                    if page.Shapes.Count > 0:
                        raw_text = page.Shapes(1).Characters.Text
                        source_code = strip_watermark(raw_text)
                    break

            doc.Close()
            visio.Quit()

            if source_code:
                self.text_extracted.emit(source_code)
            else:
                self.error_occurred.emit("Could not find 'PlantUML Source' page in this Visio file.")
        except Exception as e:
            if visio: visio.Quit()
            self.error_occurred.emit(f"Error reading Visio file: {str(e)}")
        finally:
            # Destroy any remaining win32com wrappers while this
            # thread's COM apartment still exists.
            gc.collect()
            pythoncom.CoUninitialize()


class ConverterThread(QThread):
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        pythoncom.CoInitialize()
        try:
            self._emit_log(f"\n⚙️ Processing: {self.puml_path.name}", logging.INFO)

            svg_path = generate_cleaned_svg(self.puml_path, self.jar_path, self._emit_log)
            self._sanitize_svg_for_visio(svg_path)  # Clean SVG before Visio import
            self._convert_to_vsdx(svg_path)

            vsdx_path = self.puml_path.with_suffix(".vsdx")
            self.finished_path.emit(str(vsdx_path.resolve()))

            if svg_path.exists():
                svg_path.unlink()

        except Exception as e:
            self._emit_log(f"❌ Error: {str(e)}\n{'-' * 45}", logging.ERROR)
            self.finished_path.emit("")
        finally:
            pythoncom.CoUninitialize()

    def _sanitize_svg_for_visio(self, svg_path: Path) -> None:
        """Sanitizes SVG tags to prevent Visio import crashes."""
        if not svg_path.exists():
            return
        try:
            with open(svg_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()

            content = re.sub(r'(<tspan[^>]*?)\s+font-family="[^"]*"', r'\1', content)
            content = re.sub(r'(<tspan[^>]*?)\s+baseline-shift="[^"]*"', r'\1', content)
            content = re.sub(r'(<tspan[^>]*?)\s+dy="[^"]*"', r'\1', content)
            content = re.sub(r'(<tspan[^>]*?)\s+dx="[^"]*"', r'\1', content)

            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(content)
        except Exception as e:
            self._emit_log(f"⚠️ Could not sanitize SVG: {e}", logging.WARNING)

    def _emit_log(self, message: str, level: int):
        # QueueManager -> DragDropUI.log_message() owns Python logging.
        # Logging here as well causes every worker message to be logged twice.
        self.ui_log_msg.emit(message)

    def _convert_to_vsdx(self, svg_path: Path):
        vsdx_path = svg_path.with_suffix(".vsdx")
        if vsdx_path.exists():
            try:
                vsdx_path.unlink()
            except PermissionError:
                raise PermissionError("Close file in Visio first.")

        with open(self.puml_path, "r", encoding="utf-8") as f:
            raw_code = f.read()

        final_source_code = (
                PLANTUML_WATERMARK
                + "\n\n"
                + strip_watermark(raw_code)
        )

        visio = None
        doc = None

        def stage(message):
            self._emit_log(f"   🔎 {message}", logging.INFO)

        try:
            # ---------------------------------------------------------
            # STAGE 1: Start Visio
            # ---------------------------------------------------------
            stage("Starting Microsoft Visio COM server...")

            visio = win32com.client.DispatchEx("Visio.Application")
            visio.Visible = False
            visio.AlertResponse = 7

            stage("Visio started successfully.")

            # ---------------------------------------------------------
            # STAGE 2: Create blank document
            # ---------------------------------------------------------
            stage("Creating blank Visio document...")

            doc = visio.Documents.Add("")

            stage("Blank Visio document created.")

            # ---------------------------------------------------------
            # STAGE 3: Import SVG
            # ---------------------------------------------------------
            page = doc.Pages(1)
            page.Name = "Sequence Diagram"

            stage(f"Importing SVG: {svg_path.name}")

            try:
                page.Import(str(svg_path.resolve()))
            except Exception as e:
                raise RuntimeError(
                    "SVG_IMPORT_FAILED\n"
                    f"Visio rejected the generated SVG file:\n"
                    f"{svg_path.resolve()}\n\n"
                    f"Original COM error: {e}"
                ) from e

            stage(
                f"SVG imported successfully "
                f"({page.Shapes.Count} top-level shape(s))."
            )

            # ---------------------------------------------------------
            # STAGE 4: Flatten imported SVG
            # ---------------------------------------------------------
            if page.Shapes.Count > 0:
                stage("Reading imported SVG dimensions...")

                orig_w = page.Shapes(1).CellsU("Width").ResultIU
                orig_h = page.Shapes(1).CellsU("Height").ResultIU

                stage(
                    f"Imported SVG dimensions: "
                    f"{orig_w:.3f} × {orig_h:.3f} in."
                )

                stage("Ungrouping imported SVG shapes...")

                ungroup_count = 0
                peeling = True

                while peeling:
                    peeling = False

                    for i in range(page.Shapes.Count, 0, -1):
                        s = page.Shapes(i)

                        try:
                            if s.Type == 2:  # visTypeGroup
                                s.Ungroup()
                                ungroup_count += 1
                                peeling = True
                        except Exception as e:
                            self._emit_log(
                                f"   ⚠️ Could not ungroup shape {i}: {e}",
                                logging.WARNING
                            )

                stage(
                    f"Ungrouping completed "
                    f"({ungroup_count} group operation(s))."
                )

                # -----------------------------------------------------
                # STAGE 5: Remove PlantUML background
                # -----------------------------------------------------
                stage("Checking for PlantUML background rectangle...")

                deleted_backgrounds = 0

                for i in range(page.Shapes.Count, 0, -1):
                    s = page.Shapes(i)

                    try:
                        w = s.CellsU("Width").ResultIU
                        h = s.CellsU("Height").ResultIU

                        if (
                                w >= orig_w * 0.75
                                and h >= orig_h * 0.75
                        ):
                            if len(s.Characters.Text.strip()) == 0:
                                if (
                                        s.CellsU("LinePattern").ResultIU
                                        == 0
                                ):
                                    s.Delete()
                                    deleted_backgrounds += 1

                    except Exception:
                        # Some imported SVG shapes do not expose all
                        # ShapeSheet cells. They are intentionally
                        # ignored here.
                        pass

                stage(
                    f"Background cleanup completed "
                    f"({deleted_backgrounds} shape(s) removed)."
                )

                # -----------------------------------------------------
                # STAGE 6: Clean text shapes
                # -----------------------------------------------------
                stage("Normalizing imported text shapes...")

                text_shape_count = 0
                text_shape_errors = 0

                def clean_and_shrink_text(shapes):
                    nonlocal text_shape_count
                    nonlocal text_shape_errors

                    for i in range(1, shapes.Count + 1):
                        s = shapes(i)

                        try:
                            text = s.Characters.Text

                            if len(text.strip()) > 0:
                                text_shape_count += 1

                                s.CellsU(
                                    "TopMargin"
                                ).FormulaU = "0 pt"
                                s.CellsU(
                                    "BottomMargin"
                                ).FormulaU = "0 pt"
                                s.CellsU(
                                    "LeftMargin"
                                ).FormulaU = "0 pt"
                                s.CellsU(
                                    "RightMargin"
                                ).FormulaU = "0 pt"

                                line_pattern = (
                                    s.CellsU(
                                        "LinePattern"
                                    ).ResultIU
                                )

                                fill_pattern = (
                                    s.CellsU(
                                        "FillPattern"
                                    ).ResultIU
                                )

                                if (
                                        line_pattern == 0
                                        and fill_pattern == 0
                                ):
                                    pin_x = (
                                        s.CellsU(
                                            "PinX"
                                        ).ResultIU
                                    )

                                    pin_y = (
                                        s.CellsU(
                                            "PinY"
                                        ).ResultIU
                                    )

                                    loc_pin_x = (
                                        s.CellsU(
                                            "LocPinX"
                                        ).ResultIU
                                    )

                                    loc_pin_y = (
                                        s.CellsU(
                                            "LocPinY"
                                        ).ResultIU
                                    )

                                    h = (
                                        s.CellsU(
                                            "Height"
                                        ).ResultIU
                                    )

                                    left = pin_x - loc_pin_x
                                    top = pin_y + (
                                            h - loc_pin_y
                                    )

                                    s.CellsU(
                                        "LocPinX"
                                    ).FormulaU = "0 in"

                                    s.CellsU(
                                        "LocPinY"
                                    ).FormulaU = "Height"

                                    s.CellsU(
                                        "PinX"
                                    ).FormulaU = (
                                        f"{left} in"
                                    )

                                    s.CellsU(
                                        "PinY"
                                    ).FormulaU = (
                                        f"{top} in"
                                    )

                                    s.CellsU(
                                        "Width"
                                    ).FormulaU = (
                                        "TEXTWIDTH(TheText)"
                                    )

                                    s.CellsU(
                                        "Height"
                                    ).FormulaU = (
                                        "TEXTHEIGHT("
                                        "TheText, Width)"
                                    )

                        except Exception:
                            text_shape_errors += 1

                        try:
                            if s.Type == 2:
                                clean_and_shrink_text(
                                    s.Shapes
                                )
                        except Exception:
                            pass

                clean_and_shrink_text(page.Shapes)

                stage(
                    f"Text normalization completed "
                    f"({text_shape_count} text shape(s), "
                    f"{text_shape_errors} skipped/error shape(s))."
                )

                # -----------------------------------------------------
                # STAGE 7: Resize page
                # -----------------------------------------------------
                stage("Resizing Visio page to contents...")

                page_sheet = page.PageSheet

                page_sheet.CellsU(
                    "PageLeftMargin"
                ).FormulaU = "0.05 in"

                page_sheet.CellsU(
                    "PageRightMargin"
                ).FormulaU = "0.05 in"

                page_sheet.CellsU(
                    "PageTopMargin"
                ).FormulaU = "0.05 in"

                page_sheet.CellsU(
                    "PageBottomMargin"
                ).FormulaU = "0.05 in"

                try:
                    page.ResizeToFitContents()
                    stage("Page resized successfully.")
                except Exception as e:
                    self._emit_log(
                        f"   ⚠️ ResizeToFitContents failed: {e}",
                        logging.WARNING
                    )

            # ---------------------------------------------------------
            # STAGE 8: Add PlantUML source page
            # ---------------------------------------------------------
            stage("Creating embedded PlantUML source page...")

            src_page = doc.Pages.Add()

            src_page.PageSheet.CellsU(
                "PageWidth"
            ).FormulaU = "8.27 in"

            src_page.PageSheet.CellsU(
                "PageHeight"
            ).FormulaU = "11.69 in"

            src_page.Name = "PlantUML Source"

            text_box = src_page.DrawRectangle(
                0.5,
                0.5,
                7.77,
                11.19
            )

            text_box.CellsU(
                "LinePattern"
            ).FormulaU = "0"

            text_box.CellsU(
                "FillPattern"
            ).FormulaU = "0"

            text_box.CellsU(
                "Para.HorzAlign"
            ).FormulaU = "0"

            text_box.CellsU(
                "VerticalAlign"
            ).FormulaU = "0"

            text_box.Characters.Text = final_source_code

            stage("PlantUML source page created successfully.")

            # ---------------------------------------------------------
            # STAGE 9: Select diagram page
            # ---------------------------------------------------------
            try:
                if visio.ActiveWindow:
                    visio.ActiveWindow.Page = page
            except Exception:
                pass

            # ---------------------------------------------------------
            # STAGE 10: Save VSDX
            # ---------------------------------------------------------
            stage(f"Saving Visio document: {vsdx_path.name}")

            try:
                doc.SaveAs(str(vsdx_path.resolve()))
            except Exception as e:
                raise RuntimeError(
                    "VSDX_SAVE_FAILED\n"
                    f"Visio successfully imported and processed the "
                    f"SVG, but failed while saving:\n"
                    f"{vsdx_path.resolve()}\n\n"
                    f"Original COM error: {e}"
                ) from e

            stage("Visio document saved successfully.")

            # ---------------------------------------------------------
            # Explicitly release child COM proxies BEFORE closing the
            # document / terminating the Visio COM server.
            #
            # win32com proxies can otherwise outlive Visio.Quit().
            # Accessing/releasing those disconnected proxies during
            # Python cleanup can produce RPC_E_DISCONNECTED (0x80010108)
            # and, in some pywin32/Office combinations, a fatal process
            # exception rather than a normal Python exception.
            # ---------------------------------------------------------

            try:
                text_box = None
            except Exception:
                pass

            try:
                src_page = None
            except Exception:
                pass

            try:
                page_sheet = None
            except Exception:
                pass

            try:
                s = None
            except Exception:
                pass

            try:
                page = None
            except Exception:
                pass

            # Close the document while the Visio application is alive.
            doc.Close()
            doc = None

            # Now terminate Visio.
            visio.Quit()
            visio = None

            # Encourage immediate destruction of any remaining Python
            # COM wrappers while COM is still initialized on this thread.
            gc.collect()

            self.ui_log_msg.emit(
                f"✅ Saved: {vsdx_path.name}"
            )

        except Exception as e:
            # Keep the generated SVG after a failure. This is
            # deliberate because it gives us the exact artifact
            # Visio attempted to import.

            if doc is not None:
                try:
                    doc.Close()
                except Exception:
                    pass

            if visio is not None:
                try:
                    visio.Quit()
                except Exception:
                    pass

            raise RuntimeError(
                f"Visio COM Error: {e}"
            ) from e


class SvgConverterThread(QThread):
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        try:
            self.ui_log_msg.emit(f"\n⚙️ Generating SVG for: {self.puml_path.name}")

            svg_path = generate_cleaned_svg(self.puml_path, self.jar_path, self.ui_log_msg.emit)

            if svg_path.exists():
                self.ui_log_msg.emit(f"✅ Success! SVG saved: {svg_path.name}\n{'-' * 45}")
                self.finished_path.emit(str(svg_path.resolve()))
            else:
                self.ui_log_msg.emit("❌ Error: PlantUML finished but SVG was not created.")
                self.finished_path.emit("")
        except Exception as e:
            self.ui_log_msg.emit(f"❌ Error: {str(e)}\n{'-' * 45}")
            self.finished_path.emit("")