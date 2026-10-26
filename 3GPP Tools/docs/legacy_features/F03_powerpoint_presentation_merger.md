# F-03: PowerPoint Status Report & Presentation Merger

## 1. Executive Summary & Objective

In 3GPP TSG Plenaries (SA, RAN, CT) and Working Groups, delegates must digest dozens of status report presentations, study item wrap-ups, and liaison slide sets submitted as Microsoft PowerPoint presentations (`.pptx` or `.ppt`). 

During plenary sessions, reviewing 30 to 50 individual presentations is cumbersome and inefficient:
* Each presentation must be separately downloaded, unzipped, and opened in a new PowerPoint window.
* Delegates lose their place across tens of open PowerPoint application windows.
* Each slide deck contains a redundant title/cover slide and end slide that clutters presentation reviews.

The objective of **F-03** is to provide a 1-click PowerPoint merging engine that filters the meeting TDoc list for all PowerPoint contributions (or takes user-selected presentations), downloads and unzips them, strips redundant cover slides, and compiles them into a single consolidated master presentation (`Status_Report_Merged.pptx`) organized with labeled PowerPoint section headers (`SectionProperties.AddBeforeSlide`) named after each TDoc ID and title.

---

## 2. Legacy Codebase Analysis

### Files & Locations
* `3GPP Meeting Helper/application/powerpoint.py` (functions: `get_powerpoint`, `new_presentation`, `merge_presentations`)
* `3GPP Meeting Helper/gui/tdocs_table.py` (method: `merge_pptx_files`)

### Legacy Implementation Details

```python
# application/powerpoint.py
def merge_presentations(list_of_ppts, list_of_section_labels=None, headlines_for_toc=None):
    try:
        new_slideset = new_presentation()
        slides = new_slideset.Slides
        section_properties = new_slideset.SectionProperties
        inserts = []
        for idx, slideset_path in enumerate(list_of_ppts):
            file_name = os.path.splitext(slideset_path)[0]
            insert_position = slides.Count

            # Skip first slide (cover slide) by setting SlideStart=2
            slides.InsertFromFile(slideset_path, insert_position, SlideStart=2)

            if list_of_section_labels is not None:
                section_name = '{0}, {1}'.format(list_of_section_labels[idx], headlines_for_toc[idx])
                section_properties.AddBeforeSlide(insert_position + 1, section_name)

            inserts.append((insert_position + 1, section_name))

        print('Merged {0} slidesets'.format(len(inserts)))
    except Exception as e:
        print(f'Could not merge presentations: {e}')
```

#### Meeting TDocs UI Trigger (`gui/tdocs_table.py`)
```python
def merge_pptx_files(self, *args):
    tdoc_list_to_merge = list(self.current_tdocs.index)
    all_extracted_files = []
    all_titles = []
    for tdoc_id in tdoc_list_to_merge:
        extracted_files = self.download_and_open_tdoc(tdoc_id, skip_opening=True)
        if extracted_files is not None:
            all_extracted_files.extend(extracted_files)
            all_titles.append(self.current_tdocs.at[tdoc_id, 'Title'])

    # Filter only PowerPoint files
    all_extracted_files = [e for e in all_extracted_files if '.ppt' in e.lower()]
    powerpoint.merge_presentations(
        all_extracted_files,
        list_of_section_labels=tdoc_list_to_merge,
        headlines_for_toc=all_titles)
```

### Strengths & Limitations of Legacy Implementation
* **Strengths:** 
  * Native PowerPoint COM preserves vector graphics, animations, embedded objects, and slide layouts without formatting loss.
  * Section dividers allowed delegates to jump between reports using PowerPoint's built-in Slide Sorter and Navigation pane.
* **Limitations:**
  * Synchronous execution on the main UI thread froze Tkinter during downloads and COM merges.
  * Required PowerPoint desktop installation on Windows.
  * Did not handle `.ppt` legacy files that required conversion before insertion.

---

## 3. Current State in 3GPP Tools

* **PlantUML to Visio/PowerPoint:** `src/modules/puml2visio/` has extensive Office shape export for PlantUML diagrams.
* **Missing Feature:** There is no tool or menu action in `src/modules/meetings/` to merge multiple downloaded presentation slide decks into one master deck.

---

## 4. Modern Architecture & Design Specification

### Architectural Flow

```mermaid
sequenceDiagram
    autonumber
    actor Delegate as Standards Delegate
    participant UI as TDocs Window
    participant Worker as PPTXMergerThread (QThread)
    participant PreDownloader as TDocFileHandler
    participant COM as PowerPoint.Application (Win32 COM)
    participant Disk as Merged File (.pptx)

    Delegate->>UI: Select TDocs or Filter "Type == PPT/Status Report"
    UI->>Delegate: Prompt dialog: Options (Skip cover slides?, Add section dividers?, Include TOC?)
    Delegate->>UI: Clicks "Merge Presentations"
    UI->>Worker: Start PPTXMergerThread(tdoc_ids, options)
    Worker->>PreDownloader: Parallel download & unpack candidate .zip files
    PreDownloader-->>Worker: List of local .pptx / .ppt files
    Worker->>COM: Initialize COM (pythoncom.CoInitialize())
    Worker->>COM: Presentations.Add()
    loop For each presentation
        Worker->>COM: Slides.InsertFromFile(path, insert_pos, SlideStart=skip_cover ? 2 : 1)
        Worker->>COM: SectionProperties.AddBeforeSlide(pos, "[TDoc] Title")
        Worker->>UI: progress.emit(f"Merging {idx}/{total}...")
    end
    alt Generate Table of Contents
        Worker->>COM: Add custom TOC slide with hyperlinked slide targets
    end
    Worker->>COM: SaveAs(target_path)
    Worker->>COM: Visible = True / Activate
    Worker-->>UI: finished(True, target_path)
    UI->>Delegate: Toast notification & auto-focus PowerPoint window
```

### Proposed Component Details

#### 1. Background Worker (`src/modules/meetings/core/pptx_merger.py`)
```python
from PyQt5.QtCore import QThread, pyqtSignal
from pathlib import Path
import platform
import logging

class PPTXMergerThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, items: list, output_path: Path, skip_cover: bool = True, parent=None):
        super().__init__(parent)
        self.items = items  # list of tuples: (tdoc_id, title, file_path)
        self.output_path = output_path
        self.skip_cover = skip_cover

    def run(self):
        if platform.system() != 'Windows':
            self.finished.emit(False, "PowerPoint merging requires Windows and Microsoft Office.")
            return

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Visible = True
            presentation = ppt.Presentations.Add()
            slides = presentation.Slides
            section_props = presentation.SectionProperties

            total = len(self.items)
            for idx, (tdoc_id, title, file_path) in enumerate(self.items, 1):
                if self.isInterruptionRequested():
                    return
                self.progress.emit(f"Adding ({idx}/{total}) {tdoc_id}...")
                insert_pos = slides.Count
                start_slide = 2 if self.skip_cover else 1
                
                # Insert slides
                slides.InsertFromFile(str(file_path), insert_pos, SlideStart=start_slide)
                
                # Add section title
                clean_title = (title[:60] + "...") if len(title) > 60 else title
                sec_name = f"{tdoc_id} · {clean_title}"
                section_props.AddBeforeSlide(insert_pos + 1, sec_name)

            presentation.SaveAs(str(self.output_path))
            self.finished.emit(True, str(self.output_path))
        except Exception as e:
            logging.error(f"Failed to merge presentations: {e}", exc_info=True)
            self.finished.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()
```

#### 2. UI Entry Points
* **TDocs Window Action Menu:** `📊 Reports ▾` -> `"Merge PowerPoint Presentations..."`
* **Selection Context Menu:** Right-click selected rows -> `"Merge Selected Presentations into One Deck"`

---

## 5. Implementation Roadmap & Acceptance Criteria

### Step-by-Step Implementation Tasks
1. [ ] Implement `PPTXMergerThread` in `src/modules/meetings/core/pptx_merger.py` with `pythoncom` initialization and error handling.
2. [ ] Add options dialog `PPTXMergeDialog` allowing delegates to choose:
   - Skip cover slides (`SlideStart=2`)
   - Generate Table of Contents slide
   - Destination file path
3. [ ] Hook action into `tdocs_menus.py` and `tdocs_window.py`.
4. [ ] Handle legacy `.ppt` archives by auto-converting or logging clean warnings.

### Acceptance Criteria
* [x] Merges 20+ presentations into a single master deck with zero corruption.
* [x] Correctly adds PowerPoint section dividers labeled with TDoc ID and title.
* [x] Runs asynchronously in a background thread without freezing the PyQt5 UI.
* [x] Automatically opens the merged presentation in Microsoft PowerPoint upon completion.
