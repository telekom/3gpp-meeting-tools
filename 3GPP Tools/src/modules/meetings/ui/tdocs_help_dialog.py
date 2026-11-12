# --- File: src/modules/meetings/ui/tdocs_help_dialog.py ---
"""
TDocs Cockpit Quick Guide & Keyboard Shortcuts Dialog (`TDocsHelpDialog`).

Provides comprehensive in-app guidance for delegates and engineers:
- Toolbar actions, dropdown operations, and background caching tools
- Interactive table navigation, double-click gestures, and cell popup readers
- Advanced search syntax (quotes, negation, column filters)
- Full keyboard shortcuts catalog with keycap badges
"""
import html
import re
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont, QKeySequence
from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTabWidget,
    QTextBrowser,
    QFrame,
    QShortcut,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE,
)
from core.ui.theme import C


HELP_CSS = f"""
<style>
    body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        color: {C.TEXT_STRONG};
        background-color: {C.BG_CARD};
        font-size: 13px;
        line-height: 1.5;
        margin: 12px;
    }}
    h2 {{
        color: {C.TEXT_PRIMARY};
        font-size: 16px;
        margin-top: 14px;
        margin-bottom: 8px;
        border-bottom: 1.5px solid {C.BORDER_SUBTLE};
        padding-bottom: 4px;
    }}
    .lead {{
        color: {C.TEXT_SECONDARY};
        font-size: 13px;
        margin-bottom: 14px;
    }}
    .card {{
        background: {C.BG_SURFACE};
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 12px;
    }}
    .card-title {{
        font-weight: bold;
        font-size: 14px;
        color: {C.TEXT_PRIMARY};
        margin-bottom: 4px;
    }}
    .badge {{
        display: inline-block;
        padding: 2px 7px;
        font-size: 11px;
        font-weight: 600;
        border-radius: 4px;
        margin-right: 6px;
    }}
    .badge-blue {{ background-color: {C.INFO_SOFT_STRONG}; color: {C.INFO_DARK}; border: 1px solid {C.INFO_BORDER}; }}
    .badge-green {{ background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS_DARK}; border: 1px solid {C.SUCCESS_BORDER}; }}
    .badge-amber {{ background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_BORDER}; }}
    .badge-purple {{ background-color: {C.VIOLET_SOFT}; color: {C.VIOLET_DARK}; border: 1px solid {C.VIOLET_BORDER}; }}
    .badge-slate {{ background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; }}

    kbd {{
        display: inline-block;
        padding: 2px 6px;
        font-family: SFMono-Regular, Consolas, "Liberation Mono", Menlo, Courier, monospace;
        font-size: 11px;
        font-weight: 600;
        color: {C.TEXT_STRONG};
        background-color: {C.BG_CANVAS};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        box-shadow: 0 1px 0 rgba(0,0,0,0.1);
        margin: 0 2px;
    }}

    table.help-table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 8px;
        margin-bottom: 14px;
    }}
    table.help-table th {{
        background-color: {C.BG_CANVAS};
        color: {C.TEXT_BODY};
        font-size: 12px;
        font-weight: 600;
        text-align: left;
        padding: 8px 10px;
        border-bottom: 2px solid {C.BORDER_STRONG};
    }}
    table.help-table td {{
        padding: 8px 10px;
        border-bottom: 1px solid {C.BORDER_SUBTLE};
        font-size: 12.5px;
        vertical-align: top;
    }}
    table.help-table tr:hover td {{
        background-color: {C.BG_SURFACE};
    }}

    .tip-box {{
        background-color: {C.SUCCESS_SOFT};
        border-left: 4px solid #22C55E;
        padding: 8px 12px;
        border-radius: 0 6px 6px 0;
        margin-top: 6px;
        font-size: 12px;
        color: {C.SUCCESS_DARK};
    }}
    .note-box {{
        background-color: {C.INFO_SOFT};
        border-left: 4px solid {C.INFO_ACCENT};
        padding: 8px 12px;
        border-radius: 0 6px 6px 0;
        margin-top: 6px;
        font-size: 12px;
        color: {C.INFO_DARK};
    }}
    .warn-box {{
        background-color: {C.WARNING_SOFT};
        border-left: 4px solid {C.WARNING_ACCENT};
        padding: 8px 12px;
        border-radius: 0 6px 6px 0;
        margin-top: 6px;
        font-size: 12px;
        color: {C.WARNING_DARK};
    }}
</style>
"""


def _build_toolbar_tab_html() -> str:
    return HELP_CSS + """
<body>
    <h2>🎛️ Toolbar Actions & Operational Tools</h2>
    <div class="lead">
        The top ribbon provides quick access to synchronization, meeting resources, export briefings, and live TDoc tools.
    </div>

    <div class="card">
        <div class="card-title">🔄 Refresh ▾ (Meeting Synchronizer)</div>
        <p>Reloads meeting data from the official 3GPP FTP server or local cache files. Click the arrow ▾ for full options:</p>
        <table class="help-table">
            <tr>
                <th style="width: 32%;">Action</th>
                <th>Description & Usage</th>
            </tr>
            <tr>
                <td><b>📗 Refresh Excel List</b></td>
                <td>Downloads the latest official meeting Excel spreadsheet (<code>TDocs_List.xlsx</code>) from the 3GPP FTP and updates the table.</td>
            </tr>
            <tr>
                <td><b>📋 Download Agenda CSV</b></td>
                <td>Downloads <code>agenda.csv</code> to populate high-level Agenda Item titles, acronyms, and committee structure.</td>
            </tr>
            <tr>
                <td><b>📥 Cache All TDocs (Docs/)</b></td>
                <td>Batch-downloads all TDoc ZIP archives into your local cache directory in the background. Enables 100% offline work during flights or meetings.</td>
            </tr>
            <tr>
                <td><b>📥 Download Chairman's Notes</b><br><span class="badge badge-blue">All Working Groups</span></td>
                <td>Scans <code>/Inbox/Chair_Notes</code> on the 3GPP server and opens the <b>Multi-Chair Notes Picker</b>. Automatically pre-selects the newest file per chair/track (e.g. Andy, Hector, Ruicheng, Jian in SA2; EOM/Final in RAN) and indicates which files are already downloaded locally. Downloads files cleanly to <code>Agenda/ChairNotes/</code> with 1-click options to open in Microsoft Word.</td>
            </tr>
            <tr>
                <td><b>🔍 Scan Existing Chairman's Notes</b><br><span class="badge badge-amber">RAN Groups Only</span></td>
                <td>Scans locally saved Chairman's Notes in <code>Agenda/ChairNotes</code>, extracts normative text agreements && working assumptions, and updates table badges. (In SA/CT groups, agreements are tracked directly via formal pCR/CR contribution status rather than inline text blocks).</td>
            </tr>
            <tr>
                <td><b>📄 Import TdocsByAgenda.htm</b><br><span class="badge badge-blue">SA2 Only</span></td>
                <td>Fetches <code>TdocsByAgenda.htm</code> and merges live secretary remarks, revisions, and approval statuses into the table.</td>
            </tr>
            <tr>
                <td><b>📝 Refresh Revisions</b><br><span class="badge badge-purple">Electronic Meetings</span></td>
                <td>Quickly syncs latest electronic meeting revision metadata without re-downloading the entire Excel spreadsheet.</td>
            </tr>
        </table>
    </div>

    <div class="card">
        <div class="card-title">🗂️ Resources ▾ & 📗 Original Excel</div>
        <p><b>🗂️ Resources ▾:</b> 1-click shortcuts to open local meeting folders (<code>Docs/</code>, <code>Agenda/</code>, <code>ChairNotes/</code>), launch a Windows Terminal in the cache directory, view export debriefs, or open remote 3GPP FTP web folders.</p>
        <p><b>📗 Original Excel:</b> Opens the official unmodified 3GPP Excel spreadsheet in Microsoft Excel for external calculations or custom formatting.</p>
    </div>

    <div class="card">
        <div class="card-title">📤 Export ▾ (Reports & Dossiers) &nbsp;<kbd>Ctrl+Shift+E</kbd></div>
        <p>Generates rich Markdown and HTML dossiers for delegate reporting:</p>
        <ul>
            <li><b>Pre-Meeting Briefing:</b> Summary of submitted contributions, incoming LS, and draft CR allocations.</li>
            <li><b>Post-Meeting Debrief:</b> Complete inventory of agreed CRs/pCRs, noted papers, and meeting conclusions.</li>
            <li><b>Agreements Table [RAN]:</b> Structured table of all Chairman's Notes text agreements and working assumptions.</li>
            <li><b>Reveal Export Folder:</b> Shortcut <kbd>Ctrl+Shift+E</kbd> immediately opens the <code>output/</code> directory in Windows File Explorer.</li>
        </ul>
    </div>

    <div class="card">
        <div class="card-title">📅 Session Plan & ⚖️ Agreements (RAN)</div>
        <p><b>📅 Session Plan:</b> Interactive visual timeline and track allocation viewer. Parses daily schedule documents and displays parallel session tracks, room numbers, and agenda timings.</p>
        <p><b>⚖️ Agreements [RAN Meetings]:</b> Opens the dedicated <b>Meeting Agreements Console</b>. Filter by agreement kind (Agreements, Working Assumptions, FFS), search text, view full rationale, and copy citations.</p>
    </div>

    <div class="card">
        <div class="card-title">📝 Draft CR / pCR & 📊 Statistics</div>
        <p><b>📝 Draft CR / pCR:</b> Launches the integrated 3GPP Contribution Workbench. Directly generates official Change Request (CR) or Pseudo-CR (pCR) Word documents pre-filled with cover page headers, hanging indents, and revision markers.</p>
        <p><b>📊 Statistics:</b> Produces an interactive HTML dashboard illustrating company contribution share, document type distribution (CR, pCR, Discussion, LS), and agenda item heatmaps. Click <b>⚙️</b> to configure company alias mappings.</p>
    </div>

    <div class="card">
        <div class="card-title">🚀 Instant TDoc Fetch & ⚖️ vs Word Comparison</div>
        <p><b>🚀 Instant Fetch Input &nbsp;<kbd>Enter</kbd>:</b> Type any TDoc number (e.g. <code>R2-2601234</code> or <code>S2-2605678</code>) and press <kbd>Enter</kbd> to immediately download, unpack, and launch the Word document without scrolling through the table.</p>
        <p><b>⚖️ vs Word:</b> Downloads the TDoc specified in Instant Fetch and immediately runs a Microsoft Word Track Changes comparison against your currently active foreground Word document.</p>
    </div>
</body>
"""


def _build_table_tab_html() -> str:
    return HELP_CSS + f"""
<body>
    <h2>🖱️ Table Navigation & Interactive Cell Actions</h2>
    <div class="lead">
        The TDocs table is highly interactive. Many cells provide powerful double-click readers, 1-click filters, and contextual utilities.
    </div>

    <table class="help-table">
        <tr>
            <th style="width: 22%;">Column / Area</th>
            <th style="width: 28%;">Click / Interaction</th>
            <th>Result & Productivity Tip</th>
        </tr>
        <tr>
            <td><b>TDoc ID & Action Button</b></td>
            <td>
                • Click <b>📥 Download</b><br>
                • Click <b>✔ Open</b><br>
                • <b>Double-Click</b> cell
            </td>
            <td>
                • <b>Download:</b> Fetches ZIP from 3GPP FTP, extracts DOCX, and opens in Word.<br>
                • <b>Open:</b> Instantly opens locally cached document.<br>
                • <b>Double-Click:</b> Opens the <b>TDoc Information & Revision Tree Dialog</b>, showing full metadata, revision chain (previous / subsequent revisions), and full abstract.
            </td>
        </tr>
        <tr>
            <td><b>Title Column</b></td>
            <td><b>Double-Click</b> cell</td>
            <td>
                Opens the comfortable <b>Title & Abstract Reader</b> popup window. Ideal for reading long titles, abstracts, and contribution summaries without horizontal scrolling.
            </td>
        </tr>
        <tr>
            <td><b>Secretary Remarks / Comments</b></td>
            <td><b>Double-Click</b> cell</td>
            <td>
                • If the row contains an <b>Agreement</b>, double-clicking immediately opens the <b>Meeting Agreements Viewer</b> pinned to that exact agreement!<br>
                • Otherwise, opens the full <b>Remarks & Notes Reader</b> popup to view multi-line discussion transcripts and offline email conclusions.
            </td>
        </tr>
        <tr>
            <td><b>Source (Company)</b></td>
            <td><b>Double-Click</b> cell</td>
            <td>
                Instantly filters the entire meeting table to show all contributions from that specific company (resolving multinational company aliases automatically).
            </td>
        </tr>
        <tr>
            <td><b>Agenda Item (AI)</b></td>
            <td><b>Double-Click</b> cell</td>
            <td>
                Instantly filters the entire table to show all contributions under that specific Agenda Item number.
            </td>
        </tr>
        <tr>
            <td><b>Related TDocs</b></td>
            <td>Click link (e.g. <code>Rev to S2-26...</code>)</td>
            <td>
                Instantly navigates, scrolls, and highlights the referenced target revision directly in the table.
            </td>
        </tr>
        <tr>
            <td><b>Emails Column</b></td>
            <td>Click badge (e.g. <code>💬 4</code>)</td>
            <td>
                Opens the <b>Email Discussion Inspector</b> for this TDoc, showing full email thread history, delegate sender stances (Support / Object / Discuss), timestamps, and extracted quotes.
            </td>
        </tr>
        <tr>
            <td><b>Row Context Menu</b></td>
            <td><b>Right-Click</b> any row</td>
            <td>
                Context menu with options to:
                <ul>
                    <li>📋 Copy TDoc ID, Title, or Full Citation (<code>[TDoc] Title (Source)</code>)</li>
                    <li>🌐 Open on official 3GPP Portal or in local file folder</li>
                    <li>⭐ Add to Starred Watchlist / 🎯 Followed Work Items</li>
                    <li>🏷️ Set My Status (Approved, Objected, Discuss, etc.)</li>
                    <li>🤖 Run Local Ollama AI Technical Triage & Summary</li>
                </ul>
            </td>
        </tr>
    </table>

    <div class="note-box">
        💡 <b>Visual Color Highlights:</b>
        <br>• <span class="badge badge-green">Soft Green Tint ({C.SUCCESS_SOFT})</span> : Rows with approved text agreements or agreed contributions.
        <br>• <span class="badge badge-amber">Soft Amber Tint ({C.WARNING_SOFT_LIGHT})</span> : Rows with 3GPP Working Assumptions.
        <br>• <span class="badge badge-blue">Soft Blue Tint ({C.INFO_SOFT})</span> : Rows with FFS (For Further Study) or open issues.
    </div>
</body>
"""


def _build_search_tab_html() -> str:
    return HELP_CSS + """
<body>
    <h2>🔍 Search Bar, Filters & Query Syntax</h2>
    <div class="lead">
        Quickly locate relevant documents among thousands of meeting contributions using advanced search operators and multidimensional filters.
    </div>

    <div class="card">
        <div class="card-title">🔍 Text Search Operators</div>
        <p>The main search box searches across TDoc numbers, titles, contributing companies, and secretary remarks simultaneously:</p>
        <table class="help-table">
            <tr>
                <th style="width: 25%;">Syntax / Query</th>
                <th style="width: 35%;">Example</th>
                <th>Behavior</th>
            </tr>
            <tr>
                <td><b>Free Text</b></td>
                <td><code>mobility handover</code></td>
                <td>Matches rows containing both terms anywhere in TDoc, Title, Source, or Remarks.</td>
            </tr>
            <tr>
                <td><b>Exact Phrase</b></td>
                <td><code>"interim agreement"</code></td>
                <td>Wrap terms in double quotes to match the exact sequence of words.</td>
            </tr>
            <tr>
                <td><b>Negation (Exclude)</b></td>
                <td><code>discussion -draft</code></td>
                <td>Prefix a term with a minus sign (<code>-</code>) to exclude matching rows.</td>
            </tr>
            <tr>
                <td><b>Case-Insensitive</b></td>
                <td><code>r2-2601234</code></td>
                <td>Matches <code>R2-2601234</code> regardless of case.</td>
            </tr>
        </table>
    </div>

    <div class="card">
        <div class="card-title">🎯 Quick-Filter Toggle Chips</div>
        <table class="help-table">
            <tr>
                <th style="width: 15%;">Chip</th>
                <th style="width: 25%;">Name</th>
                <th>Description</th>
            </tr>
            <tr>
                <td><button style="font-size: 13px;">🎯</button></td>
                <td><b>Followed WIs</b></td>
                <td>Filters table to show only TDocs tagged with Work Items in your <b>Followed Work Items</b> watchlist. Right-click to manage your followed list.</td>
            </tr>
            <tr>
                <td><button style="font-size: 13px;">⭐</button></td>
                <td><b>Starred WIs</b></td>
                <td>Filters table to high-priority <b>Starred Work Items</b>. Right-click to toggle inclusion of unassigned incoming Liaison Statements (LS).</td>
            </tr>
            <tr>
                <td><button style="font-size: 13px;">⚖️</button><br><span class="badge badge-amber">RAN Only</span></td>
                <td><b>Agreements Only</b></td>
                <td>One-click filter to display only TDocs that contain Chairman's Notes text agreements or working assumptions.</td>
            </tr>
        </table>
    </div>

    <div class="card">
        <div class="card-title">🎚️ Multidimensional Dropdown Filters</div>
        <ul>
            <li><b>Company:</b> Filter by one or multiple contributing companies (supports alias resolution, e.g. "Qualcomm Inc" and "Qualcomm UK" grouped together).</li>
            <li><b>Type:</b> Filter by document type: <code>CR</code>, <code>pCR</code>, <code>discussion</code>, <code>report</code>, <code>LS in</code>, <code>LS out</code>, etc.</li>
            <li><b>AI (Agenda Item):</b> Filter by specific meeting agenda item section.</li>
            <li><b>Sub Item:</b> Filter by chairman/secretary sub-agenda dividers.</li>
            <li><b>WI (Work Item):</b> Filter by 3GPP Work Item acronym (e.g. <code>FS_Ambient_IoT</code>, <code>NR_pos_enh</code>).</li>
            <li><b>TDoc Status:</b> Filter by official outcome: <code>agreed</code>, <code>approved</code>, <code>noted</code>, <code>postponed</code>, <code>revised</code>, <code>withdrawn</code>.</li>
            <li><b>My Status:</b> Filter by your personal delegate workflow status (<code>To Read</code>, <code>Discuss</code>, <code>Object</code>, etc.).</li>
            <li><b>Order:</b> Toggle between <b>Agenda Order</b> (preserves live meeting progression and nests revisions under parents) or standard column sorting.</li>
        </ul>
    </div>
</body>
"""


def _build_shortcuts_tab_html() -> str:
    return HELP_CSS + """
<body>
    <h2>⌨️ Keyboard Shortcuts & Accelerators</h2>
    <div class="lead">
        Navigate and control the TDocs Cockpit efficiently using standard keyboard shortcuts.
    </div>

    <table class="help-table">
        <tr>
            <th style="width: 28%;">Shortcut</th>
            <th style="width: 32%;">Action</th>
            <th>Description & Scope</th>
        </tr>
        <tr>
            <td><kbd>F1</kbd></td>
            <td><b>Help & Shortcuts</b></td>
            <td>Opens this comprehensive in-app guide dialog from anywhere in the window.</td>
        </tr>
        <tr>
            <td><kbd>Ctrl</kbd> + <kbd>F</kbd></td>
            <td><b>Focus Search Bar</b></td>
            <td>Immediately jumps focus to the search input and highlights existing text for quick re-typing.</td>
        </tr>
        <tr>
            <td><kbd>Esc</kbd></td>
            <td><b>Clear / Close</b></td>
            <td>Clears the search input when focused; dismisses active dropdowns or closes modal dialogs.</td>
        </tr>
        <tr>
            <td><kbd>Enter</kbd> (in Instant Fetch)</td>
            <td><b>Launch TDoc</b></td>
            <td>Immediately fetches, unpacks, and launches the entered TDoc without touching the mouse.</td>
        </tr>
        <tr>
            <td><kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>E</kbd></td>
            <td><b>Reveal Export Folder</b></td>
            <td>Opens the meeting's Markdown / HTML export directory directly in Windows File Explorer.</td>
        </tr>
        <tr>
            <td><kbd>Ctrl</kbd> + <kbd>C</kbd></td>
            <td><b>Copy Citation</b></td>
            <td>Copies selected TDoc metadata or full citation (<code>[TDoc] Title (Company)</code>) to the clipboard.</td>
        </tr>
        <tr>
            <td><kbd>↑</kbd> / <kbd>↓</kbd></td>
            <td><b>Navigate Rows</b></td>
            <td>Smoothly moves selection up and down through the TDocs table.</td>
        </tr>
        <tr>
            <td><kbd>Page Up</kbd> / <kbd>Page Down</kbd></td>
            <td><b>Scroll Page</b></td>
            <td>Pages through long lists of meeting contributions quickly.</td>
        </tr>
        <tr>
            <td><kbd>Home</kbd> / <kbd>End</kbd></td>
            <td><b>Jump to Top / Bottom</b></td>
            <td>Instantly jumps to the first or last document in the table.</td>
        </tr>
    </table>

    <div class="tip-box">
        💡 <b>Pro-Tip:</b> You can also right-click any column header to sort in ascending or descending order, or use the <b>Order</b> dropdown to restore standard <b>Agenda Order</b> at any time!
    </div>
</body>
"""


class TDocsHelpDialog(QDialog):
    """
    Comprehensive in-app help and keyboard shortcuts dialog for the 3GPP TDocs Cockpit.
    """

    def __init__(self, parent=None, is_ran: bool = False):
        super().__init__(parent)
        self.is_ran = is_ran
        self.setWindowTitle("❓ 3GPP TDocs Cockpit — Quick Guide & Shortcuts")
        self.resize(920, 680)
        self.setMinimumSize(780, 520)

        # Allow pressing F1 or Esc to close/toggle
        QShortcut(QKeySequence("Esc"), self, self.close)
        QShortcut(QKeySequence("F1"), self, self.close)

        self._init_ui()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 14, 16, 14)
        main_layout.setSpacing(10)

        # Header Ribbon
        header_layout = QHBoxLayout()
        header_layout.setSpacing(12)

        icon_lbl = QLabel("❓")
        icon_lbl.setStyleSheet("font-size: 28px;")
        header_layout.addWidget(icon_lbl)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)
        dlg_title = QLabel("3GPP TDocs Cockpit — Quick Guide & Reference")
        dlg_title.setStyleSheet(f"font-size: 17px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        dlg_sub = QLabel("Operational tools, double-click actions, search operators, and keyboard shortcuts.")
        dlg_sub.setStyleSheet(f"font-size: 12px; color: {C.TEXT_MUTED};")
        title_vbox.addWidget(dlg_title)
        title_vbox.addWidget(dlg_sub)
        header_layout.addLayout(title_vbox)

        header_layout.addStretch()

        # In-dialog live filter
        self.search_filter = QLineEdit()
        self.search_filter.setPlaceholderText("🔍 Filter help topics...")
        self.search_filter.setStyleSheet(SEARCH_INPUT_STYLE)
        self.search_filter.setFixedWidth(220)
        self.search_filter.textChanged.connect(self._on_filter_changed)
        header_layout.addWidget(self.search_filter)

        main_layout.addLayout(header_layout)

        # Divider
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setFrameShadow(QFrame.Plain)
        divider.setStyleSheet(f"color: {C.BORDER_SUBTLE}; background-color: {C.BORDER_SUBTLE}; height: 1px;")
        main_layout.addWidget(divider)

        # Tab Widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                background-color: {C.BG_CARD};
                top: -1px;
            }}
            QTabBar::tab {{
                background: {C.BG_CANVAS};
                color: {C.TEXT_SECONDARY};
                padding: 8px 16px;
                margin-right: 4px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                border: 1px solid {C.BORDER_STRONG};
                border-bottom: none;
                font-size: 12.5px;
                font-weight: 500;
            }}
            QTabBar::tab:selected {{
                background: {C.BG_CARD};
                color: {C.TEXT_PRIMARY};
                font-weight: bold;
                border-bottom: 2px solid {C.ACCENT_ALT};
            }}
            QTabBar::tab:hover:!selected {{
                background: {C.BORDER_SUBTLE};
                color: {C.TEXT_STRONG};
            }}
        """)

        # Tab 1: Toolbar
        self.browser_toolbar = QTextBrowser()
        self.browser_toolbar.setOpenExternalLinks(False)
        self.browser_toolbar.setHtml(_build_toolbar_tab_html())
        self.tab_widget.addTab(self.browser_toolbar, "🎛️ Toolbar && Tools")

        # Tab 2: Table
        self.browser_table = QTextBrowser()
        self.browser_table.setOpenExternalLinks(False)
        self.browser_table.setHtml(_build_table_tab_html())
        self.tab_widget.addTab(self.browser_table, "🖱️ Table && Double-Clicks")

        # Tab 3: Search
        self.browser_search = QTextBrowser()
        self.browser_search.setOpenExternalLinks(False)
        self.browser_search.setHtml(_build_search_tab_html())
        self.tab_widget.addTab(self.browser_search, "🔍 Search && Filtering")

        # Tab 4: Shortcuts
        self.browser_shortcuts = QTextBrowser()
        self.browser_shortcuts.setOpenExternalLinks(False)
        self.browser_shortcuts.setHtml(_build_shortcuts_tab_html())
        self.tab_widget.addTab(self.browser_shortcuts, "⌨️ Keyboard Shortcuts")

        main_layout.addWidget(self.tab_widget)

        # Footer
        footer_layout = QHBoxLayout()
        tip_lbl = QLabel(f"💡 <b>Quick Access:</b> Press <kbd style='background:{C.BG_CANVAS}; padding:1px 5px; border-radius:3px; border:1px solid {C.BORDER_STRONG};'>F1</kbd> anytime in the TDocs window to open this guide.")
        tip_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_SECONDARY};")
        footer_layout.addWidget(tip_lbl)
        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        main_layout.addLayout(footer_layout)

    def _on_filter_changed(self, text: str):
        """Finds and highlights text in the active tab's browser."""
        query = text.strip()
        current_browser: QTextBrowser = self.tab_widget.currentWidget()
        if not current_browser:
            return

        if not query:
            # Clear selection by moving cursor
            cursor = current_browser.textCursor()
            cursor.clearSelection()
            current_browser.setTextCursor(cursor)
            return

        # Find first matching occurrence
        found = current_browser.find(query)
        if not found:
            # Wrap around from top
            cursor = current_browser.textCursor()
            cursor.movePosition(cursor.Start)
            current_browser.setTextCursor(cursor)
            current_browser.find(query)
