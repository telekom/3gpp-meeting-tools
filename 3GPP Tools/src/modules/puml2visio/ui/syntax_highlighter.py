# --- File: src/modules/puml2visio/ui/syntax_highlighter.py ---
"""PlantUML Syntax Highlighter for PyQt5 QPlainTextEdit/QTextDocument.

Provides high-contrast, theme-consistent syntax highlighting for PlantUML diagrams:
- Keywords, lifelines, control fragments (alt/opt/loop/group)
- Preprocessors (@startuml, skinparam, autonumber)
- Participants, actors, entities
- Arrows, messages, and connectors
- Multi-line (/ ' ... ' /) and single-line (' ...) comments
- Quoted strings and hex color codes
"""

from PyQt5.QtCore import QRegularExpression, Qt
from PyQt5.QtGui import QColor, QFont, QSyntaxHighlighter, QTextCharFormat


class PlantUMLHighlighter(QSyntaxHighlighter):
    """Syntax highlighter for PlantUML source code in CodeDropTextEdit."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlighting_rules = []

        # --- Formats Configuration ---
        def create_format(color: str, bold: bool = False, italic: bool = False) -> QTextCharFormat:
            fmt = QTextCharFormat()
            fmt.setForeground(QColor(color))
            if bold:
                fmt.setFontWeight(QFont.Bold)
            if italic:
                fmt.setFontItalic(True)
            return fmt

        fmt_preprocessor = create_format("#7C3AED", bold=True)      # Vibrant Purple
        fmt_keyword = create_format("#1E5C99", bold=True)           # Application Primary Blue
        fmt_declaration = create_format("#0F766E", bold=True)       # Dark Teal
        fmt_arrow = create_format("#D97706", bold=True)             # Dark Amber/Orange
        fmt_string = create_format("#15803D")                       # Forest Green
        fmt_color_tag = create_format("#0284C7", bold=True)         # Sky Blue
        fmt_divider = create_format("#475569", bold=True)           # Slate
        fmt_comment = create_format("#94A3B8", italic=True)         # Muted Slate Gray

        # 1. Preprocessor & Directives (@startuml, @enduml, !pragma, skinparam, etc.)
        self.highlighting_rules.append((
            QRegularExpression(r"^\s*(@\w+|!\w+|skinparam\b|hide\b|show\b|autonumber\b|title\b|header\b|footer\b|legend\b|endlegend\b|scale\b)"),
            fmt_preprocessor
        ))

        # 2. Control flow & Lifeline keywords
        keywords = [
            r"\balt\b", r"\belse\b", r"\bopt\b", r"\bloop\b", r"\bpar\b",
            r"\bcritical\b", r"\bbreak\b", r"\bgroup\b", r"\bbox\b", r"\bend\b",
            r"\bactivate\b", r"\bdeactivate\b", r"\bdestroy\b", r"\bcreate\b",
            r"\breturn\b", r"\bref\b", r"\bover\b", r"\bleft\b", r"\bright\b",
            r"\bof\b", r"\bnote\b", r"\bhnote\b", r"\brnote\b", r"\bacross\b",
            r"\bnewpage\b", r"\bpartition\b", r"\bfork\b", r"\bagain\b",
            r"\bsplit\b", r"\bif\b", r"\bthen\b", r"\belse\b", r"\bendif\b",
            r"\bwhile\b", r"\bendwhile\b", r"\brepeat\b"
        ]
        for kw in keywords:
            self.highlighting_rules.append((QRegularExpression(kw), fmt_keyword))

        # 3. Declarations (participant, actor, class, state, etc.)
        declarations = [
            r"\bparticipant\b", r"\bactor\b", r"\bboundary\b", r"\bcontrol\b",
            r"\bentity\b", r"\bdatabase\b", r"\bcollections\b", r"\bqueue\b",
            r"\bclass\b", r"\binterface\b", r"\benum\b", r"\bstate\b",
            r"\bobject\b", r"\bpackage\b", r"\bnode\b", r"\bcomponent\b",
            r"\bartifact\b", r"\bcloud\b", r"\brectangle\b", r"\brobust\b",
            r"\bconcise\b", r"\bbinary\b", r"\bnwdiag\b", r"\bnetwork\b",
            r"\bpacketdiag\b", r"\brackdiag\b", r"\bas\b"
        ]
        for decl in declarations:
            self.highlighting_rules.append((QRegularExpression(decl), fmt_declaration))

        # 4. Dividers & Delays (== Title == or ... delay ...)
        self.highlighting_rules.append((
            QRegularExpression(r"^\s*==.*==\s*$"),
            fmt_divider
        ))
        self.highlighting_rules.append((
            QRegularExpression(r"^\s*\.\.\..*\.\.\.\s*$"),
            fmt_divider
        ))

        # 5. Arrows & Connectors (e.g. ->, -->, <-, <--, <->, ->x, -[#red]->)
        self.highlighting_rules.append((
            QRegularExpression(r"[-=~.]+\[#[A-Za-z0-9_]+\][-=~.>x*o\\\/]+|[-=~.]*<[-=~.]+[-=~.>x*o\\\/]*|[-=~.]+[-=~.>x*o\\\/]+"),
            fmt_arrow
        ))

        # 6. Stereotypes (<<...>>) and Hex Colors (#FFFFFF or #Red)
        self.highlighting_rules.append((
            QRegularExpression(r"<<[^>]+>>|#[0-9A-Fa-f]{3,8}\b|#[A-Za-z]+\b"),
            fmt_color_tag
        ))

        # 7. Strings ("...")
        self.highlighting_rules.append((
            QRegularExpression(r'"[^"\\]*(\\.[^"\\]*)*"'),
            fmt_string
        ))

        # 8. Single-line comments (' ...)
        self.highlighting_rules.append((
            QRegularExpression(r"'.*$"),
            fmt_comment
        ))

        # Multi-line comment state format
        self.multi_line_comment_format = fmt_comment
        self.comment_start_expression = QRegularExpression(r"/''*")
        self.comment_end_expression = QRegularExpression(r"''*/")

    def highlightBlock(self, text: str):
        """Applies syntax highlighting rules to the given text block."""
        for pattern, fmt in self.highlighting_rules:
            match_iterator = pattern.globalMatch(text)
            while match_iterator.hasNext():
                match = match_iterator.next()
                self.setFormat(match.capturedStart(), match.capturedLength(), fmt)

        # Multi-line block comments (/' ... '/)
        self.setCurrentBlockState(0)
        start_index = 0
        if self.previousBlockState() != 1:
            start_match = self.comment_start_expression.match(text)
            start_index = start_match.capturedStart() if start_match.hasMatch() else -1

        while start_index >= 0:
            end_match = self.comment_end_expression.match(text, start_index)
            end_index = end_match.capturedStart() if end_match.hasMatch() else -1
            comment_length = 0
            if end_index == -1:
                self.setCurrentBlockState(1)
                comment_length = len(text) - start_index
            else:
                comment_length = end_index - start_index + end_match.capturedLength()

            self.setFormat(start_index, comment_length, self.multi_line_comment_format)
            next_match = self.comment_start_expression.match(text, start_index + comment_length)
            start_index = next_match.capturedStart() if next_match.hasMatch() else -1
