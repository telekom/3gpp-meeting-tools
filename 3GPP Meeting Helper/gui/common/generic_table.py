import tkinter
from tkinter import ttk
from tkinter.ttk import Treeview
from typing import List

from application.tkinter_config import table_text_color, table_bg_color_odd, table_bg_color_even
from gui.common.tkinter_widget import TkWidget
from gui.common.utils import get_new_style

column_separator_str = "   "

def set_column(
        tree: Treeview,
        col: str,
        label: str = None,
        width=None,
        sort=True,
        center=True):
    """
    Sets a Treeview's column
    Args:
        tree: The Treeview to which to apply this column to
        col: The column name
        label: Label for the column (if any)
        width: Set column width (if any)
        sort: Whether to sort the column (asc)
        center:  Whether to center the column
    """

    if label is None:
        label = col
    if sort:
        tree.heading(col, text=label, command=lambda: treeview_sort_column(tree, col, False))
    else:
        tree.heading(col, text=label)
    if width is not None:
        tree.column(col, minwidth=0, width=width, stretch=False)
    if center:
        tree.column(col, anchor="center")


def treeview_sort_column(tree: Treeview, col, reverse=False):
    l = [(tree.set(k, col), k) for k in tree.get_children('')]
    l.sort(reverse=reverse)

    # rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tree.move(k, '', index)

    # reverse sort next time
    tree.heading(col, command=lambda: treeview_sort_column(tree, col, not reverse))


def treeview_set_row_formatting(tree: Treeview):

    tree.tag_configure(
        'odd',
        background=table_bg_color_odd,
        foreground=table_text_color
    )
    tree.tag_configure(
        'even',
        background=table_bg_color_even,
        foreground=table_text_color
    )


class GenericTable(TkWidget):

    def __init__(
            self,
            parent_widget: tkinter.Tk | None,
            widget_title: str,
            favicon,
            column_names: List[str],
            row_height=55,
            display_rows=10,
            root_widget: tkinter.Tk | None = None,
            treeview_show=('headings',)
    ):
        """
        Base class for table GUIs in this application
        Args:
            root_widget: The root widget
            display_rows: Number of rows to display in widget
            row_height: Row height for each row in the widget
            parent_widget: The caller GUI (e.g. tools dialog)
            widget_title: The title of this GUI. Will appear at the top of the GUI
            favicon: Icon to show in the top-left corner of this GUI. File path
        """

        super().__init__(
            root_widget=root_widget,
            parent_widget=parent_widget,
            widget_title=widget_title,
            favicon=favicon
        )

        self.style_name = f'mystyle.Treeview.{self.class_type}'
        self.style = None
        self.init_style(row_height=row_height)

        self.top_frame = tkinter.Frame(self.tk_top)
        self.main_frame = tkinter.Frame(self.tk_top)
        self.bottom_frame = tkinter.Frame(self.tk_top)

        # https://stackoverflow.com/questions/42074654/avoid-the-status-bar-footer-from-disappearing-in-a-gui-when-reducing-the-size
        self.top_frame.pack(side=tkinter.TOP, fill=tkinter.X)
        self.bottom_frame.pack(side=tkinter.BOTTOM, fill=tkinter.X)
        self.main_frame.pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=tkinter.YES)

        self.column_names = column_names

        # https://stackoverflow.com/questions/50625306/what-is-the-best-way-to-show-data-in-a-table-in-tkinter
        self.tree = ttk.Treeview(
            self.main_frame,
            columns=tuple(column_names),
            show=treeview_show,
            selectmode="browse",
            style=self.style_name,
            padding=[-5, -25, -5, -25],
            height=display_rows
        )  # Left, top, right, bottom

        self.tree_scroll = ttk.Scrollbar(self.main_frame)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        # tree.grid(row=0, column=0)

    def init_style(self, row_height):
        if self.style is None:
            self.style = get_new_style(self.style_name)
            self.style.configure(
                self.style_name,
                highlightthickness=0,
                bd=0,
                rowheight=row_height)
            # Modify the font of the headings
            # style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))
            self.style.layout(self.style_name,
                              [(self.style_name + '.treearea', {'sticky': 'nswe'})])  # Remove the borders

    def clear_tree(self):
        if self.tree is not None:
            self.tree.delete(*self.tree.get_children())

    def set_column(self, col: str, label: str = None, width=None, sort=True, center=True):
        set_column(self.tree, col=col, label=label, width=width, sort=sort, center=center)
