import tkinter
from tkinter import ttk
from tkinter.ttk import Treeview
from typing import List


def set_column(tree: Treeview, col: str, label: str = None, width=None, sort=True, center=True):
    """
    Sets a Treeview's column
    Args:
        tree: The Treeview to which to apply this column to
        col: The column name
        label: Label for the column (if any)
        width: Set column width (if any)
        sort: Whether to sort the column (asc)
        center:  Whether to certer the column
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
    tree.tag_configure('odd', background='#E8E8E8')
    tree.tag_configure('even', background='#DFDFDF')


class GenericTable:

    def __init__(self, parent, title: str, favicon, column_names: List[str]):
        """
        Base class for table GUIs in this application
        Args:
            parent: The caller GUI (e.g. tools dialog)
            title: The title of this GUI. Will appear at the top of the GUI
            favicon: Icon to show in the top-left corner of this GUI
        """
        self.style_name = 'mystyle.Treeview'
        self.style = None
        self.init_style()
        self.top = tkinter.Toplevel(parent)
        self.top.title(title)
        self.top.iconbitmap(favicon)
        self.favicon = favicon

        self.top_frame = tkinter.Frame(self.top)
        self.top_frame.pack(anchor='w')
        self.main_frame = tkinter.Frame(self.top)
        self.main_frame.pack()
        self.bottom_frame = tkinter.Frame(self.top)
        self.bottom_frame.pack(anchor='w')

        self.column_names = column_names

        # https://stackoverflow.com/questions/50625306/what-is-the-best-way-to-show-data-in-a-table-in-tkinter
        self.tree = ttk.Treeview(
            self.main_frame,
            columns=tuple(column_names),
            show='headings',
            selectmode="browse",
            style=self.style_name,
            padding=[-5, -25, -5, -25])  # Left, top, right, bottom

        self.tree_scroll = ttk.Scrollbar(self.main_frame)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        # tree.grid(row=0, column=0)

    # See https://bugs.python.org/issue36468
    def fixed_map(self, option):
        # Fix for setting text colour for Tkinter 8.6.9
        # From: https://core.tcl.tk/tk/info/509cafafae
        #
        # Returns the style map for 'option' with any styles starting with
        # ('!disabled', '!selected', ...) filtered out.

        # style.map() returns an empty list for missing options, so this
        # should be future-safe.
        return [elm for elm in self.style.map(self.style_name, query_opt=option) if
                elm[:2] != ('!disabled', '!selected')]

    def init_style(self):
        if self.style is None:
            self.style = ttk.Style()
            self.style.map(
                self.style_name,
                foreground=self.fixed_map('foreground'),
                background=self.fixed_map('background'))
            self.style.configure(
                self.style_name,
                highlightthickness=0,
                bd=0,
                rowheight=60)
            # Modify the font of the headings
            # style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))
            self.style.layout(self.style_name,
                              [(self.style_name + '.treearea', {'sticky': 'nswe'})])  # Remove the borders
