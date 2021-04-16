import tkinter
from tkinter import ttk
import application
import gui.main
import pyperclip

class TdocsTable:

    def __init__(self, parent, favicon):
        top = self.top = tkinter.Toplevel(parent)
        top.title("TDoc Table")
        top.iconbitmap(favicon)

        # https://stackoverflow.com/questions/50625306/what-is-the-best-way-to-show-data-in-a-table-in-tkinter
        self.tree = ttk.Treeview(
            top,
            columns=('TDoc', 'AI', 'Type', 'Title', 'Source', 'Revs', 'Emails', 'Subject', 'Result'),
            show='headings',
            selectmode="browse")

        def set_column(col, label=None, width=None, sort=True, center=True):
            if label is None:
                label = col
            if sort:
                self.tree.heading(col, text=label, command=lambda: self.treeview_sort_column(col, False))
            else:
                self.tree.heading(col, text=label)
            if width is not None:
                self.tree.column(col, minwidth=0, width=width, stretch=False)
            if center:
                self.tree.column(col, anchor="center")

        set_column('TDoc', "TDoc #", width=110)
        set_column('AI', width=50)
        set_column('Type', width=120)
        set_column('Title', center=False)
        set_column('Source', center=False)
        set_column('Revs', width=50)
        set_column('Emails', width=50)
        set_column('Subject', width=50, sort=False)
        set_column('Result', width=100)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.insert_rows(application.current_tdocs_by_agenda.tdocs)

        self.tree_scroll = ttk.Scrollbar(top)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        # tree.grid(row=0, column=0)

        # Can also do this: https://stackoverflow.com/questions/33781047/tkinter-drop-down-list-of-check-boxes-combo-boxes
        self.combo_ai = ttk.Combobox(top, values=list(application.current_tdocs_by_agenda.tdocs["AI"].unique()), state="readonly")
        self.combo_ai.pack()
        self.combo_ai.bind("<<ComboboxSelected>>", self.select_ai)
        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side='right', fill='y')

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def insert_rows(self, df):
        for idx, row in df.iterrows():
            self.tree.insert("", "end", values=(idx, row['AI'], row['Type'], row['Title'], row['Source'], '', '', 'X', row['Result']))

    def select_ai(self, event=None):
        self.tree.delete(*self.tree.get_children())
        tdocs_for_ai = application.current_tdocs_by_agenda.tdocs
        tdocs_for_ai = tdocs_for_ai[tdocs_for_ai['AI'] == self.combo_ai.get()]
        self.insert_rows(tdocs_for_ai)

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:])-1 # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except:
            actual_value = None
        tdoc_id = item_values[0]
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))
        if column == 0:
            print('Opening {0}'.format(actual_value))
            gui.main.download_and_open_tdoc(actual_value)
        if column == 5:
            print('Opening revisions for {0}'.format(tdoc_id))
            gui.tdocs_table.RevisionsTable(gui.main.root, gui.main.favicon, tdoc_id)
        if column == 6:
            print('Opening emails for {0}'.format(tdoc_id))
            gui.tdocs_table.EmailsTable(gui.main.root, gui.main.favicon, tdoc_id)
        if column == 7:
            print('Generating subject for email approval for {0}. Copied to clipboard'.format(tdoc_id))
            subject = '[SA2#{3}, AI#{1}, {0}] {2}'.format(tdoc_id, item_values[1], item_values[3], application.current_tdocs_by_agenda.meeting_number)
            print(subject)
            pyperclip.copy(subject)

    def treeview_sort_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)

        # reverse sort next time
        self.tree.heading(col, command=lambda: self.treeview_sort_column(col, not reverse))

class RevisionsTable:

    def __init__(self, parent, favicon, tdoc_id):
        top = self.top = tkinter.Toplevel(parent)
        top.title("Revisions for {0}".format(tdoc_id))
        top.iconbitmap(favicon)


class EmailsTable:

    def __init__(self, parent, favicon, tdoc_id):
        top = self.top = tkinter.Toplevel(parent)
        top.title("Emails for {0}".format(tdoc_id))
        top.iconbitmap(favicon)