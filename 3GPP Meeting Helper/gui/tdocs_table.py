import tkinter
from tkinter import ttk
import application
import gui.main
import gui.tools
import pyperclip
import re
import textwrap
import webbrowser
import gui.main
from parsing.html_revisions import revisions_file_to_dataframe
import traceback
import pandas as pd

style_name = 'mystyle.Treeview'

# See https://bugs.python.org/issue36468
def fixed_map(option):
    # Fix for setting text colour for Tkinter 8.6.9
    # From: https://core.tcl.tk/tk/info/509cafafae
    #
    # Returns the style map for 'option' with any styles starting with
    # ('!disabled', '!selected', ...) filtered out.

    # style.map() returns an empty list for missing options, so this
    # should be future-safe.
    return [elm for elm in style.map(style_name, query_opt=option) if
            elm[:2] != ('!disabled', '!selected')]

style = None

def init_style():
    global style
    if style is None:
        style = ttk.Style()
        style.map(style_name, foreground=fixed_map('foreground'),
                  background=fixed_map('background'))
        style.configure(style_name, highlightthickness=0, bd=0, rowheight=60)
        # style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))  # Modify the font of the headings
        style.layout(style_name, [(style_name + '.treearea', {'sticky': 'nswe'})])  # Remove the borders


def set_column(tree, col, label=None, width=None, sort=True, center=True):
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


def treeview_sort_column(tree, col, reverse=False):
    l = [(tree.set(k, col), k) for k in tree.get_children('')]
    l.sort(reverse=reverse)

    # rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tree.move(k, '', index)

    # reverse sort next time
    tree.heading(col, command=lambda: treeview_sort_column(tree, col, not reverse))

class TdocsTable:

    current_tdocs = None
    source_width = 200
    title_width = 550

    def __init__(self, parent, favicon, parent_gui_tools):
        init_style()
        top = self.top = tkinter.Toplevel(parent)
        top.title("TDoc Table")
        top.iconbitmap(favicon)
        self.parent_gui_tools = parent_gui_tools

        frame_1 = tkinter.Frame(top)
        frame_1.pack(anchor='w')
        frame_2 = tkinter.Frame(top)
        frame_2.pack()
        frame_3 = tkinter.Frame(top)
        frame_3.pack(anchor='w')

        self.tdoc_count = tkinter.StringVar()
        self.revisions_list = None
        self.revisions = None

        # https://stackoverflow.com/questions/50625306/what-is-the-best-way-to-show-data-in-a-table-in-tkinter
        self.tree = ttk.Treeview(
            frame_2,
            columns=('TDoc', 'AI', 'Type', 'Title', 'Source', 'Revs', 'Emails', 'Send @', 'Result'),
            show='headings',
            selectmode="browse",
            style=style_name)

        set_column(self.tree, 'TDoc', "TDoc #", width=110)
        set_column(self.tree, 'AI', width=50)
        set_column(self.tree, 'Type', width=120)
        set_column(self.tree, 'Title', width=TdocsTable.title_width, center=False)
        set_column(self.tree, 'Source', width=TdocsTable.source_width, center=False)
        set_column(self.tree, 'Revs', width=50)
        set_column(self.tree, 'Emails', width=50)
        set_column(self.tree, 'Send @', width=50, sort=False)
        set_column(self.tree, 'Result', width=100)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.load_data(reload=True)
        self.reload_revisions = False
        self.insert_current_tdocs()

        self.tree_scroll = ttk.Scrollbar(frame_2)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        # tree.grid(row=0, column=0)

        # Can also do this: https://stackoverflow.com/questions/33781047/tkinter-drop-down-list-of-check-boxes-combo-boxes
        self.search_text = tkinter.StringVar()
        self.search_text
        self.search_entry = tkinter.Entry(frame_1, textvariable=self.search_text, width=25, font='TkDefaultFont')
        self.search_text.trace_add(['write', 'unset'], self.select_text)

        tkinter.Label(frame_1, text="Search: ").pack(side=tkinter.LEFT)
        self.search_entry.pack(side=tkinter.LEFT)

        all_ais = ['All']
        all_ais.extend(list(application.current_tdocs_by_agenda.tdocs["AI"].unique()))
        self.combo_ai = ttk.Combobox(frame_1, values=all_ais, state="readonly")
        self.combo_ai.set('All')
        self.combo_ai.bind("<<ComboboxSelected>>", self.select_ai)

        tkinter.Label(frame_1, text="  Select AI: ").pack(side=tkinter.LEFT)
        self.combo_ai.pack(side=tkinter.LEFT)

        tkinter.Label(frame_1, text="  ").pack(side=tkinter.LEFT)
        tkinter.Button(
            frame_1,
            text='Clear filters',
            command=self.clear_filters).pack(side=tkinter.LEFT)
        tkinter.Button(
            frame_1,
            text='Reload data',
            command=self.reload_data).pack(side=tkinter.LEFT)

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        tkinter.Label(frame_3, textvariable=self.tdoc_count).pack(side=tkinter.LEFT)

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def load_data(self, reload=False):
        if reload:
            print('Loading revision data for table')
            current_selection = gui.main.tkvar_meeting.get()
            meeting_server_folder = application.sa2_meeting_data.get_server_folder_for_meeting_choice(current_selection)
            tdocs_by_agenda_file, revisions_file = gui.main.get_tdocs_by_agenda_for_selected_meeting(
                meeting_server_folder,
                return_revisions_file=True)
            self.revisions, self.revisions_list = revisions_file_to_dataframe(revisions_file, self.current_tdocs)
        self.meeting_number = application.current_tdocs_by_agenda.meeting_number
        self.current_tdocs = application.current_tdocs_by_agenda.tdocs

    def insert_current_tdocs(self):
        self.insert_rows(self.current_tdocs)

    def insert_rows(self, df):
        count = 0

        for idx, row in df.iterrows():
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            if self.revisions is None:
                revision_count = ''
            else:
                number_format = '{0:02d}'
                try:
                    rev_number = self.revisions.loc[idx, 'Revisions']
                    if rev_number < 1:
                        revision_count = ''
                    else:
                        revision_count = number_format.format(rev_number)
                except KeyError:
                    # Not found
                    revision_count = '' # Zero is left empty
                    pass
                except:
                    revision_count = '' # Error is left empty
                    traceback.print_exc()

            self.tree.insert("", "end", tags=(tag,), values=(
                idx,
                row['AI'],
                row['Type'],
                textwrap.fill(row['Title'], width=70),
                textwrap.fill(row['Source'], width=25),
                revision_count,
                '',
                'Click',
                row['Result']))

        self.tree.tag_configure('odd', background='#E8E8E8')
        self.tree.tag_configure('even', background='#DFDFDF')
        self.tdoc_count.set('{0} documents'.format(count))


    def clear_filters(self, *args):
        self.combo_ai.set('All')
        self.search_text.set('')
        self.load_data(reload=False)
        self.select_ai() # One will call the other


    def reload_data(self, *args):
        self.load_data(reload=True)
        self.select_ai() # One will call the other

    def select_ai(self, load_data=True, event=None):
        if load_data:
            self.load_data()

        tdocs_for_ai = self.current_tdocs
        selected_ai = self.combo_ai.get()
        print('Filtering by AI "{0}"'.format(selected_ai))
        if selected_ai == 'All':
            tdocs_for_ai = tdocs_for_ai
        else:
            tdocs_for_ai = tdocs_for_ai[tdocs_for_ai['AI'] == self.combo_ai.get()]

        self.current_tdocs = tdocs_for_ai

        self.tree.delete(*self.tree.get_children())
        self.insert_current_tdocs()

        if load_data:
            self.select_text(load_data=False)

    def select_text(self, load_data=True, *args):
        if load_data:
            self.load_data()

        # Filter based on current TDocs
        text_search = self.search_text.get()
        if text_search is None or text_search == '':
            return

        try:
            re.compile(text_search)
            is_regex = True
            print('Filtering by Regex "{0}"'.format(text_search))
        except re.error:
            is_regex = False
            print('Filtering by Text "{0}"'.format(text_search))

        text_search = text_search.lower()
        tdocs_for_text = self.current_tdocs.copy()
        tdocs_for_text['search_column'] = tdocs_for_text.index + tdocs_for_text['Title']
        tdocs_for_text['search_column'] = tdocs_for_text['search_column'].str.lower()
        tdocs_for_text = tdocs_for_text[tdocs_for_text['search_column'].str.contains(text_search, regex=is_regex)]
        self.current_tdocs = tdocs_for_text

        self.tree.delete(*self.tree.get_children())
        self.insert_current_tdocs()

        if load_data:
            self.select_ai(load_data=False)

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
        if actual_value is None or actual_value == '':
            print("Empty value")
            return
        if column == 0:
            print('Opening {0}'.format(actual_value))
            gui.main.download_and_open_tdoc(actual_value)
        if column == 5:
            print('Opening revisions for {0}'.format(tdoc_id))
            gui.tdocs_table.RevisionsTable(gui.main.root, gui.main.favicon, tdoc_id, self.revisions_list, self.parent_gui_tools)
        if column == 6:
            print('Opening emails for {0}'.format(tdoc_id))
            gui.tdocs_table.EmailsTable(gui.main.root, gui.main.favicon, tdoc_id)
        if column == 7:
            print('Generating subject for email approval for {0}. Copying to clipboard and generating empty email'.format(tdoc_id))
            subject = '[SA2#{3}, AI#{1}, {0}] {2}'.format(tdoc_id, item_values[1], item_values[3], self.meeting_number)
            subject = subject.replace('\n', ' ').replace('  ', ' ')
            print(subject)
            webbrowser.open('mailto:{0}?subject={1}'.format('3GPP_TSG_SA_WG2_EMEET@LIST.ETSI.ORG', subject), new = 1)
            pyperclip.copy(subject)

class RevisionsTable:

    def __init__(self, parent, favicon, tdoc_id, revisions_df, parent_gui_tools):
        top = self.top = tkinter.Toplevel(parent)
        top.title("Revisions for {0}".format(tdoc_id))
        top.iconbitmap(favicon)
        revisions = revisions_df.loc[tdoc_id, :]
        self.tdoc_id = tdoc_id
        self.parent_gui_tools = parent_gui_tools
        print('{0} Revisions'.format(len(revisions)))

        frame_1 = tkinter.Frame(top)
        frame_1.pack()
        frame_2 = tkinter.Frame(top)
        frame_2.pack(anchor='w')
        frame_3 = tkinter.Frame(top)
        frame_3.pack(anchor='w')

        self.compare_a = tkinter.StringVar()
        self.compare_b = tkinter.StringVar()

        self.tree = ttk.Treeview(
            frame_1,
            columns=('TDoc', 'Rev.', 'Add to compare A', 'Add to compare B'),
            show='headings',
            selectmode="browse",
            style=style_name)

        set_column(self.tree, 'TDoc', "TDoc #", width=110)
        set_column(self.tree, 'Rev.', width=50)
        set_column(self.tree, 'Add to compare A', width=110)
        set_column(self.tree, 'Add to compare B', width=110)
        self.tree.bind("<Double-Button-1>", self.on_double_click)
        self.insert_rows(revisions)

        self.tree_scroll = ttk.Scrollbar(frame_1)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        tkinter.Label(frame_2, text="{0} Documents".format(self.count)).pack(side=tkinter.LEFT)

        tkinter.Label(frame_3, textvariable=self.compare_a).pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, text='  vs.  ').pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, textvariable=self.compare_b).pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, text='  ').pack(side=tkinter.LEFT)

        tkinter.Button(
            frame_3,
            text='Compare!',
            command=self.compare_tdocs).pack(side=tkinter.LEFT)


    def insert_rows(self, df):
        count = 0
        if df is None:
            return

        if isinstance(df, pd.Series):
            rows = [(self.tdoc_id, df)]
        else:
            rows = df.iterrows()

        for idx, row in rows:
            if count == 0:
                count = count + 1
                mod = count % 2
                if mod > 0:
                    tag = 'odd'
                else:
                    tag = 'even'
                self.tree.insert("", "end", tags=(tag,), values=(
                    idx,
                    '00',
                    'Click',
                    'Click'))

            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            self.tree.insert("", "end", tags=(tag,), values=(
                idx,
                '{0:02d}'.format(row['Revisions']),
                'Click',
                'Click'))

        self.count = count

        self.tree.tag_configure('odd', background='#E8E8E8')
        self.tree.tag_configure('even', background='#DFDFDF')
        treeview_sort_column(self.tree, 'Rev.')

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except:
            actual_value = None
        tdoc_id = item_values[0]
        revision = 'r' + '{0:02d}'.format(item_values[1])
        if revision == 'r00':
            tdoc_to_search = tdoc_id
        else:
            tdoc_to_search = tdoc_id + revision
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))
        if column == 0:
            print('Opening {0}'.format(actual_value))
            gui.main.download_and_open_tdoc(actual_value)
        if column == 1:
            print('Opening {0}'.format(tdoc_to_search))
            gui.main.download_and_open_tdoc(tdoc_to_search)
        if column == 2:
            self.compare_a.set(tdoc_to_search)
        if column == 3:
            self.compare_b.set(tdoc_to_search)

    def compare_tdocs(self):
        compare_a = self.compare_a.get()
        compare_b = self.compare_b.get()
        print('Comparing {0} vs. {1}'.format(compare_a, compare_b))
        self.parent_gui_tools.compare_tdocs(entry_1=compare_a, entry_2=compare_b, )

class EmailsTable:

    def __init__(self, parent, favicon, tdoc_id):
        top = self.top = tkinter.Toplevel(parent)
        top.title("Emails for {0}".format(tdoc_id))
        top.iconbitmap(favicon)