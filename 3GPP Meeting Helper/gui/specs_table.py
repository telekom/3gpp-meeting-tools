import re
import textwrap
import tkinter
import traceback
import webbrowser
from tkinter import ttk

import pandas as pd
import pyperclip

import application
from server import specs
from server.specs import file_version_to_version, version_to_file_version, download_spec

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


class SpecsTable:
    current_specs = None
    all_specs = None
    spec_metadata = None
    title_width = 550

    filter_release = None
    filter_series = None
    filter_text = None

    def __init__(self, parent, favicon, parent_gui_tools):
        init_style()
        top = self.top = tkinter.Toplevel(parent)
        top.title("Specs Table. Double-Click on Spec # or Release # to open")
        top.iconbitmap(favicon)
        self.parent_gui_tools = parent_gui_tools

        frame_1 = tkinter.Frame(top)
        frame_1.pack(anchor='w')
        frame_2 = tkinter.Frame(top)
        frame_2.pack()
        frame_3 = tkinter.Frame(top)
        frame_3.pack(anchor='w')

        self.spec_count = tkinter.StringVar()

        # https://stackoverflow.com/questions/50625306/what-is-the-best-way-to-show-data-in-a-table-in-tkinter
        self.tree = ttk.Treeview(
            frame_2,
            columns=('Spec', 'Title', 'Versions', 'Last'),
            show='headings',
            selectmode="browse",
            style=style_name,
            padding=[-5, -25, -5, -25])  # Left, top, right, bottom

        set_column(self.tree, 'Spec', "Spec #", width=110)
        set_column(self.tree, 'Title', width=SpecsTable.title_width, center=False)
        set_column(self.tree, 'Versions', width=120)
        set_column(self.tree, 'Last', width=120)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.load_data(initial_load=True)
        self.insert_current_specs()

        self.tree_scroll = ttk.Scrollbar(frame_2)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        # tree.grid(row=0, column=0)

        # Can also do this: https://stackoverflow.com/questions/33781047/tkinter-drop-down-list-of-check-boxes-combo-boxes
        self.search_text = tkinter.StringVar()
        self.search_entry = tkinter.Entry(frame_1, textvariable=self.search_text, width=25, font='TkDefaultFont')
        self.search_text.trace_add(['write', 'unset'], self.select_text)

        tkinter.Label(frame_1, text="Search: ").pack(side=tkinter.LEFT)
        self.search_entry.pack(side=tkinter.LEFT)

        # Filter by specification series
        all_series = ['All']
        spec_series = self.current_specs['series'].unique()
        spec_series.sort()
        all_series.extend(list(spec_series))

        self.combo_series = ttk.Combobox(frame_1, values=all_series, state="readonly")
        self.combo_series.set('All')
        self.combo_series.bind("<<ComboboxSelected>>", self.select_series)

        tkinter.Label(frame_1, text="  Filter by Series: ").pack(side=tkinter.LEFT)
        self.combo_series.pack(side=tkinter.LEFT)

        # Filter by specification release
        all_releases = ['All']
        spec_releases = self.current_specs['release'].unique()
        spec_releases.sort()
        all_releases.extend(list(spec_releases))

        self.combo_releases = ttk.Combobox(frame_1, values=all_releases, state="readonly")
        self.combo_releases.set('All')
        self.combo_releases.bind("<<ComboboxSelected>>", self.select_releases)

        tkinter.Label(frame_1, text="  Filter by Release: ").pack(side=tkinter.LEFT)
        self.combo_releases.pack(side=tkinter.LEFT)

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

        tkinter.Label(frame_3, textvariable=self.spec_count).pack(side=tkinter.LEFT)

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def load_data(self, initial_load=False):
        """
        Loads specifications frm the 3GPP website
        """
        # Load specs data
        print('Loading revision data for table')
        if initial_load:
            self.all_specs, self.spec_metadata = specs.get_specs()
            self.current_specs = self.all_specs
            self.filter_text = self.all_specs
            self.filter_release = self.all_specs
            self.filter_series = self.all_specs
        else:
            self.current_specs = self.all_specs
        print('Finished loading specs')

    def insert_current_specs(self):
        self.insert_rows(self.current_specs)

    def insert_rows(self, df):
        # print(df.to_string())
        df_release_count = df.groupby(by='spec')['release'].nunique()
        df_version_max = df.groupby(by='spec')['version'].max()
        df_version_count = df.groupby(by='spec')['version'].nunique()
        df_to_plot = pd.concat(
            [df_release_count, df_version_max, df_version_count],
            axis=1,
            keys=['releases', 'max_version', 'version_count'])
        df_to_plot.sort_index(inplace=True)
        # print(df_to_plot.to_string())

        count = 0
        for idx, row in df_to_plot.iterrows():
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            # Double brackets so that it always returns a series. If not, sometimes a series will be returned,
            # sometimes a single element
            # spec_entries = df.loc[[idx], :]

            # Faster alternative
            current_spec = self.spec_metadata[idx]
            title = current_spec.title

            # 'Spec', 'Title', 'Releases', 'Last'
            self.tree.insert("", "end", tags=(tag,), values=(
                idx,
                textwrap.fill(title, width=70),
                row['version_count'],
                file_version_to_version(row['max_version'])))

        self.tree.tag_configure('odd', background='#E8E8E8')
        self.tree.tag_configure('even', background='#DFDFDF')
        self.spec_count.set('{0} specifications'.format(count))

    def clear_filters(self, *args):
        self.combo_series.set('All')
        self.combo_releases.set('All')
        self.search_text.set('')

        # Reset filters
        self.filter_text = self.all_specs
        self.filter_release = self.all_specs
        self.filter_series = self.all_specs

        # Refill list
        self.apply_filters()

    def reload_data(self, *args):
        self.load_data(initial_load=True)
        self.select_series()  # One will call the other

    def apply_filters(self):
        self.tree.delete(*self.tree.get_children())
        merged_df = pd.merge(
            self.filter_release.reset_index(),
            self.filter_series.reset_index(),
            how="inner").set_index('spec')
        merged_df = pd.merge(
            merged_df.reset_index(),
            self.filter_text.reset_index(),
            how="inner").set_index('spec')
        self.current_specs = merged_df
        self.insert_current_specs()

    def select_series(self, *args):
        specs_for_series = self.all_specs
        selected_series = self.combo_series.get()
        print('Filtering by Series "{0}"'.format(selected_series))
        if selected_series == 'All':
            specs_for_series = specs_for_series
        else:
            specs_for_series = specs_for_series[specs_for_series['series'] == selected_series]

        self.filter_series = specs_for_series
        self.apply_filters()

    def select_releases(self, *args):
        specs_for_release = self.all_specs
        selected_release = self.combo_releases.get()
        print('Filtering by Release "{0}"'.format(selected_release))
        if selected_release == 'All':
            specs_for_release = specs_for_release
        else:
            specs_for_release = specs_for_release[specs_for_release['release'] == selected_release]

        self.filter_release = specs_for_release
        self.apply_filters()

    def select_text(self, *args):
        # Filter based on current TDocs
        text_search = self.search_text.get()
        if text_search is None or text_search == '':
            return

        is_regex = False
        print('Filtering by Text "{0}"'.format(text_search))

        self.filter_text = self.all_specs[
            self.all_specs['search_column'].str.contains(text_search, regex=is_regex)]
        self.apply_filters()

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except:
            actual_value = None

        spec_id = item_values[0]
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))
        if actual_value is None or actual_value == '':
            print("Empty value")
            return
        if column == 0:
            print('Clicked spec ID {0}'.format(actual_value))
        if column == 1:
            print('Clicked title for spec ID {0}: {1}'.format(spec_id, actual_value))
        if column == 2:
            print('Clicked versions for spec ID {0}: {1}'.format(spec_id, actual_value))
        if column == 3:
            file_version = version_to_file_version(actual_value)
            print('Clicked last for spec ID {0}: {1}->{2}'.format(spec_id, actual_value, file_version))
            spec_entries = self.all_specs.loc[spec_id, :]
            # To-do: solve case when only one entry is returned
            entry_to_load = spec_entries.loc[spec_entries.version == file_version].iloc[0]
            downloaded_files = download_spec(spec_id, entry_to_load.spec_url)
            application.word.open_files(downloaded_files)


class SpecVersionsTable:

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
            style=style_name,
            height=8)  # Height in rows

        set_column(self.tree, 'TDoc', "TDoc #", width=110)
        set_column(self.tree, 'Rev.', width=50)
        set_column(self.tree, 'Add to compare A', width=110)
        set_column(self.tree, 'Add to compare B', width=110)
        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.count = 0
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
                row['Revisions'],
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

        # Some issues with automatic conversion which we solve here
        tdoc_id = item_values[0]
        if isinstance(item_values[1], int):
            revision = 'r' + '{0:02d}'.format(item_values[1])
        else:
            revision = 'r' + item_values[1]

        if revision == 'r00':
            tdoc_to_search = tdoc_id
        else:
            tdoc_to_search = tdoc_id + revision
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))
        if column == 0:
            print('Opening {0}'.format(actual_value))
            # gui.main.download_and_open_tdoc(actual_value, copy_to_clipboard=True)
        if column == 1:
            print('Opening {0}'.format(tdoc_to_search))
            # gui.main.download_and_open_tdoc(tdoc_to_search, copy_to_clipboard=True)
        if column == 2:
            self.compare_a.set(tdoc_to_search)
        if column == 3:
            self.compare_b.set(tdoc_to_search)

    def compare_tdocs(self):
        compare_a = self.compare_a.get()
        compare_b = self.compare_b.get()
        print('Comparing {0} vs. {1}'.format(compare_a, compare_b))
        self.parent_gui_tools.compare_tdocs(entry_1=compare_a, entry_2=compare_b, )
