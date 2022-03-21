import os
import re
import textwrap
import tkinter
from tkinter import ttk

import pandas as pd
import pyperclip

import application
import application.word
import parsing.word as word_parser
from parsing.html_specs import extract_spec_files_from_spec_folder, cleanup_spec_name
from parsing.spec_types import get_spec_full_name
from server import specs
from server.specs import file_version_to_version, version_to_file_version, download_spec_if_needed, \
    get_url_for_spec_page, get_spec_archive_remote_folder, get_specs_folder, get_url_for_crs_page

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


def open_url_and_copy_to_clipboard(url_to_open: str):
    """
    Opens a given URL and copies it to the clipboard
    Args:
        url_to_open: A URL
    """
    pyperclip.copy(url_to_open)
    os.startfile(url_to_open)
    print('Opened {0} and copied to clipboard'.format(url_to_open))


class SpecsTable:
    current_specs = None
    all_specs = None
    spec_metadata = None
    title_width = 550

    filter_release = None
    filter_series = None
    filter_text = None
    filter_group = None

    def __init__(self, parent, favicon, parent_gui_tools):
        init_style()
        self.top = tkinter.Toplevel(parent)
        self.top.title("Specs Table. Double-Click on Spec # or Release # to open")
        self.top.iconbitmap(favicon)
        self.favicon = favicon
        self.parent_gui_tools = parent_gui_tools

        frame_1 = tkinter.Frame(self.top)
        frame_1.pack(anchor='w')
        frame_2 = tkinter.Frame(self.top)
        frame_2.pack()
        frame_3 = tkinter.Frame(self.top)
        frame_3.pack(anchor='w')

        self.spec_count = tkinter.StringVar()

        # https://stackoverflow.com/questions/50625306/what-is-the-best-way-to-show-data-in-a-table-in-tkinter
        self.tree = ttk.Treeview(
            frame_2,
            columns=('Spec', 'Title', 'Versions', 'Last', 'Local Cache', 'Group', 'CRs'),
            show='headings',
            selectmode="browse",
            style=style_name,
            padding=[-5, -25, -5, -25])  # Left, top, right, bottom

        set_column(self.tree, 'Spec', "Spec #", width=110)
        set_column(self.tree, 'Title', width=SpecsTable.title_width, center=False)
        set_column(self.tree, 'Versions', width=70)
        set_column(self.tree, 'Last', width=80)
        set_column(self.tree, 'Local Cache', width=80)
        set_column(self.tree, 'Group', width=70)
        set_column(self.tree, 'CRs', width=70)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.load_data(initial_load=True)
        self.insert_current_specs()

        self.tree_scroll = ttk.Scrollbar(frame_2)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        # tree.grid(row=0, column=0)

        # Can also do this:
        # https://stackoverflow.com/questions/33781047/tkinter-drop-down-list-of-check-boxes-combo-boxes
        self.search_text = tkinter.StringVar()
        self.search_entry = tkinter.Entry(frame_1, textvariable=self.search_text, width=25, font='TkDefaultFont')
        self.search_text.trace_add(['write', 'unset'], self.select_text)

        tkinter.Label(frame_1, text="Search: ").pack(side=tkinter.LEFT)
        self.search_entry.pack(side=tkinter.LEFT)

        # Filter by specification series
        all_series = ['All']
        spec_series = self.all_specs['series'].unique()
        spec_series.sort()
        all_series.extend(list(spec_series))

        self.combo_series = ttk.Combobox(frame_1, values=all_series, state="readonly", width=8)
        self.combo_series.set('All')
        self.combo_series.bind("<<ComboboxSelected>>", self.select_series)

        tkinter.Label(frame_1, text="  Filter by Series: ").pack(side=tkinter.LEFT)
        self.combo_series.pack(side=tkinter.LEFT)

        # Filter by specification release
        all_releases = ['All']
        spec_releases = self.all_specs['release'].unique()
        spec_releases.sort()
        all_releases.extend(list(spec_releases))

        self.combo_releases = ttk.Combobox(frame_1, values=all_releases, state="readonly", width=8)
        self.combo_releases.set('All')
        self.combo_releases.bind("<<ComboboxSelected>>", self.select_releases)

        tkinter.Label(frame_1, text="  Filter by Release: ").pack(side=tkinter.LEFT)
        self.combo_releases.pack(side=tkinter.LEFT)

        # Filter by group responsibility release
        all_groups = ['All']
        spec_groups = self.all_specs['responsible_group'].unique()
        spec_groups.sort()
        all_groups.extend(list(spec_groups))

        self.combo_groups = ttk.Combobox(frame_1, values=all_groups, state="readonly", width=8)
        self.combo_groups.set('All')
        self.combo_groups.bind("<<ComboboxSelected>>", self.select_groups)

        tkinter.Label(frame_1, text="  Filter by Group: ").pack(side=tkinter.LEFT)
        self.combo_groups.pack(side=tkinter.LEFT)

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
        print('Loading revision data for LATEST specs per release for table')
        if initial_load:
            self.all_specs, self.spec_metadata = specs.get_specs()
            self.current_specs = self.all_specs
            self.filter_text = self.all_specs
            self.filter_release = self.all_specs
            self.filter_series = self.all_specs
            self.filter_group = self.all_specs
        else:
            self.current_specs = self.all_specs
        print('Finished loading specs')

    def insert_current_specs(self):
        self.insert_rows(self.current_specs)

    def insert_rows(self, df):
        # print(df.to_string())
        # df_release_count = df.groupby(by='spec')['release'].nunique()
        df_version_max = df.groupby(by='spec')['version'].max()
        # df_version_count = df.groupby(by='spec')['version'].nunique()
        df_to_plot = pd.concat(
            [df_version_max],
            axis=1,
            keys=['max_version'])
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
            current_spec_metadata = self.spec_metadata[idx]
            title = current_spec_metadata.title
            spec_type = current_spec_metadata.type
            responsible_group = current_spec_metadata.responsible_group

            # Construct 23.501 -> TS 23.501
            spec_name = get_spec_full_name(idx, spec_type)

            # 'Spec', 'Title', 'Releases', 'Last'
            self.tree.insert("", "end", tags=(tag,), values=(
                spec_name,
                textwrap.fill(title, width=70),
                'Click',
                file_version_to_version(row['max_version']),
                'Click',
                responsible_group,
                'Click'
            ))

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
        merged_df = pd.merge(
            merged_df.reset_index(),
            self.filter_group.reset_index(),
            how="inner").set_index('spec')
        self.current_specs = merged_df
        self.insert_current_specs()

    def select_series(self, *args):
        specs_for_series = self.all_specs
        selected_series = self.combo_series.get()
        print('Filtering by Series "{0}"'.format(selected_series))
        if selected_series != 'All':
            specs_for_series = specs_for_series[specs_for_series['series'] == selected_series]

        self.filter_series = specs_for_series
        self.apply_filters()

    def select_releases(self, *args):
        specs_for_release = self.all_specs
        selected_release = self.combo_releases.get()
        print('Filtering by Release "{0}"'.format(selected_release))
        if selected_release != 'All':
            specs_for_release = specs_for_release[specs_for_release['release'] == selected_release]

        self.filter_release = specs_for_release
        self.apply_filters()

    def select_groups(self, *args):
        specs_for_group = self.all_specs
        selected_group = self.combo_groups.get()
        print('Filtering by Group "{0}"'.format(selected_group))
        if selected_group != 'All':
            specs_for_group = specs_for_group[specs_for_group['responsible_group'] == selected_group]

        self.filter_group = specs_for_group
        self.apply_filters()

    def select_text(self, *args):
        # Filter based on current TDocs
        text_search = self.search_text.get()
        if text_search is None or text_search == '':
            self.filter_text = self.all_specs
            self.apply_filters()
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

        spec_id = cleanup_spec_name(item_values[0], clean_type=True, clean_dots=False)
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))

        # Select entries for this spec
        # Use '[[]]' so that .loc returns a DataFrame
        # https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.loc.html
        spec_entries = self.current_specs.loc[[spec_id], :]

        if actual_value is None or actual_value == '':
            print("Empty value")
            return

        if column == 0:
            print('Clicked spec ID {0}. Opening 3GPP spec page'.format(actual_value))
            url_to_open = get_url_for_spec_page(spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 1:
            print('Clicked title for spec ID {0}: {1}. Opening 3GPP spec page'.format(spec_id, actual_value))
            url_to_open = get_url_for_spec_page(spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 2:
            print('Clicked versions for spec ID {0}: {1}'.format(spec_id, actual_value))
            current_spec_metadata = self.spec_metadata[spec_id]
            spec_type = current_spec_metadata.type
            SpecVersionsTable(self.top, self.favicon, spec_entries, spec_id, spec_type)
        if column == 3:
            spec_url = get_url_for_version_text(spec_entries, actual_value)
            downloaded_files = download_spec_if_needed(spec_id, spec_url)
            application.word.open_files(downloaded_files)
        if column == 4:
            print('Clicked local folder for spec ID {0}'.format(spec_id))
            url_to_open = get_specs_folder(spec_id=spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 6:
            print('Clicked CRs link for spec ID {0}'.format(spec_id))
            url_to_open = get_url_for_crs_page(spec_id)
            open_url_and_copy_to_clipboard(url_to_open)


def get_url_for_version_text(spec_entries: pd.DataFrame, version_text: str) -> str:
    """
    Returns the URL for the matching version. It is assumed that all rows in the DataFrame are for a single Spec.
    Args:
        spec_entries: DataFrame containing entries for a given specification
        version_text: The version to be retrieved, e.g. 16.0.0

    Returns:
        The URL of the given specification/version.
    """
    file_version = version_to_file_version(version_text)

    # Because of using '[[]]', it is sure that the returned object is a DataFrame and not a Series
    entry_to_load = spec_entries.loc[spec_entries.version == file_version, ['spec_url']]
    entry_to_load = entry_to_load.iloc[0]
    return entry_to_load.spec_url


class SpecVersionsTable:
    spec_entries = None
    spec_id = None

    def __init__(self, parent, favicon, spec_entries, spec_id, spec_type):
        top = self.top = tkinter.Toplevel(parent)
        top.title("All Spec versions for {0}".format(spec_id))
        top.iconbitmap(favicon)

        self.spec_id = spec_id
        self.spec_type = spec_type

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
            columns=('Spec', 'Version', 'Open Word', 'Open PDF', 'Open HTML', 'Add to compare A', 'Add to compare B'),
            show='headings',
            selectmode="browse",
            style=style_name,
            height=8)  # Height in rows

        set_column(self.tree, 'Spec', "Spec #", width=110)
        set_column(self.tree, 'Version', width=60)
        set_column(self.tree, 'Open Word', width=100)
        set_column(self.tree, 'Open PDF', width=100)
        set_column(self.tree, 'Open HTML', width=100)
        set_column(self.tree, 'Add to compare A', width=110)
        set_column(self.tree, 'Add to compare B', width=110)
        self.tree.bind("<Double-Button-1>", self.on_double_click)

        # Before we start inserting rows, we need to load the spec archive for this specification
        # Done here because probably not all specs will be equally accessed. Thus, new versions can be reloaded
        # whenever needed
        spec_markup, archive_page_url, series_number = get_spec_archive_remote_folder(spec_id, cache=True)
        specs_from_archive = extract_spec_files_from_spec_folder(spec_markup, archive_page_url, None, series_number)
        specs_df = pd.DataFrame(specs_from_archive)
        specs_df.set_index("spec", inplace=True)
        self.spec_entries = specs_df

        self.count = 0
        self.insert_rows()

        self.tree_scroll = ttk.Scrollbar(frame_1)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        tkinter.Label(frame_2, text="{0} Spec versions".format(self.count)).pack(side=tkinter.LEFT)

        tkinter.Label(frame_3, textvariable=self.compare_a).pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, text='  vs.  ').pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, textvariable=self.compare_b).pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, text='  ').pack(side=tkinter.LEFT)

        tkinter.Button(
            frame_3,
            text='Compare!',
            command=self.compare_spec_versions).pack(side=tkinter.LEFT)

        tkinter.Button(
            frame_3,
            text='Open local folder',
            command=self.open_cache_folder).pack(side=tkinter.LEFT)

    def insert_rows(self):
        df = self.spec_entries.sort_values(by='version', ascending=False)

        count = 0
        if df is None:
            return

        if isinstance(df, pd.Series):
            rows = [(self.tdoc_id, df)]
        else:
            rows = df.iterrows()

        spec_name = get_spec_full_name(self.spec_id, self.spec_type)

        # 'Spec', 'Version', 'Open Word', 'Open PDF', 'Add to compare A', 'Add to compare B'
        for spec_id, row in rows:
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            self.tree.insert("", "end", tags=(tag,), values=(
                spec_name,
                file_version_to_version(row['version']),
                'Open Word',
                'Open PDF',
                'Open HTML',
                'Click',
                'Click'))

        self.count = count

        self.tree.tag_configure('odd', background='#E8E8E8')
        self.tree.tag_configure('even', background='#DFDFDF')

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except:
            actual_value = None

        spec_id = cleanup_spec_name(item_values[0], clean_type=True, clean_dots=False)
        row_version = item_values[1]
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))

        # 'Spec', 'Version', 'Open Word', 'Open PDF', 'Add to compare A', 'Add to compare B'
        if column == 0:
            print('Clicked spec ID {0}. Opening 3GPP spec page'.format(actual_value))
            os.startfile(get_url_for_spec_page(spec_id))
        if column == 1:
            print('Clicked spec ID {0}, version {1}'.format(spec_id, actual_value))
        if column == 2:
            print('Opening Word {0}, version {1}'.format(spec_id, row_version))
            spec_url = get_url_for_version_text(self.spec_entries, row_version)
            downloaded_files = download_spec_if_needed(spec_id, spec_url)
            application.word.open_files(downloaded_files)
        if column == 3:
            print('Opening PDF {0}, version {1}'.format(spec_id, row_version))
            spec_url = get_url_for_version_text(self.spec_entries, row_version)
            downloaded_files = download_spec_if_needed(spec_id, spec_url)
            pdf_files = application.word.export_document(
                downloaded_files,
                export_format=application.word.ExportType.PDF)
            for pdf_file in pdf_files:
                os.startfile(pdf_file)
        if column == 4:
            print('Opening HTML {0}, version {1}'.format(spec_id, row_version))
            spec_url = get_url_for_version_text(self.spec_entries, row_version)
            downloaded_files = download_spec_if_needed(spec_id, spec_url)
            pdf_files = application.word.export_document(
                downloaded_files,
                export_format=application.word.ExportType.HTML)
            for pdf_file in pdf_files:
                os.startfile(pdf_file)
        if column == 5:
            print('Added Compare A: {0}, version {1}'.format(spec_id, row_version))
            self.compare_a.set(row_version)
        if column == 6:
            print('Added Compare B: {0}, version {1}'.format(spec_id, row_version))
            self.compare_b.set(row_version)

    # Used to identify specs within unzipped files
    spec_regex = re.compile('[\d]{5}-[\w]{3}\.doc[x]?')

    def compare_spec_versions(self):
        version_a = self.compare_a.get()
        version_b = self.compare_b.get()
        file_version_a = version_to_file_version(version_a)
        file_version_b = version_to_file_version(version_b)
        print('Comparing {0} {1} ({3}) vs. {2} ({4})'.format(
            self.spec_id,
            version_a,
            version_b,
            file_version_a,
            file_version_b))
        spec_id = self.spec_id

        comparison_name = '{0}-{1}-to-{2}.docx'.format(spec_id, file_version_a, file_version_b)
        spec_folder = get_specs_folder(spec_id=spec_id)
        comparison_file = os.path.join(spec_folder, comparison_name)

        # ToDo: check if file already exists. If yes, open and return document
        # Check if document exists
        # Open document
        # Return document

        spec_a_url = get_url_for_version_text(self.spec_entries, version_a)
        spec_b_url = get_url_for_version_text(self.spec_entries, version_b)

        downloaded_a_files = download_spec_if_needed(spec_id, spec_a_url)
        downloaded_b_files = download_spec_if_needed(spec_id, spec_b_url)

        downloaded_a_files = [e for e in downloaded_a_files if self.spec_regex.search(e) is not None]
        downloaded_b_files = [e for e in downloaded_b_files if self.spec_regex.search(e) is not None]

        if len(downloaded_a_files) == 0 or len(downloaded_b_files) == 0:
            print('Need two TDocs to compare. One of them does not contain TDocs')
            return None

        downloaded_a_files = downloaded_a_files[0]
        downloaded_b_files = downloaded_b_files[0]

        comparison_document = word_parser.compare_documents(
            downloaded_a_files,
            downloaded_b_files,
            compare_formatting=False,
            compare_case_changes=False,
            compare_whitespace=False)
        comparison_document.Activate()
        comparison_window = comparison_document.ActiveWindow

        wdRevisionsMarkupAll = 2
        # wdRevisionsMarkupNone = 0
        # wdRevisionsMarkupSimple = 1
        comparison_window.View.RevisionsFilter.Markup = wdRevisionsMarkupAll
        comparison_window.View.ShowFormatChanges = False
        # wdBalloonRevisions = 0
        wdInLineRevisions = 1
        # wdMixedRevisions = 2
        comparison_window.View.RevisionsMode = wdInLineRevisions

        # ToDo: Save comparison document
        return comparison_document

    def open_cache_folder(self):
        folder_name = get_specs_folder(spec_id=self.spec_id)
        os.startfile(folder_name)
