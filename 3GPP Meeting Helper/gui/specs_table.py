import os
import re
import textwrap
import tkinter
from tkinter import ttk
from typing import NamedTuple

import pandas as pd

import application
import application.common
import application.word
from application.os import open_url_and_copy_to_clipboard, startfile
from gui.common.generic_table import GenericTable, treeview_set_row_formatting
from gui.common.gui_elements import TTKHoverHelpButton
from gui.common.icons import refresh_icon, folder_icon
from parsing.html.specs import extract_spec_files_from_spec_folder, cleanup_spec_name
from parsing.spec_types import get_spec_full_name, SpecType
from server import specs
from server.specs import file_version_to_version, version_to_file_version, download_spec_if_needed, \
    get_url_for_spec_page, get_spec_archive_remote_folder, get_specs_folder, get_url_for_crs_page, \
    get_spec_page, file_version_to_version_metadata
from utils.local_cache import file_exists


class SpecsTable(GenericTable):
    # Specs DataFrame has columns based on SpecFile
    current_specs = None
    all_specs = None
    spec_metadata = None
    title_width = 550

    def __init__(self, root_widget, favicon, parent_widget):
        super().__init__(
            parent_widget,
            "Specs Table. Double-Click on Spec # or Release # to open",
            favicon,
            ['Spec', 'Title', 'Versions', 'Local Cache', 'Group', 'CRs', 'WIs']
        )
        self.root_widget = root_widget
        self.spec_count = tkinter.StringVar()

        self.set_column('Spec', "Spec #", width=110)
        self.set_column('Title', width=SpecsTable.title_width, center=False)
        self.set_column('Versions', width=70)
        self.set_column('Local Cache', width=80)
        self.set_column('Group', width=70)
        self.set_column('CRs', width=70)
        self.set_column('WIs', width=70)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.load_data(initial_load=True, check_for_new_specs=False)
        self.insert_current_specs()

        # Can also do this:
        # https://stackoverflow.com/questions/33781047/tkinter-drop-down-list-of-check-boxes-combo-boxes
        self.search_text = tkinter.StringVar()
        self.search_entry = tkinter.Entry(
            self.top_frame,
            textvariable=self.search_text,
            width=20,
            font='TkDefaultFont')
        self.search_text.trace_add(['write', 'unset'], self.apply_filters)

        ttk.Label(self.top_frame, text="Search: ").pack(side=tkinter.LEFT)
        self.search_entry.pack(side=tkinter.LEFT)

        # Filter by specification series
        all_series = ['All']
        spec_series = self.all_specs['series'].unique()
        spec_series.sort()
        all_series.extend(list(spec_series))

        self.combo_series = ttk.Combobox(self.top_frame, values=all_series, state="readonly", width=6)
        self.combo_series.set('All')
        self.combo_series.bind("<<ComboboxSelected>>", self.apply_filters)

        ttk.Label(self.top_frame, text="  Series: ").pack(side=tkinter.LEFT)
        self.combo_series.pack(side=tkinter.LEFT)

        # Filter by specification release
        all_types = ['All']
        spec_types = self.all_specs['type'].unique()
        spec_types.sort()
        all_types.extend(list(spec_types))

        self.combo_spec_type = ttk.Combobox(self.top_frame, values=all_types, state="readonly", width=6)
        self.combo_spec_type.set('All')
        self.combo_spec_type.bind("<<ComboboxSelected>>", self.apply_filters)

        ttk.Label(self.top_frame, text="  Type: ").pack(side=tkinter.LEFT)
        self.combo_spec_type.pack(side=tkinter.LEFT)

        # Filter by group responsibility release
        all_groups = ['All']
        spec_groups = self.all_specs['responsible_group'].unique()
        spec_groups.sort()
        all_groups.extend(list(spec_groups))

        self.combo_groups = ttk.Combobox(self.top_frame, values=all_groups, state="readonly", width=8)
        self.combo_groups.set('All')
        self.combo_groups.bind("<<ComboboxSelected>>", self.apply_filters)

        ttk.Label(self.top_frame, text="  WG: ").pack(side=tkinter.LEFT)
        self.combo_groups.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  ").pack(side=tkinter.LEFT)
        ttk.Button(
            self.top_frame,
            text='Clear filters',
            command=self.clear_filters).pack(side=tkinter.LEFT)
        TTKHoverHelpButton(
            self.top_frame,
            help_text='Load specification data for all (2k+) specs, download if not available',
            command=self.load_all_specs_from_server,
            image=refresh_icon,
        ).pack(side=tkinter.LEFT)
        TTKHoverHelpButton(
            self.top_frame,
            help_text='Open local cache folder for specifications',
            image=folder_icon,
            command=lambda: startfile(get_specs_folder())).pack(side=tkinter.LEFT)

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        ttk.Label(self.bottom_frame, textvariable=self.spec_count).pack(side=tkinter.LEFT)

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def load_data(self,
                  initial_load=False,
                  check_for_new_specs=False,
                  override_pickle_cache=False,
                  load_only: list[str] | None = None):
        """
        Loads specifications frm the 3GPP website

        Args:
            load_only: If the list is not empty, it contains a list of specifications that should be re-loaded.
                Otherwise, all specifications will be reloaded
            initial_load: Loads everything
            check_for_new_specs: Whether the spec series page should be checked for new specs
            override_pickle_cache: Whether an existing cache should not be used
        """
        if load_only is None:
            load_only = []

        # Load specs data
        print('Loading revision data for LATEST specs per release for table')
        if initial_load:
            self.all_specs, self.spec_metadata = specs.get_specs(
                cache=True,
                check_for_new_specs=check_for_new_specs,
                override_pickle_cache=override_pickle_cache,
                load_only_spec_list=load_only)

            # Merge spec data and spec metadata (I do not want to have to refactor the cache...)

            # Construct 23.501 -> TS 23.501
            def try_get_spec_type(idx):
                try:
                    return self.spec_metadata[idx].type
                except Exception as e:
                    print(f'Could not retrieve spec. type for {idx}, {e}')
                    return SpecType.TS
            self.all_specs['full_name'] = self.all_specs.index.map(lambda idx: get_spec_full_name(idx, try_get_spec_type(idx)))

            # TR/TS
            self.all_specs['type'] = self.all_specs.index.map(lambda idx: try_get_spec_type(idx).to_string())

            self.current_specs = self.all_specs
        else:
            self.current_specs = self.all_specs
        print('Finished loading specs')

    def insert_current_specs(self):
        self.insert_rows(self.current_specs)

    def insert_rows(self, df):
        print('Populating specifications table')
        # print(df.to_string())
        # df_release_count = df.groupby(by='spec')['release'].nunique()
        df_version_max = df.groupby(by='spec')['version'].max()
        df_full_name = df.groupby(by='spec')['full_name'].first()
        # df_version_count = df.groupby(by='spec')['version'].nunique()
        df_to_plot = pd.concat(
            [df_version_max, df_full_name],
            axis=1,
            keys=['max_version', 'full_name'])
        df_to_plot.sort_index(inplace=True)

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
            if idx in self.spec_metadata:
                current_spec_metadata = self.spec_metadata[idx]
                title = current_spec_metadata.title
                spec_type = current_spec_metadata.type
                responsible_group = current_spec_metadata.responsible_group
            else:
                print('Could not read metadata from spec {0}, skipping'.format(idx))
                continue

            # 'Spec', 'Title', 'Releases', 'Last'
            self.tree.insert("", "end", tags=(tag,), values=(
                row['full_name'],
                textwrap.fill(title, width=70),
                'Click',
                'Click',
                responsible_group,
                'Click',
                'Click'
            ))

        treeview_set_row_formatting(self.tree)
        self.spec_count.set('{0} specifications'.format(count))

    def clear_filters(self, *args):
        # Reset filters
        self.combo_series.set('All')
        self.combo_spec_type.set('All')
        self.search_text.set('')

        # Refill list
        self.apply_filters()

    def load_all_specs_from_server(self, *args):
        self.load_data(initial_load=True, check_for_new_specs=True)
        self.apply_filters()

    def apply_filters(self, *args):
        self.tree.delete(*self.tree.get_children())
        self.current_specs = self.all_specs

        # Apply all filters
        selected_series = self.combo_series.get()
        if selected_series != 'All':
            print('Filtering by Series "{0}"'.format(selected_series))
            self.current_specs = self.current_specs[self.current_specs['series'] == selected_series]

        selected_spec_type = self.combo_spec_type.get()
        if selected_spec_type != 'All':
            print('Filtering by Spec. type "{0}"'.format(selected_spec_type))
            self.current_specs = self.current_specs[self.current_specs['type'] == selected_spec_type]

        selected_group = self.combo_groups.get()
        if selected_group != 'All':
            print('Filtering by Group "{0}"'.format(selected_group))
            self.current_specs = self.current_specs[self.current_specs['responsible_group'] == selected_group]

        text_search = self.search_text.get()
        if text_search is not None and text_search != '':
            print('Filtering by Text "{0}"'.format(text_search))
            is_regex = False
            self.current_specs = self.current_specs[
                self.current_specs['search_column'].str.contains(text_search, regex=is_regex)]

        self.insert_current_specs()

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except Exception as e:
            print(f'Error on double click: {e}')
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
            SpecVersionsTable(
                self.tk_top,
                self.favicon,
                spec_id,
                current_spec_metadata.type,
                current_spec_metadata.spec_initial_release,
                self)
        if column == 3:
            print('Clicked local folder for spec ID {0}'.format(spec_id))
            url_to_open = get_specs_folder(spec_id=spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 5:
            print('Clicked CRs link for spec ID {0}'.format(spec_id))
            url_to_open = get_url_for_crs_page(spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 6:
            print('Clicked WIs link for spec ID {0}'.format(spec_id))
            current_spec_metadata = self.spec_metadata[spec_id]
            if current_spec_metadata.related_wis is None:
                return
            print(f'Related WIs to open: {current_spec_metadata.related_wis}')
            for related_wi in current_spec_metadata.related_wis:
                url_to_open = related_wi.wid_page_url
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


def get_3gpp_guru_url_for_version_text(spec_id: str, version_text: str) -> str:
    """
        Returns the URL for the matching specification in 3GPP Guru.
        Args:
            spec_id: specification name, e.g. 23.501
            version_text: The version to be retrieved, e.g. 16.0.0

        Returns:
            The URL of the given specification/version.
    """
    file_version = version_to_file_version(version_text)
    version_metadata = file_version_to_version_metadata(file_version)
    spec_name_without_dots = spec_id.replace('.', '')

    # e.g. https://3gpp.guru/trts/Rel-19/23502-j00.html
    url = f"https://3gpp.guru/trts/Rel-{version_metadata.release}/{spec_name_without_dots}-{file_version}.html"
    return url


class SpecVersionsTable(GenericTable):
    spec_entries = None

    class SpecLocallyAvailable(NamedTuple):
        zip: bool
        pdf: bool
        html: bool
        pdf_mcc_clean: bool
        html_mcc_clean: bool

    def __init__(
            self,
            parent: tkinter.Tk,
            favicon,
            spec_id: str,
            spec_type: SpecType,
            initial_release: str,
            parent_specs_table: SpecsTable):
        super().__init__(
            root_widget=None,
            parent_widget=parent,
            widget_title="{0}, initial planned release: {1}".format(spec_id, initial_release),
            favicon=favicon,
            column_names=[
                'Spec',
                'Version',
                'Upload Date',
                'Open Word',
                'Open PDF',
                'Open HTML',
                'Open 3GPP Guru']
        )

        self.spec_id: str = spec_id
        self.spec_type = spec_type

        self.parent_specs_table = parent_specs_table
        self.spec_local_file_exists: dict[str, SpecVersionsTable.SpecLocallyAvailable] = {}

        # Before we start inserting rows, we need to load the spec archive for this specification
        # Done here because probably not all specs will be equally accessed. Thus, new versions can be reloaded
        # whenever needed
        self.count = 0
        self.spec_entries = self.load_spec_data()

        self.set_column('Spec', "Spec #", width=110, center=True)
        self.set_column('Version', width=60, center=True)
        self.set_column('Upload Date', width=100, center=True)
        self.set_column('Open Word', width=122, center=True)
        self.set_column('Open PDF', width=115, center=True)
        self.set_column('Open HTML', width=127, center=True)
        self.set_column('Open 3GPP Guru', width=127, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.footer_label = tkinter.StringVar()
        self.set_footer_label()

        ttk.Label(self.bottom_frame, textvariable=self.footer_label).pack(side=tkinter.LEFT)

        TTKHoverHelpButton(
            self.bottom_frame,
            help_text=f'Open local folder for {self.spec_id}',
            image=folder_icon,
            command=lambda: startfile(get_specs_folder(spec_id=self.spec_id))).pack(side=tkinter.LEFT)

        TTKHoverHelpButton(
            self.bottom_frame,
            help_text='(Re-)download and load specification data from 3GPP server',
            image=refresh_icon,
            command=self.reload_spec_file).pack(side=tkinter.LEFT)

        # Main frame
        self.insert_rows()
        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

    def set_footer_label(self):
        self.footer_label.set("{0} Spec versions. ".format(self.count))

    def load_spec_data(self) -> pd.DataFrame:
        """
        Loads the specification data from the HTML file retrieved from the 3GPP servers (cached as a Markdown file) and
        returns a DataFrame containing the data from the specification versions
        Returns: DataFrame containing the data from the specification versions
        """
        spec_markup, archive_page_url, series_number = get_spec_archive_remote_folder(
            self.spec_id,
            cache=True,
            force_download=False)
        specs_from_archive = extract_spec_files_from_spec_folder(spec_markup, archive_page_url, None, series_number)
        specs_df = pd.DataFrame(specs_from_archive)
        specs_df.set_index("spec", inplace=True)
        return specs_df

    def insert_rows(self):
        print('Populating version table for spec {0}'.format(self.spec_id))
        df = self.spec_entries.sort_values(by='version', ascending=False)

        count = 0
        if df is None:
            return

        if isinstance(df, pd.Series):
            rows = [(self.spec_id, df)]
        else:
            rows = df.iterrows()

        spec_name = get_spec_full_name(self.spec_id, self.spec_type)

        # 'Spec', 'Version', 'Upload Date', 'Open Word', 'Open PDF', 'Add to compare A', 'Add to compare B'
        for idx, (spec_id, row) in enumerate(rows):
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            try:
                upload_date = self.parent_specs_table.spec_metadata[spec_id].upload_dates[idx]
            except Exception as e:
                upload_date = '0000-00-00'
                print(f'Could not parse upload date: {e}')

            version_text = file_version_to_version(row['version'])

            # Fill in whether the local file is available
            spec_url = get_url_for_version_text(self.spec_entries, version_text)
            local_zip_file_path = download_spec_if_needed(spec_id, spec_url, return_only_target_local_filename=True)
            local_zip_file_path_without_extension = os.path.splitext(local_zip_file_path)[0]
            local_zip_file_exists = file_exists(local_zip_file_path)
            local_pdf_file_exists = file_exists(os.path.splitext(local_zip_file_path)[0] + '.pdf')
            local_html_file_exists = file_exists(os.path.splitext(local_zip_file_path)[0] + '.html')
            spec_locally_available = SpecVersionsTable.SpecLocallyAvailable(
                zip=file_exists(local_zip_file_path),
                pdf=file_exists(local_zip_file_path_without_extension + '.pdf'),
                html=file_exists(local_zip_file_path_without_extension + '.html'),
                pdf_mcc_clean=file_exists(local_zip_file_path_without_extension + '_MCCclean.pdf'),
                html_mcc_clean=file_exists(local_zip_file_path_without_extension + '_MCCclean.html')
            )
            self.spec_local_file_exists[self.spec_id] = spec_locally_available

            values = (
                spec_name,
                version_text,
                upload_date,
                ('Open' if spec_locally_available.zip else 'Download') + ' Word',
                ('Open' if spec_locally_available.pdf or spec_locally_available.pdf_mcc_clean else 'Download') + ' PDF',
                (
                    'Open' if spec_locally_available.html or spec_locally_available.html_mcc_clean else 'Download') + ' HTML',
                'Link'
            )
            self.tree.insert("", "end", tags=(tag,), values=values)
            # print(f'INSERTED {values}')

        self.count = count
        treeview_set_row_formatting(self.tree)
        self.set_footer_label()

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
            url_to_open = get_url_for_spec_page(spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 1:
            print('Clicked spec ID {0}, version {1}'.format(spec_id, actual_value))
        if column == 3:
            print('Opening Word {0}, version {1}'.format(spec_id, row_version))
            spec_url = get_url_for_version_text(self.spec_entries, row_version)
            downloaded_files = download_spec_if_needed(spec_id, spec_url)
            application.word.open_files(downloaded_files)
            self.reload_table()
        if column == 4:
            print('Opening PDF {0}, version {1}'.format(spec_id, row_version))
            spec_url = get_url_for_version_text(self.spec_entries, row_version)
            downloaded_files = download_spec_if_needed(spec_id, spec_url)
            pdf_files = application.word.export_document(
                downloaded_files,
                export_format=application.common.ExportType.PDF)
            for pdf_file in pdf_files:
                startfile(pdf_file)
            self.reload_table()
        if column == 5:
            print('Opening HTML {0}, version {1}'.format(spec_id, row_version))
            spec_url = get_url_for_version_text(self.spec_entries, row_version)
            downloaded_files = download_spec_if_needed(spec_id, spec_url)
            pdf_files = application.word.export_document(
                downloaded_files,
                export_format=application.common.ExportType.HTML)
            for pdf_file in pdf_files:
                startfile(pdf_file)
            self.reload_table()
        if column == 6:
            print('Opening 3GPP Guru site for {0}, version {1}'.format(spec_id, row_version))
            url_to_open = get_3gpp_guru_url_for_version_text(self.spec_id, row_version)
            startfile(url_to_open)

    # Used to identify specs within unzipped files
    spec_regex = re.compile('[\d]{5}-[\w]{3}\.doc[x]?')

    def reload_spec_file(self):
        get_spec_page(self.spec_id, cache=True, force_download=True)
        get_spec_archive_remote_folder(self.spec_id, cache=True, force_download=True)
        self.parent_specs_table.load_data(initial_load=True, override_pickle_cache=True, load_only=[self.spec_id])
        self.spec_entries = self.load_spec_data()
        self.reload_table()

    def reload_table(self):
        super().clear_tree()
        self.insert_rows()
