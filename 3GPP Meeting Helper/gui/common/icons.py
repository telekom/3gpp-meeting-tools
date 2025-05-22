import os
import tkinter

current_folder = os.path.dirname(os.path.realpath(__file__))
cloud_png_file = os.path.join(current_folder, 'cloud.png')
cloud_download_png_file = os.path.join(current_folder, 'cloud_download.png')
folder_png_file = os.path.join(current_folder, 'folder.png')
share_png_file = os.path.join(current_folder, 'share.png')
website_png_file = os.path.join(current_folder, 'website.png')
excel_png_file = os.path.join(current_folder, 'excel.png')
filter_png_file = os.path.join(current_folder, 'filter.png')
table_png_file = os.path.join(current_folder, 'table.png')
link_png_file = os.path.join(current_folder, 'link.png')
refresh_png_file = os.path.join(current_folder, 'refresh.png')
search_png_file = os.path.join(current_folder, 'search.png')

print(f'Loading generic table icons: {cloud_png_file}, {cloud_download_png_file}')
cloud_icon = tkinter.PhotoImage(
    file=cloud_png_file,
)
cloud_download_icon = tkinter.PhotoImage(
    file=cloud_download_png_file,
)
folder_icon = tkinter.PhotoImage(
    file=folder_png_file,
)
share_icon = tkinter.PhotoImage(
    file=share_png_file,
)
excel_icon = tkinter.PhotoImage(
    file=excel_png_file,
)
website_icon = tkinter.PhotoImage(
    file=website_png_file,
)
filter_icon = tkinter.PhotoImage(
    file=filter_png_file,
)
table_icon = tkinter.PhotoImage(
    file=table_png_file,
)
link_icon = tkinter.PhotoImage(
    file=link_png_file,
)
refresh_icon = tkinter.PhotoImage(
    file=refresh_png_file,
)
search_icon = tkinter.PhotoImage(
    file=search_png_file,
)