import os
import tkinter

current_folder = os.path.dirname(os.path.realpath(__file__))
cloud_png_file = os.path.join(current_folder, 'cloud.png')
cloud_download_png_file = os.path.join(current_folder, 'cloud_download.png')
folder_png_file = os.path.join(current_folder, 'folder.png')
share_png_file = os.path.join(current_folder, 'share.png')
share_markdown_png_file = os.path.join(current_folder, 'share_markdown.png')
website_png_file = os.path.join(current_folder, 'website.png')
excel_png_file = os.path.join(current_folder, 'excel.png')
filter_png_file = os.path.join(current_folder, 'filter.png')
table_png_file = os.path.join(current_folder, 'table.png')
link_png_file = os.path.join(current_folder, 'link.png')
refresh_png_file = os.path.join(current_folder, 'refresh.png')
search_png_file = os.path.join(current_folder, 'search.png')
note_png_file = os.path.join(current_folder, 'note.png')
ftp_png_file = os.path.join(current_folder, 'ftp.png')
markdown_png_file = os.path.join(current_folder, 'markdown.png')
compare_png_file = os.path.join(current_folder, 'compare.png')

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
share_markdown_icon = tkinter.PhotoImage(
    file=share_markdown_png_file,
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
note_icon = tkinter.PhotoImage(
    file=note_png_file,
)
ftp_icon = tkinter.PhotoImage(
    file=ftp_png_file,
)
markdown_icon = tkinter.PhotoImage(
    file=markdown_png_file,
)
compare_icon = tkinter.PhotoImage(
    file=compare_png_file,
)