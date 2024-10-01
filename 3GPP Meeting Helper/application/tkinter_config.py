import tkinter
from tkinter.font import Font

from gui.common.utils import get_new_style

root = tkinter.Tk()

print(f'Initialized Tkinter root. Configured Parameters:')
for e in root.configure().items():
    print(f'  {e}')

font_big = Font(root, size=25)
font_medium = Font(root, size=15)
font_normal = Font(root, size=12)

ttk_style_tbutton_big = 'my.big.TButton'
ttk_style_tbutton_medium = 'my.medium.TButton'
get_new_style().configure(ttk_style_tbutton_big, font=font_big)
get_new_style().configure(ttk_style_tbutton_medium, font=font_medium)
main_frame = tkinter.Frame(root)

table_text_color = '#000000'
table_bg_color_odd = '#E8E8E8'
table_bg_color_even = '#DFDFDF'
