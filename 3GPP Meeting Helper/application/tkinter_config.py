import tkinter
from tkinter import ttk
from tkinter.font import Font

root = tkinter.Tk()
font_big = Font(root, size=25)
font_medium = Font(root, size=15)
font_normal = Font(root, size=12)

ttk_style_tbutton_big = 'my.big.TButton'
ttk_style_tbutton_medium = 'my.medium.TButton'
ttk.Style().configure(ttk_style_tbutton_big, font=font_big)
ttk.Style().configure(ttk_style_tbutton_medium, font=font_medium)