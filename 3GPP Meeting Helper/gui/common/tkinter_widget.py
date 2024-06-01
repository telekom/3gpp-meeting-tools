import tkinter


class TkWidget:
    def __init__(
            self,
            root_widget: tkinter.Tk | None,
            parent_widget: tkinter.Tk | None,
            widget_title: str | None = None,
            favicon: str | None = None
    ):
        """

        Args:
            root_widget: The root widget
            parent_widget: The caller GUI (e.g. tools dialog)
            widget_title: The title of this GUI. Will appear at the top of the GUI
            favicon: Icon to show in the top-left corner of this GUI. File path
        """
        self.root_widget: tkinter.Tk = root_widget
        self.parent_widget: tkinter.Tk = parent_widget
        self.class_type = type(self).__name__

        if parent_widget is not None:
            self.tk_top = tkinter.Toplevel(parent_widget)
        elif root_widget is not None:
            self.tk_top = tkinter.Toplevel(root_widget)
        else:
            self.tk_top = tkinter.Toplevel()

        self.favicon = favicon
        self.widget_title = widget_title
        if self.widget_title is not None:
            self.tk_top.title(widget_title)
        if self.favicon is not None:
            self.tk_top.iconbitmap(favicon)
