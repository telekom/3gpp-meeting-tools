import tkinter as tk
from tkinter import ttk

class TTKHoverHelpButton(ttk.Button):
    def __init__(self, master=None, text="Button", help_text="Helpful information here.", hover_delay=1000, **kwargs):
        super().__init__(master, text=text, **kwargs)
        self.help_text = help_text
        self.hover_delay = hover_delay
        self.hover_timer = None
        self.tooltip_window = None

        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _show_tooltip(self):
        if self.tooltip_window is None:
            x = self.winfo_rootx() + self.winfo_width() // 2
            y = self.winfo_rooty() - 25  # Position above the button

            self.tooltip_window = tk.Toplevel(self)
            self.tooltip_window.wm_overrideredirect(True)  # Remove window decorations
            self.tooltip_window.wm_geometry(f"+{x}+{y}")

            tooltip_label = ttk.Label(self.tooltip_window, text=self.help_text,
                                      background="lightyellow", relief="solid", borderwidth=1)
            tooltip_label.pack(padx=5, pady=3)

    def _hide_tooltip(self):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        if self.hover_timer:
            self.after_cancel(self.hover_timer)
            self.hover_timer = None

    def _on_enter(self, event):
        self.hover_timer = self.after(self.hover_delay, self._show_tooltip)

    def _on_leave(self, event):
        self._hide_tooltip()

