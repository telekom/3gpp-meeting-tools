from PyQt5.QtCore import QTimer, QRect
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QPushButton


class HoverMenuButton(QPushButton):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._hover_timer = QTimer(self)
        self._hover_timer.setInterval(50)
        self._hover_timer.timeout.connect(self._check_mouse_position)

    def enterEvent(self, event):
        super().enterEvent(event)
        if self.menu() and not self.menu().isVisible():
            spawn_pos = self.mapToGlobal(self.rect().bottomLeft())
            self.menu().popup(spawn_pos)
            self._hover_timer.start()

    def _check_mouse_position(self):
        if not self.menu() or not self.menu().isVisible():
            self._hover_timer.stop()
            return
        global_pos = QCursor.pos()
        btn_rect = QRect(self.mapToGlobal(self.rect().topLeft()), self.rect().size())
        menu_rect = self.menu().geometry()
        buffered_menu_rect = menu_rect.adjusted(-10, -10, 10, 10)
        if not btn_rect.contains(global_pos) and not buffered_menu_rect.contains(global_pos):
            self.menu().hide()
            self._hover_timer.stop()
