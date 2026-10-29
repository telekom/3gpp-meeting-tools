from PyQt5.QtCore import QPoint, QRect, QTimer, Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QMenu, QPushButton


class HoverMenuButton(QPushButton):
    """
    A QPushButton with an attached QMenu that automatically opens on mouse hover
    and stays open while the mouse hovers over the button, the top-level menu,
    any open descendant submenu (e.g. Quick Copy, Dual Links, etc.), or the
    transition corridor between them.

    If clicked, the menu is pinned open until explicitly dismissed by user interaction
    (clicking outside, selecting an action, or pressing Escape).
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._pinned = False
        self._outside_count = 0
        self._hover_timer = QTimer(self)
        self._hover_timer.setInterval(80)
        self._hover_timer.timeout.connect(self._check_mouse_position)

    def setMenu(self, menu: QMenu):
        old_menu = self.menu()
        if old_menu:
            try:
                old_menu.aboutToHide.disconnect(self._on_menu_about_to_hide)
            except (TypeError, RuntimeError):
                pass
        super().setMenu(menu)
        if menu:
            menu.aboutToHide.connect(self._on_menu_about_to_hide)

    def _on_menu_about_to_hide(self):
        self._pinned = False
        self._outside_count = 0
        self._hover_timer.stop()

    def enterEvent(self, event):
        super().enterEvent(event)
        if self.menu() and not self.menu().isVisible():
            spawn_pos = self.mapToGlobal(self.rect().bottomLeft())
            self.menu().popup(spawn_pos)
            self._pinned = False
            self._outside_count = 0
            self._hover_timer.start()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.menu():
            if self.menu().isVisible() and self._pinned:
                self.menu().hide()
                self._pinned = False
                self._hover_timer.stop()
                return
            self._pinned = True
            self._hover_timer.stop()
            if not self.menu().isVisible():
                spawn_pos = self.mapToGlobal(self.rect().bottomLeft())
                self.menu().popup(spawn_pos)
            return
        super().mousePressEvent(event)

    def hideEvent(self, event):
        self._hover_timer.stop()
        if self.menu() and self.menu().isVisible():
            self.menu().hide()
        super().hideEvent(event)

    def _is_descendant_menu(self, widget) -> bool:
        cur = widget
        while cur:
            if cur is self.menu() or cur is self:
                return True
            cur = cur.parent()
        return False

    def is_pos_inside(self, global_pos: QPoint, buffer: int = 25) -> bool:
        if not self.menu() or not self.menu().isVisible():
            return False

        # 1. Button geometry (with buffer)
        btn_rect = QRect(self.mapToGlobal(self.rect().topLeft()), self.rect().size()).adjusted(
            -buffer, -buffer, buffer, buffer
        )
        if btn_rect.contains(global_pos):
            return True

        # 2. Main menu geometry (with buffer)
        main_menu_rect = self.menu().geometry().adjusted(-buffer, -buffer, buffer, buffer)
        if main_menu_rect.contains(global_pos):
            return True

        # 3. All visible descendant submenus and transition corridors
        visible_subs = [s for s in self.menu().findChildren(QMenu) if s.isVisible()]
        for sub in visible_subs:
            sub_rect = sub.geometry().adjusted(-buffer, -buffer, buffer, buffer)
            if sub_rect.contains(global_pos):
                return True
            parent_w = sub.parent()
            parent_geom = parent_w.geometry() if isinstance(parent_w, QMenu) else self.menu().geometry()
            corridor = parent_geom.united(sub.geometry()).adjusted(-10, -10, 10, 10)
            if corridor.contains(global_pos):
                return True

        # 4. Active popup widget in QApplication if belonging to our menu chain
        active = QApplication.activePopupWidget()
        if isinstance(active, QMenu) and active.isVisible() and self._is_descendant_menu(active):
            if active.geometry().adjusted(-buffer, -buffer, buffer, buffer).contains(global_pos):
                return True

        return False

    def _check_mouse_position(self):
        if self._pinned:
            return
        if not self.menu() or not self.menu().isVisible():
            self._hover_timer.stop()
            self._outside_count = 0
            return

        global_pos = QCursor.pos()
        if self.is_pos_inside(global_pos):
            self._outside_count = 0
        else:
            self._outside_count += 1
            # Grace period of 3 ticks (~240ms) before closing
            if self._outside_count >= 3:
                self.menu().hide()
                self._hover_timer.stop()
                self._outside_count = 0
