"""Cross-DPI application QIcon builder, Win32 HICON/IPropertyStore stamping, and global event filter."""

from pathlib import Path
from typing import Optional, Tuple

from PyQt5.QtCore import QEvent, QObject, Qt
from PyQt5.QtGui import QColor, QIcon, QPainter, QPen, QPixmap
from PyQt5.QtWidgets import QWidget

from core.ui.theme import C


APP_USER_MODEL_ID = "3GPP Delegate Tools.1.0"
APP_DISPLAY_NAME = "3GPP Delegate Tools"

_CACHED_APP_ICON: Optional[QIcon] = None
_CACHED_HICON_BIG: Optional[int] = None
_CACHED_HICON_SMALL: Optional[int] = None


def _get_assets_dir() -> Path:
    try:
        from core.utils.paths import get_project_root
        return get_project_root() / "assets"
    except ImportError:
        return Path("assets")


def _hydrate_file_if_needed(path: Path) -> None:
    """Force OneDrive Files On-Demand hydration by reading the file bytes."""
    try:
        if path.is_file():
            _ = path.read_bytes()
    except Exception:
        pass


def _render_fallback_pixmap(size: int = 256) -> QPixmap:
    """Render the geometric 3GPP network emblem in memory if disk assets are unhydrated or unavailable."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)

    painter.setBrush(QColor(C.TEXT_PRIMARY))
    painter.setPen(Qt.NoPen)
    corner_radius = max(2, size // 5)
    margin = max(1, size // 128)
    painter.drawRoundedRect(margin, margin, size - 2 * margin, size - 2 * margin, corner_radius, corner_radius)

    pen = QPen(QColor(C.INFO_ACCENT))
    pen.setWidth(max(2, size // 12))
    pen.setJoinStyle(Qt.RoundJoin)
    pen.setCapStyle(Qt.RoundCap)
    painter.setPen(pen)

    center_x = size * 0.50
    top_y = size * 0.28
    bl_x = size * 0.25
    bl_y = size * 0.72
    br_x = size * 0.75
    br_y = size * 0.72

    painter.drawLine(int(center_x), int(top_y), int(bl_x), int(bl_y))
    painter.drawLine(int(center_x), int(top_y), int(br_x), int(br_y))
    painter.drawLine(int(bl_x), int(bl_y), int(br_x), int(br_y))

    painter.setBrush(QColor(C.BG_CARD))
    painter.setPen(Qt.NoPen)
    node_radius = max(2, size // 10)
    painter.drawEllipse(int(center_x - node_radius), int(top_y - node_radius), node_radius * 2, node_radius * 2)
    painter.drawEllipse(int(bl_x - node_radius), int(bl_y - node_radius), node_radius * 2, node_radius * 2)
    painter.drawEllipse(int(br_x - node_radius), int(br_y - node_radius), node_radius * 2, node_radius * 2)
    painter.end()
    return pixmap


def get_app_icon_path() -> Optional[Path]:
    """Return the path to `3gpp_tools.ico` after ensuring OneDrive hydration, or a cached fallback."""
    assets_dir = _get_assets_dir()
    ico_path = assets_dir / "3gpp_tools.ico"
    _hydrate_file_if_needed(ico_path)
    if ico_path.is_file() and ico_path.stat().st_size > 0:
        return ico_path

    try:
        from core.utils.paths import get_cache_root
        fallback_ico = get_cache_root() / "3gpp_tools_fallback.ico"
        if not fallback_ico.is_file():
            fallback_ico.parent.mkdir(parents=True, exist_ok=True)
            pm = _render_fallback_pixmap(256)
            pm.save(str(fallback_ico), "ICO")
        if fallback_ico.is_file():
            return fallback_ico
    except Exception:
        pass
    return None


def create_app_icon() -> QIcon:
    """
    Build a multi-resolution application icon across all standard Windows DPI sizes
    (16, 20, 24, 32, 40, 48, 64, 128, 256 px) with in-memory fallback.
    """
    global _CACHED_APP_ICON
    if _CACHED_APP_ICON is not None and not _CACHED_APP_ICON.isNull():
        return _CACHED_APP_ICON

    assets_dir = _get_assets_dir()
    ico_path = assets_dir / "3gpp_tools.ico"
    png_path = assets_dir / "3gpp_tools.png"

    _hydrate_file_if_needed(ico_path)
    _hydrate_file_if_needed(png_path)

    icon = QIcon()
    if ico_path.is_file():
        icon.addFile(str(ico_path))

    master_pixmap = QPixmap(str(png_path)) if png_path.is_file() else QPixmap()
    if master_pixmap.isNull():
        master_pixmap = _render_fallback_pixmap(256)

    for sz in (16, 20, 24, 32, 40, 48, 64, 128, 256):
        scaled = master_pixmap.scaled(
            sz, sz, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        if not scaled.isNull():
            icon.addPixmap(scaled, QIcon.Normal, QIcon.Off)
            icon.addPixmap(scaled, QIcon.Active, QIcon.Off)
            icon.addPixmap(scaled, QIcon.Selected, QIcon.Off)

    _CACHED_APP_ICON = icon
    return icon


def _get_win32_hicons() -> Tuple[int, int]:
    """Load and cache native Win32 big and small HICON handles from `3gpp_tools.ico`."""
    global _CACHED_HICON_BIG, _CACHED_HICON_SMALL
    if _CACHED_HICON_BIG and _CACHED_HICON_SMALL:
        return _CACHED_HICON_BIG, _CACHED_HICON_SMALL

    import os
    if os.name != "nt":
        return 0, 0

    ico_path = get_app_icon_path()
    if not ico_path:
        return 0, 0

    try:
        import ctypes
        import ctypes.wintypes as wt

        user32 = ctypes.windll.user32
        user32.LoadImageW.argtypes = [
            wt.HINSTANCE,
            wt.LPCWSTR,
            wt.UINT,
            ctypes.c_int,
            ctypes.c_int,
            wt.UINT,
        ]
        user32.LoadImageW.restype = wt.HANDLE
        user32.GetSystemMetrics.argtypes = [ctypes.c_int]
        user32.GetSystemMetrics.restype = ctypes.c_int

        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        SM_CXICON = 11
        SM_CYICON = 12
        SM_CXSMICON = 49
        SM_CYSMICON = 50

        big_w = user32.GetSystemMetrics(SM_CXICON) or 32
        big_h = user32.GetSystemMetrics(SM_CYICON) or 32
        sm_w = user32.GetSystemMetrics(SM_CXSMICON) or 16
        sm_h = user32.GetSystemMetrics(SM_CYSMICON) or 16

        h_big = int(user32.LoadImageW(None, str(ico_path), IMAGE_ICON, big_w, big_h, LR_LOADFROMFILE) or 0)
        h_sm = int(user32.LoadImageW(None, str(ico_path), IMAGE_ICON, sm_w, sm_h, LR_LOADFROMFILE) or 0)
        if not h_big:
            h_big = int(user32.LoadImageW(None, str(ico_path), IMAGE_ICON, 0, 0, LR_LOADFROMFILE | 0x00000040) or 0)
        if not h_sm:
            h_sm = h_big

        _CACHED_HICON_BIG = h_big
        _CACHED_HICON_SMALL = h_sm
        return h_big, h_sm
    except Exception:
        return 0, 0


def _apply_win32_shell_window_properties(hwnd: int) -> None:
    """
    Bind explicit `IPropertyStore` shell metadata (`PKEY_AppUserModel_ID`,
    `PKEY_AppUserModel_RelaunchIconResource`, `PKEY_AppUserModel_RelaunchDisplayNameResource`)
    directly to `hwnd` so the Windows 10/11 Taskbar never falls back to a blank
    generic window icon when running as an unpinned `python.exe` process.
    """
    import os
    if os.name != "nt" or not hwnd:
        return

    ico_path = get_app_icon_path()
    if not ico_path:
        return

    try:
        import ctypes
        import ctypes.wintypes as wt
        import uuid

        class PROPERTYKEY(ctypes.Structure):
            _fields_ = [("fmtid", ctypes.c_byte * 16), ("pid", wt.DWORD)]

        class PROPVARIANT_UNION(ctypes.Union):
            _fields_ = [("pwszVal", wt.LPWSTR), ("ullVal", ctypes.c_ulonglong)]

        class PROPVARIANT(ctypes.Structure):
            _fields_ = [
                ("vt", wt.USHORT),
                ("wReserved1", wt.USHORT),
                ("wReserved2", wt.USHORT),
                ("wReserved3", wt.USHORT),
                ("u", PROPVARIANT_UNION),
            ]

        fmtid = (ctypes.c_byte * 16)(*uuid.UUID("{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}").bytes_le)
        pkey_aumid = PROPERTYKEY(fmtid, 5)
        pkey_relaunch_icon = PROPERTYKEY(fmtid, 3)
        pkey_relaunch_name = PROPERTYKEY(fmtid, 4)
        iid_prop_store = (ctypes.c_byte * 16)(*uuid.UUID("{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}").bytes_le)

        shell32 = ctypes.windll.shell32
        shell32.SHGetPropertyStoreForWindow.argtypes = [
            wt.HWND,
            ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p),
        ]
        shell32.SHGetPropertyStoreForWindow.restype = ctypes.HRESULT

        store = ctypes.c_void_p()
        hr = shell32.SHGetPropertyStoreForWindow(hwnd, iid_prop_store, ctypes.byref(store))
        if hr != 0 or not store.value:
            return

        vt = ctypes.cast(store, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        set_value = ctypes.WINFUNCTYPE(
            ctypes.HRESULT,
            ctypes.c_void_p,
            ctypes.POINTER(PROPERTYKEY),
            ctypes.POINTER(PROPVARIANT),
        )(vt[6])
        commit = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(vt[7])
        release = ctypes.WINFUNCTYPE(wt.ULONG, ctypes.c_void_p)(vt[2])

        try:
            pv = PROPVARIANT()
            pv.vt = 31  # VT_LPWSTR

            pv.u.pwszVal = APP_USER_MODEL_ID
            set_value(store, ctypes.byref(pkey_aumid), ctypes.byref(pv))

            pv.u.pwszVal = f"{ico_path},0"
            set_value(store, ctypes.byref(pkey_relaunch_icon), ctypes.byref(pv))

            pv.u.pwszVal = APP_DISPLAY_NAME
            set_value(store, ctypes.byref(pkey_relaunch_name), ctypes.byref(pv))

            commit(store)
        finally:
            release(store)
    except Exception:
        pass


def apply_native_window_icon(widget: Optional[QWidget], set_class_icon: bool = False) -> None:
    """
    Apply both Qt `QIcon` and native Win32 `WM_SETICON` / `IPropertyStore` icons
    to a top-level window widget.
    """
    if widget is None:
        return

    icon = create_app_icon()
    if not icon.isNull():
        widget.setWindowIcon(icon)

    import os
    if os.name != "nt":
        return

    try:
        hwnd = int(widget.winId())
    except Exception:
        return
    if not hwnd:
        return

    h_big, h_sm = _get_win32_hicons()
    if h_big or h_sm:
        try:
            import ctypes
            import ctypes.wintypes as wt

            user32 = ctypes.windll.user32
            user32.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
            user32.SendMessageW.restype = wt.LPARAM

            WM_SETICON = 0x0080
            ICON_SMALL = 0
            ICON_BIG = 1

            if h_big:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_big)
            if h_sm:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_sm)

            if set_class_icon and hasattr(user32, "SetClassLongPtrW"):
                user32.SetClassLongPtrW.argtypes = [wt.HWND, ctypes.c_int, ctypes.c_void_p]
                user32.SetClassLongPtrW.restype = ctypes.c_void_p
                GCLP_HICON = -14
                GCLP_HICONSM = -34
                if h_big:
                    user32.SetClassLongPtrW(hwnd, GCLP_HICON, h_big)
                if h_sm:
                    user32.SetClassLongPtrW(hwnd, GCLP_HICONSM, h_sm)
        except Exception:
            pass

    _apply_win32_shell_window_properties(hwnd)


def apply_console_window_icon() -> None:
    """Set the 3GPP Tools icon on the attached Windows Console/Terminal window if present."""
    import os
    if os.name != "nt":
        return

    h_big, h_sm = _get_win32_hicons()
    if not (h_big or h_sm):
        return

    try:
        import ctypes
        import ctypes.wintypes as wt

        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        kernel32.GetConsoleWindow.restype = wt.HWND
        user32.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
        user32.SendMessageW.restype = wt.LPARAM

        hwnd = int(kernel32.GetConsoleWindow() or 0)
        if hwnd:
            WM_SETICON = 0x0080
            if h_big:
                user32.SendMessageW(hwnd, WM_SETICON, 1, h_big)
            if h_sm:
                user32.SendMessageW(hwnd, WM_SETICON, 0, h_sm)
    except Exception:
        pass


class _WindowIconEventFilter(QObject):
    """
    Application-wide event filter that automatically stamps every top-level
    `Qt.Window` or `Qt.Dialog` with both the Qt `QIcon` and native Win32 taskbar
    icon handles (`WM_SETICON` + `IPropertyStore`) whenever its native `HWND`
    is shown or recreated (e.g. after `setWindowFlags(Qt.Window)`).
    """

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:
        et = event.type()
        if et in (QEvent.Show, QEvent.WinIdChange) and isinstance(obj, QWidget) and obj.isWindow():
            wtype = int(obj.windowFlags() & Qt.WindowType_Mask)
            if wtype in (int(Qt.Window), int(Qt.Dialog)):
                apply_native_window_icon(obj, set_class_icon=False)
        return False


def install_window_icon_manager(app) -> None:
    """Install the global top-level window icon event filter on `QApplication`."""
    if app is None:
        return
    if getattr(app, "_window_icon_filter", None) is None:
        flt = _WindowIconEventFilter(app)
        app.installEventFilter(flt)
        app._window_icon_filter = flt
    apply_console_window_icon()

