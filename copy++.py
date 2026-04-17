"""
Copy++ - 剪贴板纯文本化 (Windows, PyQt)

关键修复：
1. dataChanged 触发时不立即动手，延迟 120ms 再读剪贴板——
   避开源程序(浏览器/VSCode)分多次异步写入剪贴板的时间窗口。
2. 用 Windows 原生剪贴板 API (pywin32) 写入纯文本，
   完全模拟正常应用的写入行为，避免 Qt 跨层转换引起的数据丢失。
"""

import sys
import ctypes
from datetime import datetime

from PyQt5.QtCore import Qt, QObject, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor, QFont, QBrush
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QHBoxLayout,
    QSystemTrayIcon, QMenu, QAction, QMessageBox, QPlainTextEdit
)


# ==================== Windows 原生剪贴板写入 ====================
# 必须声明 argtypes/restype，否则在 64 位 Windows 上句柄/指针会被 ctypes
# 当作 32 位 int 处理，高位被截断，导致写入完全错乱。
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

# --- 函数签名 ---
user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.OpenClipboard.restype = wintypes.BOOL

user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = wintypes.BOOL

user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = wintypes.BOOL

user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
user32.SetClipboardData.restype = wintypes.HANDLE

kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = wintypes.HGLOBAL

kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalLock.restype = ctypes.c_void_p

kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalUnlock.restype = wintypes.BOOL

kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalFree.restype = wintypes.HGLOBAL

kernel32.Sleep.argtypes = [wintypes.DWORD]
kernel32.Sleep.restype = None


def win_set_clipboard_text(text: str):
    """用 Windows 原生 API 把纯文本写入剪贴板。
    返回 (ok: bool, err_msg: str)"""
    if not isinstance(text, str):
        return False, "非字符串"

    data = text.encode("utf-16-le") + b"\x00\x00"
    size = len(data)

    # 打开剪贴板，失败重试
    opened = False
    for _ in range(20):
        if user32.OpenClipboard(None):
            opened = True
            break
        kernel32.Sleep(10)
    if not opened:
        return False, f"OpenClipboard 失败 (err={ctypes.get_last_error()})"

    h_mem = None
    try:
        if not user32.EmptyClipboard():
            return False, f"EmptyClipboard 失败 (err={ctypes.get_last_error()})"

        h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, size)
        if not h_mem:
            return False, f"GlobalAlloc 失败 (err={ctypes.get_last_error()})"

        p_mem = kernel32.GlobalLock(h_mem)
        if not p_mem:
            err = ctypes.get_last_error()
            kernel32.GlobalFree(h_mem)
            h_mem = None
            return False, f"GlobalLock 失败 (err={err})"

        ctypes.memmove(p_mem, data, size)
        kernel32.GlobalUnlock(h_mem)

        if not user32.SetClipboardData(CF_UNICODETEXT, h_mem):
            err = ctypes.get_last_error()
            kernel32.GlobalFree(h_mem)
            h_mem = None
            return False, f"SetClipboardData 失败 (err={err})"

        # 成功：内存所有权已交给系统，不要 free
        h_mem = None
        return True, ""
    finally:
        user32.CloseClipboard()
        # 如果上面没成功，h_mem 还在我们手里，要释放
        if h_mem:
            kernel32.GlobalFree(h_mem)


# ==================== 图标 ====================
def make_icon(running: bool = False) -> QIcon:
    pix = QPixmap(64, 64)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QBrush(QColor("#4CAF50") if running else QColor("#757575")))
    p.setPen(Qt.NoPen)
    p.drawEllipse(2, 2, 60, 60)
    p.setPen(QColor("white"))
    p.setFont(QFont("Arial", 26, QFont.Bold))
    p.drawText(pix.rect(), Qt.AlignCenter, "C+")
    p.end()
    return QIcon(pix)


# ==================== 剪贴板监听器 ====================
class ClipboardWatcher(QObject):

    processed = pyqtSignal(str)
    log = pyqtSignal(str, str)

    # dataChanged 到实际处理之间的延迟（毫秒）
    # 给源程序足够时间完成它的多步写入
    PROCESS_DELAY_MS = 120
    # 写入后屏蔽自己触发信号的时间
    WRITE_GUARD_MS = 300

    def __init__(self, app: QApplication):
        super().__init__()
        self.app = app
        self.clipboard = app.clipboard()
        self._enabled = False
        self._writing_guard = False
        self._last_written_text = None

        # 延迟处理定时器
        self._pending_timer = QTimer()
        self._pending_timer.setSingleShot(True)
        self._pending_timer.timeout.connect(self._do_process)

        self.clipboard.dataChanged.connect(self._on_clipboard_changed)

    def set_enabled(self, enabled: bool):
        self._enabled = enabled
        self._writing_guard = False
        self._last_written_text = None
        self._pending_timer.stop()
        self.log.emit("INFO", f"监控已{'启动' if enabled else '停止'}")

    def _on_clipboard_changed(self):
        """只调度，不立即处理 —— 延迟 PROCESS_DELAY_MS 后再动手"""
        if not self._enabled:
            return
        if self._writing_guard:
            self.log.emit("DEBUG", "屏蔽期内的变化，忽略")
            return

        # 重置定时器：如果短时间内多次变化，只处理最后一次
        self._pending_timer.start(self.PROCESS_DELAY_MS)

    def _do_process(self):
        """定时器到期后，再读剪贴板并处理。
        此时源程序应已完成所有数据写入。"""
        if not self._enabled or self._writing_guard:
            return

        try:
            mime = self.clipboard.mimeData()
        except Exception as e:
            self.log.emit("ERROR", f"读剪贴板失败: {e}")
            return

        if mime is None:
            return

        try:
            formats = list(mime.formats())
        except Exception:
            formats = []
        fmts = ", ".join(formats) if formats else "(无)"

        # 图片 -> 跳过
        if mime.hasImage():
            self.log.emit("SKIP", f"图片，跳过 | 格式: {fmts}")
            return

        has_html = mime.hasHtml()
        has_rtf = any(
            f.lower() in ("text/rtf", "application/rtf", "text/richtext")
            for f in formats
        )

        if not (has_html or has_rtf):
            preview = (mime.text() or "")[:40].replace("\n", "\\n")
            self.log.emit("SKIP", f"非富文本，跳过 | 格式: {fmts} | {preview!r}")
            return

        try:
            plain = mime.text()
        except Exception as e:
            self.log.emit("ERROR", f"取纯文本失败: {e}")
            return

        if not plain:
            self.log.emit("SKIP", f"纯文本为空 | 格式: {fmts}")
            return

        if plain == self._last_written_text:
            self.log.emit("DEBUG", "和上次写入相同，跳过")
            return

        # 启用屏蔽期
        self._writing_guard = True
        self._last_written_text = plain

        # 用 Windows 原生 API 写入
        ok, err = win_set_clipboard_text(plain)

        preview = plain[:60].replace("\n", "\\n")
        if ok:
            self.log.emit(
                "PROCESS",
                f"已写回纯文本 | 长度: {len(plain)} | 预览: {preview!r}"
            )
            self.processed.emit(preview)
        else:
            self.log.emit("ERROR", f"Win32 写入失败: {err} | 预览: {preview!r}")

        # WRITE_GUARD_MS 后解除屏蔽
        QTimer.singleShot(self.WRITE_GUARD_MS, self._release_guard)

    def _release_guard(self):
        self._writing_guard = False


# ==================== 主窗口 ====================
class MainWindow(QWidget):

    def __init__(self, app: QApplication):
        super().__init__()
        self.app = app
        self.watcher = ClipboardWatcher(app)
        self.watcher.processed.connect(self._on_processed)
        self.watcher.log.connect(self._append_log)
        self.processed_count = 0
        self._tray_ref = None
        self._build_ui()

    def _build_ui(self):
        self.setWindowTitle("Copy++")
        self.setFixedSize(520, 520)
        self.setWindowIcon(make_icon(False))

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        title = QLabel("Copy++")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:24px;font-weight:bold;color:#333;")
        layout.addWidget(title)

        subtitle = QLabel("自动去除复制内容的格式")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color:#888;font-size:11px;")
        layout.addWidget(subtitle)

        layout.addSpacing(6)

        self.toggle_btn = QPushButton("启动")
        self.toggle_btn.setFixedHeight(42)
        self.toggle_btn.setCursor(Qt.PointingHandCursor)
        self._style_start()
        self.toggle_btn.clicked.connect(self.toggle)
        layout.addWidget(self.toggle_btn)

        self.status_label = QLabel("● 已停止")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color:#999;font-size:12px;")
        layout.addWidget(self.status_label)

        self.count_label = QLabel("已处理: 0 次")
        self.count_label.setAlignment(Qt.AlignCenter)
        self.count_label.setStyleSheet("color:#aaa;font-size:10px;")
        layout.addWidget(self.count_label)

        log_head = QHBoxLayout()
        lab = QLabel("调试日志")
        lab.setStyleSheet("color:#666;font-size:11px;font-weight:bold;")
        log_head.addWidget(lab)
        log_head.addStretch()
        clear_btn = QPushButton("清空")
        clear_btn.setFixedSize(50, 22)
        clear_btn.setStyleSheet(
            "QPushButton{background:#eee;border:none;border-radius:3px;"
            "font-size:10px;}QPushButton:hover{background:#ddd;}"
        )
        clear_btn.clicked.connect(lambda: self.log_view.clear())
        log_head.addWidget(clear_btn)
        layout.addLayout(log_head)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setStyleSheet(
            "QPlainTextEdit{background:#1e1e1e;color:#d4d4d4;"
            "font-family:Consolas,'Courier New',monospace;font-size:10px;"
            "border:1px solid #444;border-radius:4px;}"
        )
        self.log_view.setMaximumBlockCount(500)
        layout.addWidget(self.log_view, 1)

        hint = QLabel("关闭窗口 = 最小化到托盘；右键托盘图标可退出")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("color:#bbb;font-size:10px;")
        layout.addWidget(hint)

        self.setLayout(layout)

    def set_tray(self, tray):
        self._tray_ref = tray

    def _style_start(self):
        self.toggle_btn.setStyleSheet("""
            QPushButton{background:#4CAF50;color:white;font-size:15px;
                font-weight:bold;border:none;border-radius:6px;}
            QPushButton:hover{background:#45a049;}
            QPushButton:pressed{background:#3d8b40;}
        """)

    def _style_stop(self):
        self.toggle_btn.setStyleSheet("""
            QPushButton{background:#f44336;color:white;font-size:15px;
                font-weight:bold;border:none;border-radius:6px;}
            QPushButton:hover{background:#da190b;}
            QPushButton:pressed{background:#b71c1c;}
        """)

    def toggle(self):
        if self.watcher._enabled:
            self.stop_monitor()
        else:
            self.start_monitor()

    def start_monitor(self):
        self.watcher.set_enabled(True)
        self.toggle_btn.setText("停止")
        self._style_stop()
        self.status_label.setText("● 运行中")
        self.status_label.setStyleSheet("color:#4CAF50;font-size:12px;")
        if self._tray_ref:
            self._tray_ref.update_state(True)

    def stop_monitor(self):
        self.watcher.set_enabled(False)
        self.toggle_btn.setText("启动")
        self._style_start()
        self.status_label.setText("● 已停止")
        self.status_label.setStyleSheet("color:#999;font-size:12px;")
        if self._tray_ref:
            self._tray_ref.update_state(False)

    def _on_processed(self, preview):
        self.processed_count += 1
        self.count_label.setText(f"已处理: {self.processed_count} 次")

    def _append_log(self, level, msg):
        colors = {"INFO":"#4FC3F7","DEBUG":"#888",
                  "SKIP":"#FFD54F","PROCESS":"#81C784","ERROR":"#E57373"}
        c = colors.get(level, "#d4d4d4")
        ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        html = (f'<span style="color:#666">{ts}</span> '
                f'<span style="color:{c}"><b>[{level}]</b></span> '
                f'<span style="color:#d4d4d4">{self._esc(msg)}</span>')
        self.log_view.appendHtml(html)

    @staticmethod
    def _esc(s):
        return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        if self._tray_ref:
            self._tray_ref.showMessage(
                "Copy++ 仍在运行",
                "已最小化到托盘。右键图标可退出。",
                QSystemTrayIcon.Information, 2000
            )


# ==================== 托盘 ====================
class TrayIcon(QSystemTrayIcon):

    def __init__(self, window, app):
        super().__init__()
        self.window = window
        self.app = app
        self.setIcon(make_icon(False))
        self.setToolTip("Copy++ - 已停止")

        menu = QMenu()
        a_show = QAction("显示窗口", self)
        a_show.triggered.connect(self._show)
        menu.addAction(a_show)

        self.toggle_action = QAction("启动", self)
        self.toggle_action.triggered.connect(lambda: self.window.toggle())
        menu.addAction(self.toggle_action)

        menu.addSeparator()
        a_quit = QAction("退出", self)
        a_quit.triggered.connect(self._quit)
        menu.addAction(a_quit)

        self.setContextMenu(menu)
        self.activated.connect(self._on_activated)

    def _on_activated(self, reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self._show()

    def _show(self):
        self.window.show()
        self.window.raise_()
        self.window.activateWindow()

    def _quit(self):
        self.hide()
        self.app.quit()

    def update_state(self, running: bool):
        self.setIcon(make_icon(running))
        self.setToolTip("Copy++ - " + ("运行中" if running else "已停止"))
        self.toggle_action.setText("停止" if running else "启动")


# ==================== 入口 ====================
def main():
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    if not QSystemTrayIcon.isSystemTrayAvailable():
        QMessageBox.critical(None, "Copy++", "系统不支持托盘，无法运行。")
        sys.exit(1)

    window = MainWindow(app)
    tray = TrayIcon(window, app)
    window.set_tray(tray)
    tray.show()

    app._tray = tray
    app._window = window

    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()