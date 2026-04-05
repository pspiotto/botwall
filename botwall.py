"""
BotWall — live screenshot monitor for DreamBot / RuneScape clients.
Requires: PyQt5, pywin32, psutil, Pillow
"""

import sys
import time
import ctypes
import io

import psutil
import win32gui
import win32ui
import win32process
import win32con
from PIL import Image

from PyQt5.QtCore import (
    Qt, QThread, pyqtSignal, QTimer, QSize, QPoint, QUrl
)
from PyQt5.QtGui import QPixmap, QImage, QFont, QColor, QPainter, QCursor, QIcon, QDesktopServices
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QScrollArea, QGridLayout, QVBoxLayout, QHBoxLayout, QFrame,
    QSizePolicy, QMessageBox, QToolBar, QComboBox, QMenu
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
SCAN_INTERVAL_MS = 3000          # re-scan window list every 3 s
CAPTURE_INTERVAL_HIGH = 250      # High CPU: refresh every 0.25 s
CAPTURE_INTERVAL_LOW  = 1000     # Low CPU:  refresh every 1 s

CARD_W_DEFAULT = 320
CARD_H_DEFAULT = 220
HEADER_H = 30
ZOOM_FACTOR = 1.1

BG_COLOR        = "#12121e"
TOOLBAR_COLOR   = "#09090f"
CARD_COLOR      = "#1c1c2e"
HEADER_COLOR    = "#183048"
TEXT_COLOR      = "#dde1e7"
DIM_COLOR       = "#607080"
ACCENT_TEAL     = "#3dc8d8"
ACCENT_RED      = "#df4545"

KEYWORDS = ("dreambot", "runescape", "oldschool runescape")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_interesting(title: str) -> bool:
    tl = title.lower()
    return any(kw in tl for kw in KEYWORDS)


def _is_dreambot41(title: str) -> bool:
    """Returns True if the window title looks like a DreamBot 4.1.x client."""
    tl = title.lower()
    return "dreambot" in tl and "4.1" in tl


def capture_hwnd(hwnd: int) -> QPixmap | None:
    """Capture a window via PrintWindow → PIL → QPixmap. Returns None on failure."""
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bottom - top
        if w <= 0 or h <= 0:
            return None

        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc  = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bitmap  = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(mfc_dc, w, h)
        save_dc.SelectObject(bitmap)

        # PW_RENDERFULLCONTENT = 2 — captures layered/hardware-accelerated content
        result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

        bmp_info = bitmap.GetInfo()
        bmp_str  = bitmap.GetBitmapBits(True)

        img = Image.frombuffer(
            "RGB",
            (bmp_info["bmWidth"], bmp_info["bmHeight"]),
            bmp_str, "raw", "BGRX", 0, 1
        )

        # Cleanup GDI objects
        win32gui.DeleteObject(bitmap.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)

        if result == 0:
            return None  # PrintWindow failed

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        qimg = QImage()
        qimg.loadFromData(buf.read())
        return QPixmap.fromImage(qimg)

    except Exception:
        return None


# ---------------------------------------------------------------------------
# Scanner thread — emits list of (hwnd, title, pid, proc_name, cpu_pct, mem_mb)
# ---------------------------------------------------------------------------
class Scanner(QThread):
    updated = pyqtSignal(list)  # list of (hwnd, title, pid, proc_name, cpu_pct, mem_mb)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = True
        self._proc_cache: dict[int, psutil.Process] = {}  # pid → Process

    def run(self):
        while self._running:
            clients = []
            found_pids: set[int] = set()

            def _cb(hwnd, _):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                title = win32gui.GetWindowText(hwnd)
                if not title or not _is_interesting(title):
                    return
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    # Reuse cached Process object so cpu_percent() is meaningful
                    if pid not in self._proc_cache:
                        self._proc_cache[pid] = psutil.Process(pid)
                        # First call initialises the baseline; returns 0.0
                        self._proc_cache[pid].cpu_percent(interval=None)
                    proc = self._proc_cache[pid]
                    proc_name = proc.name()
                    cpu_pct   = proc.cpu_percent(interval=None)
                    mem_mb    = proc.memory_info().rss / (1024 * 1024)
                    found_pids.add(pid)
                except Exception:
                    pid, proc_name, cpu_pct, mem_mb = 0, "", 0.0, 0.0
                clients.append((hwnd, title, pid, proc_name, cpu_pct, mem_mb))

            win32gui.EnumWindows(_cb, None)

            # Evict stale entries from the cache
            stale = set(self._proc_cache) - found_pids
            for pid in stale:
                del self._proc_cache[pid]

            self.updated.emit(clients)
            # Sleep in small increments so we can stop quickly
            for _ in range(SCAN_INTERVAL_MS // 100):
                if not self._running:
                    break
                time.sleep(0.1)

    def stop(self):
        self._running = False
        self.wait(2000)


# ---------------------------------------------------------------------------
# Capturer thread — loops over known hwnds and captures screenshots
# ---------------------------------------------------------------------------
class Capturer(QThread):
    captured = pyqtSignal(int, QPixmap)  # hwnd, pixmap

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = True
        self._hwnds: list[int] = []
        self._interval_ms = CAPTURE_INTERVAL_HIGH

    def set_hwnds(self, hwnds: list[int]):
        self._hwnds = list(hwnds)

    def set_interval(self, ms: int):
        self._interval_ms = ms

    def run(self):
        while self._running:
            for hwnd in list(self._hwnds):
                if not self._running:
                    break
                px = capture_hwnd(hwnd)
                if px is not None:
                    self.captured.emit(hwnd, px)
                time.sleep(0.05)  # small pause between captures to avoid hammering
            # Wait out the remainder of the interval
            elapsed = 0
            while elapsed < self._interval_ms and self._running:
                time.sleep(0.1)
                elapsed += 100

    def stop(self):
        self._running = False
        self.wait(3000)


# ---------------------------------------------------------------------------
# ClientCard — one card per detected client window
# ---------------------------------------------------------------------------
class ClientCard(QFrame):
    pin_toggled       = pyqtSignal(int, bool)  # hwnd, is_pinned
    minimize_requested = pyqtSignal(int)        # hwnd

    def __init__(self, hwnd: int, title: str, pid: int, proc_name: str,
                 cpu_pct: float = 0.0, mem_mb: float = 0.0, parent=None):
        super().__init__(parent)
        self.hwnd = hwnd
        self.pid = pid
        self.proc_name = proc_name
        self.title = title
        self._cpu_pct = cpu_pct
        self._mem_mb  = mem_mb

        self._low_cpu = False
        self._pinned = False

        self.setFrameShape(QFrame.NoFrame)
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.setStyleSheet(f"""
            ClientCard {{
                background: {CARD_COLOR};
                border-radius: 4px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # --- Header bar ---
        self._header = QFrame()
        self._header.setFixedHeight(HEADER_H)
        self._header.setStyleSheet(f"background: {HEADER_COLOR}; border-radius: 4px 4px 0 0;")
        h_layout = QHBoxLayout(self._header)
        h_layout.setContentsMargins(8, 0, 8, 0)

        self._title_lbl = QLabel()
        self._title_lbl.setStyleSheet(f"color: {TEXT_COLOR}; font-size: 11px; font-weight: bold;")
        self._title_lbl.setTextFormat(Qt.PlainText)
        self._title_lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        self._cpu_lbl = QLabel()
        self._cpu_lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 9px;")

        self._mem_lbl = QLabel()
        self._mem_lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 9px;")

        self._pid_lbl = QLabel()
        self._pid_lbl.setStyleSheet(f"color: {ACCENT_TEAL}; font-size: 10px;")

        self._pin_btn = QPushButton("📌")
        self._pin_btn.setFixedSize(20, 20)
        self._pin_btn.setFlat(True)
        self._pin_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self._pin_btn.setToolTip("Pin to top")
        self._pin_btn.clicked.connect(self._toggle_pin)

        h_layout.addWidget(self._title_lbl)
        h_layout.addWidget(self._cpu_lbl)
        h_layout.addWidget(self._mem_lbl)
        h_layout.addWidget(self._pid_lbl)
        h_layout.addWidget(self._pin_btn)

        # --- Screenshot label ---
        self._img_lbl = QLabel()
        self._img_lbl.setAlignment(Qt.AlignCenter)
        self._img_lbl.setStyleSheet(f"background: #0a0a14;")
        self._img_lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        layout.addWidget(self._header)
        layout.addWidget(self._img_lbl)

        self._update_labels(title, pid, proc_name)
        self._update_stats(cpu_pct, mem_mb)
        self._pixmap_raw: QPixmap | None = None
        self._update_pin_visual()

    # ------------------------------------------------------------------
    def _toggle_pin(self):
        self._pinned = not self._pinned
        self._update_pin_visual()
        self.pin_toggled.emit(self.hwnd, self._pinned)

    def _update_pin_visual(self):
        if self._pinned:
            self._header.setStyleSheet(
                f"background: #1a3a55; border-top: 2px solid {ACCENT_TEAL}; border-radius: 4px 4px 0 0;"
            )
            self._pin_btn.setStyleSheet(f"color: {ACCENT_TEAL}; font-size: 12px;")
            self._pin_btn.setToolTip("Unpin")
        else:
            self._header.setStyleSheet(
                f"background: {HEADER_COLOR}; border-radius: 4px 4px 0 0;"
            )
            self._pin_btn.setStyleSheet(f"color: {DIM_COLOR}; font-size: 12px;")
            self._pin_btn.setToolTip("Pin to top")

    def set_low_cpu(self, enabled: bool):
        self._low_cpu = enabled
        self._rescale()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._bring_to_front()
        super().mousePressEvent(event)

    def _update_labels(self, title: str, pid: int, proc_name: str):
        max_chars = 26
        short = title if len(title) <= max_chars else title[:max_chars - 1] + "…"
        self._title_lbl.setText(short)
        self._pid_lbl.setText(f"PID {pid}")

    def _update_stats(self, cpu_pct: float, mem_mb: float):
        self._cpu_pct = cpu_pct
        self._mem_mb  = mem_mb

        # CPU color: dim < 40%, orange 40-80%, red > 80%
        if cpu_pct >= 80:
            cpu_color = ACCENT_RED
        elif cpu_pct >= 40:
            cpu_color = "#f0a040"
        else:
            cpu_color = DIM_COLOR
        self._cpu_lbl.setText(f"{cpu_pct:.1f}%")
        self._cpu_lbl.setStyleSheet(f"color: {cpu_color}; font-size: 9px;")
        self._cpu_lbl.setToolTip(f"CPU: {cpu_pct:.1f}%")

        # RAM color: dim < 500 MB, orange 500-1000 MB, red > 1000 MB
        if mem_mb >= 1000:
            mem_color = ACCENT_RED
        elif mem_mb >= 500:
            mem_color = "#f0a040"
        else:
            mem_color = DIM_COLOR
        if mem_mb >= 1024:
            mem_text = f"{mem_mb / 1024:.1f}GB"
        else:
            mem_text = f"{mem_mb:.0f}MB"
        self._mem_lbl.setText(mem_text)
        self._mem_lbl.setStyleSheet(f"color: {mem_color}; font-size: 9px;")
        self._mem_lbl.setToolTip(f"RAM: {mem_text}")

    def update_info(self, title: str, pid: int, proc_name: str,
                    cpu_pct: float = 0.0, mem_mb: float = 0.0):
        self.pid = pid
        self.proc_name = proc_name
        self.title = title
        self._update_labels(title, pid, proc_name)
        self._update_stats(cpu_pct, mem_mb)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{ background: #1a1a2e; color: {TEXT_COLOR}; border: 1px solid {DIM_COLOR}; }}
            QMenu::item:selected {{ background: {HEADER_COLOR}; }}
        """)
        menu.addAction("Bring to Front", lambda: self._bring_to_front())
        menu.addSeparator()
        menu.addAction("Minimize to Shelf", lambda: self.minimize_requested.emit(self.hwnd))
        menu.exec_(event.globalPos())

    def _bring_to_front(self):
        try:
            placement = win32gui.GetWindowPlacement(self.hwnd)
            if placement[1] == win32con.SW_SHOWMINIMIZED:
                win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(self.hwnd)
        except Exception:
            pass

    def update_pixmap(self, pixmap: QPixmap):
        self._pixmap_raw = pixmap
        self._rescale()

    def _rescale(self):
        if self._pixmap_raw is None:
            return
        target_w = self._img_lbl.width()
        target_h = self._img_lbl.height()
        if target_w <= 0 or target_h <= 0:
            return

        src = self._pixmap_raw
        if self._low_cpu:
            # Convert to grayscale to reduce rendering work
            gray = src.toImage().convertToFormat(QImage.Format_Grayscale8)
            src = QPixmap.fromImage(gray)

        transform = Qt.FastTransformation if self._low_cpu else Qt.SmoothTransformation
        scaled = src.scaled(target_w, target_h, Qt.KeepAspectRatio, transform)
        self._img_lbl.setPixmap(scaled)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._rescale()


# ---------------------------------------------------------------------------
# Empty-state placeholder
# ---------------------------------------------------------------------------
class EmptyPlaceholder(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        lbl = QLabel("No DreamBot clients detected")
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 15px;")
        layout.addWidget(lbl)
        sub = QLabel("Launch DreamBot and accounts will appear here automatically.")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet(f"color: {DIM_COLOR}; font-size: 11px;")
        layout.addWidget(sub)


# ---------------------------------------------------------------------------
# MinimizedStrip — compact header-only representation of a minimized client
# ---------------------------------------------------------------------------
class MinimizedStrip(QFrame):
    restore_requested = pyqtSignal(int)  # hwnd

    def __init__(self, hwnd: int, title: str, pid: int,
                 cpu_pct: float = 0.0, mem_mb: float = 0.0, parent=None):
        super().__init__(parent)
        self.hwnd = hwnd
        self.pid  = pid

        self.setFixedHeight(38)
        self.setMinimumWidth(220)
        self.setMaximumWidth(320)
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.setStyleSheet(f"""
            MinimizedStrip {{
                background: {HEADER_COLOR};
                border-radius: 4px;
                border-left: 3px solid {ACCENT_TEAL};
            }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 0, 4, 0)
        layout.setSpacing(6)

        self._title_lbl = QLabel()
        self._title_lbl.setStyleSheet(f"color: {TEXT_COLOR}; font-size: 12px; font-weight: bold;")
        self._title_lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        self._cpu_lbl = QLabel()
        self._cpu_lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 11px;")

        self._mem_lbl = QLabel()
        self._mem_lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 11px;")

        restore_btn = QPushButton("↑")
        restore_btn.setFixedSize(26, 26)
        restore_btn.setFlat(True)
        restore_btn.setCursor(QCursor(Qt.PointingHandCursor))
        restore_btn.setToolTip("Restore to grid")
        restore_btn.setStyleSheet(f"color: {ACCENT_TEAL}; font-size: 15px; font-weight: bold;")
        restore_btn.clicked.connect(lambda: self.restore_requested.emit(self.hwnd))

        layout.addWidget(self._title_lbl)
        layout.addWidget(self._cpu_lbl)
        layout.addWidget(self._mem_lbl)
        layout.addWidget(restore_btn)

        self._set_title(title)
        self.update_stats(cpu_pct, mem_mb)

    def _set_title(self, title: str):
        max_chars = 20
        short = title if len(title) <= max_chars else title[:max_chars - 1] + "…"
        self._title_lbl.setText(short)
        self._title_lbl.setToolTip(title)

    def update_stats(self, cpu_pct: float, mem_mb: float):
        if cpu_pct >= 80:
            cpu_color = ACCENT_RED
        elif cpu_pct >= 40:
            cpu_color = "#f0a040"
        else:
            cpu_color = DIM_COLOR
        self._cpu_lbl.setText(f"{cpu_pct:.1f}%")
        self._cpu_lbl.setStyleSheet(f"color: {cpu_color}; font-size: 11px;")

        if mem_mb >= 1000:
            mem_color = ACCENT_RED
        elif mem_mb >= 500:
            mem_color = "#f0a040"
        else:
            mem_color = DIM_COLOR
        mem_text = f"{mem_mb / 1024:.1f}GB" if mem_mb >= 1024 else f"{mem_mb:.0f}MB"
        self._mem_lbl.setText(mem_text)
        self._mem_lbl.setStyleSheet(f"color: {mem_color}; font-size: 11px;")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.restore_requested.emit(self.hwnd)
        super().mousePressEvent(event)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{ background: #1a1a2e; color: {TEXT_COLOR}; border: 1px solid {DIM_COLOR}; }}
            QMenu::item:selected {{ background: {HEADER_COLOR}; }}
        """)
        menu.addAction("Restore to Grid", lambda: self.restore_requested.emit(self.hwnd))
        menu.exec_(event.globalPos())


# ---------------------------------------------------------------------------
# MinimizedShelf — horizontal tray of minimized client strips
# ---------------------------------------------------------------------------
class MinimizedShelf(QWidget):
    restore_requested = pyqtSignal(int)  # hwnd

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(66)
        self.setVisible(False)
        self.setStyleSheet(f"background: #0d0d1a;")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 4, 10, 4)
        outer.setSpacing(4)

        self._label = QLabel("MINIMIZED (0)")
        self._label.setStyleSheet(
            f"color: {DIM_COLOR}; font-size: 11px; font-weight: bold; letter-spacing: 1px;"
        )
        outer.addWidget(self._label)

        scroll = QScrollArea()
        scroll.setFixedHeight(38)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("""
            QScrollArea { background: transparent; border: none; }
            QScrollBar:horizontal { background: #1a1a2e; height: 4px; }
            QScrollBar::handle:horizontal { background: #3a3a5a; border-radius: 2px; }
        """)

        self._strip_container = QWidget()
        self._strip_container.setStyleSheet("background: transparent;")
        self._strip_layout = QHBoxLayout(self._strip_container)
        self._strip_layout.setContentsMargins(0, 0, 0, 0)
        self._strip_layout.setSpacing(6)
        self._strip_layout.addStretch()
        scroll.setWidget(self._strip_container)

        outer.addWidget(scroll)

        self._strips: dict[int, MinimizedStrip] = {}

    def add_client(self, hwnd: int, title: str, pid: int,
                   cpu_pct: float = 0.0, mem_mb: float = 0.0):
        if hwnd in self._strips:
            return
        strip = MinimizedStrip(hwnd, title, pid, cpu_pct, mem_mb)
        strip.restore_requested.connect(self.restore_requested)
        # Insert before the trailing stretch
        self._strip_layout.insertWidget(self._strip_layout.count() - 1, strip)
        self._strips[hwnd] = strip
        self._refresh()

    def remove_client(self, hwnd: int):
        if hwnd not in self._strips:
            return
        strip = self._strips.pop(hwnd)
        self._strip_layout.removeWidget(strip)
        strip.deleteLater()
        self._refresh()

    def update_stats(self, hwnd: int, cpu_pct: float, mem_mb: float):
        if hwnd in self._strips:
            self._strips[hwnd].update_stats(cpu_pct, mem_mb)

    def _refresh(self):
        n = len(self._strips)
        self._label.setText(f"MINIMIZED ({n})")
        self.setVisible(n > 0)


# ---------------------------------------------------------------------------
# GridView — scrollable grid of ClientCards
# ---------------------------------------------------------------------------
class GridView(QScrollArea):
    client_minimized = pyqtSignal(int, str, int, float, float)  # hwnd, title, pid, cpu, mem
    client_removed   = pyqtSignal(int)                           # hwnd (gone from scanner)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setStyleSheet(f"""
            QScrollArea {{ background: {BG_COLOR}; border: none; }}
            QScrollBar:vertical {{ background: #1a1a2a; width: 8px; border-radius: 4px; }}
            QScrollBar::handle:vertical {{ background: #3a3a5a; border-radius: 4px; }}
            QScrollBar:horizontal {{ background: #1a1a2a; height: 8px; border-radius: 4px; }}
            QScrollBar::handle:horizontal {{ background: #3a3a5a; border-radius: 4px; }}
        """)

        self._card_w = CARD_W_DEFAULT
        self._card_h = CARD_H_DEFAULT

        self._container = QWidget()
        self._container.setStyleSheet(f"background: {BG_COLOR};")
        self._grid = QGridLayout(self._container)
        self._grid.setSpacing(8)
        self._grid.setContentsMargins(10, 10, 10, 10)
        self.setWidget(self._container)

        self._cards: dict[int, ClientCard] = {}         # hwnd → card
        self._order: list[int] = []                     # insertion order
        self._pinned: set[int] = set()                  # pinned hwnds
        self._minimized: set[int] = set()               # minimized hwnds
        self._stats: dict[int, tuple[float, float]] = {}# hwnd → (cpu_pct, mem_mb)
        self._sort_mode = "default"
        self._placeholder: EmptyPlaceholder | None = None
        self._low_cpu = False
        self._show_placeholder()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def update_clients(self, clients: list[tuple]):
        """Called from main thread with fresh 6-tuple list."""
        new_hwnds = {c[0] for c in clients}
        current_hwnds = set(self._cards.keys())

        # Remove stale cards
        removed = current_hwnds - new_hwnds
        for hwnd in removed:
            was_minimized = hwnd in self._minimized
            card = self._cards.pop(hwnd)
            self._order.remove(hwnd)
            self._pinned.discard(hwnd)
            self._minimized.discard(hwnd)
            self._stats.pop(hwnd, None)
            if not was_minimized:
                self._grid.removeWidget(card)
            card.deleteLater()
            if was_minimized:
                self.client_removed.emit(hwnd)

        # Add/update cards
        for hwnd, title, pid, proc_name, cpu_pct, mem_mb in clients:
            self._stats[hwnd] = (cpu_pct, mem_mb)
            if hwnd not in self._cards:
                card = ClientCard(hwnd, title, pid, proc_name, cpu_pct, mem_mb)
                card.set_low_cpu(self._low_cpu)
                card.pin_toggled.connect(self._on_pin_toggled)
                card.minimize_requested.connect(self._on_minimize_requested)
                self._fix_card_size(card)
                self._cards[hwnd] = card
                self._order.append(hwnd)
            else:
                self._cards[hwnd].update_info(title, pid, proc_name, cpu_pct, mem_mb)

        self._relayout()

        visible = len(self._cards) - len(self._minimized)
        if visible > 0:
            self._hide_placeholder()
        else:
            self._show_placeholder()

    def update_screenshot(self, hwnd: int, pixmap: QPixmap):
        if hwnd in self._cards:
            self._cards[hwnd].update_pixmap(pixmap)

    def zoom(self, factor: float):
        self._card_w = max(160, int(self._card_w * factor))
        self._card_h = max(110, int(self._card_h * factor))
        for card in self._cards.values():
            self._fix_card_size(card)
        self._relayout()

    def set_low_cpu(self, enabled: bool):
        self._low_cpu = enabled
        for card in self._cards.values():
            card.set_low_cpu(enabled)

    def _on_pin_toggled(self, hwnd: int, is_pinned: bool):
        if is_pinned:
            self._pinned.add(hwnd)
        else:
            self._pinned.discard(hwnd)
        self._relayout()

    def _on_minimize_requested(self, hwnd: int):
        if hwnd in self._cards and hwnd not in self._minimized:
            self._minimized.add(hwnd)
            card = self._cards[hwnd]
            self._grid.removeWidget(card)
            card.hide()
            cpu_pct, mem_mb = self._stats.get(hwnd, (0.0, 0.0))
            self.client_minimized.emit(hwnd, card.title, card.pid, cpu_pct, mem_mb)
            self._relayout()
            visible = len(self._cards) - len(self._minimized)
            if visible == 0:
                self._show_placeholder()

    def restore_client(self, hwnd: int):
        if hwnd in self._minimized:
            self._minimized.discard(hwnd)
            card = self._cards[hwnd]
            card.show()
            self._hide_placeholder()
            self._relayout()

    def set_sort_mode(self, mode: str):
        self._sort_mode = mode
        self._relayout()

    def _sorted_order(self) -> list[int]:
        visible  = [h for h in self._order if h not in self._minimized]
        pinned   = [h for h in visible if h in self._pinned]
        unpinned = [h for h in visible if h not in self._pinned]
        if self._sort_mode != "default":
            key = self._sort_key()
            pinned.sort(key=key)
            unpinned.sort(key=key)
        return pinned + unpinned

    def _sort_key(self):
        """Return a sort key function for hwnd based on current sort mode."""
        if self._sort_mode == "cpu_asc":
            return lambda h: self._stats.get(h, (0.0, 0.0))[0]
        if self._sort_mode == "cpu_desc":
            return lambda h: -self._stats.get(h, (0.0, 0.0))[0]
        if self._sort_mode == "ram_asc":
            return lambda h: self._stats.get(h, (0.0, 0.0))[1]
        if self._sort_mode == "ram_desc":
            return lambda h: -self._stats.get(h, (0.0, 0.0))[1]
        return lambda h: 0

    def all_pids(self) -> list[int]:
        return [c.pid for c in self._cards.values() if c.pid]

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _fix_card_size(self, card: ClientCard):
        card.setFixedSize(self._card_w, self._card_h)

    def _relayout(self):
        vp_w = self.viewport().width() - 20  # subtract margins
        cols = max(1, vp_w // (self._card_w + self._grid.spacing()))

        # Remove all items from grid without deleting
        for i in reversed(range(self._grid.count())):
            item = self._grid.itemAt(i)
            if item and item.widget():
                self._grid.removeWidget(item.widget())

        display_order = self._sorted_order()
        for idx, hwnd in enumerate(display_order):
            card = self._cards[hwnd]
            row, col = divmod(idx, cols)
            self._grid.addWidget(card, row, col)

        # Push cards to top-left
        self._grid.setRowStretch(len(display_order) // cols + 1, 1)
        self._grid.setColumnStretch(cols, 1)

    def _show_placeholder(self):
        if self._placeholder is None:
            self._placeholder = EmptyPlaceholder()
            self._grid.addWidget(self._placeholder, 0, 0)

    def _hide_placeholder(self):
        if self._placeholder is not None:
            self._grid.removeWidget(self._placeholder)
            self._placeholder.deleteLater()
            self._placeholder = None

    # ------------------------------------------------------------------
    # Events
    # ------------------------------------------------------------------
    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._relayout()

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            delta = event.angleDelta().y()
            factor = ZOOM_FACTOR if delta > 0 else 1.0 / ZOOM_FACTOR
            self.zoom(factor)
            event.accept()
        else:
            super().wheelEvent(event)


# ---------------------------------------------------------------------------
# BotWall — main window
# ---------------------------------------------------------------------------
class BotWall(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BotWall - Inspired by Iluminate04, Refined by Blank")
        self.resize(1280, 760)
        self.setStyleSheet(f"QMainWindow {{ background: {BG_COLOR}; }}")

        self._clients: list[tuple] = []
        self._total_opens = 0
        self._total_closes = 0
        self._active_dreambot_hwnds: dict[int, str] = {}  # hwnd → title for open 4.1.x clients
        self._self_proc = psutil.Process()
        self._self_proc.cpu_percent(interval=None)  # prime the baseline
        self._setup_ui()
        self._start_threads()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _setup_ui(self):
        # ---- Toolbar ----
        toolbar_widget = QWidget()
        toolbar_widget.setFixedHeight(54)
        toolbar_widget.setStyleSheet(f"background: {TOOLBAR_COLOR};")
        tb_layout = QHBoxLayout(toolbar_widget)
        tb_layout.setContentsMargins(12, 0, 12, 0)
        tb_layout.setSpacing(6)

        title_lbl = QLabel("BotWall")
        title_lbl.setStyleSheet(
            f"color: {ACCENT_TEAL}; font-size: 18px; font-weight: bold; letter-spacing: 1px;"
        )
        tb_layout.addWidget(title_lbl)
        tb_layout.addSpacing(10)

        def _vsep():
            s = QFrame()
            s.setFrameShape(QFrame.VLine)
            s.setFixedHeight(22)
            s.setStyleSheet(f"color: {DIM_COLOR};")
            return s

        tb_layout.addWidget(_vsep())
        tb_layout.addSpacing(8)

        # ---- BotWall self-stats ----
        self._self_cpu_lbl = QLabel("CPU: –")
        self._self_cpu_lbl.setStyleSheet("color: #5bbde8; font-size: 12px;")
        self._self_cpu_lbl.setToolTip("BotWall CPU usage")
        tb_layout.addWidget(self._self_cpu_lbl)

        tb_layout.addSpacing(10)

        self._self_mem_lbl = QLabel("MEM: –")
        self._self_mem_lbl.setStyleSheet("color: #b07ed8; font-size: 12px;")
        self._self_mem_lbl.setToolTip("BotWall memory usage")
        tb_layout.addWidget(self._self_mem_lbl)

        tb_layout.addSpacing(10)
        tb_layout.addWidget(_vsep())
        tb_layout.addSpacing(8)

        # ---- Session client stats ----
        self._opens_lbl = QLabel("↑ 0 Opened")
        self._opens_lbl.setStyleSheet("color: #4dc87a; font-size: 12px;")
        self._opens_lbl.setToolTip("DreamBot 4.1.x clients opened this session")
        tb_layout.addWidget(self._opens_lbl)

        tb_layout.addSpacing(10)

        self._closes_lbl = QLabel("↓ 0 Closed")
        self._closes_lbl.setStyleSheet("color: #e07050; font-size: 12px;")
        self._closes_lbl.setToolTip("DreamBot 4.1.x clients closed this session")
        tb_layout.addWidget(self._closes_lbl)

        tb_layout.addSpacing(10)
        tb_layout.addWidget(_vsep())
        tb_layout.addSpacing(8)

        self._count_lbl = QLabel("0 clients")
        self._count_lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 12px;")
        self._count_lbl.setToolTip("Active DreamBot clients detected")
        tb_layout.addWidget(self._count_lbl)

        tb_layout.addStretch()

        # ---- Sort control ----
        sort_combo = QComboBox()
        sort_combo.addItems(["Sort: Default", "CPU ↑", "CPU ↓", "RAM ↑", "RAM ↓"])
        sort_combo.setFixedHeight(28)
        sort_combo.setFixedWidth(120)
        sort_combo.setCursor(QCursor(Qt.PointingHandCursor))
        sort_combo.setStyleSheet(f"""
            QComboBox {{
                background: transparent;
                color: {DIM_COLOR};
                border: 1px solid {DIM_COLOR};
                border-radius: 3px;
                font-size: 11px;
                padding: 0 4px;
            }}
            QComboBox:hover {{ border-color: {ACCENT_TEAL}; color: {TEXT_COLOR}; }}
            QComboBox::drop-down {{ border: none; width: 16px; }}
            QComboBox QAbstractItemView {{
                background: #1a1a2e;
                color: {TEXT_COLOR};
                border: 1px solid {DIM_COLOR};
                selection-background-color: {HEADER_COLOR};
            }}
        """)
        sort_combo.currentIndexChanged.connect(self._on_sort_changed)
        tb_layout.addWidget(sort_combo)
        tb_layout.addSpacing(8)

        # ---- CPU mode toggle ----
        def _cpu_btn_style(active: bool) -> str:
            bg     = ACCENT_TEAL if active else "transparent"
            color  = "#000000"   if active else DIM_COLOR
            border = f"1px solid {ACCENT_TEAL}" if active else f"1px solid {DIM_COLOR}"
            return f"""
                QPushButton {{
                    background: {bg};
                    color: {color};
                    border: {border};
                    border-radius: 3px;
                    font-size: 11px;
                    font-weight: bold;
                    padding: 0 8px;
                }}
                QPushButton:hover {{ border-color: {ACCENT_TEAL}; color: {"#000" if active else TEXT_COLOR}; }}
            """

        self._btn_high_cpu = QPushButton("High CPU")
        self._btn_high_cpu.setFixedHeight(28)
        self._btn_high_cpu.setCursor(QCursor(Qt.PointingHandCursor))
        self._btn_high_cpu.setStyleSheet(_cpu_btn_style(True))
        self._btn_high_cpu.clicked.connect(lambda: self._set_cpu_mode("high"))

        self._btn_low_cpu = QPushButton("Low CPU")
        self._btn_low_cpu.setFixedHeight(28)
        self._btn_low_cpu.setCursor(QCursor(Qt.PointingHandCursor))
        self._btn_low_cpu.setStyleSheet(_cpu_btn_style(False))
        self._btn_low_cpu.clicked.connect(lambda: self._set_cpu_mode("low"))

        # store the style factory for reuse when toggling
        self._cpu_btn_style = _cpu_btn_style

        tb_layout.addWidget(self._btn_high_cpu)
        tb_layout.addWidget(self._btn_low_cpu)
        tb_layout.addSpacing(8)

        discord_btn = QPushButton("Discord")
        discord_btn.setFixedHeight(28)
        discord_btn.setCursor(QCursor(Qt.PointingHandCursor))
        discord_btn.setToolTip("Join the ETS Discord")
        discord_btn.setStyleSheet("""
            QPushButton {
                background: #5865f2;
                color: white;
                border: none;
                border-radius: 3px;
                font-size: 11px;
                font-weight: bold;
                padding: 0 10px;
            }
            QPushButton:hover { background: #6b77f5; }
            QPushButton:pressed { background: #4752c4; }
        """)
        discord_btn.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl("https://discord.gg/fEg3X3a5sh"))
        )
        tb_layout.addWidget(discord_btn)
        tb_layout.addSpacing(8)

        maximize_all_btn = QPushButton("Maximize All")
        maximize_all_btn.setFixedHeight(28)
        maximize_all_btn.setCursor(QCursor(Qt.PointingHandCursor))
        maximize_all_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {DIM_COLOR};
                border: 1px solid {DIM_COLOR};
                border-radius: 3px;
                font-size: 11px;
                font-weight: bold;
                padding: 0 8px;
            }}
            QPushButton:hover {{ border-color: {ACCENT_TEAL}; color: {TEXT_COLOR}; }}
            QPushButton:pressed {{ background: #1a3a40; }}
        """)
        maximize_all_btn.setToolTip("Maximize all DreamBot windows")
        maximize_all_btn.clicked.connect(self._maximize_all)
        tb_layout.addWidget(maximize_all_btn)

        restore_all_btn = QPushButton("Restore All")
        restore_all_btn.setFixedHeight(28)
        restore_all_btn.setCursor(QCursor(Qt.PointingHandCursor))
        restore_all_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {DIM_COLOR};
                border: 1px solid {DIM_COLOR};
                border-radius: 3px;
                font-size: 11px;
                font-weight: bold;
                padding: 0 8px;
            }}
            QPushButton:hover {{ border-color: {ACCENT_TEAL}; color: {TEXT_COLOR}; }}
            QPushButton:pressed {{ background: #1a3a40; }}
        """)
        restore_all_btn.setToolTip("Restore all DreamBot windows")
        restore_all_btn.clicked.connect(self._restore_all)
        tb_layout.addWidget(restore_all_btn)
        tb_layout.addSpacing(8)

        kill_btn = QPushButton("KILL ALL")
        kill_btn.setFixedSize(90, 32)
        kill_btn.setCursor(QCursor(Qt.PointingHandCursor))
        kill_btn.setStyleSheet(f"""
            QPushButton {{
                background: {ACCENT_RED};
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 12px;
                font-weight: bold;
            }}
            QPushButton:hover {{ background: #f05555; }}
            QPushButton:pressed {{ background: #c03535; }}
        """)
        kill_btn.clicked.connect(self._kill_all)
        tb_layout.addWidget(kill_btn)

        # ---- Grid view ----
        self._grid_view = GridView()

        # ---- Minimized shelf ----
        self._minimized_shelf = MinimizedShelf()

        # ---- Central widget ----
        central = QWidget()
        central.setStyleSheet(f"background: {BG_COLOR};")
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        main_layout.addWidget(toolbar_widget)
        main_layout.addWidget(self._grid_view)
        main_layout.addWidget(self._minimized_shelf)
        self.setCentralWidget(central)

    # ------------------------------------------------------------------
    # Threading
    # ------------------------------------------------------------------
    def _start_threads(self):
        self._scanner = Scanner()
        self._scanner.updated.connect(self._on_scan)
        self._scanner.start()

        self._capturer = Capturer()
        self._capturer.captured.connect(self._on_capture)
        self._capturer.start()

        # Self-stats refresh timer
        self._self_stats_timer = QTimer(self)
        self._self_stats_timer.setInterval(1000)
        self._self_stats_timer.timeout.connect(self._update_self_stats)
        self._self_stats_timer.start()

        # Wire minimize/restore between grid and shelf
        self._grid_view.client_minimized.connect(self._minimized_shelf.add_client)
        self._grid_view.client_removed.connect(self._minimized_shelf.remove_client)
        self._minimized_shelf.restore_requested.connect(self._on_restore_client)

    def _update_self_stats(self):
        try:
            cpu = self._self_proc.cpu_percent(interval=None)
            mem_mb = self._self_proc.memory_info().rss / (1024 * 1024)
            mem_text = f"{mem_mb / 1024:.1f}GB" if mem_mb >= 1024 else f"{mem_mb:.0f}MB"
            self._self_cpu_lbl.setText(f"CPU: {cpu:.1f}%")
            self._self_mem_lbl.setText(f"MEM: {mem_text}")
        except Exception:
            pass

    def _on_scan(self, clients: list):
        self._clients = clients
        self._capturer.set_hwnds([c[0] for c in clients])
        self._grid_view.update_clients(clients)
        # Push live stats into minimized shelf strips
        for hwnd, title, pid, proc_name, cpu_pct, mem_mb in clients:
            self._minimized_shelf.update_stats(hwnd, cpu_pct, mem_mb)
        n = len(clients)
        self._count_lbl.setText(f"{n} client{'s' if n != 1 else ''}")

        # Track DreamBot 4.1.x opens and closes
        new_hwnd_titles = {c[0]: c[1] for c in clients}
        for hwnd, title in new_hwnd_titles.items():
            if hwnd not in self._active_dreambot_hwnds and _is_dreambot41(title):
                self._active_dreambot_hwnds[hwnd] = title
                self._total_opens += 1
        for hwnd in set(self._active_dreambot_hwnds) - set(new_hwnd_titles):
            del self._active_dreambot_hwnds[hwnd]
            self._total_closes += 1
        self._opens_lbl.setText(f"↑ {self._total_opens} Opened")
        self._closes_lbl.setText(f"↓ {self._total_closes} Closed")

    def _on_capture(self, hwnd: int, pixmap: QPixmap):
        self._grid_view.update_screenshot(hwnd, pixmap)

    def _on_restore_client(self, hwnd: int):
        self._minimized_shelf.remove_client(hwnd)
        self._grid_view.restore_client(hwnd)

    def _on_sort_changed(self, idx: int):
        modes = ["default", "cpu_asc", "cpu_desc", "ram_asc", "ram_desc"]
        self._grid_view.set_sort_mode(modes[idx])

    def _set_cpu_mode(self, mode: str):
        low = (mode == "low")
        self._capturer.set_interval(CAPTURE_INTERVAL_LOW if low else CAPTURE_INTERVAL_HIGH)
        self._grid_view.set_low_cpu(low)
        self._btn_high_cpu.setStyleSheet(self._cpu_btn_style(not low))
        self._btn_low_cpu.setStyleSheet(self._cpu_btn_style(low))

    # ------------------------------------------------------------------
    # Maximize / Restore all
    # ------------------------------------------------------------------
    def _maximize_all(self):
        for hwnd, *_ in self._clients:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            except Exception:
                pass

    def _restore_all(self):
        for hwnd, *_ in self._clients:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Kill all
    # ------------------------------------------------------------------
    def _kill_all(self):
        pids = self._grid_view.all_pids()
        if not pids:
            QMessageBox.information(self, "BotWall", "No clients to kill.")
            return
        reply = QMessageBox.question(
            self, "Kill All",
            f"Kill {len(pids)} client process{'es' if len(pids) != 1 else ''}?\n"
            "This will forcefully terminate them.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        for pid in pids:
            try:
                p = psutil.Process(pid)
                p.kill()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Cleanup
    # ------------------------------------------------------------------
    def closeEvent(self, event):
        self._scanner.stop()
        self._capturer.stop()
        event.accept()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # Enable high-DPI scaling
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Apply a global dark palette so Qt widgets inherit the dark theme
    from PyQt5.QtGui import QPalette
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(BG_COLOR))
    palette.setColor(QPalette.WindowText, QColor(TEXT_COLOR))
    palette.setColor(QPalette.Base, QColor(CARD_COLOR))
    palette.setColor(QPalette.AlternateBase, QColor(TOOLBAR_COLOR))
    palette.setColor(QPalette.ToolTipBase, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ToolTipText, QColor(TOOLBAR_COLOR))
    palette.setColor(QPalette.Text, QColor(TEXT_COLOR))
    palette.setColor(QPalette.Button, QColor(HEADER_COLOR))
    palette.setColor(QPalette.ButtonText, QColor(TEXT_COLOR))
    palette.setColor(QPalette.Highlight, QColor(ACCENT_TEAL))
    palette.setColor(QPalette.HighlightedText, QColor("#000000"))
    app.setPalette(palette)

    # Set application icon (ICO for taskbar/title bar, PNG as fallback)
    import os
    # When frozen by PyInstaller, data files live in sys._MEIPASS
    if getattr(sys, "frozen", False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(base_dir, "CmCSAHz.ico")
    png_path = os.path.join(base_dir, "CmCSAHz.png")
    icon = QIcon(ico_path) if os.path.exists(ico_path) else QIcon(png_path)
    app.setWindowIcon(icon)

    window = BotWall()
    window.setWindowIcon(icon)
    window.show()
    sys.exit(app.exec_())
