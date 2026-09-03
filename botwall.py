"""
BotWall — live screenshot monitor for DreamBot / RuneScape clients.
Requires: PyQt5, pywin32, psutil (Python 3.10+)
"""

from __future__ import annotations

import sys
import re
import json
import time
import zlib
import ctypes
import winsound
import subprocess

import psutil
import win32gui
import win32ui
import win32process
import win32con

from PyQt5.QtCore import (
    Qt, QThread, pyqtSignal, QTimer, QUrl, QEvent, QSettings
)
from PyQt5.QtGui import (
    QPixmap, QImage, QColor, QCursor, QIcon, QDesktopServices, QKeySequence,
    QPainter
)
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QScrollArea, QGridLayout, QVBoxLayout, QHBoxLayout, QFrame,
    QSizePolicy, QMessageBox, QComboBox, QMenu,
    QSystemTrayIcon, QDialog, QPlainTextEdit,
    QShortcut, QInputDialog, QLineEdit, QFormLayout, QDialogButtonBox
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

# Freeze detection: a healthy game client repaints constantly, so an
# unchanged frame is a strong "stuck/disconnected" signal.
STALE_WARN_S   = 30    # amber "static" after this many unchanged seconds
STALE_FROZEN_S = 120   # red "frozen" after this many
NO_FEED_S      = 10    # dim "no feed" when captures stop arriving at all

BG_COLOR        = "#12121e"
TOOLBAR_COLOR   = "#09090f"
CARD_COLOR      = "#1c1c2e"
HEADER_COLOR    = "#183048"
TEXT_COLOR      = "#dde1e7"
DIM_COLOR       = "#607080"
ACCENT_TEAL     = "#3dc8d8"
ACCENT_RED      = "#df4545"

TB_SPACING        = 6    # toolbar inter-item gap
TB_REFLOW_SLACK   = 8    # px of headroom before a toolbar group is dropped

DISCORD_URL = "https://discord.gg/fEg3X3a5sh"

COMBO_STYLE = f"""
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
"""

MENU_STYLE = f"""
    QMenu {{ background: #1a1a2e; color: {TEXT_COLOR}; border: 1px solid {DIM_COLOR}; }}
    QMenu::item {{ padding: 4px 20px; }}
    QMenu::item:selected {{ background: {HEADER_COLOR}; }}
    QMenu::item:disabled {{ color: {DIM_COLOR}; }}
    QMenu::separator {{ height: 1px; background: {DIM_COLOR}; margin: 4px 8px; }}
"""

KEYWORDS = ("dreambot", "runescape", "twilite", "runelite")
# Defaults from earlier releases. A persisted keyword list equal to one of
# these was never customized, so it's upgraded to KEYWORDS on restore.
OLD_DEFAULT_KEYWORDS = (
    ("dreambot", "runescape"),
    ("dreambot", "runescape", "twilite"),
)

# A window only counts as a client if its process also looks like one.
# Without this, a browser tab titled "RuneScape Wiki" gets a card — and
# gets its process terminated by KILL ALL.
BOT_PROC_NAMES    = ("java.exe", "javaw.exe")
BOT_PROC_KEYWORDS = ("dreambot", "runelite", "osclient", "jagex", "twilite",
                     "onlybot")

# Some farm managers launch a stock-looking client whose window and process
# give nothing away: OnlyBot runs a RuneLite fork as a plain java.exe titled
# "RuneLite - <character>". The only fingerprint is the JVM's exe path
# (…\.onlybot\lib\jdks\…\java.exe), so the scanner tags the process name
# with the launcher — "java.exe (OnlyBot)" — and _client_kind() reads the
# tag like any other keyword. Matched case-insensitively against the exe path.
# DreamBot's bundled JRE is tagged too, so a DreamBot JVM whose title has lost
# the word (e.g. a "Fatal error starting RuneLite" dialog) stays on the
# DreamBot tab instead of drifting to RuneLite.
LAUNCHER_PATH_MARKERS = {
    ".onlybot": "OnlyBot",
    "\\dreambot\\": "DreamBot",
}

# Client kinds drive the navigation tabs. A window is classified by the
# first kind whose keyword appears in its title or process name; anything
# that passes the scan filters but matches no kind lands in "Other".
# DreamBot, TwiLite, Titan and OnlyBot always get a tab; the rest only when
# populated. OnlyBot must precede RuneLite: its titles say "RuneLite".
CLIENT_KINDS = ("DreamBot", "TwiLite", "Titan", "OnlyBot", "RuneLite")
KIND_KEYWORDS = {
    "DreamBot": ("dreambot",),
    "TwiLite":  ("twilite",),
    "Titan":    ("titan",),
    "OnlyBot":  ("onlybot",),
    "RuneLite": ("runelite",),
}
OTHER_KIND = "Other"
TITAN_KIND = "Titan"
ALWAYS_SHOWN_KINDS = ("DreamBot", "TwiLite", "Titan", "OnlyBot")

# TitanClient hosts every game client as a *tab* inside one controller
# window. Each tab is still a separate osclient.exe process, and its real
# game window (class "JagWindow") is re-parented as a hidden child of the
# controller — so it never shows up in EnumWindows and the title keyword
# scan can't find it. Scanner._scan_titan walks the controller's children
# instead. Only the ACTIVE tab actually renders: PrintWindow on an inactive
# tab's JagWindow just returns whatever is on screen at that rectangle, so
# those cards are marked "TAB HIDDEN" and dropped from the capture list.
TITAN_CONTROLLER_CLASS = "TitanClientController"
TITAN_TAB_CLASS = "JagWindow"
TITAN_TITLE = "TitanClient"
# Controller title while a game tab is active:
#   "TitanClient v0.0.99 [DEV MODE] - Pown 2194 #120612"
# (the space in the name is U+00A0; a dead tab adds " (unresponsive)").
_TITAN_ACTIVE_RE = re.compile(r"\s-\s(.+?)\s#(\d+)(?:\s*\(.*\))?\s*$")

NUM_CORES = psutil.cpu_count() or 1

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_interesting(title: str, keywords=KEYWORDS) -> bool:
    tl = title.lower()
    return any(kw in tl for kw in keywords)


def _is_bot_process(proc_name: str) -> bool:
    nl = proc_name.lower()
    return nl in BOT_PROC_NAMES or any(kw in nl for kw in BOT_PROC_KEYWORDS)


def _proc_label(proc: "psutil.Process") -> str:
    """Process name for the scanner tuple, with the launcher appended when
    the exe path reveals one (see LAUNCHER_PATH_MARKERS)."""
    name = proc.name()
    try:
        exe = proc.exe().lower()
    except Exception:  # AccessDenied, ZombieProcess…
        return name
    for marker, launcher in LAUNCHER_PATH_MARKERS.items():
        if marker in exe:
            return f"{name} ({launcher})"
    return name


def _client_kind(title: str, proc_name: str) -> str:
    """Classify a client window for the navigation tabs. `proc_name` may
    carry a launcher suffix ("java.exe (OnlyBot)") — see _proc_label."""
    haystack = f"{title} {proc_name}".lower()
    for kind in CLIENT_KINDS:
        if any(kw in haystack for kw in KIND_KEYWORDS[kind]):
            return kind
    return OTHER_KIND


def _titan_tab_title(name: str, pid: int) -> str:
    """Card title for a TitanClient tab. The account name is only known
    once the tab has been active (it's read off the controller's title),
    so fall back to the PID — the same "#<pid>" Titan shows on its tabs."""
    return f"{TITAN_TITLE} · {name}" if name else f"{TITAN_TITLE} · tab #{pid}"


def _fix_titan_title(title: str) -> str:
    """The controller sets its title as UTF-8 bytes through the ANSI API, so
    the U+00A0 in "Pown 2194" reads back as the mojibake "Â\xa0". Undo that
    (when it round-trips) and flatten NBSP to a plain space."""
    try:
        title = title.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    return title.replace("\xa0", " ")


def _titan_active_tab(ctrl_hwnd: int) -> tuple[str, int]:
    """(name, osclient pid) of the controller's active tab, or ("", 0) on
    the Home/Settings tab. Read live off the window title — cheap enough to
    call per capture, which is what makes tab switches race-free."""
    try:
        m = _TITAN_ACTIVE_RE.search(_fix_titan_title(win32gui.GetWindowText(ctrl_hwnd)))
    except Exception:
        return "", 0
    return (m.group(1).strip(), int(m.group(2))) if m else ("", 0)


def _root_hwnd(hwnd: int) -> int:
    """Top-level ancestor (GA_ROOT) — the hwnd itself for normal clients,
    the controller window for a TitanClient tab."""
    try:
        return ctypes.windll.user32.GetAncestor(hwnd, 2) or hwnd
    except Exception:
        return hwnd


# DreamBot titles read "DreamBot <ver> - <email> - <Script> v<n> - <ip> [- flag]",
# so the running script is the 3rd " - " segment with its version stripped. The
# version is dropped so v2.156 and v2.157 of the same script group together.
_SCRIPT_VER_RE = re.compile(r"\s+v\d[\d.]*$")


def _parse_script(title: str) -> str:
    """Script name from a DreamBot title, or "" if none (launcher, TwiLite…)."""
    parts = [p.strip() for p in title.split(" - ")]
    if len(parts) >= 3 and parts[0].lower().startswith("dreambot"):
        return _SCRIPT_VER_RE.sub("", parts[2]).strip()
    return ""


def _fmt_age(seconds: float) -> str:
    return f"{int(seconds)}s" if seconds < 60 else f"{int(seconds // 60)}m"


def _fmt_uptime(seconds: float) -> str:
    m = int(seconds // 60)
    if m < 60:
        return f"{m}m"
    h, m = divmod(m, 60)
    if h < 24:
        return f"{h}h{m:02d}m"
    d, h = divmod(h, 24)
    return f"{d}d{h}h"


def capture_hwnd(hwnd: int) -> QImage | None:
    """Capture a window via PrintWindow into a QImage. Returns None on failure.

    Returns QImage (not QPixmap) because this runs on the Capturer thread and
    QPixmap is GUI-thread-only.
    """
    qimg = None
    hwnd_dc = mfc_dc = save_dc = bitmap = old_bmp = None
    try:
        if win32gui.IsIconic(hwnd):
            return None  # minimized windows render as a useless titlebar sliver
        if ctypes.windll.user32.IsHungAppWindow(hwnd):
            return None  # PrintWindow would block on a hung window's message queue

        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bottom - top
        if w <= 0 or h <= 0:
            return None

        hwnd_dc = win32gui.GetWindowDC(hwnd)
        if not hwnd_dc:
            return None
        mfc_dc  = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bitmap  = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(mfc_dc, w, h)
        old_bmp = save_dc.SelectObject(bitmap)

        # PW_RENDERFULLCONTENT = 2 — captures layered/hardware-accelerated content
        result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)
        if result != 0:
            bmp_info = bitmap.GetInfo()
            bmp_str  = bitmap.GetBitmapBits(True)
            # GDI hands back BGRX rows, which is exactly Format_RGB32.
            # .copy() detaches from bmp_str before the GDI objects are freed.
            qimg = QImage(
                bmp_str, bmp_info["bmWidth"], bmp_info["bmHeight"],
                bmp_info["bmWidthBytes"], QImage.Format_RGB32
            ).copy()
    except Exception:
        qimg = None
    finally:
        # Each step guarded so one failure can't leak the remaining handles.
        # The bitmap must be deselected from the DC or DeleteObject fails.
        if old_bmp is not None:
            try:
                save_dc.SelectObject(old_bmp)
            except Exception:
                pass
        if bitmap is not None:
            try:
                win32gui.DeleteObject(bitmap.GetHandle())
            except Exception:
                pass
        if save_dc is not None:
            try:
                save_dc.DeleteDC()
            except Exception:
                pass
        if mfc_dc is not None:
            try:
                mfc_dc.DeleteDC()
            except Exception:
                pass
        if hwnd_dc:
            try:
                win32gui.ReleaseDC(hwnd, hwnd_dc)
            except Exception:
                pass
    return qimg


# ---------------------------------------------------------------------------
# Scanner thread — emits list of (hwnd, title, pid, proc_name, cpu_pct,
# mem_mb, uptime_s)
# ---------------------------------------------------------------------------
class Scanner(QThread):
    updated = pyqtSignal(list)  # list of 7-tuples, see header above
    # Emitted just before `updated` on every scan:
    #   {"controller_pid": int, "active_pid": int, "hidden_hwnds": [hwnd…]}
    # hidden_hwnds are TitanClient tabs that aren't the active tab (their
    # frames can't be captured — see TITAN_CONTROLLER_CLASS above).
    titan_state = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = True
        self._proc_cache: dict[int, psutil.Process] = {}  # pid → Process
        self._keywords: list[str] = list(KEYWORDS)
        self._titan_names: dict[int, str] = {}  # osclient pid → tab name

    def set_keywords(self, keywords: list[str]):
        """Replace the title keywords (from the settings dialog)."""
        self._keywords = [kw.lower() for kw in keywords if kw.strip()] or list(KEYWORDS)

    def run(self):
        while self._running:
            clients = []
            # cpu_percent(interval=None) measures since its own last call, so it
            # must run exactly once per process per scan — a second call in the
            # same scan (process with two matching windows) reads a ~0 ms
            # interval and returns garbage. Memoize stats per pid.
            pid_stats: dict[int, tuple[str, float, float, float]] = {}
            seen_pids: set[int] = set()
            scan_t = time.time()

            def stats_for(pid: int):
                """(proc_name, cpu_pct, mem_mb, uptime_s); raises if gone."""
                seen_pids.add(pid)
                if pid not in pid_stats:
                    # Reuse cached Process object so cpu_percent() is meaningful
                    if pid not in self._proc_cache:
                        proc = psutil.Process(pid)
                        # First call initialises the baseline; returns 0.0
                        proc.cpu_percent(interval=None)
                        self._proc_cache[pid] = proc
                    proc = self._proc_cache[pid]
                    pid_stats[pid] = (
                        _proc_label(proc),
                        proc.cpu_percent(interval=None) / NUM_CORES,
                        proc.memory_info().rss / (1024 * 1024),
                        max(0.0, scan_t - proc.create_time()),
                    )
                return pid_stats[pid]

            titan_ctrls: list[int] = []

            def _cb(hwnd, _):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                try:
                    if win32gui.GetClassName(hwnd) == TITAN_CONTROLLER_CLASS:
                        titan_ctrls.append(hwnd)
                        return  # its tabs are handled by _scan_titan
                except Exception:
                    pass
                title = win32gui.GetWindowText(hwnd)
                if not title or not _is_interesting(title, self._keywords):
                    return
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    proc_name, cpu_pct, mem_mb, uptime_s = stats_for(pid)
                except Exception:
                    return
                if not _is_bot_process(proc_name):
                    return  # e.g. a browser tab titled "RuneScape Wiki"
                clients.append((hwnd, title, pid, proc_name, cpu_pct, mem_mb, uptime_s))

            win32gui.EnumWindows(_cb, None)
            titan = self._scan_titan(titan_ctrls, clients, stats_for)

            # Evict cache entries for pids that no longer own a matching window
            stale = set(self._proc_cache) - seen_pids
            for pid in stale:
                del self._proc_cache[pid]
            for pid in [p for p in self._titan_names if p not in seen_pids]:
                del self._titan_names[pid]

            self.titan_state.emit(titan)
            self.updated.emit(clients)
            # Sleep in small increments so we can stop quickly
            for _ in range(SCAN_INTERVAL_MS // 100):
                if not self._running:
                    break
                time.sleep(0.1)

    def _scan_titan(self, ctrls: list[int], clients: list, stats_for) -> dict:
        """One client per TitanClient tab (see TITAN_CONTROLLER_CLASS).

        The tab's hwnd is its JagWindow child; the pid is the osclient.exe
        process behind it. Tab names come from the controller's own title,
        which only names the *active* tab — so names are learned as the user
        (or Titan) cycles through tabs and remembered for the session."""
        info = {"controller_pid": 0, "active_pid": 0, "hidden_hwnds": [],
                "tabs": {}}  # tab hwnd → (controller hwnd, osclient pid)
        for ctrl in ctrls:
            try:
                _, ctrl_pid = win32process.GetWindowThreadProcessId(ctrl)
            except Exception:
                continue
            info["controller_pid"] = ctrl_pid
            name, active_pid = _titan_active_tab(ctrl)
            if active_pid:
                self._titan_names[active_pid] = name
            info["active_pid"] = active_pid

            tabs: list[int] = []

            def _child(hwnd, _):
                try:
                    if win32gui.GetClassName(hwnd) == TITAN_TAB_CLASS:
                        tabs.append(hwnd)
                except Exception:
                    pass
                return True

            try:
                win32gui.EnumChildWindows(ctrl, _child, None)
            except Exception:
                pass
            for hwnd in tabs:
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid == ctrl_pid:
                        continue
                    proc_name, cpu_pct, mem_mb, uptime_s = stats_for(pid)
                except Exception:
                    continue
                title = _titan_tab_title(self._titan_names.get(pid, ""), pid)
                clients.append((hwnd, title, pid, proc_name, cpu_pct, mem_mb, uptime_s))
                info["tabs"][hwnd] = (ctrl, pid)
                if pid != active_pid:
                    info["hidden_hwnds"].append(hwnd)
        return info

    def stop(self):
        self._running = False
        if not self.wait(2000):
            # Last resort at app exit: better than Qt aborting on a
            # still-running thread during teardown.
            self.terminate()
            self.wait(1000)


# ---------------------------------------------------------------------------
# Capturer thread — loops over known hwnds and captures screenshots
# ---------------------------------------------------------------------------
class Capturer(QThread):
    # hwnd, frame, seconds since the frame content last changed
    # (QImage not QPixmap — pixmaps are GUI-thread-only)
    captured = pyqtSignal(int, QImage, float)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = True
        self._paused = False
        self._hwnds: list[int] = []
        self._interval_ms = CAPTURE_INTERVAL_HIGH
        self._target_w = CARD_W_DEFAULT * 2
        self._target_h = CARD_H_DEFAULT * 2
        self._frame_state: dict[int, tuple[int, float]] = {}  # hwnd → (crc, last_change_t)
        self._titan_tabs: dict[int, tuple[int, int]] = {}  # hwnd → (ctrl hwnd, pid)

    def set_hwnds(self, hwnds: list[int]):
        self._hwnds = list(hwnds)

    def set_titan_tabs(self, tabs: dict[int, tuple[int, int]]):
        """TitanClient tab hwnds → (controller hwnd, osclient pid). A tab's
        JagWindow only shows the game while it is the controller's active
        tab; at any other moment PrintWindow returns whatever is on screen
        (the Home page, another tab…). The Scanner only re-checks every 3 s,
        so the capture loop verifies the active tab itself, right around
        each PrintWindow, and drops the frame if the tab isn't (still) it."""
        self._titan_tabs = dict(tabs)

    def set_interval(self, ms: int):
        self._interval_ms = ms

    def set_card_size(self, w: int, h: int):
        # Frames are downscaled to 2x the card size here in the capture
        # thread, so the GUI never stores or scales full-resolution captures
        # (a 1080p client is ~8 MB/frame; 2x card size is ~0.3 MB). The 2x
        # headroom keeps cards sharp through a zoom step.
        self._target_w = max(1, w * 2)
        self._target_h = max(1, h * 2)

    def set_paused(self, paused: bool):
        """Skip captures entirely (e.g. while BotWall itself is minimized)."""
        self._paused = paused

    def run(self):
        # Deadline-based pacing: the interval is the full cycle period, not a
        # sleep appended after the work, so the refresh rate stays honest as
        # long as one pass fits in the interval.
        next_t = time.monotonic()
        while self._running:
            if not self._paused:
                for hwnd in list(self._hwnds):
                    if not self._running:
                        break
                    tab = self._titan_tabs.get(hwnd)
                    if tab is not None and _titan_active_tab(tab[0])[1] != tab[1]:
                        continue  # not the active Titan tab: nothing to see
                    img = capture_hwnd(hwnd)
                    if tab is not None and _titan_active_tab(tab[0])[1] != tab[1]:
                        continue  # tab switched mid-capture: frame is junk
                    if img is not None:
                        if (img.width() > self._target_w
                                or img.height() > self._target_h):
                            img = img.scaled(
                                self._target_w, self._target_h,
                                Qt.KeepAspectRatio, Qt.SmoothTransformation
                            )
                        self.captured.emit(hwnd, img, self._frame_age(hwnd, img))
                # Drop change-tracking state for windows we no longer capture
                alive = set(self._hwnds)
                for hwnd in [h for h in self._frame_state if h not in alive]:
                    del self._frame_state[hwnd]
            next_t += self._interval_ms / 1000.0
            now = time.monotonic()
            if next_t <= now:
                next_t = now  # pass overran the interval: restart, don't spiral
            while self._running:
                remaining = next_t - time.monotonic()
                if remaining <= 0:
                    break
                # Sleep in small chunks so stop() stays responsive
                time.sleep(min(remaining, 0.1))

    def _frame_age(self, hwnd: int, img: QImage) -> float:
        """Seconds since this window's frame content last changed."""
        ptr = img.constBits()
        ptr.setsize(img.byteCount())
        crc = zlib.crc32(ptr)
        now = time.monotonic()
        prev = self._frame_state.get(hwnd)
        if prev is None or prev[0] != crc:
            self._frame_state[hwnd] = (crc, now)
            return 0.0
        return now - prev[1]

    def stop(self):
        self._running = False
        if not self.wait(3000):
            # Likely stuck inside PrintWindow on a hung client. Better than
            # Qt aborting on a still-running thread during teardown.
            self.terminate()
            self.wait(1000)


# ---------------------------------------------------------------------------
# ClientCard — one card per detected client window
# ---------------------------------------------------------------------------
class ClientCard(QFrame):
    pin_toggled        = pyqtSignal(int, bool)     # hwnd, is_pinned
    minimize_requested = pyqtSignal(int)           # hwnd
    kill_requested     = pyqtSignal(int, str, str) # pid, title, proc_name
    restart_requested  = pyqtSignal(int, str, str) # pid, title, proc_name
    nickname_changed   = pyqtSignal(int, str)      # hwnd, new nickname ("" clears)

    def __init__(self, hwnd: int, title: str, pid: int, proc_name: str,
                 cpu_pct: float = 0.0, mem_mb: float = 0.0,
                 uptime_s: float = 0.0, parent=None):
        super().__init__(parent)
        self.hwnd = hwnd
        self.pid = pid
        self.proc_name = proc_name
        self.title = title
        self.kind = _client_kind(title, proc_name)
        self.script = _parse_script(title)
        self.nickname = ""
        self._cpu_pct = cpu_pct
        self._mem_mb  = mem_mb
        self._uptime_s = uptime_s

        self._low_cpu = False
        self._pinned = False
        # TitanClient tab that isn't the controller's active tab: no frames
        # can be captured, so the last one is kept (dimmed) and freeze /
        # no-feed detection is suspended until the tab is active again.
        self._hidden_tab = False

        self.setFrameShape(QFrame.NoFrame)
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self._set_border(None)

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

        self._status_lbl = QLabel()
        self._status_lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 9px; font-weight: bold;")
        self._status_lbl.hide()

        self._pid_lbl = QLabel()
        self._pid_lbl.setStyleSheet(f"color: {ACCENT_TEAL}; font-size: 10px;")

        self._pin_btn = QPushButton("📌")
        self._pin_btn.setFixedSize(20, 20)
        self._pin_btn.setFlat(True)
        self._pin_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self._pin_btn.setToolTip("Pin to top")
        self._pin_btn.clicked.connect(self._toggle_pin)

        h_layout.addWidget(self._title_lbl)
        h_layout.addWidget(self._status_lbl)
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

        self._update_labels()
        self._update_stats(cpu_pct, mem_mb)
        self._pixmap_raw: QPixmap | None = None
        self._last_frame_ts = time.monotonic()
        self._status_state: tuple | None = None
        self._update_pin_visual()

    # ------------------------------------------------------------------
    def _toggle_pin(self):
        self._pinned = not self._pinned
        self._update_pin_visual()
        self.pin_toggled.emit(self.hwnd, self._pinned)

    def set_pinned(self, pinned: bool):
        """Set pin state without emitting (used when restoring persisted pins)."""
        self._pinned = pinned
        self._update_pin_visual()

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

    def _update_labels(self):
        display = self.nickname or self.title
        max_chars = 26
        short = display if len(display) <= max_chars else display[:max_chars - 1] + "…"
        self._title_lbl.setText(short)
        tooltip = self.title if not self.nickname else f"{self.nickname}\n{self.title}"
        self._title_lbl.setToolTip(tooltip)
        if self._uptime_s > 0:
            self._pid_lbl.setText(f"PID {self.pid} · {_fmt_uptime(self._uptime_s)}")
        else:
            self._pid_lbl.setText(f"PID {self.pid}")
        self._pid_lbl.setToolTip("PID · process uptime (low uptime = recently restarted)")

    def set_nickname(self, nickname: str):
        self.nickname = nickname
        self._update_labels()

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
                    cpu_pct: float = 0.0, mem_mb: float = 0.0,
                    uptime_s: float = 0.0):
        self.pid = pid
        self.proc_name = proc_name
        self.title = title
        self.kind = _client_kind(title, proc_name)
        self.script = _parse_script(title)
        self._uptime_s = uptime_s
        self._update_labels()
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
        menu.addAction("Set Nickname…", self._edit_nickname)
        menu.addAction("Copy PID", self._copy_pid)
        menu.addSeparator()
        restart = menu.addAction(
            "Restart Client…",
            lambda: self.restart_requested.emit(self.pid, self.title, self.proc_name)
        )
        if self.kind == TITAN_KIND:
            # osclient.exe is spawned and injected by the Titan controller;
            # relaunching it by command line just yields an orphan process.
            restart.setEnabled(False)
            restart.setText("Restart Client (use TitanClient's Launch New Client)")
        menu.addAction(
            "Kill This Client…",
            lambda: self.kill_requested.emit(self.pid, self.title, self.proc_name)
        )
        menu.exec_(event.globalPos())

    def _copy_pid(self):
        QApplication.clipboard().setText(str(self.pid))

    def _edit_nickname(self):
        text, ok = QInputDialog.getText(
            self, "Set Nickname",
            "Nickname for this client (empty to clear):",
            QLineEdit.Normal, self.nickname
        )
        if ok:
            self.set_nickname(text.strip())
            self.nickname_changed.emit(self.hwnd, self.nickname)

    def _bring_to_front(self):
        # A Titan tab is a child window; raise its controller instead.
        hwnd = _root_hwnd(self.hwnd)
        try:
            placement = win32gui.GetWindowPlacement(hwnd)
            if placement[1] == win32con.SW_SHOWMINIMIZED:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def set_hidden_tab(self, hidden: bool):
        if hidden == self._hidden_tab:
            return
        self._hidden_tab = hidden
        if hidden:
            self._apply_status("TAB HIDDEN", DIM_COLOR)
        else:
            # Frames resume now; don't let check_feed flag the gap as NO FEED
            self._last_frame_ts = time.monotonic()
            self._apply_status(None, None)
        self._rescale()

    def update_pixmap(self, pixmap: QPixmap, age_s: float = 0.0):
        if self._hidden_tab:
            return  # in-flight capture of a tab that just went hidden: junk
        self._pixmap_raw = pixmap
        self._last_frame_ts = time.monotonic()
        self._set_freshness(age_s)
        # Cards on an inactive tab keep receiving frames (freeze detection
        # stays farm-wide) but there's no point scaling pixmaps nobody sees —
        # showEvent below catches up when the tab becomes active.
        if self.isVisible():
            self._rescale()

    # ------------------------------------------------------------------
    # Freshness / freeze indication
    # ------------------------------------------------------------------
    def _set_freshness(self, age_s: float):
        if age_s >= STALE_FROZEN_S:
            self._apply_status(f"FROZEN {_fmt_age(age_s)}", ACCENT_RED)
        elif age_s >= STALE_WARN_S:
            self._apply_status(f"static {_fmt_age(age_s)}", "#f0a040")
        else:
            self._apply_status(None, None)

    def check_feed(self):
        """Called once per scan: flag cards whose captures stopped arriving
        (window OS-minimized, or PrintWindow failing)."""
        if self._hidden_tab:
            return  # expected: nothing is captured for an inactive Titan tab
        if time.monotonic() - self._last_frame_ts > NO_FEED_S:
            self._apply_status("NO FEED", DIM_COLOR)

    def _apply_status(self, text: str | None, color: str | None):
        state = (text, color)
        if state == self._status_state:
            return
        self._status_state = state
        if text is None:
            self._status_lbl.hide()
        else:
            self._status_lbl.setText(text)
            self._status_lbl.setStyleSheet(
                f"color: {color}; font-size: 9px; font-weight: bold;"
            )
            self._status_lbl.show()
            self._status_lbl.setToolTip(
                "Not the active TitanClient tab — only the active tab renders, "
                "so this shows the last frame seen. Process stats stay live."
                if text == "TAB HIDDEN" else
                "Frame content has not changed — client may be stuck or logged out"
                if color != DIM_COLOR else
                "No captures arriving (window minimized or capture failing)"
            )
        self._set_border(color)

    def _set_border(self, color: str | None):
        border = f"1px solid {color}" if color else "1px solid transparent"
        self.setStyleSheet(f"""
            ClientCard {{
                background: {CARD_COLOR};
                border-radius: 4px;
                border: {border};
            }}
        """)

    def _rescale(self):
        if self._pixmap_raw is None:
            return
        target_w = self._img_lbl.width()
        target_h = self._img_lbl.height()
        if target_w <= 0 or target_h <= 0:
            return

        transform = Qt.FastTransformation if self._low_cpu else Qt.SmoothTransformation
        scaled = self._pixmap_raw.scaled(target_w, target_h, Qt.KeepAspectRatio, transform)
        if self._low_cpu:
            # Grayscale the card-sized result, not the full-res frame —
            # converting first costs more than the smooth scaling it replaces
            gray = scaled.toImage().convertToFormat(QImage.Format_Grayscale8)
            scaled = QPixmap.fromImage(gray)
        if self._hidden_tab:
            # Dim the stale frame so it can't pass for a live one at a glance
            scaled = QPixmap(scaled)
            painter = QPainter(scaled)
            painter.fillRect(scaled.rect(), QColor(0, 0, 0, 150))
            painter.end()
        self._img_lbl.setPixmap(scaled)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._rescale()

    def showEvent(self, event):
        super().showEvent(event)
        # Fixed-size cards get no resizeEvent on re-show, so scale the last
        # frame received while the card was hidden on an inactive tab.
        self._rescale()


# ---------------------------------------------------------------------------
# Empty-state placeholder
# ---------------------------------------------------------------------------
class EmptyPlaceholder(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        self._lbl = QLabel()
        self._lbl.setAlignment(Qt.AlignCenter)
        self._lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 15px;")
        layout.addWidget(self._lbl)
        sub = QLabel("Launch a client and it will appear here automatically.")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet(f"color: {DIM_COLOR}; font-size: 11px;")
        layout.addWidget(sub)
        self.set_kind(None)

    def set_kind(self, kind: str | None):
        """Match the empty-state text to the active tab."""
        self._lbl.setText(f"No {kind} clients detected" if kind
                          else "No clients detected")


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
    card_size_changed = pyqtSignal(int, int)                     # card w, h (zoom)
    client_kill_requested = pyqtSignal(int, str, str)            # pid, title, proc_name
    client_restart_requested = pyqtSignal(int, str, str)         # pid, title, proc_name
    nicknames_changed = pyqtSignal()                             # persist trigger

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
        self._pinned_titles: set[str] = set()           # survives client restarts
        self._nicknames: dict[str, str] = {}            # title → nickname
        self._minimized: set[int] = set()               # minimized hwnds
        self._stats: dict[int, tuple[float, float]] = {}# hwnd → (cpu_pct, mem_mb)
        self._sort_mode = "default"
        self._filter_kind: str | None = None            # None = All tab
        self._filter_script: str | None = None          # None = all scripts
        self._placeholder: EmptyPlaceholder | None = None
        self._low_cpu = False
        self._last_layout: tuple | None = None      # skip no-op relayouts
        self._last_stretch = (0, 0)                 # (row, col) stretch to clear
        self._show_placeholder()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def update_clients(self, clients: list[tuple]):
        """Called from main thread with fresh 7-tuple list (see Scanner)."""
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
        for hwnd, title, pid, proc_name, cpu_pct, mem_mb, uptime_s in clients:
            self._stats[hwnd] = (cpu_pct, mem_mb)
            if hwnd not in self._cards:
                card = ClientCard(hwnd, title, pid, proc_name, cpu_pct, mem_mb, uptime_s)
                card.set_low_cpu(self._low_cpu)
                card.pin_toggled.connect(self._on_pin_toggled)
                card.minimize_requested.connect(self._on_minimize_requested)
                card.kill_requested.connect(self.client_kill_requested)
                card.restart_requested.connect(self.client_restart_requested)
                card.nickname_changed.connect(self._on_nickname_changed)
                self._fix_card_size(card)
                self._cards[hwnd] = card
                self._order.append(hwnd)
                # Pins and nicknames are persisted by title so they survive
                # client restarts (hwnds don't)
                if title in self._pinned_titles:
                    card.set_pinned(True)
                    self._pinned.add(hwnd)
                if title in self._nicknames:
                    card.set_nickname(self._nicknames[title])
            else:
                self._cards[hwnd].update_info(
                    title, pid, proc_name, cpu_pct, mem_mb, uptime_s
                )

        self._relayout()
        self._update_placeholder()

    def set_filter_kind(self, kind: str | None):
        """Show only cards of one client kind (None = all). Filtered-out
        clients keep their captures and freeze detection — only display
        changes."""
        if kind == self._filter_kind:
            return
        self._filter_kind = kind
        self._last_layout = None  # visibility changes even if order matches
        self._relayout()
        self._update_placeholder()

    def set_filter_script(self, script: str | None):
        """Show only cards running one script (None = all). Display-only,
        same as set_filter_kind — captures keep running for hidden cards."""
        if script == self._filter_script:
            return
        self._filter_script = script
        self._last_layout = None
        self._relayout()
        self._update_placeholder()

    def _matches_filter(self, card: ClientCard) -> bool:
        return ((self._filter_kind is None or card.kind == self._filter_kind)
                and (self._filter_script is None
                     or card.script == self._filter_script))

    def _update_placeholder(self):
        if self._sorted_order():
            self._hide_placeholder()
        else:
            self._show_placeholder()

    def update_screenshot(self, hwnd: int, pixmap: QPixmap, age_s: float = 0.0):
        if hwnd in self._cards:
            self._cards[hwnd].update_pixmap(pixmap, age_s)

    def check_feeds(self):
        """Flag visible cards whose captures have stopped arriving."""
        for hwnd, card in self._cards.items():
            if hwnd not in self._minimized:
                card.check_feed()

    def set_hidden_tabs(self, hwnds: set[int]):
        """Mark inactive TitanClient tabs (see ClientCard.set_hidden_tab)."""
        for hwnd, card in self._cards.items():
            card.set_hidden_tab(hwnd in hwnds)

    def zoom(self, factor: float):
        self._card_w = max(160, min(1600, int(self._card_w * factor)))
        self._card_h = max(110, min(1100, int(self._card_h * factor)))
        for card in self._cards.values():
            self._fix_card_size(card)
        self._relayout()
        self.card_size_changed.emit(self._card_w, self._card_h)

    def card_size(self) -> tuple[int, int]:
        return (self._card_w, self._card_h)

    def set_card_size(self, w: int, h: int):
        self._card_w = max(160, min(1600, w))
        self._card_h = max(110, min(1100, h))
        for card in self._cards.values():
            self._fix_card_size(card)
        self._relayout()
        self.card_size_changed.emit(self._card_w, self._card_h)

    def set_low_cpu(self, enabled: bool):
        self._low_cpu = enabled
        for card in self._cards.values():
            card.set_low_cpu(enabled)

    def _on_pin_toggled(self, hwnd: int, is_pinned: bool):
        card = self._cards.get(hwnd)
        if is_pinned:
            self._pinned.add(hwnd)
            if card:
                self._pinned_titles.add(card.title)
        else:
            self._pinned.discard(hwnd)
            if card:
                self._pinned_titles.discard(card.title)
        self._relayout()

    def set_pinned_titles(self, titles: set[str]):
        self._pinned_titles = set(titles)
        for hwnd, card in self._cards.items():
            if card.title in self._pinned_titles and hwnd not in self._pinned:
                card.set_pinned(True)
                self._pinned.add(hwnd)
        self._relayout()

    def pinned_titles(self) -> set[str]:
        return set(self._pinned_titles)

    def _on_nickname_changed(self, hwnd: int, nickname: str):
        card = self._cards.get(hwnd)
        if card is None:
            return
        if nickname:
            self._nicknames[card.title] = nickname
        else:
            self._nicknames.pop(card.title, None)
        self.nicknames_changed.emit()

    def set_nicknames(self, nicknames: dict[str, str]):
        self._nicknames = dict(nicknames)
        for card in self._cards.values():
            card.set_nickname(self._nicknames.get(card.title, ""))

    def nicknames(self) -> dict[str, str]:
        return dict(self._nicknames)

    def _on_minimize_requested(self, hwnd: int):
        if hwnd in self._cards and hwnd not in self._minimized:
            self._minimized.add(hwnd)
            card = self._cards[hwnd]
            self._grid.removeWidget(card)
            card.hide()
            cpu_pct, mem_mb = self._stats.get(hwnd, (0.0, 0.0))
            self.client_minimized.emit(hwnd, card.title, card.pid, cpu_pct, mem_mb)
            self._relayout()
            self._update_placeholder()

    def restore_client(self, hwnd: int):
        if hwnd in self._minimized:
            self._minimized.discard(hwnd)
            # _relayout shows the card — unless it belongs to another tab,
            # in which case it stays hidden until that tab is selected.
            self._last_layout = None
            self._relayout()
            self._update_placeholder()

    def set_sort_mode(self, mode: str):
        self._sort_mode = mode
        self._relayout()

    def _sorted_order(self) -> list[int]:
        visible  = [h for h in self._order if h not in self._minimized
                    and self._matches_filter(self._cards[h])]
        pinned   = [h for h in visible if h in self._pinned]
        unpinned = [h for h in visible if h not in self._pinned]
        if self._sort_mode != "default":
            key = self._sort_key()
            pinned.sort(key=key)
            unpinned.sort(key=key)
        return pinned + unpinned

    def _sort_key(self):
        """Return a sort key function for hwnd based on current sort mode.

        Values are bucketed (CPU into 10%-steps, RAM into 200 MB-steps) so
        small fluctuations between scans don't reshuffle the grid under the
        cursor every 3 s — ties keep their previous order (stable sort).
        """
        if self._sort_mode == "cpu_asc":
            return lambda h: int(self._stats.get(h, (0.0, 0.0))[0] // 10)
        if self._sort_mode == "cpu_desc":
            return lambda h: -int(self._stats.get(h, (0.0, 0.0))[0] // 10)
        if self._sort_mode == "ram_asc":
            return lambda h: int(self._stats.get(h, (0.0, 0.0))[1] // 200)
        if self._sort_mode == "ram_desc":
            return lambda h: -int(self._stats.get(h, (0.0, 0.0))[1] // 200)
        return lambda h: 0

    def all_pids(self, kind: str | None = None,
                 script: str | None = None) -> list[int]:
        # Deduplicated: one process can own several client windows.
        # kind/script narrow the set so KILL ALL matches what's on screen
        # (active tab + script dropdown).
        return sorted({
            c.pid for c in self._cards.values()
            if c.pid
            and (kind is None or c.kind == kind)
            and (script is None or c.script == script)
        })

    def minimized_hwnds(self) -> set[int]:
        return set(self._minimized)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _fix_card_size(self, card: ClientCard):
        card.setFixedSize(self._card_w, self._card_h)

    def _relayout(self):
        vp_w = self.viewport().width() - 20  # subtract margins
        cols = max(1, vp_w // (self._card_w + self._grid.spacing()))
        display_order = self._sorted_order()

        signature = (tuple(display_order), cols, self._card_w, self._card_h)
        if signature == self._last_layout:
            return  # nothing moved — skip the remove/re-add churn
        self._last_layout = signature

        # Remove all cards from grid without deleting (placeholder stays put —
        # pulling it out of the layout would orphan it uncentered)
        for i in reversed(range(self._grid.count())):
            item = self._grid.itemAt(i)
            w = item.widget() if item else None
            if w is not None and w is not self._placeholder:
                self._grid.removeWidget(w)

        for idx, hwnd in enumerate(display_order):
            card = self._cards[hwnd]
            row, col = divmod(idx, cols)
            self._grid.addWidget(card, row, col)

        # Cards filtered out by the active tab were pulled from the layout
        # above but would still paint parented at the container's origin —
        # hide them; show cards (re-)entering the visible set.
        shown = set(display_order)
        for hwnd, card in self._cards.items():
            if hwnd in self._minimized:
                continue  # shelf handles these
            card.setVisible(hwnd in shown)

        # Clear the previous stretch before setting the new one, or stale
        # stretches accumulate and split the free space with phantom rows/cols
        old_row, old_col = self._last_stretch
        self._grid.setRowStretch(old_row, 0)
        self._grid.setColumnStretch(old_col, 0)
        stretch_row = len(display_order) // cols + 1
        self._grid.setRowStretch(stretch_row, 1)
        self._grid.setColumnStretch(cols, 1)
        self._last_stretch = (stretch_row, cols)

    def _show_placeholder(self):
        if self._placeholder is None:
            self._placeholder = EmptyPlaceholder()
            self._grid.addWidget(self._placeholder, 0, 0)
        self._placeholder.set_kind(self._filter_kind)

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
        self._titan_hidden: set[int] = set()   # inactive Titan tab hwnds
        self._titan_controller_pid = 0
        self._total_opens = 0
        self._total_closes = 0
        self._active_hwnds: dict[int, str] = {}  # hwnd → title for open clients
        self._events: list[str] = []
        self._alert_words: list[str] = []
        self._scan_keywords: list[str] = list(KEYWORDS)
        self._low_cpu_active = False
        self._active_kind: str | None = None       # None = "All" tab
        self._kind_counts: dict[str, int] = {}      # kind → client count
        self._card_sizes: dict[str, tuple[int, int]] = {}  # tab ("" = All) → (w, h)
        self._tb_measured = False  # toolbar chrome width known yet?
        self._tb_fixed_w = 0       # width of the never-hidden toolbar chrome
        self._self_proc = psutil.Process()
        self._self_proc.cpu_percent(interval=None)  # prime the baseline
        self._settings = QSettings("BotWall", "BotWall")
        self._setup_ui()
        self._setup_tray()
        self._start_threads()
        self._restore_settings()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _setup_ui(self):
        # ---- Toolbar ----
        # The toolbar is by far the widest thing in the window, so it is built
        # as a row of collapsible *groups*. When the window gets too narrow,
        # groups are hidden lowest-priority-first and their controls move into
        # the ⋯ overflow menu; otherwise the layout's natural minimum (~1300 px)
        # is a hard floor on how narrow the window can be dragged, which makes
        # a half-monitor split impossible.
        toolbar_widget = self._toolbar_widget = QWidget()
        toolbar_widget.setFixedHeight(54)
        toolbar_widget.setStyleSheet(f"background: {TOOLBAR_COLOR};")
        # Qt clamps a resize at the layout's minimum, so without an explicit
        # override the window would never get narrow enough for _reflow_toolbar
        # to be asked to do anything.
        toolbar_widget.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        toolbar_widget.setMinimumWidth(0)
        tb_layout = QHBoxLayout(toolbar_widget)
        tb_layout.setContentsMargins(12, 0, 12, 0)
        tb_layout.setSpacing(TB_SPACING)

        def _vsep():
            s = QFrame()
            s.setFrameShape(QFrame.VLine)
            s.setFixedHeight(22)
            s.setStyleSheet(f"color: {DIM_COLOR};")
            return s

        def _small_btn(text: str, tip: str) -> QPushButton:
            b = QPushButton(text)
            b.setFixedSize(28, 28)
            b.setCursor(QCursor(Qt.PointingHandCursor))
            b.setToolTip(tip)
            b.setStyleSheet(f"""
                QPushButton {{
                    background: transparent;
                    color: {DIM_COLOR};
                    border: 1px solid {DIM_COLOR};
                    border-radius: 3px;
                    font-size: 13px;
                    font-weight: bold;
                }}
                QPushButton:hover {{ border-color: {ACCENT_TEAL}; color: {TEXT_COLOR}; }}
                QPushButton:pressed {{ background: #1a3a40; }}
                QPushButton::menu-indicator {{ image: none; width: 0; }}
            """)
            return b

        def _flat_btn(text: str, tip: str) -> QPushButton:
            b = QPushButton(text)
            b.setFixedHeight(28)
            b.setCursor(QCursor(Qt.PointingHandCursor))
            b.setToolTip(tip)
            b.setStyleSheet(f"""
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
            return b

        def _group():
            """A hide-as-one section of the toolbar."""
            w = QWidget()
            lay = QHBoxLayout(w)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(TB_SPACING)
            return w, lay

        title_lbl = QLabel("BotWall")
        title_lbl.setStyleSheet(
            f"color: {ACCENT_TEAL}; font-size: 18px; font-weight: bold; letter-spacing: 1px;"
        )
        tb_layout.addWidget(title_lbl)

        # ---- Group: BotWall self-stats ----
        g_stats, g_stats_lay = _group()
        g_stats_lay.addSpacing(4)
        g_stats_lay.addWidget(_vsep())
        g_stats_lay.addSpacing(8)

        self._self_cpu_lbl = QLabel("CPU: –")
        self._self_cpu_lbl.setStyleSheet("color: #5bbde8; font-size: 12px;")
        self._self_cpu_lbl.setToolTip("BotWall CPU usage")
        g_stats_lay.addWidget(self._self_cpu_lbl)

        g_stats_lay.addSpacing(10)

        self._self_mem_lbl = QLabel("MEM: –")
        self._self_mem_lbl.setStyleSheet("color: #b07ed8; font-size: 12px;")
        self._self_mem_lbl.setToolTip("BotWall memory usage")
        g_stats_lay.addWidget(self._self_mem_lbl)
        tb_layout.addWidget(g_stats)

        # ---- Group: session client stats ----
        g_session, g_session_lay = _group()
        g_session_lay.addSpacing(4)
        g_session_lay.addWidget(_vsep())
        g_session_lay.addSpacing(8)

        self._opens_lbl = QLabel("↑ 0 Opened")
        self._opens_lbl.setStyleSheet("color: #4dc87a; font-size: 12px;")
        self._opens_lbl.setToolTip("Clients opened this session")
        g_session_lay.addWidget(self._opens_lbl)

        g_session_lay.addSpacing(10)

        self._closes_lbl = QLabel("↓ 0 Closed")
        self._closes_lbl.setStyleSheet("color: #e07050; font-size: 12px;")
        self._closes_lbl.setToolTip("Clients closed this session")
        g_session_lay.addWidget(self._closes_lbl)

        g_session_lay.addSpacing(10)
        g_session_lay.addWidget(_vsep())
        g_session_lay.addSpacing(8)

        self._count_lbl = QLabel("0 clients")
        self._count_lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 12px;")
        self._count_lbl.setToolTip("Active bot clients detected")
        g_session_lay.addWidget(self._count_lbl)
        tb_layout.addWidget(g_session)

        tb_layout.addStretch()

        # ---- Group: view controls (sort / zoom / capture rate) ----
        g_view, g_view_lay = _group()

        sort_combo = self._sort_combo = QComboBox()
        sort_combo.addItems(["Sort: Default", "CPU ↑", "CPU ↓", "RAM ↑", "RAM ↓"])
        sort_combo.setFixedHeight(28)
        sort_combo.setFixedWidth(120)
        sort_combo.setCursor(QCursor(Qt.PointingHandCursor))
        sort_combo.setStyleSheet(COMBO_STYLE)
        sort_combo.currentIndexChanged.connect(self._on_sort_changed)
        g_view_lay.addWidget(sort_combo)
        g_view_lay.addSpacing(8)

        # ---- Zoom controls (also available via Ctrl+wheel / Ctrl+= / Ctrl+- / Ctrl+0) ----
        zoom_out_btn = _small_btn("−", "Smaller cards (Ctrl+-)")
        zoom_out_btn.clicked.connect(lambda: self._grid_view.zoom(1.0 / ZOOM_FACTOR))
        zoom_reset_btn = _small_btn("⤢", "Reset card size (Ctrl+0)")
        zoom_reset_btn.clicked.connect(
            lambda: self._grid_view.set_card_size(CARD_W_DEFAULT, CARD_H_DEFAULT)
        )
        zoom_in_btn = _small_btn("+", "Larger cards (Ctrl+=)")
        zoom_in_btn.clicked.connect(lambda: self._grid_view.zoom(ZOOM_FACTOR))
        g_view_lay.addWidget(zoom_out_btn)
        g_view_lay.addWidget(zoom_reset_btn)
        g_view_lay.addWidget(zoom_in_btn)
        g_view_lay.addSpacing(8)

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

        g_view_lay.addWidget(self._btn_high_cpu)
        g_view_lay.addWidget(self._btn_low_cpu)
        tb_layout.addWidget(g_view)

        # ---- Group: secondary actions ----
        g_extra, g_extra_lay = _group()
        g_extra_lay.addSpacing(2)

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
        discord_btn.clicked.connect(self._open_discord)
        g_extra_lay.addWidget(discord_btn)
        g_extra_lay.addSpacing(8)

        maximize_all_btn = _flat_btn(
            "Maximize All", "Maximize all client windows in the current tab"
        )
        maximize_all_btn.clicked.connect(self._maximize_all)
        g_extra_lay.addWidget(maximize_all_btn)

        restore_all_btn = _flat_btn(
            "Restore All", "Restore all client windows in the current tab"
        )
        restore_all_btn.clicked.connect(self._restore_all)
        g_extra_lay.addWidget(restore_all_btn)

        log_btn = _flat_btn(
            "Log", "Session event log (opens, closes, title changes)"
        )
        log_btn.clicked.connect(self._show_log)
        g_extra_lay.addWidget(log_btn)
        tb_layout.addWidget(g_extra)

        # ---- Overflow menu (holds whatever the current width squeezed out) ----
        self._overflow_btn = _small_btn("⋯", "Controls hidden by the window width")
        self._overflow_menu = QMenu(self)
        self._overflow_menu.setStyleSheet(MENU_STYLE)
        self._overflow_menu.aboutToShow.connect(self._build_overflow_menu)
        self._overflow_btn.setMenu(self._overflow_menu)
        tb_layout.addWidget(self._overflow_btn)

        settings_btn = _small_btn("⚙", "Settings: scan keywords, alert words")
        settings_btn.clicked.connect(self._show_settings)
        tb_layout.addWidget(settings_btn)
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
        kill_btn.setToolTip("Kill every client process in the current tab")
        kill_btn.clicked.connect(self._kill_all)
        tb_layout.addWidget(kill_btn)

        # Hide order as the window narrows — first entry goes first. BotWall's
        # own CPU/MEM is the least useful thing here; the farm totals in
        # `session` are the most, so they survive longest. KILL ALL, ⚙ and ⋯
        # are never collapsible, so the toolbar always stays usable.
        self._tb_groups = [
            ("stats",   g_stats),
            ("extra",   g_extra),
            ("view",    g_view),
            ("title",   title_lbl),
            ("session", g_session),
        ]

        # ---- Client-type tabs (All / DreamBot / TwiLite / …) ----
        tabs_widget = QWidget()
        tabs_widget.setFixedHeight(36)
        tabs_widget.setStyleSheet(f"background: {TOOLBAR_COLOR};")
        tabs_layout = QHBoxLayout(tabs_widget)
        tabs_layout.setContentsMargins(12, 0, 12, 6)
        tabs_layout.setSpacing(6)

        self._tab_buttons: dict[str | None, QPushButton] = {}

        def _make_tab(kind: str | None):
            b = QPushButton(kind or "All")
            b.setFixedHeight(26)
            b.setCursor(QCursor(Qt.PointingHandCursor))
            b.clicked.connect(lambda _, k=kind: self._set_active_kind(k))
            tabs_layout.addWidget(b)
            self._tab_buttons[kind] = b

        _make_tab(None)
        for kind in (*CLIENT_KINDS, OTHER_KIND):
            _make_tab(kind)
        tabs_layout.addStretch()

        # ---- Script filter (DreamBot clients run different scripts) ----
        self._script_lbl = QLabel("Script:")
        self._script_lbl.setStyleSheet(f"color: {DIM_COLOR}; font-size: 11px;")
        tabs_layout.addWidget(self._script_lbl)

        self._script_combo = QComboBox()
        self._script_combo.setFixedHeight(24)
        self._script_combo.setMinimumWidth(150)
        self._script_combo.setCursor(QCursor(Qt.PointingHandCursor))
        self._script_combo.setToolTip("Filter DreamBot clients by the script in their title")
        self._script_combo.setStyleSheet(COMBO_STYLE)
        self._script_combo.currentIndexChanged.connect(self._on_script_changed)
        tabs_layout.addWidget(self._script_combo)
        # Hidden until a scan finds DreamBot clients with parseable scripts
        self._script_lbl.setVisible(False)
        self._script_combo.setVisible(False)

        # ---- Toolbar collapse toggle ----
        self._collapse_btn = QPushButton("▴")
        self._collapse_btn.setFixedSize(24, 24)
        self._collapse_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self._collapse_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {DIM_COLOR};
                border: 1px solid transparent;
                border-radius: 3px;
                font-size: 11px;
            }}
            QPushButton:hover {{ border-color: {DIM_COLOR}; color: {TEXT_COLOR}; }}
        """)
        self._collapse_btn.clicked.connect(
            lambda: self._set_toolbar_collapsed(self._toolbar_widget.isVisible())
        )
        tabs_layout.addWidget(self._collapse_btn)
        self._set_toolbar_collapsed(False)

        self._refresh_tabs()

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
        main_layout.addWidget(tabs_widget)
        main_layout.addWidget(self._grid_view)
        main_layout.addWidget(self._minimized_shelf)
        self.setCentralWidget(central)

        # ---- Keyboard zoom (mirrors the toolbar buttons) ----
        for seq in ("Ctrl+=", "Ctrl++"):
            QShortcut(QKeySequence(seq), self).activated.connect(
                lambda: self._grid_view.zoom(ZOOM_FACTOR)
            )
        QShortcut(QKeySequence("Ctrl+-"), self).activated.connect(
            lambda: self._grid_view.zoom(1.0 / ZOOM_FACTOR)
        )
        QShortcut(QKeySequence("Ctrl+0"), self).activated.connect(
            lambda: self._grid_view.set_card_size(CARD_W_DEFAULT, CARD_H_DEFAULT)
        )

    # ------------------------------------------------------------------
    # Responsive toolbar
    # ------------------------------------------------------------------
    def _tb_widths(self) -> dict[str, int]:
        """Width each group needs right now, incl. the gap it would free up.

        Measured live rather than cached: the stats labels start as "0 clients"
        / "CPU: –" and grow by ~150 px once real numbers arrive, so a width
        cached at construction time makes the toolbar overflow and clip."""
        return {
            name: w.minimumSizeHint().width() + TB_SPACING
            for name, w in self._tb_groups
        }

    def _measure_toolbar(self):
        """Measure the chrome that is never hidden (title through KILL ALL).

        Must run with every group visible, otherwise the remainder comes out
        too small. Unlike the groups this stays fixed, so it's measured once."""
        for _name, w in self._tb_groups:
            w.setVisible(True)
        self._overflow_btn.setVisible(True)
        self._toolbar_widget.layout().activate()
        full = self._toolbar_widget.minimumSizeHint().width()
        self._tb_fixed_w = full - sum(self._tb_widths().values())
        self._tb_measured = True

    def _reflow_toolbar(self):
        """Show as many groups as fit; the rest move to the ⋯ menu."""
        if not self._tb_measured:
            return
        widths = self._tb_widths()
        avail = self.width() - TB_REFLOW_SLACK
        used = self._tb_fixed_w
        shown = set()
        # Walk highest-priority first — _tb_groups is in hide order. Stop at the
        # first group that doesn't fit rather than squeezing in a smaller one
        # behind it, so groups always vanish in the same order and the toolbar
        # doesn't reshuffle as the window is dragged.
        for name, _w in reversed(self._tb_groups):
            cost = widths[name]
            if used + cost > avail:
                break
            used += cost
            shown.add(name)
        for name, w in self._tb_groups:
            w.setVisible(name in shown)
        self._overflow_btn.setVisible(len(shown) < len(self._tb_groups))

    def _build_overflow_menu(self):
        """Populate ⋯ with exactly the controls the current width squeezed out."""
        m = self._overflow_menu
        m.clear()
        hidden = {name for name, w in self._tb_groups if not w.isVisible()}

        if "stats" in hidden:
            m.addAction(
                f"{self._self_cpu_lbl.text()}    {self._self_mem_lbl.text()}"
            ).setEnabled(False)
        if "session" in hidden:
            m.addAction(
                f"{self._opens_lbl.text()}    {self._closes_lbl.text()}"
            ).setEnabled(False)
            m.addAction(self._count_lbl.text()).setEnabled(False)
        if hidden & {"stats", "session"} and hidden & {"view", "extra"}:
            m.addSeparator()

        if "view" in hidden:
            sort_menu = m.addMenu("Sort")
            for i in range(self._sort_combo.count()):
                a = sort_menu.addAction(self._sort_combo.itemText(i))
                a.setCheckable(True)
                a.setChecked(i == self._sort_combo.currentIndex())
                a.triggered.connect(
                    lambda _checked, ix=i: self._sort_combo.setCurrentIndex(ix)
                )

            zoom_menu = m.addMenu("Zoom")
            zoom_menu.addAction(
                "Larger cards\tCtrl+=", lambda: self._grid_view.zoom(ZOOM_FACTOR)
            )
            zoom_menu.addAction(
                "Smaller cards\tCtrl+-", lambda: self._grid_view.zoom(1.0 / ZOOM_FACTOR)
            )
            zoom_menu.addAction(
                "Reset card size\tCtrl+0",
                lambda: self._grid_view.set_card_size(CARD_W_DEFAULT, CARD_H_DEFAULT),
            )

            cpu_menu = m.addMenu("Capture rate")
            for label, mode in (("High CPU", "high"), ("Low CPU", "low")):
                a = cpu_menu.addAction(label)
                a.setCheckable(True)
                a.setChecked((mode == "low") == self._low_cpu_active)
                a.triggered.connect(
                    lambda _checked, md=mode: self._set_cpu_mode(md)
                )

        if "extra" in hidden:
            if "view" in hidden:
                m.addSeparator()
            m.addAction("Maximize All", self._maximize_all)
            m.addAction("Restore All", self._restore_all)
            m.addAction("Log", self._show_log)
            m.addAction("Discord", self._open_discord)

    def _open_discord(self):
        QDesktopServices.openUrl(QUrl(DISCORD_URL))

    def _set_toolbar_collapsed(self, collapsed: bool):
        """Hide the toolbar row entirely, leaving just the tabs."""
        self._toolbar_widget.setVisible(not collapsed)
        self._collapse_btn.setText("▾" if collapsed else "▴")
        self._collapse_btn.setToolTip(
            "Show the toolbar" if collapsed else "Hide the toolbar"
        )

    def showEvent(self, event):
        super().showEvent(event)
        if not self._tb_measured:
            self._measure_toolbar()
            self._reflow_toolbar()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._reflow_toolbar()

    # ------------------------------------------------------------------
    # Client-type tabs
    # ------------------------------------------------------------------
    def _tab_btn_style(self, active: bool) -> str:
        if active:
            return f"""
                QPushButton {{
                    background: {HEADER_COLOR};
                    color: {TEXT_COLOR};
                    border: 1px solid {ACCENT_TEAL};
                    border-radius: 3px;
                    font-size: 11px;
                    font-weight: bold;
                    padding: 0 12px;
                }}
            """
        return f"""
            QPushButton {{
                background: transparent;
                color: {DIM_COLOR};
                border: 1px solid transparent;
                border-radius: 3px;
                font-size: 11px;
                padding: 0 12px;
            }}
            QPushButton:hover {{ color: {TEXT_COLOR}; border-color: {DIM_COLOR}; }}
        """

    def _set_active_kind(self, kind: str | None):
        self._active_kind = kind
        self._grid_view.set_filter_kind(kind)
        self._apply_tab_zoom()
        self._refresh_tabs()
        # Scripts differ per tab, so start each tab unfiltered
        self._refresh_script_filter(reset=True)

    def _apply_tab_zoom(self):
        """Load the active tab's saved card size (per-tab zoom)."""
        w, h = self._card_sizes.get(self._active_kind or "",
                                    (CARD_W_DEFAULT, CARD_H_DEFAULT))
        self._grid_view.set_card_size(w, h)

    def _on_card_size_changed(self, w: int, h: int):
        """Zoom changed (buttons/wheel/reset) — remember it for this tab."""
        self._card_sizes[self._active_kind or ""] = (w, h)

    def _refresh_script_filter(self, reset: bool = False):
        """Repopulate the script dropdown from the clients on the active tab.

        Only DreamBot titles carry a script, so the control is shown on the
        DreamBot and All tabs and hidden otherwise. Preserves the current
        selection across scans unless `reset` (used on tab switch)."""
        combo = self._script_combo
        scripts = sorted({
            _parse_script(c[1]) for c in self._clients
            if _parse_script(c[1])
            and (self._active_kind is None
                 or _client_kind(c[1], c[3]) == self._active_kind)
        })
        show = self._active_kind in (None, "DreamBot") and bool(scripts)
        self._script_lbl.setVisible(show)
        combo.setVisible(show)
        if not show:
            self._grid_view.set_filter_script(None)
            return

        prev = None if reset else combo.currentData()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("All Scripts", None)
        for name in scripts:
            combo.addItem(name, name)
        idx = combo.findData(prev) if prev is not None else 0
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.blockSignals(False)
        self._grid_view.set_filter_script(combo.currentData())

    def _on_script_changed(self, _idx: int):
        self._grid_view.set_filter_script(self._script_combo.currentData())

    def _refresh_tabs(self):
        """Update tab labels/counts, highlight the active one, and hide
        never-populated optional kinds (see ALWAYS_SHOWN_KINDS)."""
        total = sum(self._kind_counts.values())
        for kind, btn in self._tab_buttons.items():
            active = kind == self._active_kind
            if kind is None:
                btn.setText(f"All ({total})")
            else:
                count = self._kind_counts.get(kind, 0)
                btn.setText(f"{kind} ({count})")
                btn.setVisible(
                    kind in ALWAYS_SHOWN_KINDS or count > 0 or active
                )
            btn.setStyleSheet(self._tab_btn_style(active))

    def _tab_clients(self) -> list[tuple]:
        """Scanner tuples currently shown (active tab + script dropdown)."""
        kind = self._active_kind
        script = self._script_combo.currentData()
        out = []
        for c in self._clients:
            if kind is not None and _client_kind(c[1], c[3]) != kind:
                continue
            if script is not None and _parse_script(c[1]) != script:
                continue
            out.append(c)
        return out

    # ------------------------------------------------------------------
    # Tray icon + event log + alerts
    # ------------------------------------------------------------------
    def _setup_tray(self):
        self._tray = QSystemTrayIcon(self)
        self._tray.setIcon(QApplication.windowIcon())
        self._tray.setToolTip("BotWall")
        self._tray.setVisible(True)
        self._tray.activated.connect(self._on_tray_activated)

    def _on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.showNormal()
            self.activateWindow()

    def _log_event(self, msg: str):
        self._events.append(f"{time.strftime('%H:%M:%S')}  {msg}")
        if len(self._events) > 500:
            del self._events[:-500]

    def _alert(self, heading: str, detail: str, *,
               flash: bool = True, beep: bool = True):
        """Notification: optional beep + taskbar flash, plus a tray toast.

        flash/beep are opt-out so routine events (a client closing) can post a
        quiet toast without the taskbar constantly blinking orange, while
        user-chosen alert words still get the full attention-grabbing treatment.
        """
        if beep:
            try:
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            except Exception:
                pass
        if flash:
            QApplication.alert(self)  # flash the taskbar entry
        if self._tray.isVisible() and QSystemTrayIcon.supportsMessages():
            self._tray.showMessage(heading, detail, QSystemTrayIcon.Warning, 5000)

    def _alert_client_closed(self, title: str):
        self._log_event(f"Client closed: {title}")
        # Clients opening/closing is normal farm churn (restarts, world hops),
        # so keep it quiet — toast only, no beep or taskbar flash. The event
        # log still records every close.
        self._alert("BotWall — client closed", title, flash=False, beep=False)

    def _check_alert_words(self, title: str):
        tl = title.lower()
        for word in self._alert_words:
            if word and word in tl:
                self._log_event(f'Alert word "{word}" in title: {title}')
                self._alert(f'BotWall — "{word}"', title)
                return

    def _show_log(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("BotWall — Event Log")
        dlg.resize(560, 420)
        dlg.setStyleSheet(f"background: {BG_COLOR};")
        layout = QVBoxLayout(dlg)
        txt = QPlainTextEdit()
        txt.setReadOnly(True)
        txt.setStyleSheet(
            f"background: {CARD_COLOR}; color: {TEXT_COLOR}; "
            f"border: none; font-size: 12px;"
        )
        txt.setPlainText("\n".join(self._events) if self._events else "No events yet.")
        txt.verticalScrollBar().setValue(txt.verticalScrollBar().maximum())
        layout.addWidget(txt)
        dlg.exec_()

    def _show_settings(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("BotWall — Settings")
        dlg.setMinimumWidth(460)
        dlg.setStyleSheet(f"""
            QDialog {{ background: {BG_COLOR}; }}
            QLabel {{ color: {TEXT_COLOR}; font-size: 12px; }}
            QLineEdit {{
                background: {CARD_COLOR}; color: {TEXT_COLOR};
                border: 1px solid {DIM_COLOR}; border-radius: 3px;
                padding: 4px; font-size: 12px;
            }}
        """)
        form = QFormLayout(dlg)

        kw_edit = QLineEdit(", ".join(self._scan_keywords))
        kw_edit.setToolTip("Window-title keywords that make a window count as a client")
        form.addRow("Title keywords:", kw_edit)

        alert_edit = QLineEdit(", ".join(self._alert_words))
        alert_edit.setToolTip(
            "If a client's title CHANGES and contains one of these words, "
            "BotWall beeps and shows a toast (e.g. \"login, stopped, error\")"
        )
        form.addRow("Alert words:", alert_edit)

        hint = QLabel("Comma-separated, case-insensitive. Keywords apply on the next scan.")
        hint.setStyleSheet(f"color: {DIM_COLOR}; font-size: 11px;")
        form.addRow(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        form.addRow(buttons)

        if dlg.exec_() != QDialog.Accepted:
            return
        self._scan_keywords = [
            w.strip().lower() for w in kw_edit.text().split(",") if w.strip()
        ] or list(KEYWORDS)
        self._alert_words = [
            w.strip().lower() for w in alert_edit.text().split(",") if w.strip()
        ]
        self._scanner.set_keywords(self._scan_keywords)
        self._settings.setValue("scan_keywords", self._scan_keywords)
        self._settings.setValue("alert_words", self._alert_words)
        self._log_event(
            f"Settings updated — keywords: {', '.join(self._scan_keywords)}"
            + (f"; alert words: {', '.join(self._alert_words)}" if self._alert_words else "")
        )

    # ------------------------------------------------------------------
    # Threading
    # ------------------------------------------------------------------
    def _start_threads(self):
        self._scanner = Scanner()
        self._scanner.titan_state.connect(self._on_titan_state)
        self._scanner.updated.connect(self._on_scan)
        self._scanner.start()

        self._capturer = Capturer()
        self._capturer.captured.connect(self._on_capture)
        self._capturer.set_card_size(*self._grid_view.card_size())
        self._grid_view.card_size_changed.connect(self._capturer.set_card_size)
        self._grid_view.card_size_changed.connect(self._on_card_size_changed)
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
        # Shelved cards are hidden — stop capturing them
        self._grid_view.client_minimized.connect(lambda *_: self._push_capture_hwnds())
        self._grid_view.client_kill_requested.connect(self._on_kill_client)
        self._grid_view.client_restart_requested.connect(self._on_restart_client_proc)
        self._grid_view.nicknames_changed.connect(self._save_nicknames)

    def _update_self_stats(self):
        try:
            cpu = self._self_proc.cpu_percent(interval=None) / NUM_CORES
            mem_mb = self._self_proc.memory_info().rss / (1024 * 1024)
            mem_text = f"{mem_mb / 1024:.1f}GB" if mem_mb >= 1024 else f"{mem_mb:.0f}MB"
            self._self_cpu_lbl.setText(f"CPU: {cpu:.1f}%")
            self._self_mem_lbl.setText(f"MEM: {mem_text}")
        except Exception:
            pass

    def _push_capture_hwnds(self):
        """Hand the Capturer only the hwnds whose cards are actually shown
        (and, for TitanClient, only the active tab — the rest can't render)."""
        skip = self._grid_view.minimized_hwnds() | self._titan_hidden
        self._capturer.set_hwnds(
            [c[0] for c in self._clients if c[0] not in skip]
        )

    def _on_titan_state(self, info: dict):
        # Arrives right before the matching `updated` signal (same thread,
        # same queue), so _on_scan sees the current hidden set.
        self._titan_hidden = set(info.get("hidden_hwnds", ()))
        self._titan_controller_pid = info.get("controller_pid", 0)
        self._capturer.set_titan_tabs(info.get("tabs", {}))

    def _on_scan(self, clients: list):
        self._clients = clients
        counts: dict[str, int] = {}
        for c in clients:
            kind = _client_kind(c[1], c[3])
            counts[kind] = counts.get(kind, 0) + 1
        self._kind_counts = counts
        self._refresh_tabs()
        self._refresh_script_filter()
        self._grid_view.update_clients(clients)
        self._grid_view.set_hidden_tabs(self._titan_hidden)
        self._grid_view.check_feeds()
        self._push_capture_hwnds()
        # Push live stats into minimized shelf strips
        for hwnd, title, pid, proc_name, cpu_pct, mem_mb, uptime_s in clients:
            self._minimized_shelf.update_stats(hwnd, cpu_pct, mem_mb)
        n = len(clients)

        # Aggregate farm load across unique processes
        per_pid = {c[2]: (c[4], c[5]) for c in clients if c[2]}
        total_cpu = sum(v[0] for v in per_pid.values())
        total_mem = sum(v[1] for v in per_pid.values())
        mem_txt = (f"{total_mem / 1024:.1f}GB" if total_mem >= 1024
                   else f"{total_mem:.0f}MB")
        if n:
            self._count_lbl.setText(
                f"{n} client{'s' if n != 1 else ''} · ΣCPU {total_cpu:.0f}% · ΣRAM {mem_txt}"
            )
        else:
            self._count_lbl.setText("0 clients")
        self._count_lbl.setToolTip(
            "Detected clients · combined CPU (% of machine) and RAM of their processes"
        )
        # The stats labels just changed width — re-check what still fits
        self._reflow_toolbar()

        # Track client opens, closes, and title changes
        new_hwnd_titles = {c[0]: c[1] for c in clients}
        for hwnd, title in new_hwnd_titles.items():
            if hwnd not in self._active_hwnds:
                self._active_hwnds[hwnd] = title
                self._total_opens += 1
                self._log_event(f"Client opened: {title}")
            elif self._active_hwnds[hwnd] != title:
                # Title changes often signal logout / script stop / world hop
                self._log_event(
                    f'Title changed: "{self._active_hwnds[hwnd]}" → "{title}"'
                )
                self._active_hwnds[hwnd] = title
                self._check_alert_words(title)
        for hwnd in set(self._active_hwnds) - set(new_hwnd_titles):
            old_title = self._active_hwnds.pop(hwnd)
            self._total_closes += 1
            self._alert_client_closed(old_title)
        self._opens_lbl.setText(f"↑ {self._total_opens} Opened")
        self._closes_lbl.setText(f"↓ {self._total_closes} Closed")

    def _on_capture(self, hwnd: int, image: QImage, age_s: float):
        # QImage → QPixmap here, on the GUI thread — the only thread where
        # QPixmap is supported.
        self._grid_view.update_screenshot(hwnd, QPixmap.fromImage(image), age_s)

    def _on_restore_client(self, hwnd: int):
        self._minimized_shelf.remove_client(hwnd)
        self._grid_view.restore_client(hwnd)
        self._push_capture_hwnds()

    def changeEvent(self, event):
        super().changeEvent(event)
        # No point burning CPU on captures nobody can see
        if event.type() == QEvent.WindowStateChange:
            self._capturer.set_paused(self.isMinimized())

    def _on_sort_changed(self, idx: int):
        modes = ["default", "cpu_asc", "cpu_desc", "ram_asc", "ram_desc"]
        self._grid_view.set_sort_mode(modes[idx])

    def _set_cpu_mode(self, mode: str):
        low = (mode == "low")
        self._low_cpu_active = low
        self._capturer.set_interval(CAPTURE_INTERVAL_LOW if low else CAPTURE_INTERVAL_HIGH)
        self._grid_view.set_low_cpu(low)
        self._btn_high_cpu.setStyleSheet(self._cpu_btn_style(not low))
        self._btn_low_cpu.setStyleSheet(self._cpu_btn_style(low))

    # ------------------------------------------------------------------
    # Settings persistence
    # ------------------------------------------------------------------
    def _restore_settings(self):
        s = self._settings
        geo = s.value("geometry")
        if geo is not None:
            self.restoreGeometry(geo)
        # Per-tab zoom: {tab ("" = All): [w, h]}. Migrate the legacy single
        # card_w/card_h into the All slot for installs from before per-tab zoom.
        self._card_sizes = {}
        try:
            for k, v in json.loads(s.value("card_sizes", "{}")).items():
                self._card_sizes[k] = (int(v[0]), int(v[1]))
        except (TypeError, ValueError, IndexError):
            self._card_sizes = {}
        if "" not in self._card_sizes:
            try:
                self._card_sizes[""] = (
                    int(s.value("card_w", CARD_W_DEFAULT)),
                    int(s.value("card_h", CARD_H_DEFAULT)),
                )
            except (TypeError, ValueError):
                self._card_sizes[""] = (CARD_W_DEFAULT, CARD_H_DEFAULT)
        self._apply_tab_zoom()  # active tab is still All here
        try:
            sort_idx = int(s.value("sort_index", 0))
        except (TypeError, ValueError):
            sort_idx = 0
        if 0 <= sort_idx < self._sort_combo.count():
            self._sort_combo.setCurrentIndex(sort_idx)
        if s.value("cpu_mode", "high") == "low":
            self._set_cpu_mode("low")
        pinned = s.value("pinned_titles", []) or []
        if isinstance(pinned, str):
            pinned = [pinned]
        self._grid_view.set_pinned_titles(set(pinned))

        def _str_list(key: str) -> list[str]:
            v = s.value(key, []) or []
            return [v] if isinstance(v, str) else list(v)

        kw = _str_list("scan_keywords")
        if kw:
            if any(sorted(kw) == sorted(old) for old in OLD_DEFAULT_KEYWORDS):
                # An earlier default was persisted verbatim — upgrade it so
                # newer client kinds (TwiLite, OnlyBot's RuneLite) are
                # detected. Customized lists are kept.
                kw = list(KEYWORDS)
            self._scan_keywords = kw
            self._scanner.set_keywords(kw)
        self._alert_words = _str_list("alert_words")
        try:
            self._grid_view.set_nicknames(json.loads(s.value("nicknames", "{}")))
        except (TypeError, ValueError):
            pass
        kind = s.value("active_kind", "") or ""
        if kind in (*CLIENT_KINDS, OTHER_KIND):
            self._set_active_kind(kind)
        if s.value("toolbar_collapsed", "false") == "true":
            self._set_toolbar_collapsed(True)

    def _save_nicknames(self):
        self._settings.setValue(
            "nicknames", json.dumps(self._grid_view.nicknames())
        )

    def _save_settings(self):
        s = self._settings
        s.setValue("geometry", self.saveGeometry())
        # Capture the active tab's current size, then persist the whole map
        self._card_sizes[self._active_kind or ""] = self._grid_view.card_size()
        s.setValue("card_sizes", json.dumps(
            {k: list(v) for k, v in self._card_sizes.items()}
        ))
        # Legacy key kept so an older build still reads a sane size
        card_w, card_h = self._card_sizes.get("", self._grid_view.card_size())
        s.setValue("card_w", card_w)
        s.setValue("card_h", card_h)
        s.setValue("sort_index", self._sort_combo.currentIndex())
        s.setValue("cpu_mode", "low" if self._low_cpu_active else "high")
        s.setValue("pinned_titles", sorted(self._grid_view.pinned_titles()))
        s.setValue("nicknames", json.dumps(self._grid_view.nicknames()))
        s.setValue("scan_keywords", self._scan_keywords)
        s.setValue("alert_words", self._alert_words)
        s.setValue("active_kind", self._active_kind or "")
        s.setValue("toolbar_collapsed",
                   "true" if not self._toolbar_widget.isVisible() else "false")

    # ------------------------------------------------------------------
    # Maximize / Restore all
    # ------------------------------------------------------------------
    def _tab_root_hwnds(self) -> list[int]:
        # Titan tabs are child windows — act on their controller (once).
        return list(dict.fromkeys(_root_hwnd(c[0]) for c in self._tab_clients()))

    def _maximize_all(self):
        for hwnd in self._tab_root_hwnds():
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            except Exception:
                pass

    def _restore_all(self):
        for hwnd in self._tab_root_hwnds():
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Kill one / kill all
    # ------------------------------------------------------------------
    def _on_kill_client(self, pid: int, title: str, proc_name: str):
        if not pid:
            return
        # One process can own several client windows — killing it takes
        # all of them down; warn if that's the case.
        siblings = sum(1 for c in self._clients if c[2] == pid) - 1
        extra = (
            f"\nNote: {siblings} other window{'s' if siblings != 1 else ''} "
            "of this process will close too."
            if siblings > 0 else ""
        )
        reply = QMessageBox.question(
            self, "Kill Client",
            f'Kill "{title}"?\n'
            f"Process {proc_name} (PID {pid}) will be forcefully terminated."
            f"{extra}",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        try:
            psutil.Process(pid).kill()
            self._log_event(f"Killed client: {title} (PID {pid})")
        except Exception:
            self._log_event(f"Failed to kill PID {pid} ({title})")

    def _on_restart_client_proc(self, pid: int, title: str, proc_name: str):
        if not pid:
            return
        try:
            proc = psutil.Process(pid)
            cmdline = proc.cmdline()
            try:
                cwd = proc.cwd()
            except Exception:
                cwd = None
        except Exception:
            QMessageBox.warning(
                self, "Restart Client",
                f"Can't read the launch command of PID {pid} — restart unavailable."
            )
            return
        if not cmdline:
            QMessageBox.warning(
                self, "Restart Client",
                f"PID {pid} has no readable command line — restart unavailable."
            )
            return
        reply = QMessageBox.question(
            self, "Restart Client",
            f'Restart "{title}"?\n'
            f"Process {proc_name} (PID {pid}) will be killed and relaunched with "
            "its original command line.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        try:
            proc.kill()
            proc.wait(5)
        except Exception:
            pass
        try:
            subprocess.Popen(cmdline, cwd=cwd)
            self._log_event(f"Restarted client: {title} (was PID {pid})")
        except Exception as e:
            self._log_event(f"Restart failed for {title}: {e}")
            QMessageBox.warning(
                self, "Restart Client",
                f"Killed PID {pid} but relaunch failed:\n{e}"
            )

    def _kill_all(self):
        # Scoped to what's on screen: active tab + script dropdown. On the
        # TwiLite tab it only touches TwiLite; with a script selected it only
        # touches DreamBot clients running that script.
        kind = self._active_kind
        script = self._script_combo.currentData()
        scope_parts = [p for p in (kind, script) if p]
        scope = (" / ".join(scope_parts) + " client") if scope_parts else "client"
        pids = self._grid_view.all_pids(kind, script)
        if not pids:
            QMessageBox.information(self, "BotWall", f"No {scope}s to kill.")
            return
        names = sorted({c[3] for c in self._clients if c[2] in pids and c[3]})
        name_line = f"Processes: {', '.join(names)}\n" if names else ""
        # Killing every Titan tab leaves a controller full of dead tabs (its
        # "Launch New Client" then lands on an "(unresponsive)" tab), so take
        # the controller down with them — but only when ALL its tabs are in
        # scope; a script/tab-scoped kill must not touch the other tabs.
        titan_pids = {c[2] for c in self._clients
                      if _client_kind(c[1], c[3]) == TITAN_KIND}
        ctrl_line = ""
        if (self._titan_controller_pid and titan_pids
                and titan_pids <= set(pids)):
            pids.append(self._titan_controller_pid)  # last: tabs die first
            ctrl_line = "Plus the TitanClient controller window.\n"
        reply = QMessageBox.question(
            self, "Kill All",
            f"Kill {len(pids)} {scope} process{'es' if len(pids) != 1 else ''}?\n"
            f"{name_line}{ctrl_line}"
            "This will forcefully terminate them.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        killed = 0
        for pid in pids:
            try:
                p = psutil.Process(pid)
                p.kill()
                killed += 1
            except Exception:
                pass
        self._log_event(
            f"KILL ALL ({' / '.join(scope_parts) or 'All'}): "
            f"terminated {killed}/{len(pids)} processes"
        )

    # ------------------------------------------------------------------
    # Cleanup
    # ------------------------------------------------------------------
    def closeEvent(self, event):
        self._save_settings()
        self._tray.setVisible(False)  # avoid a ghost tray icon after exit
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
