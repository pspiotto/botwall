# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

**Run the app:**
```bash
python botwall.py
```

**Install dependencies:**
```bash
pip install -r requirements.txt
```

**Build standalone EXE (Windows only):**
```bash
pyinstaller BotWall.spec
```
Output lands in `dist/BotWall.exe`. The spec bundles `CmCSAHz.ico` and `CmCSAHz.png` as data files and produces a single-file, no-console executable with UPX compression.

## Architecture

The entire application is a single file: `botwall.py` (~1080 lines).

**Threading model:**
- `Scanner(QThread)` — runs every 3 s, calls `win32gui.EnumWindows` to find visible windows whose titles contain "dreambot", "runescape", or "oldschool runescape". Caches `psutil.Process` objects by PID to get meaningful `cpu_percent()` readings across calls. Emits `updated` signal with a list of 6-tuples: `(hwnd, title, pid, proc_name, cpu_pct, mem_mb)`.
- `Capturer(QThread)` — iterates the current hwnd list, calls `capture_hwnd()` per window, emits `captured(hwnd, QPixmap)`. Refresh rate is 250 ms (High CPU) or 1000 ms (Low CPU).

**Window capture:**
`capture_hwnd()` uses `ctypes.windll.user32.PrintWindow(hwnd, dc, 2)` (flag 2 = `PW_RENDERFULLCONTENT`) to capture hardware-accelerated windows. The GDI bitmap is read as raw BGRX bytes, converted via `PIL.Image.frombuffer("RGB", ..., "raw", "BGRX")`, then serialized through a PNG buffer into a `QPixmap`.

**UI hierarchy:**
```
BotWall (QMainWindow)
├── toolbar_widget (QWidget, fixed 42px)
│   ├── sort QComboBox  → GridView.set_sort_mode()
│   └── High/Low CPU buttons → Capturer.set_interval() + GridView.set_low_cpu()
├── GridView (QScrollArea)
│   ├── ClientCard × N  (one per detected window)
│   └── EmptyPlaceholder (shown when no visible cards)
└── MinimizedShelf (QWidget, hidden until cards are minimized)
    └── MinimizedStrip × N
```

**Card lifecycle:**
- Scanner emits → `BotWall._on_scan` → `GridView.update_clients()` adds/removes/updates `ClientCard`s and calls `_relayout()`.
- Capturer emits → `BotWall._on_capture` → `GridView.update_screenshot()` → `ClientCard.update_pixmap()` → `_rescale()` scales to current label size.
- Right-click "Minimize to Shelf" hides the card from the grid and adds a `MinimizedStrip` to `MinimizedShelf`. Clicking the strip restores it.
- Pin button (📌) moves card to the front of the grid (pinned hwnds come first in `_sorted_order()`).

**Grid layout:**
Column count is computed dynamically: `cols = viewport_width // (card_w + spacing)`. Ctrl+Scroll calls `GridView.zoom()` which clamps minimum card size to 160×110 px.

**Low CPU mode:**
Sets capture interval to 1 s and converts each frame to grayscale (`QImage.Format_Grayscale8`) with `Qt.FastTransformation` scaling instead of smooth.

## Key Constants (top of botwall.py)

| Name | Value | Purpose |
|------|-------|---------|
| `SCAN_INTERVAL_MS` | 3000 | Window list refresh |
| `CAPTURE_INTERVAL_HIGH` | 250 ms | Frame rate, High CPU mode |
| `CAPTURE_INTERVAL_LOW` | 1000 ms | Frame rate, Low CPU mode |
| `CARD_W_DEFAULT` / `CARD_H_DEFAULT` | 320 / 220 | Initial card dimensions |
| `KEYWORDS` | `("dreambot", "runescape", "oldschool runescape")` | Window title filter |

## Platform Notes

- **Windows only** — depends on `win32gui`, `win32ui`, `win32process`, `win32con` (pywin32) and `ctypes.windll.user32`.
- The PyInstaller spec targets a no-console window app; icon files `CmCSAHz.ico` / `CmCSAHz.png` must be present alongside `botwall.py` when running from source (or bundled via `sys._MEIPASS` when frozen).
