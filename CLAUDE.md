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

The entire application is a single file: `botwall.py` (~2000 lines).

**Threading model:**
- `Scanner(QThread)` — runs every 3 s, calls `win32gui.EnumWindows` to find visible windows whose titles contain a scan keyword (default "dreambot"/"runescape"/"twilite"/"runelite", user-configurable via the ⚙ settings dialog → `set_keywords()`) AND whose owning process looks like a bot client (`_is_bot_process`: java/javaw or dreambot/runelite/osclient/jagex/twilite/onlybot in the name) — this keeps browser tabs mentioning RuneScape out of the grid and out of KILL ALL's reach. Caches `psutil.Process` objects by PID; stats are memoized per PID per scan (one `cpu_percent()` call per process, normalized by `NUM_CORES`). Emits `updated` signal with a list of 7-tuples: `(hwnd, title, pid, proc_name, cpu_pct, mem_mb, uptime_s)`. `proc_name` comes from `_proc_label()`: the plain process name, plus a launcher suffix when the exe path reveals one (`LAUNCHER_PATH_MARKERS`). A `titan_state` dict signal (`controller_pid`, `active_pid`, `hidden_hwnds`) is emitted right before it.
- **OnlyBot** is a fleet manager (`OnlyBot.exe`, a Tauri app) that launches a RuneLite fork per account as a stock-looking `java.exe` titled `RuneLite - <character>` — nothing in the window or process name says OnlyBot. The fingerprint is the JVM path (`%APPDATA%\.onlybot\lib\jdks\…\java.exe`), so `_proc_label()` reports it as `java.exe (OnlyBot)` and `_client_kind()` classifies it as OnlyBot (listed before RuneLite in `CLIENT_KINDS`, since the title also matches "runelite"). Detection needed `runelite` added to the default `KEYWORDS`; a vanilla RuneLite therefore also gets a card now, on the RuneLite tab. The script name (e.g. "Dentist Walker") lives only in the OnlyBot fleet table, not the title, so the script dropdown does not apply. Restart Client works the same as for any java client (cmdline/cwd relaunch), and the manager window is not detected (its title has no keyword).
- **TitanClient** is special-cased (`Scanner._scan_titan`): it runs one `controller.exe` window (class `TitanClientController`) with every game client as a *tab*, each tab being a separate `osclient.exe` whose real game window (class `JagWindow`) is re-parented as a hidden child of the controller — invisible to `EnumWindows` and the keyword scan. The scanner walks the controller's children instead and emits one client per `JagWindow` (hwnd = the JagWindow, pid = the osclient). **Only the active tab renders**: `PrintWindow` on an inactive tab's JagWindow just returns whatever is on screen at that rectangle, so inactive tabs are reported in `hidden_hwnds`, excluded from the capture list (and, because the scan is only every 3 s, the Capturer re-reads the controller title right before *and* after each Titan-tab `PrintWindow` via `set_titan_tabs()` / `_titan_active_tab()`, dropping the frame unless that tab is the active one both times — otherwise a tab switch leaves a Home-page frame on the card), and their cards show "TAB HIDDEN" with the last frame dimmed (`ClientCard.set_hidden_tab`; freeze/no-feed checks are suspended). Tab names exist only in the controller's imgui; the one source is the controller title suffix `- <name> #<pid>` for the *active* tab (set as UTF-8 through the ANSI API, so `_fix_titan_title` undoes the `Â ` mojibake), learned into `Scanner._titan_names` for the session — a never-activated tab is titled `TitanClient · tab #<pid>`. Restart Client is disabled for Titan tabs (osclient is injected by the controller); Bring to Front / Maximize / Restore act on the controller (`_root_hwnd`, GA_ROOT); KILL ALL adds the controller PID when every Titan tab is in scope. There is no Titan SDK screenshot/file API, so a plugin-side frame exporter is not an option.
- `Capturer(QThread)` — iterates the current hwnd list, calls `capture_hwnd()` per window, downscales each frame to 2× the current card size in-thread (the GUI never holds full-res frames), and emits `captured(hwnd, QImage, age_s)` where `age_s` is seconds since the frame content last changed (crc32 per frame; drives freeze detection). QImage, not QPixmap — pixmaps are GUI-thread-only; conversion happens in `BotWall._on_capture`. Deadline-based pacing: the interval (250 ms High CPU / 1000 ms Low CPU) is the full cycle period. Pausable via `set_paused()`; BotWall pauses it while its own window is minimized, and shelf-minimized clients are excluded from the hwnd list (`_push_capture_hwnds`).

**Monitoring features:**
- Freeze detection: cards show amber "static Ns" after `STALE_WARN_S` (30 s) of unchanged frames, red "FROZEN" after `STALE_FROZEN_S` (120 s), and dim "NO FEED" when captures stop arriving for `NO_FEED_S` (10 s) — checked per scan via `ClientCard.check_feed()`.
- Alerts: `_alert()` posts a tray toast, with opt-out `beep`/`flash`. A client disappearing is routine farm churn, so `_alert_client_closed` calls it toast-only (no beep, no `QApplication.alert` taskbar flash) — otherwise the taskbar blinks constantly as clients restart. User-configured alert words keep the full beep + taskbar flash + toast. Client opens are logged, never alerted. All opens/closes/title-changes/kills go to a session event log (toolbar "Log" button, capped at 500 entries).
- Per-client actions in the card context menu: Copy PID, Set Nickname (persisted by title), Restart Client (reads `cmdline()`/`cwd()`, kills, relaunches), Kill This Client (with confirm; warns when the process owns other windows).
- Alert words (settings dialog): if a client's title *changes* and contains one of the words, BotWall beeps + flashes the taskbar + toasts.
- Cards show process uptime next to the PID ("PID 1234 · 6h12m") — low uptime means a recent restart.
- Toolbar shows aggregate farm load (ΣCPU/ΣRAM across unique client processes); zoom via toolbar −/⤢/+ buttons, Ctrl+wheel, or Ctrl+= / Ctrl+- / Ctrl+0.
- Responsive toolbar: the toolbar is a row of collapsible *groups* (`self._tb_groups`, listed in hide order: `stats` → `extra` → `view` → `title` → `session`). `_reflow_toolbar()` runs on every `resizeEvent` and after each stats refresh, showing groups highest-priority-first and stopping at the first that doesn't fit (a strict cutoff, not a greedy fill, so groups always vanish in the same order). Whatever is hidden is rebuilt into the `⋯` overflow menu on `aboutToShow` (`_build_overflow_menu()`). Group widths are measured **live** via `_tb_widths()` — the stats labels start as `"0 clients"` / `"CPU: –"` and grow ~150 px once real numbers land, so a cached width clips them; only the never-hidden chrome (`_tb_fixed_w`: title through KILL ALL) is measured once, on first `showEvent`. KILL ALL, ⚙ and ⋯ never collapse.
- `toolbar_widget` carries `QSizePolicy.Ignored` horizontally and `setMinimumWidth(0)`. Without that, Qt clamps the window resize at the layout's own minimum and `_reflow_toolbar()` never gets asked to do anything — the toolbar used to impose a hard 1311 px floor on the window, which made a half-monitor split impossible. The floor is now the tab strip (~500 px).
- Toolbar collapse: the `▴`/`▾` button at the right of the tab strip hides `toolbar_widget` outright (`_set_toolbar_collapsed()`), persisted as `toolbar_collapsed`.
- Navigation tabs (All / DreamBot / TwiLite / Titan / OnlyBot / RuneLite / Other) filter the grid by client kind (`_client_kind()`: first `KIND_KEYWORDS` match in title or process name, else "Other"). Filtering is display-only — off-tab clients keep their captures and freeze detection (`ClientCard.update_pixmap` skips scaling while hidden; `showEvent` catches up on tab switch). KILL ALL, Maximize All, and Restore All are scoped to the active tab. DreamBot/TwiLite/Titan/OnlyBot tabs always show (`ALWAYS_SHOWN_KINDS`); RuneLite/Other only when populated.
- Script dropdown (right of the tab bar): filters within the current tab by the script named in a DreamBot title (`_parse_script()`: 3rd `" - "` segment, version suffix stripped, so `P2P Master AI v2.156` → `P2P Master AI`). Shown only on the DreamBot/All tabs and only when scripts are present; repopulated each scan (selection preserved) and reset on tab switch. It further narrows KILL ALL / Maximize All / Restore All (`_tab_clients()` / `GridView.all_pids(kind, script)`).
- Per-tab zoom: card size is stored per tab in `self._card_sizes` (`""` = All), applied on tab switch via `_apply_tab_zoom()` and recorded on every zoom via the `card_size_changed` → `_on_card_size_changed` signal. Persisted as `card_sizes` JSON; a pre-per-tab-zoom install migrates its old `card_w`/`card_h` into the All slot.
- CPU/RAM sort keys are bucketed (10% / 200 MB) so cards don't reshuffle on every scan.
- Settings persistence via `QSettings("BotWall", "BotWall")`: window geometry, card size (zoom), sort mode, CPU mode, pinned titles, nicknames (JSON, keyed by title), scan keywords, alert words, active tab (`active_kind`), per-tab card sizes (`card_sizes` JSON), toolbar collapse state (`toolbar_collapsed`). On restore, a persisted keyword list equal to any earlier default (`OLD_DEFAULT_KEYWORDS`) is upgraded to the current `KEYWORDS` so newer kinds (TwiLite, OnlyBot) get detected; customized lists are kept.

**Window capture:**
`capture_hwnd()` uses `ctypes.windll.user32.PrintWindow(hwnd, dc, 2)` (flag 2 = `PW_RENDERFULLCONTENT`) to capture hardware-accelerated windows. The GDI bitmap's raw BGRX bytes are wrapped directly in a `QImage(..., Format_RGB32).copy()` — no PIL, no PNG round-trip. Iconic (OS-minimized) and hung windows are skipped (`IsIconic` / `IsHungAppWindow`). All GDI handles are released in a per-handle-guarded `finally`; the bitmap must be deselected from the DC before `DeleteObject`, and pywin32's bitmap wrapper does NOT free the handle in its destructor — the manual delete is required.

**UI hierarchy:**
```
BotWall (QMainWindow)
├── toolbar_widget (QWidget, fixed 54px high, width-collapsible)
│   ├── g_stats   — BotWall's own CPU/MEM          (hidden first)
│   ├── g_extra   — Discord / Maximize / Restore / Log
│   ├── g_view    — sort QComboBox → GridView.set_sort_mode(); zoom −/⤢/+;
│   │               High/Low CPU → Capturer.set_interval() + GridView.set_low_cpu()
│   ├── title_lbl — "BotWall" wordmark
│   ├── g_session — ΣCPU/ΣRAM + opens/closes       (hidden last)
│   └── ⋯ overflow · ⚙ settings · KILL ALL         (never hidden)
├── tabs_widget (client-kind tabs → GridView.set_filter_kind(); script QComboBox → GridView.set_filter_script(); ▴/▾ toolbar collapse)
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
| `KEYWORDS` | `("dreambot", "runescape", "twilite", "runelite")` | Window title filter (`OLD_DEFAULT_KEYWORDS` lists superseded defaults for settings migration) |
| `BOT_PROC_NAMES` / `BOT_PROC_KEYWORDS` | java/javaw exact; dreambot/runelite/osclient/jagex/twilite/onlybot substring | Process-name filter (guards KILL ALL) |
| `LAUNCHER_PATH_MARKERS` | `{".onlybot": "OnlyBot", "\\dreambot\\": "DreamBot"}` | Exe-path substring → launcher suffix on `proc_name` (`java.exe (OnlyBot)`); DreamBot's JRE is tagged so a "Fatal error starting RuneLite" dialog from a DreamBot JVM stays on the DreamBot tab |
| `CLIENT_KINDS` / `KIND_KEYWORDS` | DreamBot / TwiLite / Titan / OnlyBot / RuneLite (+ `OTHER_KIND` fallback) | Client classification for the navigation tabs |
| `TITAN_CONTROLLER_CLASS` / `TITAN_TAB_CLASS` | `TitanClientController` / `JagWindow` | Window classes used to discover TitanClient tabs |

## Platform Notes

- **Windows only** — depends on `win32gui`, `win32ui`, `win32process`, `win32con` (pywin32) and `ctypes.windll.user32`.
- The PyInstaller spec targets a no-console window app; icon files `CmCSAHz.ico` / `CmCSAHz.png` must be present alongside `botwall.py` when running from source (or bundled via `sys._MEIPASS` when frozen).
