# BotWall

A live screenshot monitor for DreamBot, TwiLite, TitanClient, OnlyBot, and RuneLite/RuneScape clients on Windows. View all your running clients in a single dashboard with real-time screen capture, CPU/memory stats, and quick window management.

## Features

- **Live screen capture** of all detected DreamBot / TwiLite / Titan / OnlyBot / RuneLite windows
- **TitanClient tabs** — each tab of a TitanClient controller gets its own card (CPU/RAM/uptime, kill, nickname). Only Titan's *active* tab renders, so other tabs show their last frame dimmed with a "TAB HIDDEN" badge; tab names are picked up as tabs become active
- **Navigation tabs** — filter the wall by client type (All / DreamBot / TwiLite / …); KILL ALL and Maximize/Restore All apply to the active tab only
- **Script filter** — narrow DreamBot clients by the script in their title (e.g. P2P Master AI vs MaxTutorialIsland)
- **Per-tab zoom** — each tab remembers its own card size
- **Real-time stats** — CPU usage, memory, process name per client
- **Grid layout** with dynamic column sizing and Ctrl+Scroll zoom
- **Sort clients** by title, CPU, memory, or PID
- **Pin clients** to keep important windows at the front
- **Minimize to shelf** — hide cards from the grid without losing track
- **Maximize / Restore All** toolbar buttons for quick window management
- **High / Low CPU modes** — toggle between 250ms and 1s capture intervals; Low CPU mode switches to grayscale rendering

### High CPU Mode
![High CPU Mode](high_detail.PNG)

### Low CPU Mode
![Low CPU Mode](low_detail.PNG)

## Requirements

- **Windows** (uses Win32 API for window capture)
- Python 3.8+

## Installation

```bash
git clone https://github.com/pspiotto/botwall.git
cd botwall
pip install -r requirements.txt
```

## Usage

```bash
python botwall.py
```

BotWall will automatically detect visible windows with "dreambot", "twilite", "runescape", or "runelite" in the title (configurable in ⚙ settings) and begin capturing screenshots. Each client is classified as DreamBot, TwiLite, Titan, OnlyBot (a RuneLite client launched from OnlyBot's `.onlybot` install), RuneLite, or Other, and the tab bar under the toolbar filters the wall by type.

### Controls

| Action | How |
|--------|-----|
| Filter by client type | Tabs under the toolbar (All / DreamBot / TwiLite / …) |
| Filter DreamBot by script | Script dropdown at the right of the tab bar |
| Zoom in/out | Ctrl + Scroll |
| Sort clients | Dropdown in toolbar |
| Pin a client | Click the pin button on the card header |
| Minimize a card | Right-click the card > "Minimize to Shelf" |
| Toggle CPU mode | High CPU / Low CPU buttons in toolbar |
| Maximize/Restore all windows | Toolbar buttons |

## Building a Standalone EXE

```bash
pip install pyinstaller
pyinstaller BotWall.spec
```

The output lands in `dist/BotWall.exe`. The icon files (`CmCSAHz.ico`, `CmCSAHz.png`) must be present alongside `botwall.py` when building.

## License

MIT
