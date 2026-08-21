# FocusShield

A Windows system-tray utility that lets programs finish loading in the background instead of taking over your screen. When an app puts its window up while you are working, FocusShield drops that window behind your work and hands your keyboard focus straight back — the program keeps starting up as normal, and its taskbar button flashes so you can switch to it whenever you are ready.

## Features

- **Background loading**: Apps that finish starting up stay behind whatever you are working on — no stolen focus, no window covering your screen
- **Keeps focus where it is**: Your window stays on top and keeps the keyboard, mid-sentence and mid-shortcut
- **Smart detection**: Recognises the switches you asked for — Alt+Tab, Win+Tab, clicking a window or a taskbar button — and lets those through
- **Whitelist**: Allow specific apps (Discord, Slack, etc.) to always come to the front
- **Blacklist**: Always keep known offenders (Steam, updaters, popups) in the background
- **Load on Boot**: Enable auto-start on Windows startup from the tray menu
- **Minimal Overhead**: Runs silently in the system tray with negligible CPU/memory usage

## Requirements

- Windows 10 or later (Vista+ minimum)
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

## Installation

### From Binary
1. Download the latest `FocusShield.exe` from [Releases](https://github.com/jrgrafisk/FocusShield/releases)
2. Run the executable — it will appear in your system tray immediately
3. (Optional) Right-click tray icon → **Load on Boot** to auto-start on Windows startup

### From Source
```powershell
git clone https://github.com/jrgrafisk/FocusShield.git
cd FocusShield
dotnet build -c Release
.\FocusShield\bin\Release\net8.0-windows\FocusShield.exe
```

## Usage

### Tray Menu Options

| Menu Item | Function |
|-----------|----------|
| **Protection Enabled** | Toggle protection on/off (checkbox). Double-clicking the tray icon does the same |
| **Keep New Windows in Background** | Send windows that have just finished loading to the back (checkbox, on by default) |
| **Show Notifications** | Show a balloon tip when a window is held back (checkbox) |
| **Whitelist...** | Add/remove apps that may always come to the front |
| **Blacklist...** | Add/remove apps that are always kept in the background |
| **Load on Boot** | Auto-start FocusShield when Windows boots (checkbox) |
| **Exit** | Quit FocusShield and restore original system settings |

The tray icon is a green shield while protecting, amber for a couple of seconds after it has held a window back, and grey while paused.

### Common Scenarios

**A program you launched pops up over your work when it finishes loading:**
- This is handled out of the box — it opens behind your window with its taskbar button flashing
- Click that taskbar button whenever you want to switch to it

**Steam keeps interrupting your work:**
- Right-click tray → **Blacklist...** → Add → type `steam` → OK
- Steam is now kept in the background even if you were typing at the time

**Discord notifications should interrupt you:**
- Right-click tray → **Whitelist...** → Add → type `discord` → OK
- Discord will always come to the front when it needs to notify you

**You want launched apps to appear in front again:**
- Right-click tray → uncheck **Keep New Windows in Background**
- Apps stealing focus while you are idle are still held back

## How It Works

1. **Windows-level blocking**: Raises `FOREGROUNDLOCKTIMEOUT` to 30 seconds so the OS itself suppresses most focus steals. The setting is changed in memory only, so a crash can never leave it stuck on your machine
2. **Shell hook detection**: Registers for shell notifications to see windows being created and activated, backed by a lightweight foreground poll for the notifications Windows drops
3. **Loading detection**: A window activating within seconds of being created is a program that has just finished loading, whatever you happened to be doing at that moment
4. **Intent heuristics**: Alt+Tab, Win+Tab, and clicks on a window or its taskbar button are recognised as switches you asked for and are let through
5. **Background placement**: The window is pushed to the bottom of the z-order without being activated, so it is out of the way but still loading normally, and your window is brought back to the front
6. **Configuration**: Settings and app lists are stored in `%APPDATA%\FocusShield\config.json`

An app that refuses to stay put is given up on after several attempts, so FocusShield never ends up fighting a window back and forth across your screen.

## Configuration

Settings live in `%APPDATA%\FocusShield\config.json` and are written whenever you change something from the tray menu:

```json
{
  "whitelist": ["discord", "slack"],
  "blacklist": ["steam"],
  "protectionEnabled": true,
  "backgroundNewWindows": true,
  "showNotifications": true,
  "idleThresholdMs": 300
}
```

| Key | Meaning |
|-----|---------|
| `whitelist` | Apps that may always come to the front |
| `blacklist` | Apps that are always kept in the background |
| `protectionEnabled` | Master switch, mirrors the tray menu |
| `backgroundNewWindows` | Keep windows that have just finished loading behind your work |
| `showNotifications` | Show balloon tips |
| `idleThresholdMs` | How soon after your last keypress or mouse move an activation still counts as something you asked for (0–5000, default 300) |

- App names are **case-insensitive**
- Use the process name — a trailing `.exe` is stripped for you
- Find process names in Task Manager (Details tab)

## Troubleshooting

**An app still comes to the front:**
- Add it to the blacklist — blacklisted apps are held back even if you were typing at the time
- Or lower `idleThresholdMs` in `config.json` and restart FocusShield

**An app you want to see is being held back:**
- Add it to the whitelist, or uncheck **Keep New Windows in Background**

**FocusShield doesn't start on boot:**
- Ensure you clicked **Load on Boot** in the tray menu
- Check `HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run` in Registry Editor

**Need to remove from startup:**
- Right-click tray → **Load on Boot** (uncheck it), or
- Delete the `FocusShield` value from the registry key above

## Contributing

Pull requests welcome. Please open an issue first for major changes.

## License

[MIT](LICENSE)
