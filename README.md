# FocusShield

A Windows system-tray utility that prevents background applications from stealing keyboard focus. When an app tries to force itself to the foreground, FocusShield blocks the takeover, flashes the offending app's taskbar button, and keeps you in whatever you were working on.

## Features

- **Automatic Protection**: Blocks background apps from stealing focus while you work
- **Smart Detection**: Uses idle timing to distinguish between accidental focus steals and user-initiated window switches
- **Whitelist**: Allow specific apps (Discord, Slack, etc.) to always grab focus when needed
- **Blacklist**: Permanently block known problematic apps from ever stealing focus
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
| **Protection Enabled** | Toggle protection on/off (checkbox) |
| **Whitelist...** | Add/remove apps that can always steal focus |
| **Blacklist...** | Add/remove apps that should always be blocked |
| **Load on Boot** | Auto-start FocusShield when Windows boots (checkbox) |
| **Exit** | Quit FocusShield and restore original system settings |

### Common Scenarios

**Steam keeps interrupting your work:**
- Right-click tray → **Blacklist...** → Add → type `steam` → OK
- Steam will now never steal focus

**Discord notifications should interrupt you:**
- Right-click tray → **Whitelist...** → Add → type `discord` → OK
- Discord will always grab focus when it needs to notify you

**Prevent any app from stealing focus:**
- Right-click tray → **Protection Enabled** (checked) — this is the default
- Left-click taskbar buttons to manually switch to apps

## How It Works

1. **Windows-level blocking**: Raises `FOREGROUNDLOCKTIMEOUT` to 30 seconds so the OS itself suppresses most focus steals
2. **Shell hook detection**: Registers for `HSHELL_RUDEAPPACTIVATED` notifications when apps force activation
3. **Idle-aware heuristics**: Blocks apps that activate while you're idle (150ms+), allowing legitimate user-initiated switches
4. **Smart restoration**: Returns focus to your active window and flashes the blocked app's taskbar button
5. **Configuration**: Whitelist/blacklist stored in `%APPDATA%\FocusShield\config.json`

## Configuration

App lists are stored in `%APPDATA%\FocusShield\config.json`:

```json
{
  "whitelist": ["discord", "slack"],
  "blacklist": ["steam"]
}
```

- App names are **case-insensitive**
- Use the process name **without** `.exe` extension
- Find process names in Task Manager (Details tab)

## Troubleshooting

**An app still steals focus:**
- Lower threshold by editing code (line 37 in `FocusShieldForm.cs`), or
- Add the app to the blacklist

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
