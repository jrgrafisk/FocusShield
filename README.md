# FocusShield

A Windows system-tray utility that prevents background applications from stealing keyboard focus. When an app tries to force itself to the foreground, FocusShield blocks the takeover, flashes the offending app's taskbar button, and keeps you in whatever you were working on.

## Requirements

- Windows 10 or later (Vista+ minimum)
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

## Build

```powershell
# Clone
git clone https://github.com/jrgrafisk/FocusShield.git
cd FocusShield

# Build (Debug)
dotnet build

# Build (Release)
dotnet build -c Release
```

The output binary is placed in `FocusShield\bin\Release\net8.0-windows\`.

## Run

```powershell
dotnet run --project FocusShield
```

Or double-click the compiled `FocusShield.exe`. A shield icon appears in the system tray.

## Usage

| Action | Result |
|--------|--------|
| App starts | FocusShield runs silently in the tray, protecting focus |
| Background app tries to steal focus | Takeover is blocked; offending app's taskbar button flashes |
| Right-click tray icon → **Protection Enabled** | Toggle protection on/off |
| Right-click tray icon → **Exit** | Quit FocusShield and restore original system settings |

## How It Works

1. Raises `FOREGROUNDLOCKTIMEOUT` to 30 seconds so Windows itself suppresses focus steals.
2. Registers a shell hook (`RegisterShellHookWindow`) to receive `HSHELL_RUDEAPPACTIVATED` notifications when an app forces activation.
3. On detection: flashes the rude window's taskbar button and calls `SetForegroundWindow` to restore the previous window.
4. On exit: restores the original `FOREGROUNDLOCKTIMEOUT` value.

## Contributing

Pull requests are welcome. Please open an issue first for major changes.

## License

[MIT](LICENSE)
