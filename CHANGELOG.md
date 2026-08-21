# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2026-08-21

### Added
- **Background loading**: A window that has just finished loading is pushed to the bottom of the z-order instead of only having its focus taken back, so the program keeps starting up without covering your screen
- **Intent detection**: Alt+Tab, Win+Tab and clicks on a window or its taskbar button are recognised as deliberate switches and let through
- **Foreground poll**: A 200 ms backstop catches focus changes the shell hook never reports
- **Keep New Windows in Background** and **Show Notifications** tray toggles
- **Settings persistence**: Protection state, the new toggles and a tunable `idleThresholdMs` are stored in `config.json`
- **Back-off**: An app that refuses to stay put is left alone after several attempts instead of fighting over the foreground
- Double-clicking the tray icon toggles protection

### Fixed
- **Focus restore was blocked by our own setting**: The raised `FOREGROUNDLOCKTIMEOUT` also stopped FocusShield's own `SetForegroundWindow` call; it is now lifted for the duration of the restore
- **Full-screen apps were always blocked**: `HSHELL_RUDEAPPACTIVATED` means the window is full-screen, not that it forced activation, so it is judged by the normal rules
- **Whitelist/blacklist editor was unusable**: Add and Remove carried dialog results and closed the window on the first click, and the list was rebound to a stale snapshot so edits never appeared
- **Load on Boot pointed at the wrong file**: The registry entry used `Assembly.Location` (the managed dll, empty for single-file builds) instead of the executable, and was not quoted for paths containing spaces
- **Tray icons pointed at a freed handle**: Icons are built from an in-memory `.ico` rather than a cloned `HICON` that was destroyed immediately afterwards
- **Crash on exit**: `ApplicationExit` touched the tray icon after `FormClosing` had disposed it
- Shell hook registration now survives a window handle recreation
- Restoring focus no longer targets a window that has been closed or minimised, and a window that is still loading can no longer hang FocusShield through `AttachThreadInput`
- Balloon tips are rate-limited per window, so a persistent app cannot flood the tray
- Process names are read with `QueryFullProcessImageName`, which also works for elevated processes
- The foreground lock timeout is no longer written to the registry, so a crash cannot leave the system setting changed

### Changed
- App list entries are normalised (trimmed, `.exe` stripped, de-duplicated) and an app can no longer sit on both lists
- Windows that cannot take over the screen — tool windows, cloaked windows, the shell, zero-sized helpers — are ignored outright
- `AppListManager` is now `ConfigManager`, since it holds settings as well as the app lists
- Replaced the fixed 150 ms idle threshold with the configurable `idleThresholdMs` (default 300 ms), now only one of several signals

## [1.0.0] - 2026-03-20

### Added
- **Whitelist system**: Allow specific apps (Discord, Slack, etc.) to always steal focus when needed
- **Blacklist system**: Permanently block known problematic apps (Steam, popups, etc.) from stealing focus
- **Load on Boot option**: Auto-start FocusShield on Windows startup via tray menu
- **App configuration storage**: Whitelist/blacklist stored in `%APPDATA%\FocusShield\config.json`
- **Friendly notifications**: Changed alert tone from "blocked a pop-up" to "{App} is ready"
- **Shell hook detection**: Catches both rude app activation attempts and idle-triggered focus steals
- **Focus restoration**: Intelligently returns focus to your previous active window
- **Taskbar flashing**: Visual indicator when an app's focus attempt is blocked

### Changed
- Lowered idle threshold from 300ms to 150ms to catch more aggressive apps like Steam
- Improved notification messages for better user experience

### Technical Details
- Built on .NET 8 with Windows Forms
- Uses Windows shell hooks (`RegisterShellHookWindow`) for focus monitoring
- Registry-based startup configuration
- JSON-based app list persistence
