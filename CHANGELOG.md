# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
