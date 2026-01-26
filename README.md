# Clipboard Sync (Windows Tray App)

> Former name: **Clipboard Sender** (renamed in v1.0.0 after adding *receive* support)

**Clipboard Sync** is a lightweight Windows tray app for **Notemod-selfhosted** that makes clipboard sharing *bidirectional*:

- **Send**: Windows clipboard → Notemod-selfhosted (write)
- **Receive**: Notemod-selfhosted → Windows clipboard → *(optional)* auto **paste (Ctrl+V)** with one hotkey

It also supports scheduled **INBOX cleanup** and (if backups are enabled on the server) a one-click **backup file purge** from the Settings screen.

## Features

- Tray app (no main window), click tray icon to enable/disable sending
- Global hotkey to toggle sending
- **Send** clipboard text to Notemod-selfhosted (POST)
- **Receive** latest note via `read_api.php` (GET) and copy to clipboard
  - Optional **auto paste** right after receive (Ctrl+V)
  - Optional “clipboard stabilize wait (ms)” before pasting
- Cleanup settings:
  - Manual **Delete all INBOX notes**
  - Auto delete **daily** or **every X minutes**
- Backup maintenance (when Notemod-selfhosted backup is enabled):
  - Show **current backup file count**
  - One-click **purge all backup files**
- Optional **Basic Auth** support
- UI languages: English / 日本語 / Türkçe

## Intended use

This app is designed to work with **Notemod-selfhosted** (server) + **Clipboard Sync** (Windows) + iPhone Shortcuts / Back Tap.

Example:
- Copy text on iPhone → (Notemod-selfhosted saves it) → press *Receive hotkey* on Windows → paste into the active app.

## Requirements

- Windows 10/11
- .NET 8 Desktop Runtime (for running the released EXE)
- A working **Notemod-selfhosted** installation providing:
  - `api.php` (write / send)
  - `read_api.php` (read / receive)
  - `cleanup_api.php` (INBOX delete + backup purge)

> If your Notemod-selfhosted is protected by Basic Auth, configure it in Settings.

## Installation

1. Download the latest release from GitHub Releases.
2. Run the EXE.
3. Open **Settings** from the tray menu and set URLs / tokens.

## Usage

### Startup & basic behavior

- The app runs in the system tray.
- Left-click the tray icon to enable/disable **sending**.
- Right-click the tray icon to open the menu:
  - Delete all INBOX notes
  - Settings
  - Help / About

### Settings screen

Settings are split into two tabs:

- **Send/Delete**
  - POST URL + token (send)
  - Basic Auth (optional)
  - Cleanup_API URL + token (delete INBOX)
  - Auto delete schedule (daily / every X minutes)
  - Backup file count + “purge all backup files”
  - Language
- **Receive**
  - Read API URL
  - Receive hotkey
  - Auto paste after receive
  - Clipboard stabilize wait (ms)

> Note: Auto paste may fail when the target application is running as **Administrator**.  
> Try pasting into Notepad first to verify basic operation.

## Build (for developers)

### Prerequisites

- Visual Studio 2022
- .NET 8 SDK
- Windows Forms workload

### Steps

1. Open the solution in Visual Studio.
2. Build / publish.
3. Ensure icons are included correctly (see next section).

## About embedding icons/resources (Resources)

Tray icons are embedded resources (so the tray icon works even when running from a single folder).

- `Assets/tray_on.ico` → EmbeddedResource
- `Assets/tray_off.ico` → EmbeddedResource
- `Assets/app.ico` → ApplicationIcon

## Security / Privacy

- The app sends only the clipboard text you copy (when sending is enabled).
- Tokens/passwords are stored in `settings.json` with **DPAPI encryption** (Windows user profile scope).
- Consider using HTTPS and Basic Auth if you expose the server on the internet.

## Acknowledgements

- Notemod-selfhosted: https://github.com/StayHomeLabNet/Notemod-selfhosted

## License

MIT License

## Contributing

Issues / PRs are welcome.

## Links

- Notemod-selfhosted: https://github.com/StayHomeLabNet/Notemod-selfhosted
- Help: https://stayhomelab.net/Clipboardsync-en
- GitHub: https://github.com/StayHomeLabNet/ClipboardSync
