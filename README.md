# Google Voice Dialer

**Call phone numbers from any app in Windows using Google Voice.**

Click a `tel:` or `callto:` link and have it open directly in the Google Voice web app (PWA) or Chrome — no copy/paste, no fumbling with your browser.

## Why use it?

- **Seamless calling** – Instantly dial numbers from emails, documents, or CRM tools.
- **Works anywhere** – Any `tel:` or `callto:` link triggers Google Voice.
- **No admin rights required** – Registers as your default TEL handler at the user level.
- **Lightweight & self-contained** – Small executable, no background services.

## What it does

- Registers itself as the default `tel:`/`callto:` handler in Windows.
- Launches the Google Voice PWA via Chrome for quick dialing.
- Automatically detects your installed Chrome/Voice app.
- Supports both portable script use and one-click install to `%APPDATA%`.

## Quick start

```powershell
# Install dependencies (Python 3.11+)
pip install pywin32 pyinstaller pillow pyinstaller_versionfile

# Install as default TEL handler
py google_voice_dialer.py --install

# Remove
py google_voice_dialer.py --uninstall
