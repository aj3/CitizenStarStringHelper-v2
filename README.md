# Citizen StarString Helper

Citizen StarString Helper is a Windows desktop utility for keeping a StarStrings install up to date for Star Citizen.

It checks the configured StarStrings GitHub repository for releases, installs updates into the Star Citizen `LIVE` folder, creates recoverable backups before overwriting files, preserves `USER.cfg`, and can update the application itself from GitHub Releases.

## Download

**[Download the latest Citizen StarString Helper.exe](https://github.com/aj3/CitizenStarStringHelper-v2/releases/download/v2.0.1/Citizen.StarString.Helper.exe)**

Or browse all releases: [GitHub Releases](https://github.com/aj3/CitizenStarStringHelper/releases)

## What The App Does

- Checks the configured StarStrings repository for new releases
- Downloads and installs StarStrings into the selected Star Citizen `LIVE` folder
- Creates timestamped backups before changes are applied
- Preserves and merges `USER.cfg` safely
- Supports scheduled automatic checks for StarStrings updates
- Supports in-place application self-updates through GitHub Releases
- Uses the currently running EXE path for self-updates, so the app can be run from different folders

## Backups

Backups are stored under:

`%LOCALAPPDATA%\CitizenStarStringHelper\Backups`

Each backup snapshot preserves the StarStrings files about to be replaced, plus `USER.cfg` if it exists. The restore dialog can be used to roll back to a previous snapshot.

## App Update Behavior

Application updates are delivered through this repository’s GitHub Releases.

When a newer version is available, the app can:

- detect the new release
- prompt the user to install it
- download the latest EXE
- replace the currently running copy
- relaunch the updated application

## Default Target Path

The default Star Citizen install path used by the app is:

`C:\Program Files\Roberts Space Industries\StarCitizen\LIVE`

## App Data Location

The app stores writable runtime data in:

`%LOCALAPPDATA%\CitizenStarStringHelper`

That folder contains:

- `starstrings_settings.json`
- `starstrings_state.json`
- `starstrings_updater.log`
- `Backups\`
- `PendingAppUpdate\`

## Included Source Files

This production source folder only includes the required project files:

- `starstrings_updater.py`
- `updater_helper.py`
- `build_exe.bat`
- `app_icon.ico`
- `app_icon.png`

## Notes

- This repository is for the desktop helper application itself, not the StarStrings content repository.
- Built executables, logs, backups, and runtime state are intentionally excluded from source control.
