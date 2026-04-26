# Nextion Stream Deck

Turn a Nextion HMI panel into a Windows desktop control surface with Stream Deck-style pages, app launch buttons, hotkeys, media controls, and custom artwork.

## Highlights

- Serial bridge for Nextion touch events over COM
- Multi-page deck profiles mapped to Nextion page ids
- Fixed-size tiles with light and dark editor themes
- Layout presets including `5 x 3` and `3 x 2`
- App import for `.exe`, `.lnk`, `.url`, `.bat`, `.cmd`, and `.ps1`
- Per-tile custom name, icon/photo, shortcut keys, and label sync target
- Media controls such as `play_pause`, `next_track`, and `previous_track`
- Profile storage in JSON for easy backup and editing

## Requirements

- Windows
- Python 3.13 or newer
- A Nextion HMI connected over serial
- Matching baud settings on both the Nextion side and the app side

## Quick Start

1. Install dependencies:

```powershell
pip install -r requirements.txt
```

2. Start the app:

```powershell
python app.py
```

If Windows starts you in the wrong folder, use:

```powershell
.\run_app.ps1
```

or double-click `run_app.bat`.

To run it without leaving a terminal window open, use:

```powershell
.\run_app_silent.bat
```

or double-click `run_app_silent.vbs`.

## How It Works

The app listens for standard Nextion touch event packets:

```text
0x65 page component event 0xFF 0xFF 0xFF
```

When a touch press arrives, the app matches the `page + component` pair to a configured tile and runs the tile action on your PC.

## Features

### App Tiles

- Launch desktop apps, files, scripts, and URLs
- Send keyboard shortcuts
- Launch an app and then send follow-up shortcut keys
- Import metadata from Windows shortcuts and executables

### Visual Customization

- Custom tile names
- Custom icons or photos per tile
- Automatic square crop/resize for imported artwork
- Fixed tile sizes so the layout stays stable

### Pages And Layouts

- Multiple named deck pages
- Nextion page id mapping per page
- Layout presets for different panel densities
- Page duplication for quick iteration

### Nextion Integration

- Label sync back to named HMI components
- Manual sync per tile or batch sync for all labels

## Setup Your Nextion Project

For each tappable control on the HMI:

- Enable `Touch Press Event`
- Note the page id and component id
- Map that pair in the app editor

If a tile uses label sync, set a `label_target` like:

```text
page0.b0
```

The app will send commands like:

```text
page0.b0.txt="OBS"
```

## Adding Apps

Use `Import App` in the editor and choose one of:

- `.exe`
- `.lnk`
- `.url`
- `.bat`
- `.cmd`
- `.ps1`

The app will try to collect:

- a display name
- the launch target or URL
- an icon source

You can override all of that afterward with:

- `Custom Name`
- `Choose Photo/Icon`
- `Shortcut Keys`

## Media Controls

For Spotify and other media apps, use a tile with `Action Type` set to `hotkey` and one of these payloads:

- `play_pause`
- `next_track`
- `previous_track`
- `media_stop`

You can also put those values in `Shortcut Keys` if you want them to fire after a launch action.

## Profile Format

Profiles live in `profiles/default.json`.

Each profile stores:

- `rows`
- `cols`
- `pages`
- `active_page`
- `theme_mode`

Each tile stores:

- `slot`
- `page_id`
- `component_id`
- `label`
- `label_target`
- `action_type`
- `payload`
- `icon_path`
- `source_path`
- `shortcut_keys`

## Project Layout

```text
app.py
nextion_stream_deck/
  actions.py
  config.py
  metadata.py
  protocol.py
  serial_bridge.py
  ui.py
profiles/
tests/
```

## Development

Run the tests with:

```powershell
python -m unittest discover -s tests -v
```

## Roadmap

- Better icon handling for more image formats
- Export/import profile presets
- Richer app-state integrations
- Optional packaging into a standalone Windows executable
